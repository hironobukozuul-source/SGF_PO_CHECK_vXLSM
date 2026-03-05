import streamlit as st
import pandas as pd
import io

# --- カラム定義 ---
PLAN_MAT_COL = "品目コード"
PLAN_PROD_COL = "製品記号"
PLAN_START_COL = "製造開始日"
PLAN_QTY_COL = "Quantity"

MASTER_KEY = "Parent material number"
MATERIAL_TYPE_COL = "Material Type"
TARGET_TYPE = "VERP"

MASTER_PARENT_QTY_COL = "Parent Material Quantity"
MASTER_COMP_QTY_COL = "Component Quantity"
MASTER_DESC_COL = "Component Description"
MASTER_COMP_NUM_COL = "Component Number"

def get_plan_data(uploaded_file):
    if uploaded_file is None: return {}
    plans_dict = {}
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
    target_sheets = [s for s in xls.sheet_names if s.startswith("Day bucket plan_12")]
    for sheet in target_sheets:
        meta_df = pd.read_excel(xls, sheet_name=sheet, nrows=1, header=None)
        plan_name = str(meta_df.iloc[0, 0]).strip() if not meta_df.empty else sheet
        data_df = pd.read_excel(xls, sheet_name=sheet, skiprows=5)
        data_df = data_df.dropna(how='all')
        if not data_df.empty:
            plans_dict[plan_name] = data_df
    return plans_dict

def compute_necessary_qty(row, plan_qty):
    desc = str(row.get(MASTER_DESC_COL, '')).upper()
    if "BOTTLE" in desc or "PUMP" in desc:
        return plan_qty
    p_qty = row.get(MASTER_PARENT_QTY_COL, 1)
    c_qty = row.get(MASTER_COMP_QTY_COL, 0)
    if pd.isna(p_qty) or p_qty == 0: return 0
    return (plan_qty / p_qty) * c_qty

def process_bom_structure(plan_df, cu_df, du_df):
    """親と子(VERP)を紐付け、1つの構造化されたリストを返す"""
    if plan_df.empty: return pd.DataFrame()

    # マスタのクリーニング
    for df in [cu_df, du_df]:
        if MATERIAL_TYPE_COL in df.columns:
            df.query(f"`{MATERIAL_TYPE_COL}` == '{TARGET_TYPE}'", inplace=True)

    results = []
    for _, row in plan_df.iterrows():
        # 1. 親品目自体を登録
        parent_info = {
            PLAN_MAT_COL: row[PLAN_MAT_COL],
            PLAN_PROD_COL: row[PLAN_PROD_COL],
            PLAN_START_COL: row[PLAN_START_COL],
            'Component Number': row[PLAN_MAT_COL],
            'Component Description': "(Parent Item)",
            'Plan Quantity': row[PLAN_QTY_COL],
            'Necessary Quantity': row[PLAN_QTY_COL],
            'Sort Key': 0
        }
        results.append(parent_info)

        # 2. CU/DUから子品目を抽出
        prod_code = str(row[PLAN_PROD_COL]).strip()
        for m_df in [cu_df, du_df]:
            if MASTER_KEY in m_df.columns:
                children = m_df[m_df[MASTER_KEY].astype(str) == prod_code]
                for _, child in children.iterrows():
                    results.append({
                        PLAN_MAT_COL: row[PLAN_MAT_COL],
                        PLAN_PROD_COL: row[PLAN_PROD_COL],
                        PLAN_START_COL: row[PLAN_START_COL],
                        'Component Number': child[MASTER_COMP_NUM_COL],
                        'Component Description': child[MASTER_DESC_COL],
                        'Plan Quantity': row[PLAN_QTY_COL],
                        'Necessary Quantity': compute_necessary_qty(child, row[PLAN_QTY_COL]),
                        'Sort Key': 1
                    })
    
    return pd.DataFrame(results)

# --- Streamlit UI ---
st.set_page_config(page_title="SAP PO Auditor Pro", layout="wide")
st.title("📊 SAP製造指示 階層型監査レポート作成")

with st.sidebar:
    st.header("1. マスタ読み込み")
    cu_file = st.file_uploader("CUリスト", type=["xlsx"])
    du_file = st.file_uploader("DUリスト", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("旧計画ファイル (.xlsm)", type=["xlsm", "xlsx"])
with col2:
    new_file = st.file_uploader("新計画ファイル (.xlsm)", type=["xlsm", "xlsx"])

if st.button("🔍 レポートを生成"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("ファイルをすべてアップロードしてください。")
    else:
        try:
            cu_m = pd.read_excel(cu_file)
            du_m = pd.read_excel(du_file)
            old_plans = get_plan_data(old_file)
            new_plans = get_plan_data(new_file)
            
            all_names = set(old_plans.keys()).union(set(new_plans.keys()))
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                diff_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                for name in all_names:
                    old_bom = process_bom_structure(old_plans.get(name, pd.DataFrame()), cu_m, du_m)
                    new_bom = process_bom_structure(new_plans.get(name, pd.DataFrame()), cu_m, du_m)

                    # 旧・新を外部結合（品目、日付、構成品、製品記号で紐付け）
                    merge_keys = [PLAN_MAT_COL, PLAN_PROD_COL, PLAN_START_COL, 'Component Number']
                    comparison = pd.merge(old_bom, new_bom, on=merge_keys, how='outer', suffixes=('_旧', '_新'))

                    # 数値・テキストの欠損埋め
                    comparison.fillna({'Necessary Quantity_旧': 0, 'Necessary Quantity_新': 0, 
                                      'Plan Quantity_旧': 0, 'Plan Quantity_新': 0,
                                      'Component Description_旧': '(Deleted)', 'Component Description_新': '(Added)'}, inplace=True)

                    # 並び替え（製品記号 > 日付 > 親品目 > Sort Key）
                    comparison['Sort_Temp'] = comparison['Sort Key_新'].fillna(comparison['Sort Key_旧'])
                    comparison.sort_values(by=[PLAN_PROD_COL, PLAN_START_COL, PLAN_MAT_COL, 'Sort_Temp'], inplace=True)
                    
                    # 不要なSort列を削除して出力
                    final_df = comparison.drop(columns=['Sort Key_旧', 'Sort Key_新', 'Sort_Temp'])
                    
                    safe_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                    final_df.to_excel(writer, index=False, sheet_name=safe_name)

                    # 差異行のハイライト（Necessary Quantityの差）
                    worksheet = writer.sheets[safe_name]
                    for i, row in enumerate(final_df.itertuples()):
                        if abs(row._6 - row._10) > 0.1: # インデックスは列数に合わせて調整
                            worksheet.set_row(i + 1, None, diff_format)

            st.success("レポートの作成が完了しました。")
            st.download_button("📥 階層レポートをダウンロード", output.getvalue(), "SAP_PO_Audit_Comparison.xlsx")
        except Exception as e:
            st.error(f"システムエラー: {e}")
