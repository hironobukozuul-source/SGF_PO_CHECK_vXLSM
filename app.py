import streamlit as st
import pandas as pd
import io

# --- カラム定義（環境に合わせて調整可能） ---
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
    """Excelの指定シートからRow 6以降を取得"""
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
    """BOTTLE/PUMP判定を含む数量計算ロジック"""
    desc = str(row.get(MASTER_DESC_COL, '')).upper()
    if "BOTTLE" in desc or "PUMP" in desc:
        return plan_qty
    p_qty = row.get(MASTER_PARENT_QTY_COL, 1)
    c_qty = row.get(MASTER_COMP_QTY_COL, 0)
    if pd.isna(p_qty) or p_qty == 0: return 0
    return (plan_qty / p_qty) * c_qty

def create_structured_bom(plan_df, cu_df, du_df):
    """親品目の直後に子品目(VERP)を配置したリストを作成"""
    if plan_df.empty: return pd.DataFrame()

    # マスタをVERPのみに事前フィルタ
    masters = []
    for df in [cu_df, du_df]:
        if df is not None and not df.empty:
            if MATERIAL_TYPE_COL in df.columns:
                masters.append(df[df[MATERIAL_TYPE_COL] == TARGET_TYPE].copy())
            else:
                masters.append(df.copy())
    
    full_master = pd.concat(masters, ignore_index=True) if masters else pd.DataFrame()
    if not full_master.empty:
        full_master[MASTER_KEY] = full_master[MASTER_KEY].astype(str).str.strip()

    structured_data = []
    for _, row in plan_df.iterrows():
        parent_mat = str(row[PLAN_MAT_COL]).strip()
        plan_qty = row[PLAN_QTY_COL]

        # 1. 親品目行の追加
        parent_row = {
            'Material Code': parent_mat,
            'Product Code': row[PLAN_PROD_COL],
            'Production Start': row[PLAN_START_COL],
            'Component Number': parent_mat,
            'Component Description': "(Parent Item)",
            'Plan Quantity': plan_qty,
            'Necessary Quantity': plan_qty,
            'Sort Order': 0 # 親
        }
        structured_data.append(parent_row)

        # 2. CU/DUマスタから、この親の品目コードに紐づく子を検索
        if not full_master.empty:
            children = full_master[full_master[MASTER_KEY] == parent_mat]
            for _, child in children.iterrows():
                child_row = {
                    'Material Code': parent_mat,
                    'Product Code': row[PLAN_PROD_COL],
                    'Production Start': row[PLAN_START_COL],
                    'Component Number': child[MASTER_COMP_NUM_COL],
                    'Component Description': child[MASTER_DESC_COL],
                    'Plan Quantity': plan_qty,
                    'Necessary Quantity': compute_necessary_qty(child, plan_qty),
                    'Sort Order': 1 # 子
                }
                structured_data.append(child_row)

    return pd.DataFrame(structured_data)

# --- Streamlit UI ---
st.set_page_config(page_title="SAP Audit Tool V5", layout="wide")
st.title("📊 SAP監査レポート (親・子 階層構造版)")
st.info("親品目のすぐ下に、CU/DUリストから抽出された子品目(VERP)が自動で並びます。")

with st.sidebar:
    st.header("1. マスタデータ")
    cu_file = st.file_uploader("CUリスト", type=["xlsx"])
    du_file = st.file_uploader("DUリスト", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("旧計画ファイル (.xlsm)", type=["xlsm", "xlsx"])
with col2:
    new_file = st.file_uploader("新計画ファイル (.xlsm)", type=["xlsm", "xlsx"])

if st.button("🔍 比較レポートを作成"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("すべてのファイルをアップロードしてください。")
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
                red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                for name in sorted(all_names):
                    # 各計画ごとに親子構造リストを作成
                    old_bom = create_structured_bom(old_plans.get(name, pd.DataFrame()), cu_m, du_m)
                    new_bom = create_structured_bom(new_plans.get(name, pd.DataFrame()), cu_m, du_m)

                    # 旧と新を外部結合（品目、日付、構成品番号で紐付け）
                    merge_keys = ['Material Code', 'Production Start', 'Component Number']
                    comparison = pd.merge(old_bom, new_bom, on=merge_keys, how='outer', suffixes=('_旧', '_新'))

                    # 欠損値の穴埋め (削除/追加対応)
                    comparison['Necessary Quantity_旧'] = comparison['Necessary Quantity_旧'].fillna(0)
                    comparison['Necessary Quantity_新'] = comparison['Necessary Quantity_新'].fillna(0)
                    comparison['Product Code_旧'] = comparison['Product Code_旧'].fillna(comparison['Product Code_新'])
                    comparison['Component Description_旧'] = comparison['Component Description_旧'].fillna('(New Item)')
                    comparison['Component Description_新'] = comparison['Component Description_新'].fillna('(Deleted Item)')

                    # 並び替え: 製品記号 > 日付 > 親品目コード > Sort Order (親0、子1)
                    # 結合後にSort Orderが2つできるため、統合して使用
                    comparison['Sort_Key'] = comparison['Sort Order_新'].fillna(comparison['Sort Order_旧'])
                    comparison = comparison.sort_values(by=['Product Code_旧', 'Production Start', 'Material Code', 'Sort_Key'])

                    # 不要な列を整理して出力
                    final_df = comparison.drop(columns=['Sort Order_旧', 'Sort Order_新', 'Sort_Key'])
                    
                    safe_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                    final_df.to_excel(writer, index=False, sheet_name=safe_name)

                    # 数量に差がある行をハイライト
                    worksheet = writer.sheets[safe_name]
                    for i, row in enumerate(final_df.itertuples()):
                        # 必要数量(旧)と(新)の比較
                        # final_dfの列順に注意 (Necessary Quantity_旧 と _新 の位置)
                        if abs(row._7 - row._12) > 0.001: 
                            worksheet.set_row(i + 1, None, red_format)

            st.success("レポートが完成しました。")
            st.download_button("📥 階層型レポートをダウンロード", output.getvalue(), "SAP_PO_Hierarchical_Audit.xlsx")
        except Exception as e:
            st.error(f"システムエラー: {e}")
