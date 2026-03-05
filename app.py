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

def create_structured_bom(plan_df, cu_df, du_df):
    """品目コードをキーにCU/DUから子アイテムを検索し、階層構造を作成"""
    if plan_df.empty: return pd.DataFrame()

    # マスタ結合
    masters = [df for df in [cu_df, du_df] if df is not None and not df.empty]
    full_master = pd.concat(masters, ignore_index=True) if masters else pd.DataFrame()
    
    if not full_master.empty:
        full_master[MASTER_KEY] = full_master[MASTER_KEY].astype(str).str.strip()
        if MATERIAL_TYPE_COL in full_master.columns:
            full_master = full_master[full_master[MATERIAL_TYPE_COL] == TARGET_TYPE]

    structured_data = []
    for _, row in plan_df.iterrows():
        p_mat = str(row[PLAN_MAT_COL]).strip()
        p_qty = row[PLAN_QTY_COL]

        # 1. 親行
        structured_data.append({
            'Parent Mat': p_mat,
            'Product Code': row[PLAN_PROD_COL],
            'Start Date': row[PLAN_START_COL],
            'Comp Number': p_mat,
            'Comp Name': "(Parent Item)",
            'Plan Qty': p_qty,
            'Need Qty': p_qty,
            'Level': 0
        })

        # 2. 子行 (品目コードで検索)
        if not full_master.empty:
            children = full_master[full_master[MASTER_KEY] == p_mat]
            for _, child in children.iterrows():
                structured_data.append({
                    'Parent Mat': p_mat,
                    'Product Code': row[PLAN_PROD_COL],
                    'Start Date': row[PLAN_START_COL],
                    'Comp Number': child[MASTER_COMP_NUM_COL],
                    'Comp Name': child[MASTER_DESC_COL],
                    'Plan Qty': p_qty,
                    'Need Qty': compute_necessary_qty(child, p_qty),
                    'Level': 1
                })

    return pd.DataFrame(structured_data)

# --- UI ---
st.set_page_config(page_title="SAP Auditor Pro", layout="wide")
st.title("📊 SAP監査レポート作成")

with st.sidebar:
    cu_file = st.file_uploader("CUリスト", type=["xlsx"])
    du_file = st.file_uploader("DUリスト", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("旧計画ファイル", type=["xlsm", "xlsx"])
with col2:
    new_file = st.file_uploader("新計画ファイル", type=["xlsm", "xlsx"])

if st.button("🔍 レポート作成"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("全ファイルをアップロードしてください")
    else:
        try:
            cu_m = pd.read_excel(cu_file)
            du_m = pd.read_excel(du_file)
            old_plans = get_plan_data(old_file)
            new_plans = get_plan_data(new_file)
            
            all_names = sorted(set(old_plans.keys()).union(set(new_plans.keys())))
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                red_format = writer.book.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                for name in all_names:
                    old_bom = create_structured_bom(old_plans.get(name, pd.DataFrame()), cu_m, du_m)
                    new_bom = create_structured_bom(new_plans.get(name, pd.DataFrame()), cu_m, du_m)

                    # 外部結合
                    m_keys = ['Parent Mat', 'Start Date', 'Comp Number']
                    df = pd.merge(old_bom, new_bom, on=m_keys, how='outer', suffixes=('_旧', '_新'))

                    # クリーニングと並び替え
                    df.fillna({'Need Qty_旧': 0, 'Need Qty_新': 0}, inplace=True)
                    df['Sort_L'] = df['Level_新'].fillna(df['Level_旧'])
                    df['Prod_C'] = df['Product Code_旧'].fillna(df['Product Code_新'])
                    df = df.sort_values(['Prod_C', 'Start Date', 'Parent Mat', 'Sort_L'])
                    
                    # 不要列削除
                    df = df.drop(columns=['Sort_L', 'Prod_C'])
                    
                    sheet_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                    df.to_excel(writer, index=False, sheet_name=sheet_name)

                    # --- エラー回避：列名で直接インデックスを取得してハイライト ---
                    ws = writer.sheets[sheet_name]
                    cols = df.columns.tolist()
                    try:
                        idx_old = cols.index('Need Qty_旧')
                        idx_new = cols.index('Need Qty_新')
                        for i, row in enumerate(df.itertuples(index=False)):
                            # row[idx_old] のようにアクセス
                            if abs(row[idx_old] - row[idx_new]) > 0.1:
                                ws.set_row(i + 1, None, red_format)
                    except ValueError:
                        continue

            st.success("作成完了")
            st.download_button("📥 ダウンロード", output.getvalue(), "SAP_Audit_Final.xlsx")
        except Exception as e:
            st.error(f"システムエラー: {e}")
