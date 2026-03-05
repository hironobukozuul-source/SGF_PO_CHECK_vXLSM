import streamlit as st
import pandas as pd
import io

# --- 定義 ---
PLAN_MAT_COL = "品目コード"
PLAN_PROD_COL = "製品記号"
PLAN_START_COL = "製造開始日"
PLAN_QTY_COL = "Quantity"

MASTER_KEY = "Parent material number"
MASTER_COMP_NUM_COL = "Component Number"
MASTER_DESC_COL = "Component Description"
MATERIAL_TYPE_COL = "Material Type"
TARGET_TYPE = "VERP"

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

def compute_qty(row, plan_qty):
    desc = str(row.get(MASTER_DESC_COL, '')).upper()
    if "BOTTLE" in desc or "PUMP" in desc:
        return plan_qty
    p_qty = row.get("Parent Material Quantity", 1)
    c_qty = row.get("Component Quantity", 0)
    if pd.isna(p_qty) or p_qty == 0: return 0
    return (plan_qty / p_qty) * c_qty

def create_structured_bom(plan_df, cu_df, du_df):
    if plan_df.empty: return pd.DataFrame()

    # マスタのクリーニング
    for df in [cu_df, du_df]:
        if df is not None and not df.empty:
            df[MASTER_KEY] = df[MASTER_KEY].astype(str).str.strip()
            df[MASTER_COMP_NUM_COL] = df[MASTER_COMP_NUM_COL].astype(str).str.strip()
            df[MASTER_DESC_COL] = df[MASTER_DESC_COL].astype(str).str.strip()

    structured_data = []

    for _, row in plan_df.iterrows():
        p_mat = str(row[PLAN_MAT_COL]).strip()
        p_qty = row[PLAN_QTY_COL]
        
        # 1. 親行
        structured_data.append({
            'Parent Mat': p_mat, 'Product Code': row[PLAN_PROD_COL], 'Start Date': row[PLAN_START_COL],
            'Comp Number': p_mat, 'Comp Name': "(Parent Item)", 'Need Qty': p_qty, 'Level': 0
        })

        # 2. DUリストからの探索 (Step 1)
        if du_df is not None and not du_df.empty:
            du_children = du_df[du_df[MASTER_KEY] == p_mat]
            
            for _, child in du_children.iterrows():
                comp_num = child[MASTER_COMP_NUM_COL]
                comp_desc = child[MASTER_DESC_COL]
                desc_upper = comp_desc.upper()
                
                # 除外フィルタ (Tape/Glue)
                if "TAPE" in desc_upper or "GLUE" in desc_upper: continue

                # Step 2: Component Description が "_CU" で終わるか判定
                if comp_desc.endswith("_CU"):
                    # その行の Component Number を使って CUリストを検索
                    cu_search_key = comp_num
                    
                    if cu_df is not None and not cu_df.empty:
                        # 計算された中間数量（DUベース）
                        intermediate_qty = compute_qty(child, p_qty)
                        
                        cu_items = cu_df[cu_df[MASTER_KEY] == cu_search_key]
                        for _, cu_item in cu_items.iterrows():
                            cu_desc = str(cu_item[MASTER_DESC_COL]).upper()
                            if "TAPE" in cu_desc or "GLUE" in cu_desc: continue
                            
                            # VERPのみ追加
                            if str(cu_item.get(MATERIAL_TYPE_COL)) == TARGET_TYPE:
                                structured_data.append({
                                    'Parent Mat': p_mat, 'Product Code': row[PLAN_PROD_COL], 'Start Date': row[PLAN_START_COL],
                                    'Comp Number': cu_item[MASTER_COMP_NUM_COL], 'Comp Name': cu_item[MASTER_DESC_COL],
                                    'Need Qty': compute_qty(cu_item, intermediate_qty), 'Level': 1
                                })
                else:
                    # _CUでない通常のVERP子アイテムをDUから追加
                    if str(child.get(MATERIAL_TYPE_COL)) == TARGET_TYPE:
                        structured_data.append({
                            'Parent Mat': p_mat, 'Product Code': row[PLAN_PROD_COL], 'Start Date': row[PLAN_START_COL],
                            'Comp Number': comp_num, 'Comp Name': comp_desc,
                            'Need Qty': compute_qty(child, p_qty), 'Level': 1
                        })

    return pd.DataFrame(structured_data)

# --- UI ---
st.set_page_config(page_title="SAP Audit Tool V9", layout="wide")
st.title("📊 SAP監査レポート (Description末尾判定版)")

with st.sidebar:
    st.info("新ロジック:\n1. DU内で品目コード検索\n2. Nameが'_CU'で終わる行のNumberをキーにCU内を検索")
    cu_file = st.file_uploader("CUリスト", type=["xlsx"])
    du_file = st.file_uploader("DUリスト", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("旧計画ファイル", type=["xlsm", "xlsx"])
with col2:
    new_file = st.file_uploader("新計画ファイル", type=["xlsm", "xlsx"])

if st.button("🔍 レポート作成"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("ファイルをすべてアップロードしてください")
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
                    
                    m_keys = ['Parent Mat', 'Start Date', 'Comp Number']
                    df = pd.merge(old_bom, new_bom, on=m_keys, how='outer', suffixes=('_旧', '_新'))
                    df.fillna({'Need Qty_旧': 0, 'Need Qty_新': 0}, inplace=True)
                    
                    # 並び替え用の一時列
                    df['Prod_C'] = df['Product Code_旧'].fillna(df['Product Code_新'])
                    df['Lvl_S'] = df['Level_新'].fillna(df['Level_旧'])
                    df = df.sort_values(['Prod_C', 'Start Date', 'Parent Mat', 'Lvl_S'])
                    df = df.drop(columns=['Prod_C', 'Lvl_S'])

                    sheet_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                    df.to_excel(writer, index=False, sheet_name=sheet_name)

                    ws = writer.sheets[sheet_name]
                    cols = df.columns.tolist()
                    idx_o, idx_n = cols.index('Need Qty_旧'), cols.index('Need Qty_新')
                    for i, r in enumerate(df.itertuples(index=False)):
                        if abs(r[idx_o] - r[idx_n]) > 0.1:
                            ws.set_row(i + 1, None, red_format)

            st.success("作成完了。CUリストのアイテムが抽出されているか確認してください。")
            st.download_button("📥 ダウンロード", output.getvalue(), "SAP_Audit_DescriptionSearch.xlsx")
        except Exception as e:
            st.error(f"システムエラー: {e}")
