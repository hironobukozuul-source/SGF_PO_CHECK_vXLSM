import streamlit as st
import pandas as pd
import io
import math

# --- Configuration & Column Definitions [cite: 6, 12, 17] ---
PLAN_MAT_COL = "品目コード"
PLAN_PROD_COL = "製品記号"
PLAN_START_COL = "製造開始日"
PLAN_QTY_COL = "Quantity"

MASTER_KEY = "Parent material number"
MASTER_COMP_NUM_COL = "Component Number"
MASTER_DESC_COL = "Component Description"
MATERIAL_TYPE_COL = "Material Type"
TARGET_TYPE = "VERP"

# Exclusion keywords per specification [cite: 12]
EXCLUDE_KEYWORDS = ["TAPE", "GLUE", "INK", "SOLVENT"]

def get_plan_data(uploaded_file):
    """Extracts plan data from Excel sheets starting with 'Day bucket plan_12' [cite: 5]"""
    if uploaded_file is None: return {}
    plans_dict = {}
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
    target_sheets = [s for s in xls.sheet_names if s.startswith("Day bucket plan_12")]
    for sheet in target_sheets:
        meta_df = pd.read_excel(xls, sheet_name=sheet, nrows=1, header=None)
        plan_name = str(meta_df.iloc[0, 0]).strip() if not meta_df.empty else sheet
        # Skip top 5 rows as per file structure [cite: 5]
        data_df = pd.read_excel(xls, sheet_name=sheet, skiprows=5)
        data_df = data_df.dropna(how='all')
        if not data_df.empty:
            plans_dict[plan_name] = data_df
    return plans_dict

def compute_qty(row, plan_qty):
    """Calculates quantity with BOTTLE/PUMP exceptions and math.ceil rounding [cite: 14, 15, 16]"""
    desc = str(row.get(MASTER_DESC_COL, '')).upper()
    # Rule: BOTTLE or PUMP equals Plan Qty [cite: 16]
    if "BOTTLE" in desc or "PUMP" in desc:
        base_qty = plan_qty
    else:
        p_qty = row.get("Parent Material Quantity", 1)
        c_qty = row.get("Component Quantity", 0)
        if pd.isna(p_qty) or p_qty == 0: 
            base_qty = 0
        else:
            # Rule: (Plan Qty / Parent Material Qty) * Comp Qty [cite: 15]
            base_qty = (plan_qty / p_qty) * c_qty
    return math.ceil(base_qty)

def is_excluded(description):
    """Filters out items based on keyword list [cite: 12]"""
    desc_upper = str(description).upper()
    return any(kw in desc_upper for kw in EXCLUDE_KEYWORDS)

def create_structured_bom(plan_df, cu_df, du_df):
    """Main Search Logic: DU -> CU -> VERP Filter [cite: 7, 8, 9, 10]"""
    # Define columns to prevent KeyError during empty merges
    cols = ['Parent Mat', 'Product Code', 'Start Date', 'Comp Number', 'Comp Name', 'Need Qty', 'Level']
    if plan_df.empty: return pd.DataFrame(columns=cols)

    for df in [cu_df, du_df]:
        if df is not None and not df.empty:
            for c in [MASTER_KEY, MASTER_COMP_NUM_COL, MASTER_DESC_COL]:
                if c in df.columns: df[c] = df[c].astype(str).str.strip()

    structured_data = []

    for _, row in plan_df.iterrows():
        p_mat = str(row[PLAN_MAT_COL]).strip()
        p_qty = row[PLAN_QTY_COL]
        
        # Level 0 (Parent) [cite: 14]
        structured_data.append({
            'Parent Mat': p_mat, 'Product Code': row[PLAN_PROD_COL], 'Start Date': row[PLAN_START_COL],
            'Comp Number': p_mat, 'Comp Name': "(Parent Item)", 'Need Qty': math.ceil(p_qty), 'Level': 0
        })

        # Step 1: DU Search [cite: 7]
        if du_df is not None and not du_df.empty:
            du_children = du_df[du_df[MASTER_KEY] == p_mat]
            
            for _, child in du_children.iterrows():
                comp_num = child[MASTER_COMP_NUM_COL]
                comp_desc = child[MASTER_DESC_COL]
                
                if is_excluded(comp_desc): continue

                # Step 2: CU Search logic [cite: 8, 9]
                if comp_desc.endswith("_CU"):
                    if cu_df is not None and not cu_df.empty:
                        intermediate_qty = compute_qty(child, p_qty)
                        cu_items = cu_df[cu_df[MASTER_KEY] == comp_num]
                        for _, cu_item in cu_items.iterrows():
                            if is_excluded(cu_item[MASTER_DESC_COL]): continue
                            # Step 3: Target VERP Only [cite: 10]
                            if str(cu_item.get(MATERIAL_TYPE_COL)) == TARGET_TYPE:
                                structured_data.append({
                                    'Parent Mat': p_mat, 'Product Code': row[PLAN_PROD_COL], 'Start Date': row[PLAN_START_COL],
                                    'Comp Number': cu_item[MASTER_COMP_NUM_COL], 'Comp Name': cu_item[MASTER_DESC_COL],
                                    'Need Qty': compute_qty(cu_item, intermediate_qty), 'Level': 1
                                })
                else:
                    if str(child.get(MATERIAL_TYPE_COL)) == TARGET_TYPE:
                        structured_data.append({
                            'Parent Mat': p_mat, 'Product Code': row[PLAN_PROD_COL], 'Start Date': row[PLAN_START_COL],
                            'Comp Number': comp_num, 'Comp Name': comp_desc,
                            'Need Qty': compute_qty(child, p_qty), 'Level': 1
                        })
    
    return pd.DataFrame(structured_data) if structured_data else pd.DataFrame(columns=cols)

# --- UI Layout [cite: 21, 22] ---
st.set_page_config(page_title="SAP Audit Tool V12", layout="wide")
st.title("📊 SAP監査レポート作成 (Stability v12.2)")

with st.sidebar:
    st.header("1. マスターデータ")
    cu_file = st.file_uploader("CUリスト (Parent material number 検索用)", type=["xlsx"])
    du_file = st.file_uploader("DUリスト (品目コード 検索用)", type=["xlsx"])
    st.divider()
    st.write("**除外設定:**", ", ".join(EXCLUDE_KEYWORDS))

st.header("2. 計画ファイル比較")
c1, c2 = st.columns(2)
with c1: old_file = st.file_uploader("旧計画 (Old Plan)", type=["xlsm", "xlsx"])
with c2: new_file = st.file_uploader("新計画 (New Plan)", type=["xlsm", "xlsx"])

if st.button("🔍 レポート作成"):
    if not (cu_file and du_file and old_file and new_file):
        st.warning("すべてのファイルをアップロードしてください。")
    else:
        try:
            cu_m = pd.read_excel(cu_file)
            du_m = pd.read_excel(du_file)
            old_plans = get_plan_data(old_file)
            new_plans = get_plan_data(new_file)
            
            all_names = sorted(set(old_plans.keys()).union(set(new_plans.keys())))
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                red_fmt = writer.book.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                
                for name in all_names:
                    old_bom = create_structured_bom(old_plans.get(name, pd.DataFrame()), cu_m, du_m)
                    new_bom = create_structured_bom(new_plans.get(name, pd.DataFrame()), cu_m, du_m)
                    
                    m_keys = ['Parent Mat', 'Start Date', 'Comp Number']
                    
                    # Merge with safety handling for empty dataframes
                    df = pd.merge(old_bom, new_bom, on=m_keys, how='outer', suffixes=('_旧', '_新'))
                    df.fillna({'Need Qty_旧': 0, 'Need Qty_新': 0}, inplace=True)
                    
                    if df.empty: continue

                    # Sorting logic [cite: 20]
                    df['P_Sort'] = df['Product Code_旧'].fillna(df['Product Code_新'])
                    df['L_Sort'] = df['Level_新'].fillna(df['Level_旧'])
                    df = df.sort_values(['P_Sort', 'Start Date', 'Parent Mat', 'L_Sort']).drop(columns=['P_Sort', 'L_Sort'])

                    # Excel output [cite: 18, 19]
                    sheet_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                    df.to_excel(writer, index=False, sheet_name=sheet_name)

                    ws = writer.sheets[sheet_name]
                    cols = df.columns.tolist()
                    idx_o, idx_n = cols.index('Need Qty_旧'), cols.index('Need Qty_新')
                    
                    for i, r in enumerate(df.itertuples(index=False)):
                        # Highlight if difference >= 1 [cite: 19]
                        if abs(r[idx_o] - r[idx_n]) >= 1:
                            ws.set_row(i + 1, None, red_fmt)

            st.success("レポート作成成功！")
            st.download_button(
                label="📥 レポートをダウンロード",
                data=output.getvalue(),
                file_name="プラン変更_PO確認.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"システムエラー: {e}")
