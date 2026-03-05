import streamlit as st
import pandas as pd
import io

# --- CONFIGURATION & HEADERS ---
# Headers in the Plan Sheets (Row 6)
PLAN_MAT_COL = "品目コード"
PLAN_PROD_COL = "製品記号"
PLAN_START_COL = "製造開始日"
PLAN_QTY_COL = "Quantity"

# Header in the Master Data (CU/DU Lists)
MASTER_KEY = "Parent material number"

# --- CORE LOGIC FUNCTIONS ---

def get_plan_data(uploaded_file):
    """
    Parses an Excel file (.xlsm).
    Key = Cell A1 (Plan Name)
    Value = DataFrame starting from Row 6 (Header row)
    """
    if uploaded_file is None:
        return {}
    
    plans_dict = {}
    # Use 'openpyxl' engine for .xlsm files
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
    
    # Process only sheets starting with the specific prefix
    target_sheets = [s for s in xls.sheet_names if s.startswith("Day bucket plan_12")]
    
    for sheet in target_sheets:
        # 1. Get Plan Name from Cell A1 (Row 0, Col 0)
        meta_df = pd.read_excel(xls, sheet_name=sheet, nrows=1, header=None)
        plan_name = str(meta_df.iloc[0, 0]).strip() if not meta_df.empty else sheet
        
        # 2. Get Data starting from Row 6 (skiprows=5)
        # Row 6 becomes the header
        data_df = pd.read_excel(xls, sheet_name=sheet, skiprows=5)
        data_df = data_df.dropna(how='all') # Remove empty spacer rows
        
        if not data_df.empty:
            plans_dict[plan_name] = data_df
            
    return plans_dict

def calculate_bom(plan_df, cu_df, du_df):
    """
    Maps Master Data 'Parent material number' to Plan '製品記号'.
    """
    # 1. Check for required Plan headers
    required_plan = [PLAN_MAT_COL, PLAN_PROD_COL, PLAN_QTY_COL, PLAN_START_COL]
    missing_plan = [c for c in required_plan if c not in plan_df.columns]
    if missing_plan:
        st.error(f"Plan Format Error: Row 6 is missing: {', '.join(missing_plan)}")
        return pd.DataFrame()

    # 2. Check for required Master headers
    if MASTER_KEY not in cu_df.columns:
        st.error(f"Master Error: CU List is missing '{MASTER_KEY}' column.")
        return pd.DataFrame()
    if MASTER_KEY not in du_df.columns:
        st.error(f"Master Error: DU List is missing '{MASTER_KEY}' column.")
        return pd.DataFrame()

    # --- Calculations ---
    # CU Components: Join Plan[製品記号] with Master[Parent material number]
    plan_cu = plan_df.merge(cu_df, left_on=PLAN_PROD_COL, right_on=MASTER_KEY, how='left')
    plan_cu['Component Number'] = plan_cu.get('Component Number_CU', "N/A")
    plan_cu['Necessary Quantity'] = plan_cu[PLAN_QTY_COL] * plan_cu.get('CU_Ratio', 0)

    # DU Components: Same join logic
    plan_du = plan_df.merge(du_df, left_on=PLAN_PROD_COL, right_on=MASTER_KEY, how='left')
    plan_du['Component Number'] = plan_du.get('Component Number_DU', "N/A")
    plan_du['Necessary Quantity'] = plan_du[PLAN_QTY_COL] * plan_du.get('DU_Ratio', 0)

    # Finished Good (Self) Entry
    plan_self = plan_df.copy()
    plan_self['Component Number'] = plan_self[PLAN_MAT_COL]
    plan_self['Necessary Quantity'] = plan_self[PLAN_QTY_COL]

    return pd.concat([plan_cu, plan_du, plan_self], ignore_index=True)

# --- STREAMLIT UI ---

st.set_page_config(page_title="SAP PO Auditor Pro", layout="wide")
st.title("📊 SAP Production Plan Comparison Tool")

st.markdown(f"""
### ⚙️ Mapping Rules:
- **Plan Header (Row 6):** Looks for `{PLAN_MAT_COL}`, `{PLAN_PROD_COL}`, `{PLAN_START_COL}`, `{PLAN_QTY_COL}`.
- **Master Link:** Plan's `{PLAN_PROD_COL}` matches Master's `{MASTER_KEY}`.
- **Output Names:** Generated sheets are named after the value in **Cell A1** of the input files.
""")

# Sidebar: Masters
with st.sidebar:
    st.header("1. Master Data")
    cu_file = st.file_uploader("Upload CU List (Excel)", type=["xlsx"])
    du_file = st.file_uploader("Upload DU List (Excel)", type=["xlsx"])

# Main: Plan Files
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("Upload Old Plan (.xlsm)", type=["xlsm", "xlsx"])
with col2:
    new_file = st.file_uploader("Upload New Plan (.xlsm)", type=["xlsm", "xlsx"])

if st.button("🔍 Run Multi-Sheet Comparison"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("Please upload all 4 files (2 Masters and 2 Plans).")
    else:
        try:
            with st.spinner("Processing files..."):
                # Load Masters
                cu_master = pd.read_excel(cu_file)
                du_master = pd.read_excel(du_file)

                # Parse Plans (Organized by Cell A1 content)
                old_plans = get_plan_data(old_file)
                new_plans = get_plan_data(new_file)

                # Match common plans based on the Cell A1 values
                common_names = set(old_plans.keys()).intersection(set(new_plans.keys()))

                if not common_names:
                    st.error("No matching plan names found in Cell A1 of both workbooks.")
                else:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        # Light red background for discrepancies
                        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                        for name in common_names:
                            # 1. Calculate BOMs for this specific plan
                            old_bom = calculate_bom(old_plans[name], cu_master, du_master)
                            new_bom = calculate_bom(new_plans[name], cu_master, du_master)

                            if old_bom.empty or new_bom.empty:
                                continue

                            # 2. Outer Join to find differences
                            merge_cols = [PLAN_MAT_COL, PLAN_START_COL, "Component Number"]
                            comparison = pd.merge(
                                old_bom, new_bom, on=merge_cols, how='outer', suffixes=('_OLD', '_NEW')
                            )
                            
                            comparison['Necessary Quantity_OLD'] = comparison['Necessary Quantity_OLD'].fillna(0)
                            comparison['Necessary Quantity_NEW'] = comparison['Necessary Quantity_NEW'].fillna(0)

                            # 3. Create Excel Sheet (named after Cell A1)
                            # Clean invalid Excel chars and truncate to 31 chars
                            safe_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                            comparison.to_excel(writer, index=False, sheet_name=safe_name)
                            
                            # 4. Apply Highlighting
                            worksheet = writer.sheets[safe_name]
                            for row_num, (old_q, new_q) in enumerate(zip(comparison['Necessary Quantity_OLD'], comparison['Necessary Quantity_NEW'])):
                                if old_q != new_q:
                                    worksheet.set_row(row_num + 1, None, red_format)

                    st.success(f"Audit Complete: Processed {len(common_names)} sheets.")
                    st.download_button(
                        label="📥 Download Detailed Audit Excel",
                        data=output.getvalue(),
                        file_name="SAP_Comparison_Audit_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"An error occurred: {e}")
