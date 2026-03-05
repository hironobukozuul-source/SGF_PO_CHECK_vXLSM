import streamlit as st
import pandas as pd
import io

# --- HELPER FUNCTIONS ---

def load_excel_plans(uploaded_file):
    """
    Reads all sheets starting with 'Day bucket plan_12' and combines them.
    Assumes standard SAP export format within the Excel sheets.
    """
    if uploaded_file is None:
        return None
    
    try:
        xls = pd.ExcelFile(uploaded_file)
        # Filter sheets starting with the specific prefix
        target_sheets = [s for s in xls.sheet_names if s.startswith("Day bucket plan_12")]
        
        if not target_sheets:
            st.warning(f"No sheets starting with 'Day bucket plan_12' found in {uploaded_file.name}")
            return None

        all_data = []
        for sheet in target_sheets:
            # Note: You may need to adjust 'header' or 'skiprows' if your Excel 
            # export has specific top-row padding.
            df = pd.read_excel(xls, sheet_name=sheet)
            
            # Basic cleaning: Drop rows where critical SAP identifiers are missing
            if 'Material Code' in df.columns:
                df = df.dropna(subset=['Material Code'])
            
            df['Source_Sheet'] = sheet
            all_data.append(df)
        
        return pd.concat(all_data, ignore_index=True)
    except Exception as e:
        st.error(f"Error reading {uploaded_file.name}: {e}")
        return None

def calculate_bom(plan_df, cu_df, du_df):
    """
    Maps Master Data to the Production Plan to calculate Necessary Quantities.
    """
    # 1. Map CU Components
    plan_cu = plan_df.merge(cu_df, on='Product Code', how='left')
    plan_cu['Component Number'] = plan_cu['Component Number_CU']
    plan_cu['Component Description'] = plan_cu['Component Description_CU']
    plan_cu['Necessary Quantity'] = plan_cu['Volume(pcs)'] * plan_cu['CU_Ratio']

    # 2. Map DU Components
    plan_du = plan_df.merge(du_df, on='Product Code', how='left')
    plan_du['Component Number'] = plan_du['Component Number_DU']
    plan_du['Component Description'] = plan_du['Component Description_DU']
    plan_du['Necessary Quantity'] = plan_du['Volume(pcs)'] * plan_du['DU_Ratio']

    # 3. Self-reference (FG as a component)
    plan_self = plan_df.copy()
    plan_self['Component Number'] = plan_self['Material Code']
    plan_self['Component Description'] = plan_self['Product Code']
    plan_self['Necessary Quantity'] = plan_self['Volume(pcs)']

    return pd.concat([plan_cu, plan_du, plan_self], ignore_index=True)

# --- STREAMLIT UI ---

st.set_page_config(page_title="SAP PO Auditor Pro", layout="wide")

st.title("📑 SAP PO Auditor (Excel Upgrade)")
st.markdown("Compare Rev06 and Rev07 Production Plans directly from XLSM files.")

# Sidebar: Master Data
with st.sidebar:
    st.header("1. Master Data (Excel)")
    cu_file = st.file_uploader("Upload CU List", type=["xlsx"])
    du_file = st.file_uploader("Upload DU List", type=["xlsx"])

# Main Area: Plan Uploads
col1, col2 = st.columns(2)
with col1:
    st.header("2. Old Plan (Rev06)")
    old_file = st.file_uploader("Upload Rev06 .xlsm", type=["xlsm", "xlsx"])

with col2:
    st.header("3. New Plan (Rev07)")
    new_file = st.file_uploader("Upload Rev07 .xlsm", type=["xlsm", "xlsx"])

if st.button("🔍 Generate Comparison Report"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("Please upload all 4 files (CU, DU, Rev06, and Rev07).")
    else:
        # Load Data
        cu_master = pd.read_excel(cu_file)
        du_master = pd.read_excel(du_file)
        
        df_old_raw = load_excel_plans(old_file)
        df_new_raw = load_excel_plans(new_file)

        if df_old_raw is not None and df_new_raw is not None:
            # Process BOMs
            old_bom = calculate_bom(df_old_raw, cu_master, du_master)
            new_bom = calculate_bom(df_new_raw, cu_master, du_master)

            # Align columns for comparison
            merge_cols = ['Material Code', 'Production Start', 'Component Number']
            comparison = pd.merge(
                old_bom, new_bom, 
                on=merge_cols, 
                how='outer', 
                suffixes=('_OLD', '_NEW')
            )

            # Fill NaNs with 0 for quantity comparison
            comparison['Necessary Quantity_OLD'] = comparison['Necessary Quantity_OLD'].fillna(0)
            comparison['Necessary Quantity_NEW'] = comparison['Necessary Quantity_NEW'].fillna(0)

            # Export with Highlighting
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                comparison.to_excel(writer, index=False, sheet_name='Comparison')
                
                workbook  = writer.book
                worksheet = writer.sheets['Comparison']

                # Define Red Format
                red_format = workbook.add_format({
                    'bg_color': '#FFC7CE', # Light Red
                    'font_color': '#9C0006' # Dark Red
                })

                # Apply static highlighting where quantities differ
                for row_num, (old_q, new_q) in enumerate(zip(comparison['Necessary Quantity_OLD'], comparison['Necessary Quantity_NEW'])):
                    if old_q != new_q:
                        worksheet.set_row(row_num + 1, None, red_format)

            st.success("Analysis complete!")
            st.download_button(
                label="📥 Download Comparison Report",
                data=output.getvalue(),
                file_name="PO_Audit_Comparison.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
