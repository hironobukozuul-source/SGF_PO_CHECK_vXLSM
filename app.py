import streamlit as st
import pandas as pd
import io

# --- 設定とヘッダー定義 ---
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

# --- ロジック関数 ---

def get_plan_data(uploaded_file):
    if uploaded_file is None:
        return {}
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
    
    if pd.isna(p_qty) or p_qty == 0:
        return 0
    return (plan_qty / p_qty) * c_qty

def process_master_merge(plan_df, master_df):
    if master_df is None or master_df.empty:
        return pd.DataFrame()

    required_cols = [MASTER_KEY, MASTER_DESC_COL, MASTER_COMP_NUM_COL]
    for col in required_cols:
        if col not in master_df.columns:
            return pd.DataFrame()

    if MATERIAL_TYPE_COL in master_df.columns:
        master_df = master_df[master_df[MATERIAL_TYPE_COL] == TARGET_TYPE].copy()

    plan_df[PLAN_PROD_COL] = plan_df[PLAN_PROD_COL].astype(str).str.strip()
    master_df[MASTER_KEY] = master_df[MASTER_KEY].astype(str).str.strip()

    merged = plan_df.merge(master_df, left_on=PLAN_PROD_COL, right_on=MASTER_KEY, how='inner')
    
    if merged.empty:
        return pd.DataFrame()

    merged['Component Number'] = merged[MASTER_COMP_NUM_COL]
    merged['Necessary Quantity'] = merged.apply(lambda r: compute_necessary_qty(r, r[PLAN_QTY_COL]), axis=1)
    merged['Sort Key'] = 1  # 構成品はソート順を1にする（親の直後）
    
    return merged[[PLAN_MAT_COL, PLAN_PROD_COL, PLAN_START_COL, 'Component Number', 'Necessary Quantity', 'Sort Key']]

def calculate_bom(plan_df, cu_df, du_df):
    # 1. CU/DUの構成品取得
    res_cu = process_master_merge(plan_df, cu_df)
    res_du = process_master_merge(plan_df, du_df)
    
    # 2. 親品目（Self）の作成
    plan_self = plan_df.copy()
    plan_self['Component Number'] = plan_self[PLAN_MAT_COL].astype(str)
    plan_self['Necessary Quantity'] = plan_self[PLAN_QTY_COL]
    plan_self['Sort Key'] = 0  # 親品目はソート順を0にする（最優先）
    res_self = plan_self[[PLAN_MAT_COL, PLAN_PROD_COL, PLAN_START_COL, 'Component Number', 'Necessary Quantity', 'Sort Key']]

    # 3. 全結合と並び替え
    combined = pd.concat([res_self, res_cu, res_du], ignore_index=True)
    # 製品記号、製造開始日、品目コードでグループ化し、その中で親(0) -> 子(1)の順に並べる
    combined = combined.sort_values(by=[PLAN_PROD_COL, PLAN_START_COL, PLAN_MAT_COL, 'Sort Key'])
    
    return combined.drop(columns=['Sort Key'])

# --- Streamlit UI ---

st.set_page_config(page_title="SAP Audit Tool V3", layout="wide")
st.title("📊 SAP監査レポート作成 (製品記号・階層表示対応)")

with st.sidebar:
    st.header("1. マスタ読み込み")
    cu_file = st.file_uploader("CUリスト", type=["xlsx"])
    du_file = st.file_uploader("DUリスト", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("旧計画ファイル", type=["xlsm", "xlsx"])
with col2:
    new_file = st.file_uploader("新計画ファイル", type=["xlsm", "xlsx"])

if st.button("🔍 レポートを作成"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("全てのファイルをアップロードしてください。")
    else:
        try:
            cu_m = pd.read_excel(cu_file)
            du_m = pd.read_excel(du_file)
            old_plans = get_plan_data(old_file)
            new_plans = get_plan_data(new_file)

            common_names = set(old_plans.keys()).intersection(set(new_plans.keys()))

            if not common_names:
                st.error("一致する計画名(A1セル)が見つかりません。")
            else:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                    for name in common_names:
                        old_bom = calculate_bom(old_plans[name], cu_m, du_m)
                        new_bom = calculate_bom(new_plans[name], cu_m, du_m)

                        # 比較（製品記号を含めてマージ）
                        merge_cols = [PLAN_MAT_COL, PLAN_PROD_COL, PLAN_START_COL, 'Component Number']
                        comparison = pd.merge(old_bom, new_bom, on=merge_cols, how='outer', suffixes=('_旧', '_新'))
                        
                        comparison['Necessary Quantity_旧'] = comparison['Necessary Quantity_旧'].fillna(0)
                        comparison['Necessary Quantity_新'] = comparison['Necessary Quantity_新'].fillna(0)

                        # 再度並び替え（マージ後に崩れる可能性があるため）
                        comparison = comparison.sort_values(by=[PLAN_PROD_COL, PLAN_START_COL, PLAN_MAT_COL])

                        safe_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                        comparison.to_excel(writer, index=False, sheet_name=safe_name)
                        
                        worksheet = writer.sheets[safe_name]
                        for row_num, (o_q, n_q) in enumerate(zip(comparison['Necessary Quantity_旧'], comparison['Necessary Quantity_新'])):
                            if abs(o_q - n_q) > 0.1:
                                worksheet.set_row(row_num + 1, None, red_format)

                st.success(f"{len(common_names)} 件のシートを処理しました。")
                st.download_button("📥 最終レポートをダウンロード", output.getvalue(), "SAP_Final_Audit_Report.xlsx")
        except Exception as e:
            st.error(f"システムエラー: {e}")
