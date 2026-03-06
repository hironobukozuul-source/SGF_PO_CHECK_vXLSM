import streamlit as st
import pandas as pd
import io
import math

# --- 設定定義 [cite: 5, 7, 8, 10, 12] ---
PLAN_MAT_COL = "品目コード"
PLAN_PROD_COL = "製品記号"
PLAN_START_COL = "製造開始日"
PLAN_QTY_COL = "Quantity"

MASTER_KEY = "Parent material number"
MASTER_COMP_NUM_COL = "Component Number"
MASTER_DESC_COL = "Component Description"
MATERIAL_TYPE_COL = "Material Type"
TARGET_TYPE = "VERP"

EXCLUDE_KEYWORDS = ["TAPE", "GLUE", "INK", "SOLVENT"]

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
    """数量計算と切り上げ処理 [cite: 14, 15, 16]"""
    desc = str(row.get(MASTER_DESC_COL, '')).upper()
    if "BOTTLE" in desc or "PUMP" in desc:
        base_qty = plan_qty
    else:
        p_qty = row.get("Parent Material Quantity", 1)
        c_qty = row.get("Component Quantity", 0)
        base_qty = (plan_qty / p_qty) * c_qty if p_qty != 0 else 0
    return math.ceil(base_qty)

def is_excluded(description):
    """除外キーワード判定 [cite: 12]"""
    desc_upper = str(description).upper()
    return any(kw in desc_upper for kw in EXCLUDE_KEYWORDS)

def create_structured_bom(plan_df, cu_m, du_m):
    """BOM構成作成とデータ欠落のチェック [cite: 7, 8, 9, 10]"""
    cols = ['Parent Mat', 'Product Code', 'Start Date', 'Comp Number', 'Comp Name', 'Need Qty', 'Level']
    if plan_df.empty: return pd.DataFrame(columns=cols), set()

    structured_data = []
    missing_mats = set() # 見つからなかった品目を格納

    for _, row in plan_df.iterrows():
        p_mat = str(row[PLAN_MAT_COL]).strip()
        p_qty = row[PLAN_QTY_COL]
        
        # Level 0 追加
        structured_data.append({
            'Parent Mat': p_mat, 'Product Code': row[PLAN_PROD_COL], 'Start Date': row[PLAN_START_COL],
            'Comp Number': p_mat, 'Comp Name': "(Parent Item)", 'Need Qty': math.ceil(p_qty), 'Level': 0
        })

        # DU検索 [cite: 7]
        du_children = du_m[du_m[MASTER_KEY] == p_mat]
        
        if du_children.empty:
            missing_mats.add(p_mat) # DUリストに親品目がない場合
            continue

        found_verp = False
        for _, child in du_children.iterrows():
            comp_num = str(child[MASTER_COMP_NUM_COL]).strip()
            comp_desc = str(child[MASTER_DESC_COL]).strip()
            
            if is_excluded(comp_desc): continue

            # CU検索 [cite: 8, 9]
            if comp_desc.endswith("_CU"):
                intermediate_qty = compute_qty(child, p_qty)
                cu_items = cu_m[cu_m[MASTER_KEY] == comp_num]
                
                if cu_items.empty:
                    missing_mats.add(f"{p_mat} (CU: {comp_num} 不足)")
                    continue

                for _, cu_item in cu_items.iterrows():
                    if is_excluded(cu_item[MASTER_DESC_COL]): continue
                    if str(cu_item.get(MATERIAL_TYPE_COL)) == TARGET_TYPE:
                        structured_data.append({
                            'Parent Mat': p_mat, 'Product Code': row[PLAN_PROD_COL], 'Start Date': row[PLAN_START_COL],
                            'Comp Number': cu_item[MASTER_COMP_NUM_COL], 'Comp Name': cu_item[MASTER_DESC_COL],
                            'Need Qty': compute_qty(cu_item, intermediate_qty), 'Level': 1
                        })
                        found_verp = True
            else:
                if str(child.get(MATERIAL_TYPE_COL)) == TARGET_TYPE:
                    structured_data.append({
                        'Parent Mat': p_mat, 'Product Code': row[PLAN_PROD_COL], 'Start Date': row[PLAN_START_COL],
                        'Comp Number': comp_num, 'Comp Name': comp_desc,
                        'Need Qty': compute_qty(child, p_qty), 'Level': 1
                    })
                    found_verp = True
        
        # VERPが1つも見つからなかった場合（除外分を除く）
        if not found_verp:
            missing_mats.add(f"{p_mat} (対象VERPなし)")

    return pd.DataFrame(structured_data), missing_mats

# --- UI ---
st.set_page_config(page_title="SAP Audit Tool V12", layout="wide")
st.title("📊 SAP監査レポート作成 (エラー表示機能付)")

with st.sidebar:
    st.header("1. マスターデータ設定")
    cu_file = st.file_uploader("CUリスト", type=["xlsx"])
    du_file = st.file_uploader("DUリスト", type=["xlsx"])

st.header("2. 計画ファイル比較")
c1, c2 = st.columns(2)
with c1: old_file = st.file_uploader("旧計画", type=["xlsm", "xlsx"])
with c2: new_file = st.file_uploader("新計画", type=["xlsm", "xlsx"])

if st.button("🔍 レポート作成"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("ファイルをすべてアップロードしてください")
    else:
        try:
            cu_m = pd.read_excel(cu_file).astype(str)
            du_m = pd.read_excel(du_file).astype(str)
            # 数量計算用のカラムのみ数値変換
            for m in [cu_m, du_m]:
                for c in ["Parent Material Quantity", "Component Quantity"]:
                    if c in m.columns: m[c] = pd.to_numeric(m[c], errors='coerce').fillna(0)

            old_plans = get_plan_data(old_file)
            new_plans = get_plan_data(new_file)
            
            all_names = sorted(set(old_plans.keys()).union(set(new_plans.keys())))
            output = io.BytesIO()
            all_errors = set()

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                red_format = writer.book.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                
                for name in all_names:
                    old_bom, err_o = create_structured_bom(old_plans.get(name, pd.DataFrame()), cu_m, du_m)
                    new_bom, err_n = create_structured_bom(new_plans.get(name, pd.DataFrame()), cu_m, du_m)
                    all_errors.update(err_o)
                    all_errors.update(err_n)
                    
                    m_keys = ['Parent Mat', 'Start Date', 'Comp Number']
                    df = pd.merge(old_bom, new_bom, on=m_keys, how='outer', suffixes=('_旧', '_新'))
                    if df.empty: continue
                    df.fillna({'Need Qty_旧': 0, 'Need Qty_新': 0}, inplace=True)
                    
                    # ソート処理 [cite: 20]
                    df['P_S'] = df['Product Code_旧'].fillna(df['Product Code_新'])
                    df['L_S'] = df['Level_新'].fillna(df['Level_旧'])
                    df = df.sort_values(['P_S', 'Start Date', 'Parent Mat', 'L_S']).drop(columns=['P_S', 'L_S'])

                    sheet_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                    df.to_excel(writer, index=False, sheet_name=sheet_name)

                    # ハイライト適用 [cite: 19]
                    ws = writer.sheets[sheet_name]
                    cols = df.columns.tolist()
                    idx_o, idx_n = cols.index('Need Qty_旧'), cols.index('Need Qty_新')
                    for i, r in enumerate(df.itertuples(index=False)):
                        if abs(r[idx_o] - r[idx_n]) >= 1:
                            ws.set_row(i + 1, None, red_format)

            # エラー表示エリア
            if all_errors:
                st.warning("⚠️ 以下の品目はマスターデータ（DU/CU）で見つからなかったか、対象外でした：")
                st.code("\n".join(sorted(all_errors)))

            st.success("レポート作成が完了しました。")
            st.download_button(
                label="📥 ダウンロード: プラン変更_PO確認.xlsx [cite: 18]",
                data=output.getvalue(),
                file_name="プラン変更_PO確認.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"システムエラー: {e}")
