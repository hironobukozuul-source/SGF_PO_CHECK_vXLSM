import streamlit as st
import pandas as pd
import io

# --- 設定とヘッダー定義 ---
PLAN_MAT_COL = "品目コード"
PLAN_PROD_COL = "製品記号"
PLAN_START_COL = "製造開始日"
PLAN_QTY_COL = "Quantity"
MASTER_KEY = "Parent material number"

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

def calculate_bom(plan_df, cu_df, du_df):
    """
    データ型の不一致（object vs int64）を解消して結合するロジック
    """
    # 1. 計画ファイルの列チェック
    required_plan = [PLAN_MAT_COL, PLAN_PROD_COL, PLAN_QTY_COL, PLAN_START_COL]
    missing_plan = [c for c in required_plan if c not in plan_df.columns]
    if missing_plan:
        st.error(f"計画ファイルエラー: Row 6に以下の列が見つかりません: {', '.join(missing_plan)}")
        return pd.DataFrame()

    # --- 重要: 型の変換 (Error Fix) ---
    # 結合キーとなる列をすべて「文字列」に変換し、前後の空白を削除します
    plan_df[PLAN_PROD_COL] = plan_df[PLAN_PROD_COL].astype(str).str.strip()
    cu_df[MASTER_KEY] = cu_df[MASTER_KEY].astype(str).str.strip()
    du_df[MASTER_KEY] = du_df[MASTER_KEY].astype(str).str.strip()
    
    # 比較用のキーも文字列に統一（結合エラー防止）
    plan_df[PLAN_MAT_COL] = plan_df[PLAN_MAT_COL].astype(str).str.strip()
    # --------------------------------

    # CU結合
    plan_cu = plan_df.merge(cu_df, left_on=PLAN_PROD_COL, right_on=MASTER_KEY, how='left')
    plan_cu['Component Number'] = plan_cu.get('Component Number_CU', "N/A")
    plan_cu['Necessary Quantity'] = plan_cu[PLAN_QTY_COL] * plan_cu.get('CU_Ratio', 0)

    # DU結合
    plan_du = plan_df.merge(du_df, left_on=PLAN_PROD_COL, right_on=MASTER_KEY, how='left')
    plan_du['Component Number'] = plan_du.get('Component Number_DU', "N/A")
    plan_du['Necessary Quantity'] = plan_du[PLAN_QTY_COL] * plan_du.get('DU_Ratio', 0)

    # 自品目（Self）
    plan_self = plan_df.copy()
    plan_self['Component Number'] = plan_self[PLAN_MAT_COL]
    plan_self['Necessary Quantity'] = plan_self[PLAN_QTY_COL]

    return pd.concat([plan_cu, plan_du, plan_self], ignore_index=True)

# --- Streamlit UI ---

st.set_page_config(page_title="SAP PO Auditor Pro", layout="wide")
st.title("📊 SAP製造指示 比較ツール (型不一致修正版)")

# サイドバー
with st.sidebar:
    st.header("1. マスタデータ")
    cu_file = st.file_uploader("CUリスト", type=["xlsx"])
    du_file = st.file_uploader("DUリスト", type=["xlsx"])

# メインエリア
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("旧計画ファイル", type=["xlsm", "xlsx"])
with col2:
    new_file = st.file_uploader("新計画ファイル", type=["xlsm", "xlsx"])

if st.button("🔍 比較を実行"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("すべてのファイルをアップロードしてください。")
    else:
        try:
            with st.spinner("データを照合中..."):
                cu_master = pd.read_excel(cu_file)
                du_master = pd.read_excel(du_file)
                old_plans = get_plan_data(old_file)
                new_plans = get_plan_data(new_file)

                common_names = set(old_plans.keys()).intersection(set(new_plans.keys()))

                if not common_names:
                    st.error("A1セルに一致する計画名が見つかりませんでした。")
                else:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                        for name in common_names:
                            old_bom = calculate_bom(old_plans[name], cu_master, du_master)
                            new_bom = calculate_bom(new_plans[name], cu_master, du_master)

                            if old_bom.empty or new_bom.empty:
                                continue

                            # 比較用キーを文字列化（最終結合用）
                            for df in [old_bom, new_bom]:
                                df[PLAN_MAT_COL] = df[PLAN_MAT_COL].astype(str)
                                df['Component Number'] = df['Component Number'].astype(str)

                            merge_cols = [PLAN_MAT_COL, PLAN_START_COL, "Component Number"]
                            comparison = pd.merge(
                                old_bom, new_bom, on=merge_cols, how='outer', suffixes=('_旧', '_新')
                            )
                            
                            comparison['Necessary Quantity_旧'] = comparison['Necessary Quantity_旧'].fillna(0)
                            comparison['Necessary Quantity_新'] = comparison['Necessary Quantity_新'].fillna(0)

                            safe_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                            comparison.to_excel(writer, index=False, sheet_name=safe_name)
                            
                            worksheet = writer.sheets[safe_name]
                            for row_num, (old_q, new_q) in enumerate(zip(comparison['Necessary Quantity_旧'], comparison['Necessary Quantity_新'])):
                                if old_q != new_q:
                                    worksheet.set_row(row_num + 1, None, red_format)

                    st.success(f"照合完了！ {len(common_names)} 個の計画を処理しました。")
                    st.download_button(
                        label="📥 比較レポートをダウンロード",
                        data=output.getvalue(),
                        file_name="SAP_Audit_Comparison.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
