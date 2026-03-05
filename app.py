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

# マスタデータ（CU/DU）内の数量列名
MASTER_PARENT_QTY_COL = "Parent Material Quantity"
MASTER_COMP_QTY_COL = "Component Quantity"
MASTER_DESC_COL = "Component Description"

# --- ロジック関数 ---

def get_plan_data(uploaded_file):
    """ExcelファイルからRow 6以降のデータを取得"""
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

def compute_necessary_qty(row, list_suffix):
    """
    要件に基づいた数量計算ロジック:
    - BOTTLE/PUMP: 計画Quantityと同数
    - その他VERP: (計画Quantity / マスタ親数量) * マスタ子数量
    """
    plan_qty = row.get(PLAN_QTY_COL, 0)
    
    # マスタから取得した名称の列（suffixes付き）
    desc = str(row.get(f'{MASTER_DESC_COL}_{list_suffix}', '')).upper()
    
    # 1. BOTTLE または PUMP の判定
    if "BOTTLE" in desc or "PUMP" in desc:
        return plan_qty
    
    # 2. その他のVERP（Parent Material Quantityに基づく比例計算）
    p_qty = row.get(f'{MASTER_PARENT_QTY_COL}_{list_suffix}', 1)
    c_qty = row.get(f'{MASTER_COMP_QTY_COL}_{list_suffix}', 0)
    
    if p_qty == 0 or pd.isna(p_qty):
        return 0
        
    return (plan_qty / p_qty) * c_qty

def calculate_bom(plan_df, cu_df, du_df):
    """マスタと紐付け、数量計算を行い、リストを生成"""
    if PLAN_PROD_COL not in plan_df.columns:
        return pd.DataFrame()

    # 型の不一致を解消
    plan_df[PLAN_PROD_COL] = plan_df[PLAN_PROD_COL].astype(str).str.strip()
    cu_df[MASTER_KEY] = cu_df[MASTER_KEY].astype(str).str.strip()
    du_df[MASTER_KEY] = du_df[MASTER_KEY].astype(str).str.strip()
    plan_df[PLAN_MAT_COL] = plan_df[PLAN_MAT_COL].astype(str).str.strip()

    # VERPのみにフィルタリング
    if MATERIAL_TYPE_COL in cu_df.columns:
        cu_df = cu_df[cu_df[MATERIAL_TYPE_COL] == TARGET_TYPE]
    if MATERIAL_TYPE_COL in du_df.columns:
        du_df = du_df[du_df[MATERIAL_TYPE_COL] == TARGET_TYPE]

    # --- CU リスト処理 ---
    plan_cu = plan_df.merge(cu_df, left_on=PLAN_PROD_COL, right_on=MASTER_KEY, how='left', suffixes=('', '_CU'))
    plan_cu = plan_cu.dropna(subset=[f'{MASTER_KEY}_CU'])
    if not plan_cu.empty:
        plan_cu['Component Number'] = plan_cu.get('Component Number_CU', "N/A")
        plan_cu['Necessary Quantity'] = plan_cu.apply(lambda r: compute_necessary_qty(r, 'CU'), axis=1)

    # --- DU リスト処理 ---
    plan_du = plan_df.merge(du_df, left_on=PLAN_PROD_COL, right_on=MASTER_KEY, how='left', suffixes=('', '_DU'))
    plan_du = plan_du.dropna(subset=[f'{MASTER_KEY}_DU'])
    if not plan_du.empty:
        plan_du['Component Number'] = plan_du.get('Component Number_DU', "N/A")
        plan_du['Necessary Quantity'] = plan_du.apply(lambda r: compute_necessary_qty(r, 'DU'), axis=1)

    # --- 製品自体（Self） ---
    plan_self = plan_df.copy()
    plan_self['Component Number'] = plan_self[PLAN_MAT_COL]
    plan_self['Necessary Quantity'] = plan_self[PLAN_QTY_COL]

    # 全てを結合して一つのリストにする
    return pd.concat([plan_cu, plan_du, plan_self], ignore_index=True)

# --- Streamlit UI ---

st.set_page_config(page_title="SAP Audit Tool (Automated Qty)", layout="wide")
st.title("📊 SAP製造指示 数量自動計算・比較ツール")

st.markdown("""
### 🧮 自動適用される数量ルール:
1. **BOTTLE / PUMP**: 計画上の生産数量(`Quantity`)をそのまま必要数とします。
2. **その他 VERP**: マスタの構成比率(`Parent Material Qty`ベース)で計算します。
3. **製品自体**: 生産数量(`Quantity`)と同数をセットします。
""")

# サイドバー
with st.sidebar:
    st.header("1. マスタデータ準備")
    cu_file = st.file_uploader("CUリスト (xlsx)", type=["xlsx"])
    du_file = st.file_uploader("DUリスト (xlsx)", type=["xlsx"])

# メイン画面
col1, col2 = st.columns(2)
with col1:
    old_file = st.file_uploader("旧計画ファイル (.xlsm)", type=["xlsm", "xlsx"])
with col2:
    new_file = st.file_uploader("新計画ファイル (.xlsm)", type=["xlsm", "xlsx"])

if st.button("🔍 比較レポートを作成"):
    if not (cu_file and du_file and old_file and new_file):
        st.error("全てのファイルをアップロードしてください。")
    else:
        try:
            with st.spinner("マスタ結合と数量計算を実行中..."):
                cu_master = pd.read_excel(cu_file)
                du_master = pd.read_excel(du_file)
                old_plans = get_plan_data(old_file)
                new_plans = get_plan_data(new_file)

                common_names = set(old_plans.keys()).intersection(set(new_plans.keys()))

                if not common_names:
                    st.error("一致する計画名(A1セル)が見つかりませんでした。")
                else:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

                        for name in common_names:
                            old_bom = calculate_bom(old_plans[name], cu_master, du_master)
                            new_bom = calculate_bom(new_plans[name], cu_master, du_master)

                            if old_bom.empty and new_bom.empty:
                                continue

                            # 型の統一（結合用）
                            for df in [old_bom, new_bom]:
                                if not df.empty:
                                    df[PLAN_MAT_COL] = df[PLAN_MAT_COL].astype(str)
                                    df['Component Number'] = df['Component Number'].astype(str)

                            # 旧・新のデータをマージ
                            merge_cols = [PLAN_MAT_COL, PLAN_START_COL, "Component Number"]
                            comparison = pd.merge(
                                old_bom, new_bom, on=merge_cols, how='outer', suffixes=('_旧', '_新')
                            )
                            
                            # 数値の欠損埋め
                            comparison['Necessary Quantity_旧'] = comparison['Necessary Quantity_旧'].fillna(0)
                            comparison['Necessary Quantity_新'] = comparison['Necessary Quantity_新'].fillna(0)

                            # シート出力
                            safe_name = str(name)[:31].translate(str.maketrans("", "", r"[]:*?/\\"))
                            comparison.to_excel(writer, index=False, sheet_name=safe_name)
                            
                            # 差異行のハイライト
                            worksheet = writer.sheets[safe_name]
                            for row_num, (old_q, new_q) in enumerate(zip(comparison['Necessary Quantity_旧'], comparison['Necessary Quantity_新'])):
                                if abs(old_q - new_q) > 0.001: # 浮動小数点の誤差を考慮
                                    worksheet.set_row(row_num + 1, None, red_format)

                    st.success(f"完了: {len(common_names)} 件のシートを処理しました。")
                    st.download_button(
                        label="📥 自動計算済みレポートをダウンロード",
                        data=output.getvalue(),
                        file_name="SAP_Automated_Audit_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"システムエラー: {e}")
