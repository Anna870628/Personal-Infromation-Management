import streamlit as st
from supabase import create_client, Client
import pandas as pd
import io

# ==========================================
# 0. 網頁基本配置
# ==========================================
st.set_page_config(page_title="車美仕個資盤點系統", page_icon="🛡️", layout="wide")

# ==========================================
# 1. 定義共用的下拉選項與 Excel 匯出函式
# ==========================================
YN_OPTIONS = ["Y", "N"]
PI_AMOUNT_OPTIONS = ["每年產生大於1000筆", "每年產生100~1000筆", "每年產生小於100筆"]
PI_PURPOSE_OPTIONS = [
    "○○二 人事管理...", "○三一 社會保險...", "○四○ 行銷...",
    "○五二 內部管理...", "○六三 依法蒐集...", "○六九 契約事務...",
    "○七七 訂位購票...", "○九○ 客戶管理...", "一五七 調查分析..."
]
PI_CATEGORY_OPTIONS = ["Ｃ○○一 辨識個人", "Ｃ○○二 辨識財務", "Ｃ○一一 個人描述", "Ｃ一一一 健康紀錄", "其他(請參照法規)"]
LEGAL_BASIS_OPTIONS = ["A.法律明文規定", "B.履行法定義務", "C.已公開之個資", "D.法定職務必要", "E.書面同意"]
COLLECT_METHOD_OPTIONS = ["直接蒐集", "間接蒐集"]

def generate_excel(df, rename_dict, color_rules, sheet_name="Sheet1"):
    """將 DataFrame 轉為帶有顏色標頭的 Excel 檔案"""
    # 1. 篩選並重新命名欄位 (轉成中文)
    export_df = df.copy()
    export_df = export_df[[col for col in rename_dict.keys() if col in export_df.columns]]
    export_df = export_df.rename(columns=rename_dict)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # 2. 定義 Excel 標頭顏色樣式
        formats = {
            "default": workbook.add_format({'bg_color': '#D9E1F2', 'border': 1, 'bold': True}), # 淺藍
            "green": workbook.add_format({'bg_color': '#E2EFDA', 'border': 1, 'bold': True}),   # 淺綠
            "orange": workbook.add_format({'bg_color': '#FCE4D6', 'border': 1, 'bold': True}),  # 淺橘
            "yellow": workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'bold': True}),  # 淺黃
            "purple": workbook.add_format({'bg_color': '#E1DFED', 'border': 1, 'bold': True}),  # 淺紫
            "grey": workbook.add_format({'bg_color': '#D9D9D9', 'border': 1, 'bold': True}),    # 淺灰
            "red": workbook.add_format({'bg_color': '#F2DCDB', 'border': 1, 'bold': True}),     # 淺紅
        }
        
        # 3. 逐欄套用顏色
        for col_num, value in enumerate(export_df.columns.values):
            fmt_key = "default"
            for color, columns in color_rules.items():
                if value in columns:
                    fmt_key = color
                    break
            worksheet.write(0, col_num, value, formats[fmt_key])
            worksheet.set_column(col_num, col_num, 15) # 調整欄寬
            
    return output.getvalue()

# ==========================================
# 2. 安全防護 & 資料庫連線
# ==========================================
try:
    SYSTEM_PASSWORD = st.secrets["auth"]["admin_password"]
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
except Exception:
    st.error("❌ 找不到 Secrets 設定。")
    st.stop()

@st.cache_resource
def init_connection() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_connection()

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center;'>🛡️ 個資盤點系統</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        input_pwd = st.text_input("請輸入系統密碼", type="password")
        if st.button("登入系統"):
            if input_pwd == SYSTEM_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("密碼錯誤。")
    st.stop()

# ==========================================
# 5. 側邊欄與 CRUD 邏輯
# ==========================================
st.sidebar.title("👤 使用者管理")
user_unit = st.sidebar.selectbox("登入單位", ["業務企劃室", "科技創新發展室", "總管理員"])
is_admin = (user_unit == "總管理員")
menu = st.sidebar.radio("📂 功能導覽", ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑表", "4. 委外廠商清冊"])

def load_data(table_name):
    query = supabase.table(table_name).select("*")
    if not is_admin: query = query.eq("unit_name", user_unit)
    response = query.execute()
    return pd.DataFrame(response.data)

def save_data(table_name, df):
    if not is_admin and "unit_name" in df.columns:
        df["unit_name"] = df["unit_name"].replace(["", None], user_unit)
    df = df.where(pd.notnull(df), None)
    records = df.to_dict(orient="records")
    if records:
        supabase.table(table_name).upsert(records).execute()
        st.success("✅ 資料存檔成功！")

# ==========================================
# 7. 分頁實作 (加上短名稱與 Help 說明)
# ==========================================

if menu == "1. 自檢表":
    st.markdown("### 🛡️ 個資管理工作自檢表")
    st.markdown("🟦 `基本資訊` | 🟧 `委外管理` | 🟩 `結案專用` (滑鼠移至標題可查看完整說明)")
    
    df = load_data("self_checklist")
    expected_cols = ["id", "item_no", "project_name", "owner", "status", "pi_inventory_done", "vendor_mgmt_done", "vendor_name", "form_d001", "form_d002", "form_d003", "pi_destroyed", "unit_name"]
    for col in expected_cols:
        if col not in df.columns: df[col] = None

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "id": None,
            "item_no": st.column_config.TextColumn("項次", help="請輸入序號"),
            "unit_name": st.column_config.TextColumn("單位名稱"),
            "project_name": st.column_config.TextColumn("業務名稱", help="請填寫專案名稱"),
            "owner": st.column_config.TextColumn("負責人"),
            "status": st.column_config.SelectboxColumn("狀態", options=["進行中", "已結案"]),
            "pi_inventory_done": st.column_config.SelectboxColumn("清冊建檔", options=["v", "-"]),
            "vendor_mgmt_done": st.column_config.SelectboxColumn("委外管理", options=["有", "-"]),
            
            # 委外 (橘色)
            "vendor_name": st.column_config.TextColumn("🟧廠商名稱", help="【委外管理】請填寫委外廠商名稱"),
            "form_d001": st.column_config.SelectboxColumn("🟧D001", options=["v", "-"], help="【委外管理】D001 委外檔案清冊是否完成"),
            "form_d002": st.column_config.SelectboxColumn("🟧D002", options=["v", "-"], help="【委外管理】D002 存取單是否完成"),
            "form_d003": st.column_config.SelectboxColumn("🟧D003", options=["v", "-"], help="【委外管理】D003 銷毀單是否完成"),
            
            # 結案 (綠色)
            "pi_destroyed": st.column_config.SelectboxColumn("🟩個資銷毀", options=["v", "-"], help="【結案專用】專案結束時需確認個資是否銷毀")
        }
    )

    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("💾 儲存自檢表"):
            save_data("self_checklist", edited_df)
    with col2:
        # Excel 匯出設定
        rename_dict = {
            "item_no": "項次", "unit_name": "單位名稱", "project_name": "業務名稱",
            "owner": "負責人", "status": "狀態", "pi_inventory_done": "清冊建檔",
            "vendor_mgmt_done": "委外管理", "vendor_name": "委外廠商名稱",
            "form_d001": "D001清冊", "form_d002": "D002存取單",
            "form_d003": "D003銷毀單", "pi_destroyed": "個資已銷毀"
        }
        color_rules = {
            "orange": ["委外廠商名稱", "D001清冊", "D002存取單", "D003銷毀單"],
            "green": ["個資已銷毀"]
        }
        excel_data = generate_excel(edited_df, rename_dict, color_rules)
        st.download_button("📥 匯出 Excel 表", excel_data, f"自檢表_{user_unit}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "2. 個資清冊":
    st.markdown("### 📁 個資與機敏檔案清冊")
    st.markdown("🟦`基本資訊` 🟩`個資範圍` 🟧`使用` 🟨`傳送` 🟪`儲存` ⬜`刪除` 🟥`國際傳輸` (滑鼠移至標題看說明)")
    
    df = load_data("pi_inventory")
    pi_scope_cols = ["姓名", "出生日", "身分證", "護照", "聯絡方式", "財務", "車籍資料", "其他"]
    
    # 網頁端縮短名稱與 Info 設定
    col_cfg = {
        "id": None,
        "unit_name": st.column_config.TextColumn("🟦所屬單位"),
        "dept_name": st.column_config.TextColumn("🟦部名稱"),
        "process_desc": st.column_config.TextColumn("🟦流程說明"),
        "pi_purpose": st.column_config.SelectboxColumn("🟩特定目的", options=PI_PURPOSE_OPTIONS),
        "legal_basis": st.column_config.SelectboxColumn("🟩合法依據", options=LEGAL_BASIS_OPTIONS),
        
        "use_target": st.column_config.TextColumn("🟧對象", help="【生命週期-使用】單位內使用對象"),
        "use_purpose": st.column_config.TextColumn("🟧目的", help="【生命週期-使用】使用目的"),
        "use_method": st.column_config.TextColumn("🟧方式", help="【生命週期-使用】如列印、下載"),
        
        "trans_target": st.column_config.TextColumn("🟨對象", help="【生命週期-傳送】傳送對象"),
        "trans_purpose": st.column_config.TextColumn("🟨目的", help="【生命週期-傳送】傳送目的"),
        "trans_method": st.column_config.TextColumn("🟨方式", help="【生命週期-傳送】郵寄、系統傳輸等"),
        
        "store_loc": st.column_config.TextColumn("🟪位置", help="【生命週期-儲存】上鎖櫃、資料庫等"),
        "store_legal_time": st.column_config.TextColumn("🟪法定時限", help="【生命週期-儲存】法定保留時限"),
        
        "del_method": st.column_config.TextColumn("⬜方式", help="【生命週期-刪除】銷毀或刪除方式"),
        "intl_country": st.column_config.TextColumn("🟥國家", help="【生命週期-國際傳輸】傳送至哪個國家")
    }
    
    for col in pi_scope_cols:
        col_cfg[f"scope_{col}"] = st.column_config.SelectboxColumn(f"🟩{col}", options=YN_OPTIONS, help="【個資範圍】是否包含此資料")

    for c in col_cfg.keys():
        if c not in df.columns: df[c] = None

    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_config=col_cfg)
    
    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("💾 儲存個資清冊"): save_data("pi_inventory", edited_df)
    with col2:
        # Excel 匯出設定 (還原完整的 Excel 巢狀標題名稱)
        rename_dict = {
            "unit_name": "所屬單位", "dept_name": "I.部名稱", "process_desc": "I.流程說明",
            "pi_purpose": "II.特定目的", "legal_basis": "II.合法依據",
            "use_target": "[使用]對象", "use_purpose": "[使用]目的", "use_method": "[使用]方式",
            "trans_target": "[傳送]對象", "trans_purpose": "[傳送]目的", "trans_method": "[傳送]方式",
            "store_loc": "[儲存]位置", "store_legal_time": "[儲存]法定時限",
            "del_method": "[刪除]方式", "intl_country": "[國際]國家"
        }
        for col in pi_scope_cols: rename_dict[f"scope_{col}"] = f"[範圍]{col}"
        
        color_rules = {
            "blue": ["所屬單位", "I.部名稱", "I.流程說明"],
            "green": ["II.特定目的", "II.合法依據"] + [f"[範圍]{col}" for col in pi_scope_cols],
            "orange": ["[使用]對象", "[使用]目的", "[使用]方式"],
            "yellow": ["[傳送]對象", "[傳送]目的", "[傳送]方式"],
            "purple": ["[儲存]位置", "[儲存]法定時限"],
            "grey": ["[刪除]方式"],
            "red": ["[國際]國家"]
        }
        excel_data = generate_excel(edited_df, rename_dict, color_rules)
        st.download_button("📥 匯出 Excel 表", excel_data, f"個資清冊_{user_unit}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "3. 風險評鑑表":
    st.markdown("### ⚠️ 個人資料風險評鑑")
    
    df = load_data("risk_assessment")
    expected_columns = ["id", "unit_name", "project_name", "score_1", "score_2", "score_3", "score_4", "score_5"]
    for col in expected_columns:
        if col not in df.columns: df[col] = 1 if 'score' in col else None
            
    edited_df = st.data_editor(
        df, num_rows="dynamic", use_container_width=True,
        column_config={
            "id": None,
            "unit_name": st.column_config.TextColumn("單位", disabled=not is_admin),
            "project_name": "業務名稱",
            "score_1": st.column_config.NumberColumn("🟨(1)數量", min_value=1, max_value=5, help="1~5分：評估個資數量多寡"),
            "score_2": st.column_config.NumberColumn("🟨(2)敏感度", min_value=1, max_value=5, help="1~5分：個資敏感程度"),
            "score_3": st.column_config.NumberColumn("🟨(3)信譽損害", min_value=1, max_value=5, help="1~5分：若外洩對公司信譽影響"),
            "score_4": st.column_config.NumberColumn("🟨(4)隱私衝擊", min_value=1, max_value=5, help="1~5分：當事人隱私受損程度"),
            "score_5": st.column_config.NumberColumn("🟨(5)合作單位", min_value=1, max_value=5, help="1~5分：業務合作單位外洩風險"),
        }
    )
    
    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("💾 儲存評估"): save_data("risk_assessment", edited_df)
    with col2:
        rename_dict = {
            "unit_name": "單位", "project_name": "業務名稱",
            "score_1": "(1)數量", "score_2": "(2)敏感度", "score_3": "(3)信譽損害",
            "score_4": "(4)隱私衝擊", "score_5": "(5)合作單位"
        }
        color_rules = {"yellow": ["(1)數量", "(2)敏感度", "(3)信譽損害", "(4)隱私衝擊", "(5)合作單位"]}
        excel_data = generate_excel(edited_df, rename_dict, color_rules)
        st.download_button("📥 匯出 Excel 表", excel_data, f"風險評鑑_{user_unit}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "4. 委外廠商清冊":
    st.markdown("### 🤝 委外廠商個資檔案清冊")
    
    df = load_data("vendor_inventory")
    for col in ["id", "unit_name", "vendor_name", "file_name", "pi_scope", "trans_purpose", "trans_method"]:
        if col not in df.columns: df[col] = None
            
    edited_df = st.data_editor(
        df, num_rows="dynamic", use_container_width=True,
        column_config={
            "id": None,
            "unit_name": st.column_config.TextColumn("單位", disabled=not is_admin),
            "vendor_name": "廠商名稱",
            "file_name": "個資檔案名稱",
            "pi_scope": st.column_config.TextColumn("🟩個資範圍", help="【檔案資訊】請說明提供給廠商的個資範圍"),
            "trans_purpose": st.column_config.TextColumn("🟨傳送目的", help="【傳送資訊】為何需要傳送給該廠商"),
            "trans_method": st.column_config.TextColumn("🟨傳送方式", help="【傳送資訊】API、郵寄、加密檔案等")
        }
    )
    
    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("💾 儲存清冊"): save_data("vendor_inventory", edited_df)
    with col2:
        rename_dict = {
            "unit_name": "單位", "vendor_name": "廠商名稱", "file_name": "個資檔案名稱",
            "pi_scope": "個資範圍", "trans_purpose": "傳送目的", "trans_method": "傳送方式"
        }
        color_rules = {"green": ["個資範圍"], "yellow": ["傳送目的", "傳送方式"]}
        excel_data = generate_excel(edited_df, rename_dict, color_rules)
        st.download_button("📥 匯出 Excel 表", excel_data, f"委外清冊_{user_unit}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.sidebar.divider()
st.sidebar.caption("© 2026 Carmax Co., Ltd.")
