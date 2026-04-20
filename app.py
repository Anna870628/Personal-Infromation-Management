import streamlit as st
from supabase import create_client, Client
import pandas as pd

# 1. 網頁基本配置
st.set_page_config(page_title="車美仕個資盤點系統", page_icon="🛡️", layout="wide")

# 2. 安全防護：從 Secrets 讀取敏感資訊
# 提醒：請務必在 Streamlit Cloud 或本地 .streamlit/secrets.toml 設定這些值
try:
    SYSTEM_PASSWORD = st.secrets["auth"]["admin_password"]
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
except Exception:
    st.error("❌ 找不到 Secrets 設定。請確保已設定 admin_password, SUPABASE_URL 與 SUPABASE_KEY。")
    st.stop()

# 3. 初始化資料庫連線
@st.cache_resource
def init_connection() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_connection()

# 4. 登入驗證機制
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center;'>🛡️ 個資盤點系統管理後台</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        input_pwd = st.text_input("請輸入系統密碼 (CMX_PIM)", type="password")
        if st.button("登入系統"):
            if input_pwd == SYSTEM_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("密碼錯誤，請洽系統管理員。")
    st.stop()

# 5. 側邊欄：權限與導覽
st.sidebar.title("👤 使用者管理")
# 預設單位清單，可根據實務需求增加
units = ["業務企劃室", "科技創新發展室", "總管理員"]
user_unit = st.sidebar.selectbox("切換登入單位", units)

is_admin = (user_unit == "總管理員")
st.sidebar.info(f"當前權限：{user_unit}")

st.sidebar.divider()
menu = st.sidebar.radio(
    "📂 功能導覽", 
    ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑表", "4. 委外廠商清冊"]
)

# 6. 資料處理邏輯 (CRUD)
def load_data(table_name):
    query = supabase.table(table_name).select("*")
    if not is_admin:
        query = query.eq("unit_name", user_unit)
    response = query.execute()
    return pd.DataFrame(response.data)

def save_data(table_name, df):
    # 確保單位欄位正確
    if not is_admin:
        df["unit_name"] = user_unit
    
    records = df.to_dict(orient="records")
    if records:
        supabase.table(table_name).upsert(records).execute()
        st.success("✅ 資料存檔成功！")
    else:
        st.warning("無資料可儲存")

# 7. 各頁面介面實作
if menu == "1. 自檢表":
    st.title("📋 個資管理工作自檢表")
    st.caption("請確認各項業務的個資盤點進度。")
    
    df = load_data("self_checklist")
    
    # 根據附件定義欄位
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "unit_name": st.column_config.TextColumn("單位", disabled=not is_admin),
            "project_name": "業務名稱",
            "owner": "負責人",
            "status": st.column_config.SelectboxColumn("狀態", options=["進行中", "已結案"]),
            "has_inventory": st.column_config.SelectboxColumn("個資清冊", options=["v", "-"]),
            "has_vendor": st.column_config.SelectboxColumn("委外廠商個資管理", options=["有", "-"]),
            "vendor_name": "委外廠商名稱"
        }
    )
    
    if st.button("儲存變更"):
        save_data("self_checklist", edited_df)

elif menu == "2. 個資清冊":
    st.title("📁 個資與機敏
