import streamlit as st
from supabase import create_client, Client
import pandas as pd
import io
import time

# ==========================================
# 0. 網頁基本配置
# ==========================================
st.set_page_config(page_title="車美仕個資盤點系統", page_icon="🛡️", layout="wide")

# ==========================================
# 1. 定義共用選項
# ==========================================
YN_OPTIONS = ["Y", "N"]
PI_AMOUNT_OPTIONS = ["每年產生大於1000筆", "每年產生100~1000筆", "每年產生小於100筆"]
PI_PURPOSE_OPTIONS = [
    "○○二 人事管理（包含甄選、離職及所屬員工基本資訊...等）",
    "○三一 全民健康保險、勞工保險、農民保險...等",
    "○四○ 行銷（包含金控共同行銷業務）",
    "○五二 法人或團體對股東、會員...之內部管理",
    "○六三 非公務機關依法定義務...之蒐集",
    "○六九 契約、類似契約或其他法律關係事務",
    "○七七 訂位、住宿登記與購票業務",
    "○九○ 消費者、客戶管理與服務",
    "一五七 調查、統計與研究分析"
]
PI_CATEGORY_OPTIONS = [
    "Ｃ○○一 辨識個人者", "Ｃ○○二 辨識財務者", "Ｃ○○三 政府資料中之辨識者",
    "Ｃ○一一 個人描述", "Ｃ○二一 家庭情形", "Ｃ○三一 住家及設施",
    "Ｃ○三九 執照或其他許可", "Ｃ○五一 學校紀錄", "Ｃ○五二 資格或技術",
    "Ｃ○六一 現行之受僱情形", "Ｃ○六五 工作、差勤紀錄", "Ｃ○六六 健康與安全紀錄",
    "Ｃ○六八 薪資與預扣款", "Ｃ一一一 健康紀錄", "Ｃ一三一 書面文件之檢索",
    "Ｃ一三二 未分類之資料"
]
LEGAL_BASIS_OPTIONS = [
    "A.法律明文規定", "B.履行法定義務所必要...", "C.當事人自行公開...", "D.執行法定職務必要...", "E.經當事人書面同意"
]
COLLECT_METHOD_OPTIONS = ["直接蒐集", "間接蒐集"]

def generate_excel(df, rename_dict, color_rules):
    export_df = df.copy()
    ordered_cols = [col for col in rename_dict.keys() if col in export_df.columns]
    export_df = export_df[ordered_cols].rename(columns=rename_dict)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        formats = {
            "blue": workbook.add_format({'bg_color': '#D9E1F2', 'border': 1, 'bold': True}),
            "green": workbook.add_format({'bg_color': '#E2EFDA', 'border': 1, 'bold': True}),
            "orange": workbook.add_format({'bg_color': '#FCE4D6', 'border': 1, 'bold': True}),
            "yellow": workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'bold': True}),
            "purple": workbook.add_format({'bg_color': '#E1DFED', 'border': 1, 'bold': True}),
            "red": workbook.add_format({'bg_color': '#F2DCDB', 'border': 1, 'bold': True}),
        }
        for col_num, value in enumerate(export_df.columns.values):
            fmt = next((formats[c] for c, cols in color_rules.items() if value in cols), None)
            worksheet.write(0, col_num, value, fmt)
            worksheet.set_column(col_num, col_num, 18)
    return output.getvalue()

# ==========================================
# 2. 資料庫連線與登入
# ==========================================
try:
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
    SYSTEM_PASSWORD = st.secrets["auth"]["admin_password"]
except:
    st.error("❌ Secrets 設定錯誤。")
    st.stop()

@st.cache_resource
def init_connection():
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_connection()

if "auth" not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.title("🛡️ 車美仕個資盤點")
        pwd = st.text_input("系統密碼", type="password")
        if st.button("進入系統") and pwd == SYSTEM_PASSWORD:
            st.session_state.auth = True
            st.rerun()
    st.stop()

# ==========================================
# 3. 組織架構同步
# ==========================================
def fetch_org():
    try:
        d = supabase.table("departments").select("*").execute().data
        u = supabase.table("units").select("*").execute().data
        return pd.DataFrame(d or []), pd.DataFrame(u or [])
    except: return pd.DataFrame(columns=["dept_name"]), pd.DataFrame(columns=["dept_name", "unit_name"])

df_dept, df_unit = fetch_org()
dept_list = df_dept["dept_name"].tolist() if not df_dept.empty else []
unit_list = df_unit["unit_name"].tolist() if not df_unit.empty else []

# ==========================================
# 5. 側邊欄與資料處理邏輯
# ==========================================
st.sidebar.title("👤 使用者設定")
user_unit = st.sidebar.selectbox("當前單位", unit_list + ["總管理員"])
is_admin = (user_unit == "總管理員")

menu = st.sidebar.radio("📂 功能選單", ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑", "4. 委外廠商", "5. 組織管理"] if is_admin else ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑", "4. 委外廠商"])

def load_data(table):
    q = supabase.table(table).select("*")
    if not is_admin: q = q.eq("unit_name", user_unit)
    return pd.DataFrame(q.execute().data or [])

def save_data(table, edited_df, original_df):
    # 處理刪除
    if not original_df.empty and "id" in original_df.columns:
        deleted = list(set(original_df["id"].astype(str)) - set(edited_df["id"].dropna().astype(str)))
        if deleted: supabase.table(table).delete().in_("id", deleted).execute()

    # 背景補齊單位標籤與組織資訊
    if not is_admin:
        edited_df["unit_name"] = user_unit
        if table == "pi_inventory":
            # 自動帶入該單位所屬的部門名稱
            current_dept = df_unit[df_unit["unit_name"] == user_unit]["dept_name"].iloc[0] if not df_unit.empty else ""
            edited_df["dept_name"] = current_dept
            edited_df["room_name"] = user_unit

    records = edited_df.where(pd.notnull(edited_df), None).to_dict(orient="records")
    valid = [r for r in records if any(v and str(v).strip() != "" for k,v in r.items() if k != 'id')]
    if valid:
        supabase.table(table).upsert(valid).execute()
        st.toast("✅ 資料已同步至雲端", icon="☁️")
        return True
    return False

# ==========================================
# 7. 分頁實作
# ==========================================

if menu == "1. 自檢表":
    st.markdown("### 🛡️ 自檢表管理")
    df = load_data("self_checklist")
    cols = ["item_no", "project_name", "owner", "status", "pi_inventory_done", "vendor_mgmt_done", "vendor_name", "form_d001", "form_d002", "form_d003", "pi_destroyed"]
    for c in cols: 
        if c not in df.columns: df[c] = None
    
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=cols, column_config={
        "id": None, "item_no": "🟦項次", "project_name": "🟦業務名稱", "owner": "🟦負責人", "status": st.column_config.SelectboxColumn("🟦狀態", options=YN_OPTIONS),
        "pi_inventory_done": st.column_config.SelectboxColumn("🟦清冊", options=YN_OPTIONS), "vendor_mgmt_done": st.column_config.SelectboxColumn("🟦委外", options=YN_OPTIONS),
        "vendor_name": "🟧廠商名稱", "form_d001": st.column_config.SelectboxColumn("🟧D001", options=YN_OPTIONS),
        "form_d002": st.column_config.SelectboxColumn("🟧D002", options=YN_OPTIONS), "form_d003": st.column_config.SelectboxColumn("🟧D003", options=YN_OPTIONS),
        "pi_destroyed": st.column_config.SelectboxColumn("🟩銷毀", options=YN_OPTIONS)
    })
    if st.button("💾 儲存並重整"):
        if save_data("self_checklist", edited, df): time.sleep(0.5); st.rerun()

elif menu == "2. 個資清冊":
    st.markdown("### 📁 個資與機敏檔案清冊")
    st.caption(f"當前填報人：{user_unit} (已自動隱藏所屬單位欄位，存檔時會自動標記)")
    df = load_data("pi_inventory")
    
    # 修正順序：部/室名稱 -> 管理者 -> 流程 -> 筆數 -> 法源 -> 目的 -> 類別 -> 範圍...
    scopes = ["姓名", "出生年月日", "身分證號碼", "護照號碼", "特徵", "指紋", "婚姻", "家庭", "教育職業", "病歷", "醫療", "基因", "性生活", "健康檢查", "犯罪前科", "聯絡方式", "財務情況", "社會活動", "車籍資料", "其他"]
    order = ["dept_name", "room_name", "pi_manager", "process_desc", "pi_amount", "legal_rule", "pi_purpose", "pi_category"]
    order += [f"scope_{s}" for s in scopes]
    order += ["legal_basis", "collect_method", "sys_name", "sys_source", "use_target", "use_purpose", "use_method", "use_protect", "trans_target", "trans_purpose", "trans_method", "trans_protect", "store_loc", "store_legal_time", "store_inner_time", "store_protect", "del_method", "del_unit", "intl_country", "intl_target", "intl_purpose", "intl_method", "intl_protect"]
    
    for c in order:
        if c not in df.columns: df[c] = None

    # 非管理員自動帶入部名稱與室名稱
    if not is_admin and not df_unit.empty:
        curr_dept = df_unit[df_unit["unit_name"] == user_unit]["dept_name"].iloc[0]
        df["dept_name"] = df["dept_name"].fillna(curr_dept)
        df["room_name"] = df["room_name"].fillna(user_unit)

    cfg = {
        "id": None, "unit_name": None,
        "dept_name": st.column_config.SelectboxColumn("🟦部名稱", options=dept_list, disabled=not is_admin),
        "room_name": st.column_config.SelectboxColumn("🟦室名稱", options=unit_list, disabled=not is_admin),
        "pi_amount": st.column_config.SelectboxColumn("🟩筆數", options=PI_AMOUNT_OPTIONS),
        "pi_purpose": st.column_config.SelectboxColumn("🟩目的", options=PI_PURPOSE_OPTIONS),
        "pi_category": st.column_config.SelectboxColumn("🟩類別", options=PI_CATEGORY_OPTIONS),
        "legal_basis": st.column_config.SelectboxColumn("🟩依據", options=LEGAL_BASIS_OPTIONS),
        "collect_method": st.column_config.SelectboxColumn("🟩方式", options=COLLECT_METHOD_OPTIONS)
    }
    for s in scopes: cfg[f"scope_{s}"] = st.column_config.SelectboxColumn(f"🟩{s}", options=YN_OPTIONS)

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=order, column_config=cfg)
    if st.button("💾 儲存個資清冊"):
        if save_data("pi_inventory", edited, df): time.sleep(0.5); st.rerun()

elif menu == "5. 組織管理":
    st.markdown("### 🏢 部門與單位管理")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. 部門 CRUD")
        ed_d = st.data_editor(df_dept, num_rows="dynamic", use_container_width=True, column_config={"id":None, "dept_name":"🏢 部門名稱"})
        if st.button("💾 存部門"):
            if save_data("departments", ed_d, df_dept): time.sleep(1); st.rerun()
    with c2:
        st.subheader("2. 單位 CRUD")
        ed_u = st.data_editor(df_unit, num_rows="dynamic", use_container_width=True, column_config={"id":None, "dept_name":st.column_config.SelectboxColumn("所屬部門", options=dept_list), "unit_name":"🏠 單位名稱"})
        if st.button("💾 存單位"):
            if save_data("units", ed_u, df_unit): time.sleep(1); st.rerun()

st.sidebar.divider()
st.sidebar.caption("© 2026 Carmax Co., Ltd.")
