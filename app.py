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
# 1. 下拉選項與 Excel 匯出函式
# ==========================================
YN_OPTIONS = ["Y", "N"]
PI_AMOUNT_OPTIONS = ["每年產生大於1000筆", "每年產生100~1000筆", "每年產生小於100筆"]
PI_PURPOSE_OPTIONS = [
    "○○二 人事管理（包含甄選、離職及所屬員工基本資訊、現職、學經歷、考試分發、終身學習訓練進修、考績獎懲、銓審、薪資待遇、差勤、福利措施、褫奪公權、特殊查核或其他人事措施）",
    "○三一 全民健康保險、勞工保險、農民保險、國民年金保險或其他社會保險",
    "○四○ 行銷（包含金控共同行銷業務）",
    "○五二 法人或團體對股東、會員（含股東、會員指派之代表）、董事、監察人、理事、監事或其他成員名冊之內部管理",
    "○六三 非公務機關依法定義務所進行個人資料之蒐集處理及利用",
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
    "A.法律明文規定", "B.履行法定義務所必要，且有適當安全維護措施",
    "C.當事人自行公開或其他已合法公開之個人資料",
    "D.協助公務機關執行法定職務或非公務機關履行法定義務所必要，且有適當安全維護措施",
    "E.經當事人書面同意"
]
COLLECT_METHOD_OPTIONS = ["直接蒐集", "間接蒐集"]

def generate_excel(df, rename_dict, color_rules, sheet_name="Sheet1"):
    export_df = df.copy()
    ordered_cols = [col for col in rename_dict.keys() if col in export_df.columns]
    export_df = export_df[ordered_cols]
    export_df = export_df.rename(columns=rename_dict)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        formats = {
            "default": workbook.add_format({'bg_color': '#F2F2F2', 'border': 1, 'bold': True}), 
            "blue": workbook.add_format({'bg_color': '#D9E1F2', 'border': 1, 'bold': True}),    
            "green": workbook.add_format({'bg_color': '#E2EFDA', 'border': 1, 'bold': True}),   
            "orange": workbook.add_format({'bg_color': '#FCE4D6', 'border': 1, 'bold': True}),  
            "yellow": workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'bold': True}),  
            "purple": workbook.add_format({'bg_color': '#E1DFED', 'border': 1, 'bold': True}),  
            "red": workbook.add_format({'bg_color': '#F2DCDB', 'border': 1, 'bold': True}),     
        }
        for col_num, value in enumerate(export_df.columns.values):
            fmt_key = "default"
            for color, columns in color_rules.items():
                if value in columns: fmt_key = color; break
            worksheet.write(0, col_num, value, formats.get(fmt_key, formats["default"]))
            worksheet.set_column(col_num, col_num, 20) 
    return output.getvalue()

# ==========================================
# 2. 資料庫連線與登入
# ==========================================
try:
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
    SYSTEM_PASSWORD = st.secrets["auth"]["admin_password"]
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
# 3. 組織架構讀取 (修復空資料表無欄位 BUG)
# ==========================================
def fetch_org_data():
    try:
        depts = supabase.table("departments").select("*").execute().data
        units = supabase.table("units").select("*").execute().data
        
        # 關鍵修復：即使沒有資料，也強制建立具有正確欄位名稱的空 DataFrame
        df_d = pd.DataFrame(depts) if depts else pd.DataFrame(columns=["id", "dept_name"])
        df_u = pd.DataFrame(units) if units else pd.DataFrame(columns=["id", "dept_name", "unit_name"])
        return df_d, df_u
    except Exception:
        return pd.DataFrame(columns=["id", "dept_name"]), pd.DataFrame(columns=["id", "dept_name", "unit_name"])

df_dept, df_unit = fetch_org_data()
dept_list = df_dept["dept_name"].dropna().unique().tolist() if not df_dept.empty else []
unit_list = df_unit["unit_name"].dropna().unique().tolist() if not df_unit.empty else []

# ==========================================
# 5. 側邊欄與 CRUD 邏輯
# ==========================================
st.sidebar.title("👤 使用者管理")
available_units = unit_list + ["總管理員"] if unit_list else ["總管理員"]
user_unit = st.sidebar.selectbox("登入單位", available_units)
is_admin = (user_unit == "總管理員")

menu_options = ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑表", "4. 委外廠商清冊"]
if is_admin:
    menu_options.append("5. 組織架構管理 (管理員)")
menu = st.sidebar.radio("📂 功能導覽", menu_options)

def load_data(table_name):
    query = supabase.table(table_name).select("*")
    if not is_admin: query = query.eq("unit_name", user_unit)
    response = query.execute()
    return pd.DataFrame(response.data)

def save_data(table_name, edited_df, original_df):
    """執行新增、修改、同步刪除"""
    success = False
    try:
        # 1. 處理刪除
        if "id" in original_df.columns and "id" in edited_df.columns:
            original_ids = set(original_df["id"].dropna().astype(str).tolist())
            edited_ids = set(edited_df["id"].dropna().astype(str).tolist())
            deleted_ids = list(original_ids - edited_ids)
            if deleted_ids:
                supabase.table(table_name).delete().in_("id", deleted_ids).execute()

        # 2. 自動掛載單位 (若非組織管理表)
        if not is_admin and table_name not in ["departments", "units"]:
            edited_df["unit_name"] = user_unit 
            
        upsert_df = edited_df.where(pd.notnull(edited_df), None)
        records = upsert_df.to_dict(orient="records")
        valid_records = [r for r in records if any(v is not None and str(v).strip() != "" for k, v in r.items() if k != 'id')]
        
        if valid_records:
            supabase.table(table_name).upsert(valid_records).execute()
            st.toast("✅ 資料同步成功！", icon="🎉")
        else:
            st.toast("✅ 已執行更新", icon="🗑️")
        success = True
    except Exception as e:
        st.error(f"❌ 存檔失敗：{e}")
    return success

# ==========================================
# 7. 各分頁實作
# ==========================================

if menu == "1. 自檢表":
    st.markdown("### 🛡️ 個資管理工作自檢表")
    df = load_data("self_checklist")
    chk_order = ["item_no", "unit_name", "project_name", "owner", "status", "pi_inventory_done", "vendor_mgmt_done", "vendor_name", "form_d001", "form_d002", "form_d003", "pi_destroyed"]
    for col in chk_order:
        if col not in df.columns: df[col] = None

    edited_df = st.data_editor(
        df, num_rows="dynamic", use_container_width=True, column_order=chk_order, 
        column_config={
            "id": None, "item_no": st.column_config.TextColumn("🟦項次"), "unit_name": st.column_config.TextColumn("🟦單位名稱", disabled=True),
            "project_name": st.column_config.TextColumn("🟦業務名稱"), "owner": st.column_config.TextColumn("🟦負責人"),
            "status": st.column_config.SelectboxColumn("🟦狀態", options=YN_OPTIONS), "pi_inventory_done": st.column_config.SelectboxColumn("🟦清冊建檔", options=YN_OPTIONS),
            "vendor_mgmt_done": st.column_config.SelectboxColumn("🟦委外管理", options=YN_OPTIONS), "vendor_name": st.column_config.TextColumn("🟧廠商名稱"),
            "form_d001": st.column_config.SelectboxColumn("🟧D001", options=YN_OPTIONS), "form_d002": st.column_config.SelectboxColumn("🟧D002", options=YN_OPTIONS),
            "form_d003": st.column_config.SelectboxColumn("🟧D003", options=YN_OPTIONS), "pi_destroyed": st.column_config.SelectboxColumn("🟩個資已銷毀", options=YN_OPTIONS)
        }
    )
    if st.button("💾 儲存自檢表"):
        if save_data("self_checklist", edited_df, df):
            time.sleep(1) 
            st.rerun()

elif menu == "2. 個資清冊":
    st.markdown("### 📁 個資與機敏檔案清冊")
    st.info("💡 提示：雙擊 (連點兩下) 儲存格即可編輯文字；點擊最左側行號並按 Delete 鍵即可刪除該列資料。")
    df = load_data("pi_inventory")
    pi_scope_cols = ["姓名", "出生年月日", "身分證號碼", "護照號碼", "特徵", "指紋", "婚姻", "家庭", "教育職業", "病歷", "醫療", "基因", "性生活", "健康檢查", "犯罪前科", "聯絡方式", "財務情況", "社會活動", "車籍資料", "其他"]
    pi_order = ["dept_name", "room_name", "pi_manager", "process_desc", "pi_amount", "legal_rule", "pi_purpose", "pi_category"]
    pi_order.extend([f"scope_{col}" for col in pi_scope_cols])               
    pi_order.extend(["legal_basis", "collect_method", "sys_name", "sys_source", "use_target", "use_purpose", "use_method", "use_protect", "trans_target", "trans_purpose", "trans_method", "trans_protect", "store_loc", "store_legal_time", "store_inner_time", "store_protect", "del_method", "del_unit", "intl_country", "intl_target", "intl_purpose", "intl_method", "intl_protect"])
    
    for col in pi_order:
        if col not in df.columns: df[col] = None

    col_cfg = {
        "id": None, "unit_name": None,
        "dept_name": st.column_config.SelectboxColumn("🟦部名稱", options=dept_list),
        "room_name": st.column_config.SelectboxColumn("🟦室名稱", options=unit_list),
        "pi_amount": st.column_config.SelectboxColumn("🟩筆數/份數", options=PI_AMOUNT_OPTIONS),
        "pi_purpose": st.column_config.SelectboxColumn("🟩特定目的", options=PI_PURPOSE_OPTIONS),
        "pi_category": st.column_config.SelectboxColumn("🟩個資之類別", options=PI_CATEGORY_OPTIONS),
        "legal_basis": st.column_config.SelectboxColumn("🟩合法依據", options=LEGAL_BASIS_OPTIONS),
        "collect_method": st.column_config.SelectboxColumn("🟩蒐集方式", options=COLLECT_METHOD_OPTIONS)
    }
    for col in pi_scope_cols: col_cfg[f"scope_{col}"] = st.column_config.SelectboxColumn(f"🟩{col}", options=YN_OPTIONS)

    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=pi_order, column_config=col_cfg)
    if st.button("💾 儲存個資清冊"):
        if save_data("pi_inventory", edited_df, df):
            time.sleep(1)
            st.rerun()

elif menu == "3. 風險評鑑表":
    st.markdown("### ⚠️ 個人資料風險評鑑")
    df = load_data("risk_assessment")
    st.data_editor(df, num_rows="dynamic", use_container_width=True) 
    if st.button("💾 儲存評估"):
        if save_data("risk_assessment", df, df):
            time.sleep(1)
            st.rerun()

elif menu == "4. 委外廠商清冊":
    st.markdown("### 🤝 委外廠商個資檔案清冊")
    df = load_data("vendor_inventory")
    st.data_editor(df, num_rows="dynamic", use_container_width=True) 
    if st.button("💾 儲存清冊"):
        if save_data("vendor_inventory", df, df):
            time.sleep(1)
            st.rerun()

elif menu == "5. 組織架構管理 (管理員)":
    st.markdown("### 🏢 組織架構管理")
    st.info("💡 提示：點擊下方 `+` 新增列後，對著顯示 `None` 的格子 **連點兩下 (雙擊)** 即可開始打字輸入！")
    
    col_dept, col_unit = st.columns(2)
    with col_dept:
        st.subheader("1. 部門管理")
        edited_dept = st.data_editor(
            df_dept, num_rows="dynamic", use_container_width=True, 
            column_config={
                "id": None, 
                "dept_name": st.column_config.TextColumn("🏢 部門名稱", required=True)
            }
        )
        if st.button("💾 儲存部門異動"):
            if save_data("departments", edited_dept, df_dept):
                time.sleep(1.5)
                st.rerun()

    with col_unit:
        st.subheader("2. 單位(室)管理")
        opts = dept_list if dept_list else ["(請先建立部門)"]
        edited_unit = st.data_editor(
            df_unit, num_rows="dynamic", use_container_width=True, 
            column_config={
                "id": None, 
                "dept_name": st.column_config.SelectboxColumn("所屬部門", options=opts, required=True), 
                "unit_name": st.column_config.TextColumn("🏠 單位名稱", required=True)
            }
        )
        if st.button("💾 儲存單位異動"):
            if save_data("units", edited_unit, df_unit):
                time.sleep(1.5)
                st.rerun()

st.sidebar.divider()
st.sidebar.caption("© 2026 Carmax Co., Ltd.")
