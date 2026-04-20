import streamlit as st
from supabase import create_client, Client
import pandas as pd

# ==========================================
# 0. 網頁基本配置
# ==========================================
st.set_page_config(page_title="車美仕個資盤點系統", page_icon="🛡️", layout="wide")

# ==========================================
# 1. 定義共用的下拉選單選項 (保持程式碼整潔)
# ==========================================
YN_OPTIONS = ["Y", "N"]

PI_AMOUNT_OPTIONS = [
    "每年產生大於1000筆", 
    "每年產生100~1000筆", 
    "每年產生小於100筆"
]

PI_PURPOSE_OPTIONS = [
    "○○二 人事管理（包含甄選、離職及所屬員工基本資訊...等）",
    "○三一 全民健康保險、勞工保險、農民保險、國民年金保險或其他社會保險",
    "○四○ 行銷（包含金控共同行銷業務）",
    "○五二 法人或團體對股東、會員（含股東、會員指派之代表）、董事、監察人...之內部管理",
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
    "A.法律明文規定",
    "B.履行法定義務所必要，且有適當安全維護措施",
    "C.當事人自行公開或其他已合法公開之個人資料",
    "D.協助公務機關執行法定職務或非公務機關履行法定義務所必要...",
    "E.經當事人書面同意"
]

COLLECT_METHOD_OPTIONS = ["直接蒐集", "間接蒐集"]

# ==========================================
# 2. 安全防護：從 Secrets 讀取敏感資訊
# ==========================================
try:
    SYSTEM_PASSWORD = st.secrets["auth"]["admin_password"]
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
except Exception:
    st.error("❌ 找不到 Secrets 設定。請確保 Streamlit Cloud Secrets 中已設定 [auth] 與 [supabase] 區塊。")
    st.stop()

# ==========================================
# 3. 初始化資料庫連線
# ==========================================
@st.cache_resource
def init_connection() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_connection()

# ==========================================
# 4. 登入驗證機制
# ==========================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h2 style='text-align: center;'>🛡️ 個資盤點系統管理後台</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        input_pwd = st.text_input("請輸入系統密碼", type="password")
        if st.button("登入系統"):
            if input_pwd == SYSTEM_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("密碼錯誤，請重新輸入。")
    st.stop()

# ==========================================
# 5. 側邊欄：權限與導覽
# ==========================================
st.sidebar.title("👤 使用者管理")
units = ["業務企劃室", "科技創新發展室", "總管理員"]
user_unit = st.sidebar.selectbox("切換登入單位", units)

is_admin = (user_unit == "總管理員")
st.sidebar.info(f"當前權限：{user_unit}")

st.sidebar.divider()
menu = st.sidebar.radio(
    "📂 功能導覽", 
    ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑表", "4. 委外廠商清冊"]
)

# ==========================================
# 6. 資料處理邏輯 (CRUD)
# ==========================================
def load_data(table_name):
    query = supabase.table(table_name).select("*")
    if not is_admin:
        query = query.eq("unit_name", user_unit)
    response = query.execute()
    return pd.DataFrame(response.data)

def save_data(table_name, df):
    # 確保不會讓非管理員覆寫到別人的單位
    if not is_admin and "unit_name" in df.columns:
        df["unit_name"] = df["unit_name"].replace(["", None], user_unit)
    
    # 將 pandas 的 NaN 或 pd.NA 轉換為 None 以符合 Supabase 寫入格式
    df = df.where(pd.notnull(df), None)
    records = df.to_dict(orient="records")
    
    if records:
        try:
            supabase.table(table_name).upsert(records).execute()
            st.success("✅ 資料存檔成功！")
        except Exception as e:
            st.error(f"存檔失敗，請確認資料庫欄位是否設定正確：{e}")
    else:
        st.warning("無資料可儲存")

# ==========================================
# 7. 各頁面介面實作
# ==========================================

if menu == "1. 自檢表":
    # --- 頁面標題與單位資訊 ---
    st.markdown("### 🛡️ 個資管理工作自檢表")
    
    header_col1, header_col2, header_col3 = st.columns(3)
    with header_col1:
        st.write(f"**單位：** {user_unit}")
    with header_col2:
        st.write("**日期：** 系統當前日期")
    with header_col3:
        st.write("**狀態：** 盤點執行中")
    
    st.divider()

    df = load_data("self_checklist")
    
    expected_columns = [
        "id", "item_no", "project_name", "owner", "status", "pi_inventory_done",
        "vendor_mgmt_done", "vendor_name", "form_d001", "form_d002", "form_d003", "pi_destroyed", "unit_name"
    ]
    for col in expected_columns:
        if col not in df.columns:
            df[col] = None

    st.info("💡 點擊表格下方「+」新增業務。存檔時若未填寫單位名稱，系統會自動帶入您的當前單位。")
    
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        key="self_check_editor",
        column_config={
            "id": st.column_config.Column("系統編號", disabled=True, hidden=True),
            "item_no": st.column_config.TextColumn("項次", help="請手動輸入序號 (例如 1, 2...)"),
            "unit_name": st.column_config.TextColumn("單位名稱"),
            "project_name": st.column_config.TextColumn("業務名稱", placeholder="請輸入業務或專案名稱"),
            "owner": st.column_config.TextColumn("業務負責人"),
            "status": st.column_config.SelectboxColumn("業務狀態", options=["進行中", "已結案"]),
            "pi_inventory_done": st.column_config.SelectboxColumn("個資清冊建檔", options=["v", "-"]),
            "vendor_mgmt_done": st.column_config.SelectboxColumn("委外廠商個資管理", options=["有", "-"]),
            
            # --- [巢狀結構模擬] 委外管理區段 ---
            "vendor_name": st.column_config.TextColumn("※委外管理 | 廠商名稱", help="若有委外廠商個資管理需填寫此欄位"),
            "form_d001": st.column_config.SelectboxColumn("※委外管理 | D001 清冊", options=["v", "-"]),
            "form_d002": st.column_config.SelectboxColumn("※委外管理 | D002 存取單", options=["v", "-"]),
            "form_d003": st.column_config.SelectboxColumn("※委外管理 | D003 銷毀單", options=["v", "-"]),
            
            # --- 結案區段 ---
            "pi_destroyed": st.column_config.SelectboxColumn("💡結案專用 | 個資已銷毀", options=["v", "-"], help="專案已結束需填寫")
        }
    )

    btn_col1, btn_col2 = st.columns([1, 6])
    with btn_col1:
        if st.button("💾 儲存盤點結果"):
            # 存檔前：如果使用者新增了列卻沒打單位，系統自動補上當前登入單位
            if "unit_name" in edited_df.columns:
                edited_df["unit_name"] = edited_df["unit_name"].replace(["", None], user_unit)
            save_data("self_checklist", edited_df)
    with btn_col2:
        # 下載功能
        csv = edited_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 匯出 CSV 檔",
            data=csv,
            file_name=f"個資自檢表_{user_unit}.csv",
            mime="text/csv",
        )

elif menu == "2. 個資清冊":
    st.title("📁 個資與機敏檔案清冊")
    st.caption("包含單位流程、個資資訊與生命週期管理。 (表格較寬，請左右滑動填寫)")
    
    df = load_data("pi_inventory")
    
    pi_scope_cols = ["姓名", "出生年月日", "身分證號碼", "護照號碼", "特徵", "指紋", "婚姻", "家庭", 
                     "教育職業", "病歷", "醫療", "基因", "性生活", "健康檢查", "犯罪前科", 
                     "聯絡方式", "財務情況", "社會活動", "車籍資料", "其他"]
    
    column_config_dict = {
        "id": st.column_config.Column("編號", disabled=True, hidden=True),
        "unit_name": st.column_config.TextColumn("所屬單位", disabled=not is_admin),
        # I. 單位及業務流程資訊
        "dept_name": "I.部名稱",
        "room_name": "I.室名稱",
        "pi_manager": "I.個資檔案管理者",
        "process_desc": "I.業務流程說明",
        
        # II. 個人資料資訊
        "pi_amount": st.column_config.SelectboxColumn("II.筆數/份數", options=PI_AMOUNT_OPTIONS),
        "legal_rule": "II.法源/內部規範依據",
        "pi_purpose": st.column_config.SelectboxColumn("II.特定目的", options=PI_PURPOSE_OPTIONS),
        "pi_category": st.column_config.SelectboxColumn("II.個資之類別", options=PI_CATEGORY_OPTIONS),
        "legal_basis": st.column_config.SelectboxColumn("II.合法蒐集依據", options=LEGAL_BASIS_OPTIONS),
        "collect_method": st.column_config.SelectboxColumn("II.蒐集方式", options=COLLECT_METHOD_OPTIONS),
        
        # III. 個人資料生命週期 
        "sys_name": "III.應用系統名稱",
        "sys_source": "III.來源",
        
        "use_target": "[使用] 對象",
        "use_purpose": "[使用] 目的",
        "use_method": "[使用] 方式",
        "use_protect": "[使用] 保護方式",
        
        "trans_target": "[傳送] 對象",
        "trans_purpose": "[傳送] 目的",
        "trans_method": "[傳送] 方式",
        "trans_protect": "[傳送] 保護方式",
        
        "store_loc": "[儲存] 位置",
        "store_legal_time": "[儲存] 法定時限",
        "store_inner_time": "[儲存] 內部時限",
        "store_protect": "[儲存] 保護措施",
        
        "del_method": "[刪除] 方式",
        "del_unit": "[刪除] 單位",
        
        "intl_country": "[國際傳輸] 國家",
        "intl_target": "[國際傳輸] 對象",
        "intl_purpose": "[國際傳輸] 目的",
        "intl_method": "[國際傳輸] 方式",
        "intl_protect": "[國際傳輸] 保護方式",
    }

    # 批次加入個資範圍 Y/N 下拉選單
    for col in pi_scope_cols:
        db_col_name = f"scope_{col}" 
        column_config_dict[db_col_name] = st.column_config.SelectboxColumn(f"II.[範圍]{col}", options=YN_OPTIONS)

    # 確保 DataFrame 包含所有設定的欄位
    for col in column_config_dict.keys():
        if col not in df.columns:
            df[col] = None

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config=column_config_dict
    )
    
    if st.button("💾 確認儲存個資清冊"):
        save_data("pi_inventory", edited_df)

elif menu == "3. 風險評鑑表":
    st.title("⚠️ 個人資料風險評鑑")
    st.caption("針對各專案進行外洩風險分數評估。")
    
    df = load_data("risk_assessment")
    
    expected_columns = ["id", "unit_name", "project_name", "score_1", "score_2", "score_3", "score_4", "score_5"]
    for col in expected_columns:
        if col not in df.columns:
            df[col] = 1 if 'score' in col else None
            
    st.markdown("""
    **評分標準：** 1(低) ~ 5(高)  
    *(1)個資數量、(2)敏感度、(3)信譽損害、(4)隱私衝擊、(5)業務合作單位*
    """)
    
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "id": st.column_config.Column("系統編號", disabled=True, hidden=True),
            "unit_name": st.column_config.TextColumn("單位", disabled=not is_admin),
            "project_name": "業務子流程名稱",
            "score_1": st.column_config.NumberColumn("(1)數量", min_value=1, max_value=5, step=1),
            "score_2": st.column_config.NumberColumn("(2)敏感度", min_value=1, max_value=5, step=1),
            "score_3": st.column_config.NumberColumn("(3)信譽損害", min_value=1, max_value=5, step=1),
            "score_4": st.column_config.NumberColumn("(4)隱私衝擊", min_value=1, max_value=5, step=1),
            "score_5": st.column_config.NumberColumn("(5)業務合作單位", min_value=1, max_value=5, step=1),
        }
    )
    
    if not edited_df.empty and 'score_1' in edited_df.columns:
        score_cols = ['score_1', 'score_2', 'score_3', 'score_4', 'score_5']
        # 自動計算分數
        edited_df['total_score'] = edited_df[score_cols].sum(axis=1)
        edited_df['risk_level'] = edited_df['total_score'].apply(
            lambda x: '高' if pd.notnull(x) and x >= 18 else ('中' if pd.notnull(x) and x >= 10 else '低')
        )
        st.dataframe(edited_df[["project_name", "total_score", "risk_level"]], use_container_width=True)

    if st.button("💾 儲存風險評估結果"):
        save_data("risk_assessment", edited_df)

elif menu == "4. 委外廠商清冊":
    st.title("🤝 委外廠商個資檔案清冊")
    st.caption("管理委外廠商可接觸到的個資範圍。")
    
    df = load_data("vendor_inventory")
    
    expected_columns = ["id", "unit_name", "vendor_name", "file_name", "pi_scope", "trans_purpose", "trans_method"]
    for col in expected_columns:
        if col not in df.columns:
            df[col] = None
            
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "id": st.column_config.Column("系統編號", disabled=True, hidden=True),
            "unit_name": st.column_config.TextColumn("單位", disabled=not is_admin),
            "vendor_name": "廠商名稱",
            "file_name": "個資檔案名稱",
            "pi_scope": "個資範圍",
            "trans_purpose": "傳送目的",
            "trans_method": "傳送方式"
        }
    )
    
    if st.button("💾 儲存廠商資料"):
        save_data("vendor_inventory", edited_df)

# ==========================================
# 8. 頁尾資訊
# ==========================================
st.sidebar.divider()
st.sidebar.caption("© 2026 Carmax Co., Ltd. 版權所有")
