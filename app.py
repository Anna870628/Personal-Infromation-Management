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
    "A.法律明文規定",
    "B.履行法定義務所必要，且有適當安全維護措施",
    "C.當事人自行公開或其他已合法公開之個人資料",
    "D.協助公務機關執行法定職務或非公務機關履行法定義務所必要，且有適當安全維護措施",
    "E.經當事人書面同意"
]

COLLECT_METHOD_OPTIONS = ["直接蒐集", "間接蒐集"]

def generate_excel(df, rename_dict, color_rules, sheet_name="Sheet1"):
    """將 DataFrame 轉為帶有顏色標頭的 Excel 檔案"""
    export_df = df.copy()
    # 嚴格依照 rename_dict 的順序來排列 Excel 欄位
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
            "grey": workbook.add_format({'bg_color': '#D9D9D9', 'border': 1, 'bold': True}),    
            "red": workbook.add_format({'bg_color': '#F2DCDB', 'border': 1, 'bold': True}),     
        }
        
        for col_num, value in enumerate(export_df.columns.values):
            fmt_key = "default"
            for color, columns in color_rules.items():
                if value in columns:
                    fmt_key = color
                    break
            if fmt_key not in formats:
                fmt_key = "default"
                
            worksheet.write(0, col_num, value, formats[fmt_key])
            worksheet.set_column(col_num, col_num, 20) 
            
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
        try:
            supabase.table(table_name).upsert(records).execute()
            st.success("✅ 資料存檔成功！")
        except Exception as e:
            st.error(f"存檔失敗，請確認資料庫是否已新增對應欄位：{e}")

# ==========================================
# 7. 分頁實作
# ==========================================

if menu == "1. 自檢表":
    st.markdown("### 🛡️ 個資管理工作自檢表")
    st.markdown("🟦 `基本資訊` | 🟧 `委外管理` | 🟩 `結案專用`")
    
    df = load_data("self_checklist")
    
    # 嚴格定義顯示順序
    chk_order = [
        "item_no", "unit_name", "project_name", "owner", "status", 
        "pi_inventory_done", "vendor_mgmt_done", "vendor_name", 
        "form_d001", "form_d002", "form_d003", "pi_destroyed"
    ]
    for col in chk_order:
        if col not in df.columns: df[col] = None

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_order=chk_order, # 強制鎖定 UI 顯示順序
        column_config={
            "id": None,
            "item_no": st.column_config.TextColumn("項次", help="請輸入序號"),
            "unit_name": st.column_config.TextColumn("單位名稱"),
            "project_name": st.column_config.TextColumn("業務名稱", help="請填寫專案名稱"),
            "owner": st.column_config.TextColumn("業務負責人"),
            "status": st.column_config.SelectboxColumn("業務狀態", options=YN_OPTIONS), 
            "pi_inventory_done": st.column_config.SelectboxColumn("個資清冊建檔", options=YN_OPTIONS),
            "vendor_mgmt_done": st.column_config.SelectboxColumn("委外廠商個資管理", options=YN_OPTIONS),
            
            "vendor_name": st.column_config.TextColumn("🟧委外廠商名稱", help="【委外管理】請填寫委外廠商名稱"),
            "form_d001": st.column_config.SelectboxColumn("🟧D001清冊", options=YN_OPTIONS, help="【委外管理】D001 委外檔案清冊是否完成"),
            "form_d002": st.column_config.SelectboxColumn("🟧D002存取單", options=YN_OPTIONS, help="【委外管理】D002 存取單是否完成"),
            "form_d003": st.column_config.SelectboxColumn("🟧D003銷毀單", options=YN_OPTIONS, help="【委外管理】D003 銷毀單是否完成"),
            
            "pi_destroyed": st.column_config.SelectboxColumn("🟩個資已銷毀", options=YN_OPTIONS, help="專案已結束需填寫")
        }
    )

    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("💾 儲存自檢表"):
            save_data("self_checklist", edited_df)
    with col2:
        rename_dict = {
            "item_no": "項次", "unit_name": "單位名稱", "project_name": "業務名稱",
            "owner": "業務負責人", "status": "業務狀態", "pi_inventory_done": "個資清冊建檔",
            "vendor_mgmt_done": "委外廠商個資管理", "vendor_name": "委外廠商名稱",
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
    st.markdown("🟦`基本資訊` 🟩`個資範圍` 🟧`使用` 🟨`傳送` 🟪`儲存` ⬜`刪除` 🟥`國際傳輸`")
    
    df = load_data("pi_inventory")
    
    pi_scope_cols = ["姓名", "出生年月日", "身分證號碼", "護照號碼", "特徵", "指紋", "婚姻", "家庭", 
                     "教育職業", "病歷", "醫療", "基因", "性生活", "健康檢查", "犯罪前科", 
                     "聯絡方式", "財務情況", "社會活動", "車籍資料", "其他"]
    
    # 嚴格定義顯示順序 (完全對齊您的要求)
    pi_order = [
        "unit_name", "dept_name", "room_name", "pi_manager", "process_desc", # I
        "pi_amount", "legal_rule", "pi_purpose", "pi_category"               # II 前半
    ]
    pi_order.extend([f"scope_{col}" for col in pi_scope_cols])               # II 範圍
    pi_order.extend(["legal_basis", "collect_method"])                       # II 後半
    pi_order.extend([                                                        # III
        "sys_name", "sys_source",
        "use_target", "use_purpose", "use_method", "use_protect",
        "trans_target", "trans_purpose", "trans_method", "trans_protect",
        "store_loc", "store_legal_time", "store_inner_time", "store_protect",
        "del_method", "del_unit",
        "intl_country", "intl_target", "intl_purpose", "intl_method", "intl_protect"
    ])
    
    for col in pi_order:
        if col not in df.columns: df[col] = None

    col_cfg = {
        "id": None,
        "unit_name": st.column_config.TextColumn("🟦所屬單位"),
        "dept_name": st.column_config.TextColumn("🟦部名稱"),
        "room_name": st.column_config.TextColumn("🟦室名稱"),
        "pi_manager": st.column_config.TextColumn("🟦個資檔案管理者"),
        "process_desc": st.column_config.TextColumn("🟦業務流程說明"),
        
        "pi_amount": st.column_config.SelectboxColumn("🟩筆數/份數", options=PI_AMOUNT_OPTIONS),
        "legal_rule": st.column_config.TextColumn("🟩法源/內部規範依據"),
        "pi_purpose": st.column_config.SelectboxColumn("🟩特定目的", options=PI_PURPOSE_OPTIONS),
        "pi_category": st.column_config.SelectboxColumn("🟩個資之類別", options=PI_CATEGORY_OPTIONS),
        
        # 範圍在此插入...
        
        "legal_basis": st.column_config.SelectboxColumn("🟩合法蒐集依據", options=LEGAL_BASIS_OPTIONS),
        "collect_method": st.column_config.SelectboxColumn("🟩蒐集方式", options=COLLECT_METHOD_OPTIONS),
        
        "sys_name": st.column_config.TextColumn("🟦應用系統名稱"),
        "sys_source": st.column_config.TextColumn("🟦來源"),
        
        "use_target": st.column_config.TextColumn("🟧使用對象"),
        "use_purpose": st.column_config.TextColumn("🟧使用目的"),
        "use_method": st.column_config.TextColumn("🟧使用方式"),
        "use_protect": st.column_config.TextColumn("🟧保護方式"),
        
        "trans_target": st.column_config.TextColumn("🟨傳送對象"),
        "trans_purpose": st.column_config.TextColumn("🟨傳送目的"),
        "trans_method": st.column_config.TextColumn("🟨傳送方式"),
        "trans_protect": st.column_config.TextColumn("🟨保護方式"),
        
        "store_loc": st.column_config.TextColumn("🟪儲存位置"),
        "store_legal_time": st.column_config.TextColumn("🟪法定保留時限"),
        "store_inner_time": st.column_config.TextColumn("🟪內部保存期限"),
        "store_protect": st.column_config.TextColumn("🟪保護措施"),
        
        "del_method": st.column_config.TextColumn("⬜刪除/銷毀方式"),
        "del_unit": st.column_config.TextColumn("⬜刪除/銷毀單位"),
        
        "intl_country": st.column_config.TextColumn("🟥傳送國家"),
        "intl_target": st.column_config.TextColumn("🟥傳送對象"),
        "intl_purpose": st.column_config.TextColumn("🟥傳送目的"),
        "intl_method": st.column_config.TextColumn("🟥傳送方式"),
        "intl_protect": st.column_config.TextColumn("🟥保護方式")
    }
    
    for col in pi_scope_cols:
        col_cfg[f"scope_{col}"] = st.column_config.SelectboxColumn(f"🟩{col}", options=YN_OPTIONS, help="【個資範圍】是否包含此資料")

    edited_df = st.data_editor(
        df, 
        num_rows="dynamic", 
        use_container_width=True, 
        column_order=pi_order, # 強制鎖定 UI 顯示順序
        column_config=col_cfg
    )
    
    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("💾 儲存個資清冊"): save_data("pi_inventory", edited_df)
    with col2:
        # Excel 匯出字典 (已依照順序排列)
        rename_dict = {
            "unit_name": "所屬單位", "dept_name": "I.部名稱", "room_name": "I.室名稱", 
            "pi_manager": "I.個資檔案管理者", "process_desc": "I.業務流程說明",
            "pi_amount": "II.筆數/份數", "legal_rule": "II.法源/內部規範依據",
            "pi_purpose": "II.特定目的", "pi_category": "II.個資之類別"
        }
        for col in pi_scope_cols: rename_dict[f"scope_{col}"] = f"[範圍]{col}"
        
        rename_dict.update({
            "legal_basis": "II.合法蒐集依據", "collect_method": "II.蒐集方式",
            "sys_name": "III.應用系統名稱", "sys_source": "III.來源",
            "use_target": "[使用]對象", "use_purpose": "[使用]目的", 
            "use_method": "[使用]方式", "use_protect": "[使用]保護方式",
            "trans_target": "[傳送]對象", "trans_purpose": "[傳送]目的", 
            "trans_method": "[傳送]方式", "trans_protect": "[傳送]保護方式",
            "store_loc": "[儲存]位置", "store_legal_time": "[儲存]法定保留時限",
            "store_inner_time": "[儲存]內部保存期限", "store_protect": "[儲存]保護措施",
            "del_method": "[刪除]銷毀方式", "del_unit": "[刪除]銷毀單位",
            "intl_country": "[國際]國家", "intl_target": "[國際]對象",
            "intl_purpose": "[國際]目的", "intl_method": "[國際]方式", "intl_protect": "[國際]保護方式"
        })
        
        color_rules = {
            "blue": ["所屬單位", "I.部名稱", "I.室名稱", "I.個資檔案管理者", "I.業務流程說明", "III.應用系統名稱", "III.來源"],
            "green": ["II.筆數/份數", "II.法源/內部規範依據", "II.特定目的", "II.個資之類別", "II.合法蒐集依據", "II.蒐集方式"] + [f"[範圍]{col}" for col in pi_scope_cols],
            "orange": ["[使用]對象", "[使用]目的", "[使用]方式", "[使用]保護方式"],
            "yellow": ["[傳送]對象", "[傳送]目的", "[傳送]方式", "[傳送]保護方式"],
            "purple": ["[儲存]位置", "[儲存]法定保留時限", "[儲存]內部保存期限", "[儲存]保護措施"],
            "grey": ["[刪除]銷毀方式", "[刪除]銷毀單位"],
            "red": ["[國際]國家", "[國際]對象", "[國際]目的", "[國際]方式", "[國際]保護方式"]
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
        column_order=["unit_name", "project_name", "score_1", "score_2", "score_3", "score_4", "score_5"],
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
        column_order=["unit_name", "vendor_name", "file_name", "pi_scope", "trans_purpose", "trans_method"],
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
