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
# 1. 定義共用選項與色彩映射邏輯
# ==========================================
YN_OPTIONS = ["Y", "N"]
PI_AMOUNT_OPTIONS = [
    "每年產生大於100萬筆", 
    "每年產生10萬~100萬筆", 
    "每年產生1萬~10萬筆", 
    "每年產生1000~1萬筆", 
    "每年產生小於1000筆"
]

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
            fmt_key = "default"
            for color, columns in color_rules.items():
                if value in columns:
                    fmt_key = color
                    break
            if fmt_key in formats:
                worksheet.write(0, col_num, value, formats[fmt_key])
            worksheet.set_column(col_num, col_num, 20)
            
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
# 3. 組織資料讀取
# ==========================================
def fetch_org():
    try:
        d = supabase.table("departments").select("*").execute().data
        u = supabase.table("units").select("*").execute().data
        
        df_d = pd.DataFrame(d) if d else pd.DataFrame(columns=["id", "dept_name"])
        df_u = pd.DataFrame(u) if u else pd.DataFrame(columns=["id", "dept_name", "unit_name"])
        
        df_d["dept_name"] = df_d["dept_name"].astype("string")
        df_u["dept_name"] = df_u["dept_name"].astype("string")
        df_u["unit_name"] = df_u["unit_name"].astype("string")
        
        return df_d, df_u
    except: 
        return pd.DataFrame(columns=["id", "dept_name"]), pd.DataFrame(columns=["id", "dept_name", "unit_name"])

df_dept, df_unit = fetch_org()
dept_list = df_dept["dept_name"].dropna().unique().tolist() if not df_dept.empty else []
unit_list = df_unit["unit_name"].dropna().unique().tolist() if not df_unit.empty else []

# ==========================================
# 5. 側邊欄與權限隔離邏輯
# ==========================================
st.sidebar.title("👤 登入設定")
user_unit = st.sidebar.selectbox("切換單位", unit_list + ["總管理員"])
is_admin = (user_unit == "總管理員")

menu = st.sidebar.radio("📂 功能選單", ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑", "4. 委外廠商", "5. 組織管理"] if is_admin else ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑", "4. 委外廠商"])

def load_data(table):
    query = supabase.table(table).select("*")
    if not is_admin: query = query.eq("unit_name", user_unit)
    res = query.execute().data
    return pd.DataFrame(res or [])

def save_data(table, edited_df, original_df):
    if not original_df.empty and "id" in original_df.columns:
        orig_ids = set(original_df["id"].dropna().astype(str).tolist())
        edit_ids = set(edited_df["id"].dropna().astype(str).tolist()) if "id" in edited_df.columns else set()
        deleted = list(orig_ids - edit_ids)
        if deleted: 
            supabase.table(table).delete().in_("id", deleted).execute()

    if not is_admin and table not in ["departments", "units"]:
        edited_df["unit_name"] = user_unit

    records = edited_df.where(pd.notnull(edited_df), None).to_dict(orient="records")
    
    valid = []
    for r in records:
        meaningful_keys = [k for k in r.keys() if k not in ['id', 'unit_name']]
        if any(r[k] is not None and str(r[k]).strip() != "" for k in meaningful_keys):
            if pd.isna(r.get('id')): r.pop('id', None)
            valid.append(r)
            
    if valid:
        try:
            supabase.table(table).upsert(valid).execute()
            st.toast("✅ 資料已同步儲存", icon="☁️")
            return True
        except Exception as e:
            st.error(f"存檔失敗：{e}")
    else:
        st.toast("✅ 變更已套用 (含刪除)", icon="🗑️")
        return True
    return False

# ==========================================
# 7. 各分頁實作
# ==========================================

if menu == "1. 自檢表":
    st.markdown("### 🛡️ 自檢表管理")
    if is_admin: st.info("👁️ 目前身分：【總管理員】，可看見全公司資料。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
    else: st.info(f"🔒 目前身分：【{user_unit}】，僅顯示本單位資料。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
        
    df = load_data("self_checklist")
    
    cols = ["item_no", "unit_name", "project_name", "owner", "status", "pi_inventory_done", "vendor_mgmt_done", "vendor_name", "form_d001", "form_d002", "form_d003", "pi_destroyed"]
    for c in cols: 
        if c not in df.columns: df[c] = None

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=cols, column_config={
        "id": None,
        "item_no": "🟦項次", "unit_name": st.column_config.TextColumn("🟦單位", disabled=not is_admin),
        "project_name": "🟦業務名稱", "owner": "🟦負責人", "status": st.column_config.SelectboxColumn("🟦狀態", options=YN_OPTIONS),
        "pi_inventory_done": st.column_config.SelectboxColumn("🟦清冊建檔", options=YN_OPTIONS),
        "vendor_mgmt_done": st.column_config.SelectboxColumn("🟦委外管理", options=YN_OPTIONS),
        "vendor_name": "🟧廠商名稱", "form_d001": st.column_config.SelectboxColumn("🟧D001", options=YN_OPTIONS),
        "form_d002": st.column_config.SelectboxColumn("🟧D002", options=YN_OPTIONS), "form_d003": st.column_config.SelectboxColumn("🟧D003", options=YN_OPTIONS),
        "pi_destroyed": st.column_config.SelectboxColumn("🟩個資已銷毀", options=YN_OPTIONS)
    })
    
    c1, c2 = st.columns([1, 6])
    with c1:
        if st.button("💾 儲存"): 
            if save_data("self_checklist", edited, df): time.sleep(0.5); st.rerun()
    with c2:
        rename_dict = {"item_no":"項次","unit_name":"單位","project_name":"業務名稱","owner":"負責人","status":"狀態","pi_inventory_done":"清冊建檔","vendor_mgmt_done":"委外管理","vendor_name":"委外廠商名稱","form_d001":"D001清冊","form_d002":"D002存取單","form_d003":"D003銷毀單","pi_destroyed":"個資已銷毀"}
        rules = {"blue":["項次","單位","業務名稱","負責人","狀態","清冊建檔","委外管理"],"orange":["委外廠商名稱","D001清冊","D002存取單","D003銷毀單"],"green":["個資已銷毀"]}
        xl = generate_excel(edited, rename_dict, rules)
        st.download_button("📥 匯出 Excel", xl, "自檢表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "2. 個資清冊":
    st.markdown("### 📁 個資與機敏檔案清冊")
    if is_admin: st.info("👁️ 目前身分：【總管理員】，可看見全公司資料。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
    else: st.info(f"🔒 目前身分：【{user_unit}】，僅顯示本單位資料。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
        
    scopes = ["姓名", "出生年月日", "身分證號碼", "護照號碼", "特徵", "指紋", "婚姻", "家庭", "教育職業", "病歷", "醫療", "基因", "性生活", "健康檢查", "犯罪前科", "聯絡方式", "財務情況", "社會活動", "車籍資料", "其他"]
    order = ["dept_name", "room_name", "pi_manager", "process_desc", "pi_amount", "legal_rule", "pi_purpose", "pi_category"]
    order += [f"scope_{s}" for s in scopes]
    order += ["legal_basis", "collect_method", "sys_name", "sys_source", "use_target", "use_purpose", "use_method", "use_protect", "trans_target", "trans_purpose", "trans_method", "trans_protect", "store_loc", "store_legal_time", "store_inner_time", "store_protect", "del_method", "del_unit", "intl_country", "intl_target", "intl_purpose", "intl_method", "intl_protect"]
    
    # ------------------------------------------
    # 🌟 黃底說明列：自動換行 & 模擬合併儲存格
    # ------------------------------------------
    st.markdown("##### 💡 填寫範例與說明參考 (同 Excel 附件)")
    
    # 模擬 20 個個資欄位的「合併儲存格」說明
    st.info("📌 **【個人資料範圍 (姓名 ~ 其他) 填寫說明】**：\n請依個人資料保護法施行細則第4條及第5條之規定，就所蒐集之個人資料，於該適當欄位填列 Y ，若無則填列 N ，但其他可直接或間接方式識別個人之資料(請於「其他」欄位直接列舉)。")
    
    example_dict = {
        "dept_name": "請填列部門名稱", 
        "room_name": "請填列室名稱", 
        "pi_manager": "請填列個資檔案管理者人員名稱", 
        "process_desc": "請填列業務子流程名稱",
        "pi_amount": "請選擇約略數量", 
        "legal_rule": "外部法規依據/內部規範依據/NA", 
        "pi_purpose": "請下拉選擇", 
        "pi_category": "請下拉選擇",
        "legal_basis": "列示合法蒐集個資之依據，如：個資授權同意書、隱私權條款\n(僅資料蒐集單位須填寫)", 
        "collect_method": "屬於直接蒐集或間接蒐集\n(僅資料蒐集單位須填寫)",
        "sys_name": "該筆個人資料涉及的系統或檔案名稱", 
        "sys_source": "請填寫個人資料(包括紙本文件跟電子檔案)的來源對象，不限公司內外單位；若個人資料來自於資訊系統，則填寫資訊系統名稱",
        "use_target": "資料單位內使用者\n(如無請填列N/A)", 
        "use_purpose": "使用目的如：資料建檔、廣告投放等\n(如無請填列N/A)",
        "use_method": "如使用者及目的之欄位有填列，請說明使用資料的方式，如列印、下載 。\n(如無請填列N/A)", 
        "use_protect": "如有填寫使用方式，應一併說明保護方式，如: 權限控管、刷卡等\n(如無請填列N/A)",
        "trans_target": "資料傳送之對象(如:XXX委外廠商、XXX主管機關或XXX內部單位等)\n(如無請填列N/A)", 
        "trans_purpose": "傳送目的\n(如無請填列N/A)",
        "trans_method": "如傳送對象及目的之欄位有填列，請說明傳輸資料的方式，如親自提供 / 郵寄 / 掛號 / 快遞 / 傳真 / 檔案傳遞 / 對外或對內系統(入口網站、FTP、其他公司系統等) 。\n(如無請填列N/A)", 
        "trans_protect": "如有填寫傳送方式，應一併說明保護方式，如: 專人親送／親取／加密等\n(如無請填列N/A)",
        "store_loc": "如:實體櫃/雲端資料庫", 
        "store_legal_time": "法定保存年限",
        "store_inner_time": "公司內部規定保存年限", 
        "store_protect": "上鎖、密碼控管等",
        "del_method": "碎紙機銷毀、系統刪除等", 
        "del_unit": "負責執行銷毀之單位",
        "intl_country": "傳送到其他國家\n(如無請填列N/A)", 
        "intl_target": "傳送對象\n(如無請填列N/A)",
        "intl_purpose": "傳送目的\n(如無請填列N/A)", 
        "intl_method": "如傳送國家及目的之欄位有填列，請說明傳輸資料的方式，例如：檔案傳輸系統、應用程式與應用程式之間傳輸等。\n(如無請填列N/A)", 
        "intl_protect": "保護方式\n(如無請填列N/A)"
    }
    
    for s in scopes: 
        example_dict[f"scope_{s}"] = "填 Y 或 N"
    example_dict["scope_其他"] = "請直接列舉"
    
    rename_mapping = {
        "dept_name": "🟦部名稱", "room_name": "🟦室名稱", "pi_manager": "🟦個資檔案管理者", "process_desc": "🟦業務流程說明",
        "pi_amount": "🟩筆數/份數", "legal_rule": "🟩法源/內部規範依據", "pi_purpose": "🟩特定目的", "pi_category": "🟩個資之類別",
        "legal_basis": "🟩合法蒐集依據", "collect_method": "🟩蒐集方式",
        "sys_name": "🟧應用系統名稱", "sys_source": "🟧來源", 
        "use_target": "🟧使用對象", "use_purpose": "🟧使用目的", "use_method": "🟧使用方式", "use_protect": "🟧使用保護方式",
        "trans_target": "🟧傳送對象", "trans_purpose": "🟧傳送目的", "trans_method": "🟧傳送方式", "trans_protect": "🟧傳送保護方式",
        "store_loc": "🟪儲存位置", "store_legal_time": "🟪法定時限", "store_inner_time": "🟪內部保存期限", "store_protect": "🟪儲存保護措施",
        "del_method": "🟪刪除方式", "del_unit": "🟪刪除單位",
        "intl_country": "🟥傳送國家", "intl_target": "🟥國際傳送對象", "intl_purpose": "🟥國際傳送目的", "intl_method": "🟥國際傳送方式", "intl_protect": "🟥國際保護方式"
    }
    for s in scopes: rename_mapping[f"scope_{s}"] = f"🟩{s}"
    
    ex_df = pd.DataFrame([example_dict])[ [c for c in order if c in rename_mapping] ].rename(columns=rename_mapping)
    
    # 加入 CSS 自動換行屬性 (white-space: pre-wrap)
    styled_ex_df = ex_df.style.set_properties(**{
        'background-color': '#FFF2CC', 
        'color': '#000000',
        'white-space': 'pre-wrap'
    })
    st.dataframe(styled_ex_df, hide_index=True)
    # ------------------------------------------

    df = load_data("pi_inventory")
    for c in order:
        if c not in df.columns: df[c] = None

    cfg = {
        "id": None, "unit_name": None, 
        "dept_name": st.column_config.SelectboxColumn("🟦部名稱", options=dept_list),
        "room_name": st.column_config.SelectboxColumn("🟦室名稱", options=unit_list),
        "pi_manager": "🟦個資檔案管理者", "process_desc": "🟦業務流程說明",
        "pi_amount": st.column_config.SelectboxColumn("🟩筆數/份數", options=PI_AMOUNT_OPTIONS),
        "legal_rule": "🟩法源/內部規範依據",
        "pi_purpose": st.column_config.SelectboxColumn("🟩特定目的", options=PI_PURPOSE_OPTIONS),
        "pi_category": st.column_config.SelectboxColumn("🟩個資之類別", options=PI_CATEGORY_OPTIONS),
        "legal_basis": st.column_config.SelectboxColumn("🟩合法蒐集依據", options=LEGAL_BASIS_OPTIONS),
        "collect_method": st.column_config.SelectboxColumn("🟩蒐集方式", options=COLLECT_METHOD_OPTIONS),
        "sys_name": "🟧應用系統名稱", "sys_source": "🟧來源",
        "use_target": "🟧使用對象", "use_purpose": "🟧使用目的", "use_method": "🟧使用方式", "use_protect": "🟧使用保護方式",
        "trans_target": "🟧傳送對象", "trans_purpose": "🟧傳送目的", "trans_method": "🟧傳送方式", "trans_protect": "🟧傳送保護方式",
        "store_loc": "🟪儲存位置", "store_legal_time": "🟪法定時限", "store_inner_time": "🟪內部保存期限", "store_protect": "🟪儲存保護措施",
        "del_method": "🟪刪除方式", "del_unit": "🟪刪除單位",
        "intl_country": "🟥傳送國家", "intl_target": "🟥國際傳送對象", "intl_purpose": "🟥國際傳送目的", "intl_method": "🟥國際傳送方式", "intl_protect": "🟥國際保護方式"
    }
    for s in scopes: cfg[f"scope_{s}"] = st.column_config.SelectboxColumn(f"🟩{s}", options=YN_OPTIONS)

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=order, column_config=cfg)
    
    c1, c2 = st.columns([1, 6])
    with c1:
        if st.button("💾 儲存清冊"):
            if save_data("pi_inventory", edited, df): time.sleep(0.5); st.rerun()
    with c2:
        rename_dict = {"dept_name":"I.部名稱","room_name":"I.室名稱","pi_manager":"I.個資檔案管理者","process_desc":"I.業務流程說明","pi_amount":"II.筆數/份數","legal_rule": "II.法源/內部規範依據", "pi_purpose":"II.特定目的","pi_category":"II.個資之類別","legal_basis":"II.合法蒐集依據","collect_method":"II.蒐集方式","sys_name":"III.應用系統名稱","sys_source":"III.來源","use_target":"[使用]對象","use_purpose":"[使用]目的","use_method":"[使用]方式","use_protect":"[使用]保護方式","trans_target":"[傳送]對象","trans_purpose":"[傳送]目的","trans_method":"[傳送]方式","trans_protect":"[傳送]保護方式","store_loc":"[儲存]位置","store_legal_time":"[儲存]法定保留時限","store_inner_time":"[儲存]內部保存期限","store_protect":"[儲存]保護措施","del_method":"[刪除]銷毀方式","del_unit":"[刪除]銷毀單位","intl_country":"[國際]國家","intl_target":"[國際]對象","intl_purpose":"[國際]目的","intl_method":"[國際]方式","intl_protect":"[國際]保護方式"}
        for s in scopes: rename_dict[f"scope_{s}"] = f"[範圍]{s}"
        rules = {
            "blue": ["I.部名稱","I.室名稱","I.個資檔案管理者","I.業務流程說明"],
            "green": ["II.筆數/份數","II.法源/內部規範依據","II.特定目的","II.個資之類別","II.合法蒐集依據","II.蒐集方式"] + [f"[範圍]{s}" for s in scopes],
            "orange": ["III.應用系統名稱","III.來源","[使用]對象","[使用]目的","[使用]方式","[使用]保護方式","[傳送]對象","[傳送]目的","[傳送]方式","[傳送]保護方式"],
            "purple": ["[儲存]位置","[儲存]法定保留時限","[儲存]內部保存期限","[儲存]保護措施","[刪除]銷毀方式","[刪除]銷毀單位"],
            "red": ["[國際]國家","[國際]對象","[國際]目的","[國際]方式","[國際]保護方式"]
        }
        xl = generate_excel(edited, rename_dict, rules)
        st.download_button("📥 匯出 Excel", xl, "個資清冊.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "3. 風險評鑑":
    st.markdown("### ⚠️ 個人資料風險評鑑")
    if is_admin: st.info("👁️ 目前身分：【總管理員】。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
    else: st.info(f"🔒 目前身分：【{user_unit}】。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
        
    df = load_data("risk_assessment")
    score_cols = ["score_1", "score_2", "score_3", "score_4", "score_5"]
    cols = ["unit_name", "project_name"] + score_cols
    for c in cols: 
        if c not in df.columns: df[c] = 1 if 'score' in c else None

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=cols, column_config={
        "id": None, 
        "unit_name": st.column_config.TextColumn("🟦單位", disabled=not is_admin), 
        "project_name": "🟦業務名稱",
        "score_1": st.column_config.NumberColumn("🟨(1)數量", min_value=1, max_value=5), 
        "score_2": st.column_config.NumberColumn("🟨(2)敏感度", min_value=1, max_value=5), 
        "score_3": st.column_config.NumberColumn("🟨(3)信譽損害", min_value=1, max_value=5), 
        "score_4": st.column_config.NumberColumn("🟨(4)隱私衝擊", min_value=1, max_value=5), 
        "score_5": st.column_config.NumberColumn("🟨(5)合作單位", min_value=1, max_value=5)
    })
    
    c1, c2 = st.columns([1, 6])
    with c1:
        if st.button("💾 儲存評估"):
            if save_data("risk_assessment", edited, df): time.sleep(0.5); st.rerun()
    with c2:
        rename_dict = {"unit_name": "單位", "project_name": "業務名稱", "score_1": "(1)數量", "score_2": "(2)敏感度", "score_3": "(3)信譽損害", "score_4": "(4)隱私衝擊", "score_5": "(5)合作單位"}
        rules = {"blue": ["單位", "業務名稱"], "yellow": ["(1)數量", "(2)敏感度", "(3)信譽損害", "(4)隱私衝擊", "(5)合作單位"]}
        xl = generate_excel(edited, rename_dict, rules)
        st.download_button("📥 匯出 Excel", xl, "風險評鑑表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "4. 委外廠商":
    st.markdown("### 🤝 委外廠商個資清冊")
    if is_admin: st.info("👁️ 目前身分：【總管理員】。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
    else: st.info(f"🔒 目前身分：【{user_unit}】。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
        
    df = load_data("vendor_inventory")
    cols = ["unit_name", "vendor_name", "file_name", "pi_scope", "trans_purpose", "trans_method"]
    for c in cols:
        if c not in df.columns: df[c] = None

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=cols, column_config={
        "id": None, 
        "unit_name": st.column_config.TextColumn("🟦單位", disabled=not is_admin), 
        "vendor_name": "🟦廠商名稱", "file_name": "🟦檔案名稱",
        "pi_scope": "🟩個資範圍", "trans_purpose": "🟧傳送目的", "trans_method": "🟧傳送方式"
    })
    
    c1, c2 = st.columns([1, 6])
    with c1:
        if st.button("💾 儲存廠商"):
            if save_data("vendor_inventory", edited, df): time.sleep(0.5); st.rerun()
    with c2:
        rename_dict = {"unit_name": "單位", "vendor_name": "廠商名稱", "file_name": "個資檔案名稱", "pi_scope": "個資範圍", "trans_purpose": "傳送目的", "trans_method": "傳送方式"}
        rules = {"blue": ["單位", "廠商名稱", "個資檔案名稱"], "green": ["個資範圍"], "orange": ["傳送目的", "傳送方式"]}
        xl = generate_excel(edited, rename_dict, rules)
        st.download_button("📥 匯出 Excel", xl, "委外廠商清冊.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "5. 組織管理":
    st.markdown("### 🏢 組織架構管理")
    st.info("💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊下方儲存。連點兩下儲存格即可開始輸入！")
    
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("1. 部門 CRUD")
        ed_d = st.data_editor(
            df_dept, num_rows="dynamic", use_container_width=True, 
            column_order=["dept_name"],
            column_config={
                "id": None, 
                "dept_name": st.column_config.TextColumn("🏢 部門名稱", required=True)
            }
        )
        if st.button("💾 存部門"):
            if save_data("departments", ed_d, df_dept): time.sleep(1); st.rerun()
            
    with c2:
        st.subheader("2. 單位 CRUD")
        opts = dept_list if dept_list else ["(請先建立部門)"]
        ed_u = st.data_editor(
            df_unit, num_rows="dynamic", use_container_width=True, 
            column_order=["dept_name", "unit_name"],
            column_config={
                "id": None, 
                "dept_name": st.column_config.SelectboxColumn("所屬部門", options=opts, required=True), 
                "unit_name": st.column_config.TextColumn("🏠 單位名稱", required=True)
            }
        )
        if st.button("💾 存單位"):
            if save_data("units", ed_u, df_unit): time.sleep(1); st.rerun()

st.sidebar.divider()
st.sidebar.caption("© 2026 Carmax Co., Ltd.")
