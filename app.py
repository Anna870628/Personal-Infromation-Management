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
FILE_TYPE_OPTIONS = ["實體紙本", "數位檔案", "影像檔案", "影音檔案"]

SCORE_1_OPTS = ["5: 每年產生大於1000筆", "3: 每年產生100~1000筆", "1: 每年產生小於100筆"]
SCORE_2_OPTS = ["5: 包含姓名、身分證號、私人連絡方式(電話+地址)、財務情況、指紋、特種個資", "3: 包含姓名、身分證號、護照、私人聯絡方式(電話及地址)、其他非特種個資欄位", "1: 僅含姓名、聯絡方式(電話)"]
SCORE_3_OPTS = ["5: 若作業發生個資外洩事故，將導致公司形象、信譽受到非常嚴重損害，如：導致國際性媒體報導負面新聞、造成民眾集結遊行抗爭或上級機關關切等情形。", "3: 若作業發生個資外洩事故，將導致公司形象、信譽受到嚴重損害，如：導致3家以上媒體報導負面新聞或造成民眾至機關抗議或陳情等情形。", "1: 若該作業發生個資外洩事故，將導致公司形象、信譽受到輕微損害，如：導致部份媒體報導負面新聞、造成多位民眾電話抱怨等情形。"]
SCORE_4_OPTS = ["5: 洩漏資訊，對個資當事人造成重大影響，如：勒索、綁架、甚至危及生命安全或造成重大財產損失。", "3: 洩漏資訊，對個資當事人有中度影響，如：身分遭冒用、詐騙、影響個人信用或造成部分財產損失。", "1: 洩漏資訊，對個資當事人影響較輕微，如：收到推銷電話、垃圾郵件等滋擾。"]
SCORE_5_OPTS = ["5: 業務作業流程涉及外部廠商或第三方，且未簽訂保密協定或缺乏安全監督機制。", "3: 業務作業流程涉及外部廠商或第三方，已簽訂保密協定但缺乏定期監督或稽核。", "1: 業務作業流程無委外(僅內部作業)，或委外且具備完善保密協定與定期監督機制。"]

def generate_excel(df, rename_dict, color_rules):
    """通用匯出引擎 (保留給自檢表、個資清冊、風險評鑑表使用)"""
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
            "grey": workbook.add_format({'bg_color': '#D9D9D9', 'border': 1, 'bold': True}),
        }
        
        for col_num, value in enumerate(export_df.columns.values):
            fmt_key = "default"
            for color, columns in color_rules.items():
                if value in columns:
                    fmt_key = color
                    break
            if fmt_key in formats:
                worksheet.write(0, col_num, value, formats[fmt_key])
            worksheet.set_column(col_num, col_num, 25) 
            
    return output.getvalue()

def generate_vendor_excel(df):
    """⭐️ 委外廠商專屬 100% 官方排版匯出引擎"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = writer.sheets.add_worksheet('委外廠商個資檔案清冊')

        # 設定各種儲存格格式
        title_fmt = workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter', 'font_size': 11})
        hdr_main_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        inst_fmt = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFF2CC', 'text_wrap': True})
        data_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        data_text_fmt = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})

        # 第 1 列：表單抬頭
        worksheet.set_row(0, 30)
        worksheet.merge_range(0, 0, 0, 35, "表單名稱：委外廠商個資檔案清冊\n表單編號：BM000-B-007-D001", title_fmt)

        # 第 2-3 列：合併標題列
        worksheet.merge_range(1, 0, 2, 0, "編號", hdr_main_fmt)
        worksheet.merge_range(1, 1, 2, 1, "委外廠商名稱", hdr_main_fmt)
        
        worksheet.merge_range(1, 2, 1, 5, "個人資料資訊", hdr_main_fmt)
        worksheet.write(2, 2, "個人資料檔案名稱", hdr_main_fmt)
        worksheet.write(2, 3, "檔案類型", hdr_main_fmt)
        worksheet.write(2, 4, "筆數/份數", hdr_main_fmt)
        worksheet.write(2, 5, "個人資料檔案使用目的", hdr_main_fmt)

        worksheet.merge_range(1, 6, 1, 27, "個人資料範圍", hdr_main_fmt)
        scopes = ["姓名", "出生年月日", "國民身分證編號", "電話", "地址", "護照號碼", "特徵", "指紋", "婚姻", "家庭", "教育", "職業", "病歷", "特種資料", "財務情況", "社會活動", "車籍資料\n(車號、引擎號碼、車身號碼、底盤號碼等)", "醫療", "基因", "性生活", "健康檢查", "犯罪前科"]
        for i, s in enumerate(scopes):
            worksheet.write(2, 6 + i, s, hdr_main_fmt)

        worksheet.merge_range(1, 28, 1, 34, "個人資料生命循環", hdr_main_fmt)
        life_cols = ["資料來源", "資料來源管道", "儲存地點及位置", "資料鍵入之資訊系統", "傳送對象", "傳送目的", "傳送方式"]
        for i, s in enumerate(life_cols):
            worksheet.write(2, 28 + i, s, hdr_main_fmt)

        worksheet.merge_range(1, 35, 2, 35, "備註", hdr_main_fmt)

        # 第 4 列：黃底說明列
        worksheet.set_row(3, 110)
        worksheet.write(3, 0, "請依流水號進行填列", inst_fmt)
        worksheet.write(3, 1, "請填寫委外廠商名稱", inst_fmt)
        worksheet.write(3, 2, "請填列含有和泰所屬個人資料之檔案名稱\n(個人資料應分筆分列填寫)", inst_fmt)
        worksheet.write(3, 3, "請填列實體紙本、數位檔案、影像檔案、影音檔案\n(不同類型請填列不同筆)", inst_fmt)
        worksheet.write(3, 4, "填列筆數(數位、影像、影音)/份數(紙本)", inst_fmt)
        worksheet.write(3, 5, "識別該資料之使用目的", inst_fmt)
        
        worksheet.merge_range(3, 6, 3, 27, "辨識檔案是否含有自然人之姓名、出生年月日、國民身分證統一編號、護照號碼、特徵…等個人資料\n(如有請填列Y，如無請填列N)", inst_fmt)
        
        worksheet.write(3, 28, "請填列個人資料的來源", inst_fmt)
        worksheet.write(3, 29, "請填列資料來源管道 ，如：親自提供 / 郵件 / 傳真 / 雲端空間 / google表單 / 對外或對內系統(入口網站、FTP、其他公司系統等)", inst_fmt)
        worksheet.write(3, 30, "識別資料儲存地點及位置：\n(1) 如為實體紙本之取得 - 請填列儲存地點 - 例如：XX人員的上鎖櫃/XX檔案室的公用櫃/一般櫃\n(2) 如為數位檔案、影像檔案、影音檔案，請填列儲存地點，例如：XXX個人電腦/XXX公用檔案系統/XXXUSB、行動硬碟/XXX個人雲端空間", inst_fmt)
        worksheet.write(3, 31, "如有將資料鍵入資訊系統(公司內部/公司外部) ，請填列，例如：XX系統\n(如無請填列N/A)", inst_fmt)
        worksheet.write(3, 32, "資料傳送之對象(如和泰、顧客、XXX廠商(含郵局)、XXX主管機關、其他Legal Entity 或 XXX內部單位等)\n(如無請填列N/A)", inst_fmt)
        worksheet.write(3, 33, "傳送目的\n(如無請填列N/A)", inst_fmt)
        worksheet.write(3, 34, "如傳送對象及目的之欄位有填列，請說明傳輸資料的方式，如親自提供 / 郵寄 / 掛號 / 快遞 / 傳真 / 檔案傳遞 / 雲端空間/google表單/對外或對內系統(入口網站、FTP、其他公司系統等)。\n(如無請填列N/A)", inst_fmt)
        worksheet.write(3, 35, "針對前面所填寫進行補充說明", inst_fmt)

        # 填入真實資料
        col_keys = ["item_no", "vendor_name", "file_name", "file_type", "pi_amount", "pi_purpose"] + [f"scope_{s.split(chr(10))[0]}" for s in scopes] + ["data_source", "source_channel", "store_loc", "sys_name", "trans_target", "trans_purpose", "trans_method", "remark"]
        
        for row_idx, row_data in enumerate(df.to_dict('records')):
            for col_idx, col_key in enumerate(col_keys):
                if "車籍資料" in col_key: col_key = "scope_車籍資料"
                val = row_data.get(col_key, "")
                val_clean = val if pd.notnull(val) else ""
                
                # Y/N 等短資料置中，長文字靠左
                if col_idx >= 6 and col_idx <= 27:
                    worksheet.write(4 + row_idx, col_idx, val_clean, data_fmt)
                else:
                    worksheet.write(4 + row_idx, col_idx, val_clean, data_text_fmt)

        # 設定欄位寬度
        worksheet.set_column(0, 0, 10)
        worksheet.set_column(1, 2, 20)
        worksheet.set_column(3, 5, 18)
        worksheet.set_column(6, 27, 6) # Y/N 欄位變窄
        worksheet.set_column(28, 35, 25) 

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
    
    st.markdown("##### 💡 填寫範例與說明參考 (同 Excel 附件)")
    st.info("📌 **【個人資料範圍 (姓名 ~ 其他) 填寫說明】**：\n請依個人資料保護法施行細則第4條及第5條之規定，就所蒐集之個人資料，於該適當欄位填列 Y ，若無則填列 N ，但其他可直接或間接方式識別個人之資料(請於「其他」欄位直接列舉)。")
    
    example_dict = {
        "dept_name": "請填列部門名稱", "room_name": "請填列室名稱", "pi_manager": "請填列個資檔案管理者人員名稱", "process_desc": "請填列業務子流程名稱",
        "pi_amount": "請選擇約略數量", "legal_rule": "外部法規依據/內部規範依據/NA", "pi_purpose": "請下拉選擇", "pi_category": "請下拉選擇",
        "legal_basis": "列示合法蒐集個資之依據，如：個資授權同意書、隱私權條款\n(僅資料蒐集單位須填寫)", "collect_method": "屬於直接蒐集或間接蒐集\n(僅資料蒐集單位須填寫)",
        "sys_name": "該筆個人資料涉及的系統或檔案名稱", "sys_source": "請填寫個人資料(包括紙本文件跟電子檔案)的來源對象，不限公司內外單位；若個人資料來自於資訊系統，則填寫資訊系統名稱",
        "use_target": "資料單位內使用者\n(如無請填列N/A)", "use_purpose": "使用目的如：資料建檔、廣告投放等\n(如無請填列N/A)",
        "use_method": "如使用者及目的之欄位有填列，請說明使用資料的方式，如列印、下載 。\n(如無請填列N/A)", "use_protect": "如有填寫使用方式，應一併說明保護方式，如: 權限控管、刷卡等\n(如無請填列N/A)",
        "trans_target": "資料傳送之對象(如:XXX委外廠商、XXX主管機關或XXX內部單位等)\n(如無請填列N/A)", "trans_purpose": "傳送目的\n(如無請填列N/A)",
        "trans_method": "如傳送對象及目的之欄位有填列，請說明傳輸資料的方式，如親自提供 / 郵寄 / 掛號 / 快遞 / 傳真 / 檔案傳遞 / 對外或對內系統(入口網站、FTP、其他公司系統等) 。\n(如無請填列N/A)", "trans_protect": "如有填寫傳送方式，應一併說明保護方式，如: 專人親送／親取／加密等\n(如無請填列N/A)",
        "store_loc": "如:實體櫃/雲端資料庫", "store_legal_time": "法定保存年限", "store_inner_time": "公司內部規定保存年限", "store_protect": "上鎖、密碼控管等",
        "del_method": "碎紙機銷毀、系統刪除等", "del_unit": "負責執行銷毀之單位",
        "intl_country": "傳送到其他國家\n(如無請填列N/A)", "intl_target": "傳送對象\n(如無請填列N/A)", "intl_purpose": "傳送目的\n(如無請填列N/A)", 
        "intl_method": "如傳送國家及目的之欄位有填列，請說明傳輸資料的方式，例如：檔案傳輸系統、應用程式與應用程式之間傳輸等。\n(如無請填列N/A)", "intl_protect": "保護方式\n(如無請填列N/A)"
    }
    for s in scopes: example_dict[f"scope_{s}"] = "填 Y 或 N"
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
    st.dataframe(ex_df.style.set_properties(**{'color': '#1E90FF', 'white-space': 'pre-wrap'}), hide_index=True)

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
    
    st.markdown("##### 💡 填寫範例與說明參考 (同 Excel 附件)")
    st.info("📌 **【風險評鑑表填寫說明】**：\n請針對各項業務流程，於下拉選單選擇對應的風險等級 (5、3、1分)，並自行加總前5項分數填入「風險評分加總」，最後說明應對風險的作法。")
    
    example_risk_dict = {
        "item_no": "請依流水號進行填列", "unit_name": "請填列單位名稱", "project_name": "請填列業務子流程名稱",
        "score_1": "請下拉選擇 (5, 3, 1)", "score_2": "請下拉選擇 (5, 3, 1)", "score_3": "請下拉選擇 (5, 3, 1)",
        "score_4": "請下拉選擇 (5, 3, 1)", "score_5": "請下拉選擇 (5, 3, 1)",
        "total_score": "將前5項評分加總(最高25分)", "risk_action": "說明風險對應作法(如: 維持現狀、增加控管)"
    }
    
    rename_risk_map = {
        "item_no": "🟦編號", "unit_name": "🟦單位名稱", "project_name": "🟦作業流程名稱",
        "score_1": "🟨(1) 個資數量", "score_2": "🟨(2) 個資敏感度", "score_3": "🟨(3) 損害組織信譽", 
        "score_4": "🟨(4) 當事人隱私衝擊", "score_5": "🟨(5) 業務合作單位",
        "total_score": "🟨風險評分加總", "risk_action": "🟩風險對應作法"
    }
    
    ex_risk_df = pd.DataFrame([example_risk_dict]).rename(columns=rename_risk_map)
    st.dataframe(ex_risk_df.style.set_properties(**{'color': '#1E90FF', 'white-space': 'pre-wrap'}), hide_index=True)

    df = load_data("risk_assessment")
    risk_cols = ["item_no", "unit_name", "project_name", "score_1", "score_2", "score_3", "score_4", "score_5", "total_score", "risk_action"]
    for c in risk_cols: 
        if c not in df.columns: df[c] = None

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=risk_cols, column_config={
        "id": None, "item_no": "🟦編號",
        "unit_name": st.column_config.TextColumn("🟦單位名稱", disabled=not is_admin), 
        "project_name": "🟦作業流程名稱",
        "score_1": st.column_config.SelectboxColumn("🟨(1) 個資數量", options=SCORE_1_OPTS, width="large"), 
        "score_2": st.column_config.SelectboxColumn("🟨(2) 個資敏感度", options=SCORE_2_OPTS, width="large"), 
        "score_3": st.column_config.SelectboxColumn("🟨(3) 損害組織信譽", options=SCORE_3_OPTS, width="large"), 
        "score_4": st.column_config.SelectboxColumn("🟨(4) 當事人隱私衝擊", options=SCORE_4_OPTS, width="large"), 
        "score_5": st.column_config.SelectboxColumn("🟨(5) 業務合作單位", options=SCORE_5_OPTS, width="large"),
        "total_score": "🟨風險評分加總", "risk_action": "🟩風險對應作法"
    })
    
    c1, c2 = st.columns([1, 6])
    with c1:
        if st.button("💾 儲存評估"):
            if save_data("risk_assessment", edited, df): time.sleep(0.5); st.rerun()
    with c2:
        rename_dict = {
            "item_no": "編號", "project_name": "作業流程名稱", "unit_name": "單位名稱",
            "score_1": "(1) 個資數量", "score_2": "(2) 個資敏感度", "score_3": "(3) 損害組織信譽", 
            "score_4": "(4) 當事人隱私衝擊", "score_5": "(5) 業務合作單位",
            "total_score": "風險評分加總", "risk_action": "風險對應作法"
        }
        rules = {"blue": ["編號", "作業流程名稱", "單位名稱"], "yellow": ["(1) 個資數量", "(2) 個資敏感度", "(3) 損害組織信譽", "(4) 當事人隱私衝擊", "(5) 業務合作單位", "風險評分加總"], "green": ["風險對應作法"]}
        xl = generate_excel(edited, rename_dict, rules)
        st.download_button("📥 匯出 Excel", xl, "風險評鑑表.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "4. 委外廠商":
    st.markdown("### 🤝 委外廠商個資清冊")
    if is_admin: st.info("👁️ 目前身分：【總管理員】。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
    else: st.info(f"🔒 目前身分：【{user_unit}】。💡 刪除方式：選取最左側行號 -> 按鍵盤 `Delete` 鍵 -> 點擊儲存。")
    
    # ------------------------------------------
    # 🌟 委外廠商清冊：100% 依據最新 Excel 欄位對齊
    # ------------------------------------------
    vendor_scopes = [
        "姓名", "出生年月日", "國民身分證編號", "電話", "地址", "護照號碼", "特徵", "指紋", 
        "婚姻", "家庭", "教育", "職業", "病歷", "特種資料", "財務情況", "社會活動", 
        "車籍資料", "醫療", "基因", "性生活", "健康檢查", "犯罪前科"
    ]
    
    st.markdown("##### 💡 填寫範例與說明參考 (同 Excel 附件)")
    st.info("📌 **【個人資料範圍填寫說明】**：\n辨識檔案是否含有自然人之姓名、出生年月日、國民身分證統一編號、護照號碼、特徵…等個人資料 (如有請下拉選擇Y，如無請選擇N)。")
    
    ex_vendor_dict = {
        "item_no": "請依流水號進行填列",
        "vendor_name": "請填寫委外廠商名稱",
        "file_name": "請填列含有和泰所屬個人資料之檔案名稱\n(個人資料應分筆分列填寫)",
        "file_type": "請下拉選擇",
        "pi_amount": "填列筆數(數位、影像、影音)/份數(紙本)",
        "pi_purpose": "識別該資料之使用目的",
        "data_source": "請填列個人資料的來源",
        "source_channel": "請填列資料來源管道 ，如：親自提供 / 郵件 / 傳真 / 雲端空間 / google表單 / 對外或對內系統(入口網站、FTP、其他公司系統等)",
        "store_loc": "識別資料儲存地點及位置：\n(1) 如為實體紙本之取得 - 請填列儲存地點 - 例如：XX人員的上鎖櫃/XX檔案室的公用櫃/一般櫃\n(2) 如為數位檔案、影像檔案、影音檔案，請填列儲存地點，例如：XXX個人電腦/XXX公用檔案系統/XXXUSB、行動硬碟/XXX個人雲端空間",
        "sys_name": "如有將資料鍵入資訊系統(公司內部/公司外部) ，請填列，例如：XX系統\n(如無請填列N/A)",
        "trans_target": "資料傳送之對象(如和泰、顧客、XXX廠商(含郵局)、XXX主管機關、其他Legal Entity 或 XXX內部單位等)\n(如無請填列N/A)",
        "trans_purpose": "傳送目的\n(如無請填列N/A)",
        "trans_method": "如傳送對象及目的之欄位有填列，請說明傳輸資料的方式，如親自提供 / 郵寄 / 掛號 / 快遞 / 傳真 / 檔案傳遞 / 雲端空間/google表單/對外或對內系統(入口網站、FTP、其他公司系統等)。\n(如無請填列N/A)",
        "remark": "針對前面所填寫進行補充說明"
    }
    
    for s in vendor_scopes: ex_vendor_dict[f"scope_{s}"] = "填 Y 或 N"
    
    rename_vendor_map = {
        "item_no": "🟦編號", "vendor_name": "🟦委外廠商名稱", "file_name": "🟦個人資料檔案名稱", "file_type": "🟦檔案類型",
        "pi_amount": "🟩筆數/份數", "pi_purpose": "🟩個人資料檔案使用目的",
        "data_source": "🟧資料來源", "source_channel": "🟧資料來源管道", "sys_name": "🟧資料鍵入之資訊系統",
        "trans_target": "🟧傳送對象", "trans_purpose": "🟧傳送目的", "trans_method": "🟧傳送方式",
        "store_loc": "🟪儲存地點及位置", "remark": "⬜備註"
    }
    
    for s in vendor_scopes:
        if s == "車籍資料": rename_vendor_map[f"scope_{s}"] = "🟩車籍資料\n(車號、引擎號碼、車身號碼、底盤號碼等)"
        else: rename_vendor_map[f"scope_{s}"] = f"🟩{s}"
    
    vendor_order = ["item_no", "vendor_name", "file_name", "file_type", "pi_amount", "pi_purpose"] + [f"scope_{s}" for s in vendor_scopes] + ["data_source", "source_channel", "store_loc", "sys_name", "trans_target", "trans_purpose", "trans_method", "remark"]
    
    ex_vendor_df = pd.DataFrame([ex_vendor_dict])[ [c for c in vendor_order if c in rename_vendor_map] ].rename(columns=rename_vendor_map)
    st.dataframe(ex_vendor_df.style.set_properties(**{'color': '#1E90FF', 'white-space': 'pre-wrap'}), hide_index=True)
    # ------------------------------------------

    df = load_data("vendor_inventory")
    for c in vendor_order:
        if c not in df.columns: df[c] = None

    cfg = {
        "id": None, "unit_name": None,
        "item_no": "🟦編號", "vendor_name": "🟦委外廠商名稱", "file_name": "🟦個人資料檔案名稱",
        "file_type": st.column_config.SelectboxColumn("🟦檔案類型", options=FILE_TYPE_OPTIONS),
        "pi_amount": "🟩筆數/份數", "pi_purpose": "🟩個人資料檔案使用目的",
        "data_source": "🟧資料來源", "source_channel": "🟧資料來源管道", "sys_name": "🟧資料鍵入之資訊系統",
        "trans_target": "🟧傳送對象", "trans_purpose": "🟧傳送目的", "trans_method": "🟧傳送方式",
        "store_loc": "🟪儲存地點及位置", "remark": "⬜備註"
    }
    
    for s in vendor_scopes:
        if s == "車籍資料":
            cfg[f"scope_{s}"] = st.column_config.SelectboxColumn("🟩車籍資料\n(車號、引擎號碼、車身號碼、底盤號碼等)", options=YN_OPTIONS)
        else:
            cfg[f"scope_{s}"] = st.column_config.SelectboxColumn(f"🟩{s}", options=YN_OPTIONS)

    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=vendor_order, column_config=cfg)
    
    c1, c2 = st.columns([1, 6])
    with c1:
        if st.button("💾 儲存廠商清冊"):
            if save_data("vendor_inventory", edited, df): time.sleep(0.5); st.rerun()
    with c2:
        # 當點擊匯出時，使用專屬的 Excel 匯出引擎
        xl = generate_vendor_excel(edited)
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
