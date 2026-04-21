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
PI_AMOUNT_OPTIONS = ["每年產生大於100萬筆", "每年產生10萬~100萬筆", "每年產生1萬~10萬筆", "每年產生1000~1萬筆", "每年產生小於1000筆"]
FILE_TYPE_OPTIONS = ["實體紙本", "數位檔案", "影像檔案", "影音檔案"]
SCORE_1_OPTS = ["5: 每年產生大於1000筆", "3: 每年產生100~1000筆", "1: 每年產生小於100筆"]
SCORE_2_OPTS = ["5: 包含姓名、身分證號、私人連絡方式(電話+地址)、財務情況、指紋、特種個資", "3: 包含姓名、身分證號、護照、私人聯絡方式(電話及地址)、其他非特種個資欄位", "1: 僅含姓名、聯絡方式(電話)"]
SCORE_3_OPTS = ["5: 若作業發生個資外洩事故，將導致公司形象、信譽受到非常嚴重損害...", "3: 若作業發生個資外洩事故，將導致公司形象、信譽受到嚴重損害...", "1: 若該作業發生個資外洩事故，將導致公司形象、信譽受到輕微損害..."]
SCORE_4_OPTS = ["5: 洩漏資訊，對個資當事人造成重大影響，如：勒索、綁架。", "3: 洩漏資訊，對個資當事人有部分影響，如：遭受不明騷擾、詐騙。", "1: 洩漏資訊，對個資當事人產生些微影響"]
SCORE_5_OPTS = ["5: 有將個資委託廠商進行蒐集、處理或利用，但廠商未取得相關資安認證。", "3: 有將個資委託廠商進行蒐集、處理或利用，該廠商有取得相關資安認證。", "1: 僅與公司內其他單位合作。"]

# ==========================================
# 2. 匯出引擎區
# ==========================================
def generate_excel_basic(df, rename_dict, color_rules):
    """通用匯出引擎 (自檢表使用)"""
    export_df = df.copy()
    ordered_cols = [col for col in rename_dict.keys() if col in export_df.columns]
    export_df = export_df[ordered_cols].rename(columns=rename_dict)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, index=False)
        workbook, worksheet = writer.book, writer.sheets['Sheet1']
        formats = {
            "blue": workbook.add_format({'bg_color': '#D9E1F2', 'border': 1, 'bold': True}),
            "green": workbook.add_format({'bg_color': '#E2EFDA', 'border': 1, 'bold': True}),
            "orange": workbook.add_format({'bg_color': '#FCE4D6', 'border': 1, 'bold': True})
        }
        for col_num, value in enumerate(export_df.columns.values):
            fmt = next((formats[c] for c, cols in color_rules.items() if value in cols), None)
            if fmt: worksheet.write(0, col_num, value, fmt)
            worksheet.set_column(col_num, col_num, 20)
    return output.getvalue()

def generate_pi_excel(df, scopes):
    """⭐️ 個資清冊專屬 100% 官方格式匯出引擎"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('個資清冊')
        title_fmt = workbook.add_format({'bold': True, 'align': 'left', 'font_size': 11})
        hdr_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        inst_fmt = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFF2CC', 'text_wrap': True})
        data_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})

        worksheet.write(0, 0, "最後更新日期：", title_fmt)
        worksheet.merge_range(1, 0, 3, 0, "編號", hdr_fmt)
        worksheet.merge_range(1, 1, 1, 4, "I. 單位及業務流程資訊", hdr_fmt)
        worksheet.merge_range(1, 5, 1, 29, "II. 個人資料資訊", hdr_fmt)
        worksheet.merge_range(1, 30, 1, 53, "III. 個人資料生命週期", hdr_fmt)

        headers_l2 = ["部名稱", "室名稱", "個資檔案管理者", "業務流程說明", "筆數/份數", "法源依據", "特定目的", "類別"]
        for i, h in enumerate(headers_l2): worksheet.merge_range(2, 1+i, 3, 1+i, h, hdr_fmt)
        
        worksheet.merge_range(2, 9, 2, 29, "個人資料範圍", hdr_fmt)
        worksheet.merge_range(2, 30, 2, 33, "蒐集/取得", hdr_fmt)
        worksheet.merge_range(2, 34, 2, 37, "使用", hdr_fmt)
        worksheet.merge_range(2, 38, 2, 41, "傳送", hdr_fmt)
        worksheet.merge_range(2, 42, 2, 45, "儲存", hdr_fmt)
        worksheet.merge_range(2, 46, 2, 48, "刪除", hdr_fmt)
        worksheet.merge_range(2, 49, 2, 53, "國際傳遞", hdr_fmt)

        for i, s in enumerate(scopes): worksheet.write(3, 9+i, s, hdr_fmt)
        sub_h = ["依據", "方式", "檔名", "來源", "對象", "目的", "方式", "保護", "對象", "目的", "方式", "保護", "位置", "法定", "內部", "保護", "方式", "單位", "日期", "國家", "對象", "目的", "方式", "保護"]
        for i, h in enumerate(sub_h): worksheet.write(3, 30+i, h, hdr_fmt)

        col_keys = ["item_no", "dept_name", "room_name", "pi_manager", "process_desc", "pi_amount", "legal_rule", "pi_purpose", "pi_category"] + [f"scope_{s}" for s in scopes] + ["legal_basis", "collect_method", "sys_name", "sys_source", "use_target", "use_purpose", "use_method", "use_protect", "trans_target", "trans_purpose", "trans_method", "trans_protect", "store_loc", "store_legal_time", "store_inner_time", "store_protect", "del_method", "del_unit", "del_date", "intl_country", "intl_target", "intl_purpose", "intl_method", "intl_protect"]
        
        for r_idx, r_data in enumerate(df.to_dict('records')):
            for c_idx, key in enumerate(col_keys):
                val = r_data.get(key, "")
                worksheet.write(5 + r_idx, c_idx, val if pd.notnull(val) else "", data_fmt)
    return output.getvalue()

def generate_risk_excel(df):
    """⭐️ 風險評鑑表專屬匯出引擎"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('風險評鑑')
        hdr_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        data_fmt = workbook.add_format({'align': 'center', 'border': 1, 'text_wrap': True})
        
        worksheet.merge_range(0, 0, 1, 0, "編號", hdr_fmt)
        worksheet.merge_range(0, 1, 1, 1, "作業流程名稱", hdr_fmt)
        worksheet.merge_range(0, 2, 0, 6, "業務流程風險分析", hdr_fmt)
        for i, h in enumerate(["個資數量", "敏感度", "信譽", "隱私", "合作單位"]): worksheet.write(1, 2+i, h, hdr_fmt)
        worksheet.merge_range(0, 7, 1, 7, "總分", hdr_fmt)
        worksheet.merge_range(0, 8, 1, 8, "對應作法", hdr_fmt)
        worksheet.merge_range(0, 9, 1, 9, "單位", hdr_fmt)
        
        col_keys = ["item_no", "project_name", "score_1", "score_2", "score_3", "score_4", "score_5", "total_score", "risk_action", "unit_name"]
        for r_idx, r_data in enumerate(df.to_dict('records')):
            for c_idx, key in enumerate(col_keys):
                val = r_data.get(key, "")
                worksheet.write(2 + r_idx, c_idx, val if pd.notnull(val) else "", data_fmt)
    return output.getvalue()

def generate_vendor_excel(df):
    """⭐️ 委外廠商專屬匯出引擎"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('委外廠商')
        hdr_fmt = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})
        data_fmt = workbook.add_format({'border': 1, 'align': 'left'})
        for c_idx, col in enumerate(df.columns):
            worksheet.write(0, c_idx, col, hdr_fmt)
            for r_idx, val in enumerate(df[col]):
                worksheet.write(r_idx+1, c_idx, str(val) if pd.notnull(val) else "", data_fmt)
    return output.getvalue()

# ==========================================
# 3. 資料庫與權限
# ==========================================
try:
    SUPABASE_URL = st.secrets["supabase"]["url"]
    SUPABASE_KEY = st.secrets["supabase"]["key"]
    SYSTEM_PASSWORD = st.secrets["auth"]["admin_password"]
    ADMIN_PWD = st.secrets["auth"]["admin_login_pwd"] # ⭐️ 讀取管理員專用密碼
except:
    st.error("❌ Secrets 設定錯誤。")
    st.stop()

@st.cache_resource
def init_connection(): return create_client(SUPABASE_URL, SUPABASE_KEY)
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
# 4. 側邊欄與管理員認證
# ==========================================
def fetch_org():
    try:
        d = supabase.table("departments").select("*").execute().data
        u = supabase.table("units").select("*").execute().data
        return pd.DataFrame(d or []), pd.DataFrame(u or [])
    except: return pd.DataFrame(), pd.DataFrame()

df_dept, df_unit = fetch_org()
unit_list = df_unit["unit_name"].dropna().unique().tolist() if not df_unit.empty else []
safe_unit_list = unit_list if len(unit_list) > 0 else ["(請先至組織管理建立單位)"]

st.sidebar.title("👤 登入設定")
selected_role = st.sidebar.selectbox("切換身分", unit_list + ["總管理員"])

# ⭐️ 管理員認證邏輯
if "admin_verified" not in st.session_state: st.session_state.admin_verified = False

if selected_role == "總管理員":
    if not st.session_state.admin_verified:
        st.sidebar.warning("🔐 請輸入管理員密碼")
        admin_input = st.sidebar.text_input("管理員密碼", type="password", key="admin_pwd_box")
        if admin_input == ADMIN_PWD:
            st.session_state.admin_verified = True
            st.sidebar.success("驗證成功")
            st.rerun()
        else:
            if admin_input != "": st.sidebar.error("密碼錯誤")
            st.info("請在左側輸入管理員專用密碼以解鎖進階功能。")
            st.stop()
    is_admin = True
    user_unit = "總管理員"
else:
    st.session_state.admin_verified = False # 切回一般單位時重置驗證
    is_admin = False
    user_unit = selected_role

menu_opts = ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑", "4. 委外廠商", "5. 組織管理"] if is_admin else ["1. 自檢表", "2. 個資清冊", "3. 風險評鑑", "4. 委外廠商"]
menu = st.sidebar.radio("📂 功能選單", menu_opts)

# ==========================================
# 5. 資料處理核心
# ==========================================
def load_data(table):
    q = supabase.table(table).select("*")
    if not is_admin: q = q.eq("unit_name", user_unit)
    res = q.execute().data
    return pd.DataFrame(res or [])

def save_data(table, edited_df, original_df):
    if not original_df.empty and "id" in original_df.columns:
        orig_ids = set(original_df["id"].dropna().astype(str).tolist())
        edit_ids = set(edited_df["id"].dropna().astype(str).tolist()) if "id" in edited_df.columns else set()
        deleted = list(orig_ids - edit_ids)
        if deleted: supabase.table(table).delete().in_("id", deleted).execute()

    if not is_admin: edited_df["unit_name"] = user_unit
    records = edited_df.where(pd.notnull(edited_df), None).to_dict(orient="records")
    valid = []
    for r in records:
        if any(r[k] is not None for k in r.keys() if k not in ['id', 'unit_name']):
            if pd.isna(r.get('id')): r.pop('id', None)
            valid.append(r)
    if valid:
        supabase.table(table).upsert(valid).execute()
        st.toast("✅ 資料已儲存", icon="☁️")
        return True
    return False

# ==========================================
# 6. 分頁內容
# ==========================================
if menu == "1. 自檢表":
    st.markdown("### 🛡️ 自檢表管理")
    df = load_data("self_checklist")
    if is_admin and not df.empty: df["item_no"] = [str(i) for i in range(1, len(df)+1)]
    cols = ["item_no", "unit_name", "project_name", "owner", "status", "pi_inventory_done", "vendor_mgmt_done", "vendor_name", "form_d001", "form_d002", "form_d003", "pi_destroyed"]
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=cols, column_config={
        "unit_name": st.column_config.SelectboxColumn("單位", options=safe_unit_list, disabled=not is_admin),
        "status": st.column_config.SelectboxColumn("狀態", options=YN_OPTIONS),
        "pi_inventory_done": st.column_config.SelectboxColumn("清冊", options=YN_OPTIONS),
        "vendor_mgmt_done": st.column_config.SelectboxColumn("委外", options=YN_OPTIONS),
        "form_d001": st.column_config.SelectboxColumn("D001", options=YN_OPTIONS),
        "form_d002": st.column_config.SelectboxColumn("D002", options=YN_OPTIONS),
        "form_d003": st.column_config.SelectboxColumn("D003", options=YN_OPTIONS),
        "pi_destroyed": st.column_config.SelectboxColumn("銷毀", options=YN_OPTIONS)
    })
    if st.button("💾 儲存"): 
        if save_data("self_checklist", edited, df): time.sleep(0.5); st.rerun()

elif menu == "2. 個資清冊":
    st.markdown("### 📁 個資與機敏檔案清冊")
    scopes = ["姓名", "出生年月日", "身分證號碼", "護照號碼", "特徵", "指紋", "婚姻", "家庭", "教育", "職業", "病歷", "醫療", "基因", "性生活", "健康檢查", "犯罪前科", "聯絡方式", "財務情況", "社會活動", "車籍資料", "其他"]
    df = load_data("pi_inventory")
    # ⭐️ 匯出
    if st.button("📥 匯出 1.2 版官方 Excel"):
        xl = generate_pi_excel(df, scopes)
        st.download_button("點此下載", xl, "個資清冊.xlsx")
    
    order = ["dept_name", "room_name", "pi_manager", "process_desc", "pi_amount", "legal_rule", "pi_purpose", "pi_category"] + [f"scope_{s}" for s in scopes] + ["legal_basis", "collect_method", "sys_name", "sys_source", "use_target", "use_purpose", "use_method", "use_protect", "trans_target", "trans_purpose", "trans_method", "trans_protect", "store_loc", "store_legal_time", "store_inner_time", "store_protect", "del_method", "del_unit", "del_date", "intl_country", "intl_target", "intl_purpose", "intl_method", "intl_protect"]
    for c in order: 
        if c not in df.columns: df[c] = None
    
    cfg = {"dept_name": st.column_config.SelectboxColumn("部", options=df_dept["dept_name"].unique().tolist() if not df_dept.empty else []), "pi_amount": st.column_config.SelectboxColumn("筆數", options=PI_AMOUNT_OPTIONS)}
    for s in scopes: cfg[f"scope_{s}"] = st.column_config.SelectboxColumn(s, options=YN_OPTIONS)
    
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=order, column_config=cfg)
    if st.button("💾 儲存清冊"):
        if save_data("pi_inventory", edited, df): time.sleep(0.5); st.rerun()

elif menu == "3. 風險評鑑":
    st.markdown("### ⚠️ 個人資料風險評鑑")
    df = load_data("risk_assessment")
    if is_admin and not df.empty: df["item_no"] = [str(i) for i in range(1, len(df)+1)]
    
    risk_cols = ["item_no", "project_name", "score_1", "score_2", "score_3", "score_4", "score_5", "total_score", "risk_action", "unit_name"]
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, column_order=risk_cols, column_config={
        "unit_name": st.column_config.SelectboxColumn("單位", options=safe_unit_list, default=user_unit, disabled=not is_admin),
        "score_1": st.column_config.SelectboxColumn("(1)", options=SCORE_1_OPTS),
        "score_2": st.column_config.SelectboxColumn("(2)", options=SCORE_2_OPTS),
        "score_3": st.column_config.SelectboxColumn("(3)", options=SCORE_3_OPTS),
        "score_4": st.column_config.SelectboxColumn("(4)", options=SCORE_4_OPTS),
        "score_5": st.column_config.SelectboxColumn("(5)", options=SCORE_5_OPTS)
    })
    if st.button("💾 儲存評估"):
        if save_data("risk_assessment", edited, df): time.sleep(0.5); st.rerun()
    if st.button("📥 匯出風險評鑑"):
        st.download_button("下載", generate_risk_excel(edited), "風險評鑑.xlsx")

elif menu == "4. 委外廠商":
    st.markdown("### 🤝 委外廠商個資清冊")
    df = load_data("vendor_inventory")
    if is_admin and not df.empty: df["item_no"] = [str(i) for i in range(1, len(df)+1)]
    
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    if st.button("💾 儲存"):
        if save_data("vendor_inventory", edited, df): time.sleep(0.5); st.rerun()
    if st.button("📥 匯出委外清冊"):
        st.download_button("下載", generate_vendor_excel(edited), "委外廠商.xlsx")

elif menu == "5. 組織管理":
    st.markdown("### 🏢 組織架構管理")
    c1, c2 = st.columns(2)
    with c1:
        ed_d = st.data_editor(df_dept, num_rows="dynamic", use_container_width=True)
        if st.button("存部門"): save_data("departments", ed_d, df_dept); st.rerun()
    with c2:
        ed_u = st.data_editor(df_unit, num_rows="dynamic", use_container_width=True)
        if st.button("存單位"): save_data("units", ed_u, df_unit); st.rerun()

st.sidebar.divider()
st.sidebar.caption("© 2026 Carmax Co., Ltd.")
