import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

# --- 1. НАСТРОЙКИ ПУТЕЙ И КОНСТАНТЫ ---
JOURNAL_DB = 'journal_db.csv'
JOBS_FILE = 'jobs_internal.xlsx'
LOGO_FILE = 'Mashav_Logo.png'
SETTINGS_FILE = 'mashav_settings.json'

MONTH_MARKERS = {
    1: ['01', 'jan', 'january', 'ינואר'], 2: ['02', 'feb', 'february', 'פברואר'],
    3: ['03', 'mar', 'march', 'מרץ'], 4: ['04', 'apr', 'april', 'אפריל'],
    5: ['05', 'may', 'מאי'], 6: ['06', 'jun', 'june', 'יוני'],
    7: ['07', 'jul', 'july', 'יולי'], 8: ['08', 'aug', 'august', 'אוגוסט'],
    9: ['09', 'sep', 'september', 'ספטמבר'], 10: ['10', 'oct', 'october', 'אוקטובר'],
    11: ['11', 'nov', 'november', 'נובמבר'], 12: ['12', 'dec', 'december', 'דצמבר']
}

st.set_page_config(page_title="יומן תפעולי משאב", layout="wide")

if 'log_date' not in st.session_state:
    st.session_state.log_date = datetime.now().date()

# --- 2. GOOGLE DRIVE API ---
@st.cache_resource
def get_drive_service():
    creds_json = json.loads(st.secrets["GOOGLE_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=['https://www.googleapis.com/auth/drive'])
    return build('drive', 'v3', credentials=creds)

try:
    drive_service = get_drive_service()
    FOLDER_ID = st.secrets["FOLDER_ID"]
except Exception as e:
    st.error(f"שגיאה בהתחברות לענן. בדוק את Secrets. Error: {e}")
    st.stop()

def get_file_id(filename):
    query = f"'{FOLDER_ID}' in parents and name = '{filename}' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    items = results.get('files', [])
    return items[0]['id'] if items else None

# --- SMTP EMAIL ENGINE ---
def send_email_core(to_email, subject, html_body, attachment_df=None, attachment_name=""):
    from_email = "mashav.journal@gmail.com"
    password = st.secrets.get("EMAIL_PASS", "").replace(" ", "")
    if not password:
        return False, "Пароль приложений לא נמצא ב-Secrets!"
    
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = Header(subject, 'utf-8')
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    
    if attachment_df is not None:
        excel_buffer = io.BytesIO()
        attachment_df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)
        part = MIMEApplication(excel_buffer.read(), Name=attachment_name)
        part['Content-Disposition'] = f'attachment; filename="{attachment_name}"'
        msg.attach(part)
        
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(from_email, password)
        server.send_message(msg)
        server.quit()
        return True, "נשלח בהצלחה למייל!"
    except Exception as e:
        return False, f"שגיאה בשרת הדואר: {str(e)}"

def send_jobs_email(to_email, data_list, date_str):
    df_mail = pd.DataFrame(data_list)
    html_table = df_mail.to_html(index=False, justify='right')
    html_body = f"""<html dir="rtl"><body style="font-family: Arial; text-align: right; direction: rtl;">
        <h2>עבודות מתוכננות להיום ({date_str})</h2>{html_table}</body></html>"""
    return send_email_core(to_email, f"עבודות מתוכננות להיום - {date_str}", html_body, df_mail, f"Jobs_{date_str}.xlsx")

def send_journal_email_func(to_email, data_list, date_str):
    if not data_list:
        html_table = "<p>אין רישומים ביומן ליום זה.</p>"
        df_mail = None
    else:
        df_mail = pd.DataFrame(data_list)
        html_table = df_mail.to_html(index=False, justify='right')
        
    html_body = f"""<html dir="rtl"><body style="font-family: Arial; text-align: right; direction: rtl;">
        <h2>דוח משמרת מרוכז ({date_str})</h2>{html_table}</body></html>"""
    return send_email_core(to_email, f"דוח משמרת - {date_str}", html_body, df_mail, f"Journal_{date_str}.xlsx" if df_mail is not None else "")

def send_warehouse_email(to_email, unit, shift, hour, desc, date_str):
    html_body = f"""<html dir="rtl"><body style="font-family: Arial; text-align: right; direction: rtl;">
        <h2>הזמנה / דיווח למחסן</h2>
        <p><b>תאריך:</b> {date_str}</p><p><b>יחידה:</b> {unit}</p>
        <p><b>משמרת:</b> {shift}</p><p><b>שעה:</b> {hour}</p>
        <p><b>תיאור:</b> {desc}</p></body></html>"""
    return send_email_core(to_email, f"הודעה למחסן - {unit} - {date_str}", html_body)

# --- 3. ТОТАЛЬНАЯ ЭКСЕЛЬ-СЕТКА (CSS) ---
st.markdown("""
    <style>
    .block-container { padding-top: 3.5rem !important; padding-bottom: 1rem !important; max-width: 98% !important; }
    .stApp { background-color: #14161d; }
    * { direction: rtl !important; text-align: right !important; }
    
    div[data-testid="stVerticalBlock"] { gap: 0px !important; }
    div[data-testid="stHorizontalBlock"] { gap: 0px !important; align-items: stretch !important; margin-bottom: 0px !important; }
    div[data-testid="column"] { padding: 0px !important; } 
    
    div.element-container { margin-bottom: 0px !important; padding-bottom: 0px !important; overflow: visible !important; }
    label[data-testid="stWidgetLabel"] { display: none !important; height: 0px !important; margin: 0px !important; }
    
    /* СДВИГ ЛОГОТИПА ВВЕРХ */
    div[data-testid="stImage"] {
        margin-top: -20px !important;
    }
    
    /* СТИЛЬ ДЛЯ ТЕКСТОВЫХ ПОЛЕЙ (ИНДУСТРИАЛЬНЫЙ ТЕМНЫЙ) */
    div[data-testid="stTextInput"] div[data-baseweb="input"] {
        border-radius: 0px !important; 
        height: 40px !important;
        min-height: 40px !important;
        background-color: transparent !important; 
        border: none !important;
        margin: 0px !important;
        padding: 0px !important;
    }
    
    div[data-testid="stTextInput"] input {
        background-color: #2d333b !important; 
        direction: rtl !important;
        text-align: right !important;
        font-size: 16px !important;
        font-weight: bold !important;
        color: #adbac7 !important;
        padding-right: 8px !important;
        border-radius: 0px !important;
        border: 1px solid #444c56 !important;
        height: 40px !important;
        margin: 0px !important;
        margin-top: -1px !important;
        margin-right: -1px !important; 
    }

    /* ЗАГОЛОВКИ СМЕН В СТИЛЕ SCADA ПАНЕЛИ */
    .header-orange, .header-blue {
        text-align: center !important;
        font-weight: bold !important;
        font-size: 16px !important;
        color: #ffffff !important;
        padding: 10px 0px !important; 
        border: 1px solid #444c56 !important;
        display: block !important;
        margin: 0px !important;
        margin-right: -1px !important;
        box-sizing: border-box !important;
        background-color: #22272e !important;
    }
    .header-orange { border-top: 3px solid #d35400 !important; }
    .header-blue { border-top: 3px solid #2980b9 !important; }
    .header-orange p, .header-blue p { margin: 0px !important; padding: 0px !important; line-height: normal !important; }

    /* НОМЕРА В עבודות */
    .num-box {
        background-color: #22272e !important;
        color: #adbac7 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-weight: bold !important;
        font-size: 16px !important;
        height: 40px !important;
        border: 1px solid #444c56 !important;
        margin: 0px !important;
        margin-top: -1px !important;
        margin-right: -1px !important;
    }
    .num-box p { margin: 0px !important; padding: 0px !important; }
    
    /* КНОПКА @ - ИНДУСТРИАЛЬНЫЙ ТЕМНЫЙ СТИЛЬ */
    div.row-widget.stButton { margin: 0px !important; padding: 0px !important; }
    button[kind="secondary"] {
        height: 40px !important;
        min-height: 40px !important;
        width: 100% !important;
        margin: 0px !important;
        padding: 0px !important;
        border-radius: 0px !important;
        border: 1px solid #444c56 !important;
        background-color: #22272e !important;
        color: #27ae60 !important;
        font-weight: bold !important;
        font-size: 18px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        margin-top: -1px !important;
        z-index: 10;
    }
    button[kind="secondary"]:hover {
        background-color: #27ae60 !important;
        color: #ffffff !important;
    }

    /* ДИЗАЙН ОСНОВНЫХ КНОПОК И ВКЛАДОК */
    .stTabs [data-baseweb="tab-list"] { background-color: #22272e; border-radius: 5px; padding: 2px; margin-bottom: 15px;}
    .stTabs [data-baseweb="tab"] { font-size: 22px !important; font-weight: bold !important; color: #adbac7 !important; padding: 10px 20px; }
    .stTabs [aria-selected="true"] { background-color: #14161d !important; color: #fff !important; border-radius: 5px; border: 1px solid #444c56 !important; }
    .stButton button[kind="primary"] { 
        background-color: #27ae60 !important; 
        color: white !important; 
        font-weight: bold; 
        font-size: 18px; 
        border: 1px solid #1e7e34 !important; 
        height: 42px !important;
        margin: 0px !important;
    }

    /* Цвета текстов для темной темы */
    h3, h4 { color: #ffffff !important; font-weight: bold !important; }
    div[data-testid="stMarkdownContainer"] p { color: #adbac7 !important; }
    
    [data-testid="stDataFrame"] { border: none !important; }
    
    /* НАСТРОЙКИ: Текстовые поля с английским языком */
    div[data-testid="stTextArea"] textarea {
        background-color: #2d333b !important;
        color: #adbac7 !important;
        border: 1px solid #444c56 !important;
        border-radius: 0px !important;
        direction: ltr !important; 
        text-align: left !important;
        font-size: 16px !important;
    }
    div[data-testid="stSelectbox"] div[role="button"] {
        background-color: #2d333b !important;
        color: #adbac7 !important;
        border: 1px solid #444c56 !important;
        border-radius: 0px !important;
    }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] {
        direction: ltr !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 4. КАНАЛЫ ДАННЫХ ОБЛАКА ---
@st.cache_data(ttl=10)
def get_schedule_file_drive(target_month):
    markers = MONTH_MARKERS.get(target_month, [])
    query = f"'{FOLDER_ID}' in parents and trashed = false"
    try:
        results = drive_service.files().list(q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        for f in results.get('files', []):
            name = f['name']
            if name.endswith(('.xlsx', '.xls', '.csv')):
                if any(m in name.lower() for m in markers): return f['id'], name
    except Exception: pass
    return None, None

@st.cache_data(ttl=10) 
def get_operators(file_id, target_day):
    if not file_id: return ["קובץ חסר"], ["קובץ חסר"]
    try:
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0); df = pd.read_excel(fh, header=None); raw = df.values.tolist()
        cal_row_idx = None; cleaned_cal_row = []
        for i, row in enumerate(raw[:15]):
            c_row = [str(x).split('.')[0].strip() for x in row]
            if '1' in c_row and '2' in c_row and '15' in c_row: cal_row_idx = i; cleaned_cal_row = c_row; break
        if cal_row_idx is None: return ["שגיאה"], []
        target_str = str(target_day)
        if target_str not in cleaned_cal_row: return ["חסר"], []
        target_col = cleaned_cal_row.index(target_str)
        known = ['אמיר', 'נתי', 'גידי', 'אודל', 'ויקטור', 'יבגני', 'ליאור', 'ודים', "ז'קה", 'סשה']
        s1, s2 = [], []
        for row in raw:
            for val in row:
                clean_val = str(val).strip()
                if clean_val in known:
                    sh_val = str(row[target_col]).split('.')[0].strip()
                    if sh_val == '1': s1.append(clean_val)
                    elif sh_val == '2': s2.append(clean_val)
                    break
        return list(dict.fromkeys(s1)), list(dict.fromkeys(s2))
    except: return ["שגיאה"], []

@st.cache_data(ttl=5)
def load_journal_db():
    file_id = get_file_id(JOURNAL_DB)
    if file_id:
        try:
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: _, done = downloader.next_chunk()
            fh.seek(0); return pd.read_csv(fh, dtype=str).fillna("")
        except: pass
    return pd.DataFrame(columns=['Date', 'Unit', 'Shift', 'RowIdx', 'Hour', 'Description'])

def get_journal_data_list(date_str, unit, shift):
    df = load_journal_db()
    sub = df[(df['Date'] == date_str) & (df['Unit'] == unit) & (df['Shift'] == shift)].copy()
    if not sub.empty:
        sub['RowIdx'] = pd.to_numeric(sub['RowIdx'])
        sub = sub.sort_values('RowIdx')
    raw_list = sub.to_dict('records')
    while len(raw_list) < 6: raw_list.append({'Hour': '', 'Description': ''})
    return raw_list[:6]

@st.cache_data(ttl=5)
def load_jobs_db(target_date):
    file_id = get_file_id(JOBS_FILE)
    if file_id:
        try:
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: _, done = downloader.next_chunk()
            fh.seek(0); df = pd.read_excel(fh).fillna("")
            
            if "Date" not in df.columns:
                return [""] * 15
            
            sub = df[df['Date'] == target_date].copy()
            if not sub.empty:
                sub['RowIdx'] = pd.to_numeric(sub['RowIdx'])
                sub = sub.sort_values('RowIdx')
                jobs_list = sub['Description'].tolist()
                while len(jobs_list) < 15: jobs_list.append("")
                return jobs_list[:15]
        except: pass
    return [""] * 15

@st.cache_data(ttl=5)
def load_settings():
    file_id = get_file_id(SETTINGS_FILE)
    if file_id:
        try:
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: _, done = downloader.next_chunk()
            fh.seek(0)
            return json.loads(fh.read().decode('utf-8'))
        except: pass
    return {"dropdown_emails": ["wider71@gmail.com"], "warehouse_email": "wider71@gmail.com"}

def generate_safe_styles(df, target_col):
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    for col in df.columns:
        for idx in df.index:
            val = str(df.at[idx, col]).split('.')[0].strip()
            css = ''
            if val == '1': 
                css += 'background-color: #a9dfbf; color: black; font-weight: bold; font-size: 18px;'
            elif val == '2': 
                css += 'background-color: #85c1e9; color: black; font-weight: bold; font-size: 18px;'
            elif val in ['8', '9']: 
                css += 'background-color: #f9e79f; color: black; font-size: 18px;'
            elif val in ['ח', 'מ']: 
                css += 'background-color: #f5b7b1; color: black; font-weight: bold; font-size: 18px;'
            else: 
                css += 'color: #adbac7; background-color: #14161d; font-size: 16px;'
            
            if str(col) == str(target_col):
                css += ' border: 3px solid #2ecc71 !important;'
                
            styles.at[idx, col] = css
    return styles

app_settings = load_settings()
warehouse_email_target = app_settings.get("warehouse_email", "wider71@gmail.com")
dropdown_emails_list = app_settings.get("dropdown_emails", ["wider71@gmail.com"])

# --- 5. ВКЛАДКИ ОКОН ---
tab_log, tab_sch, tab_jobs, tab_settings = st.tabs(["דוח משמרת", "סידור", "עבודות היום", "הגדרות"])

# ==========================================
# ОКНО 1: ОПЕРАТИВНЫЙ ЖУРНАЛ
# ==========================================
with tab_log:
    col_logo, col_title, col_cal_r, col_cal_m, col_cal_l = st.columns([1, 4, 1.2, 1.8, 1.2])
    with col_logo:
        if os.path.exists(LOGO_FILE): st.image(LOGO_FILE, width=120)
    with col_title:
        st.markdown("<h4 style='color: #ffffff; margin-top:2px; font-weight:bold; font-size:20px;'>דוח משמרת תחנת כוח משאב</h4>", unsafe_allow_html=True)
    with col_cal_r:
        if st.button("▶ יום הבא", type="primary", use_container_width=True): st.session_state.log_date += timedelta(days=1); st.rerun()
    with col_cal_m:
        new_date = st.date_input("תאריך", value=st.session_state.log_date, label_visibility="collapsed")
        if new_date != st.session_state.log_date: st.session_state.log_date = new_date; st.rerun()
    with col_cal_l:
        if st.button("יום קודם ◀", type="primary", use_container_width=True): st.session_state.log_date -= timedelta(days=1); st.rerun()

    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

    active_sch_id, sch_name = get_schedule_file_drive(st.session_state.log_date.month)
    s1_names, s2_names = get_operators(active_sch_id, st.session_state.log_date.day)

    st.markdown(f"""
        <div style="display: flex; gap: 10px; margin-bottom: 10px;">
            <div style="flex: 1; border: 1px solid #444c56; padding: 4px 12px; background-color: #22272e; border-radius: 4px; display: flex; justify-content: space-between; align-items: center;">
                <b style="color: #d35400; font-size: 14px;">🌞 משמרת בוקר:</b> <span style="font-size: 16px; font-weight: bold; color: #ffffff;">{', '.join(s1_names)}</span>
            </div>
            <div style="flex: 1; border: 1px solid #444c56; padding: 4px 12px; background-color: #22272e; border-radius: 4px; display: flex; justify-content: space-between; align-items: center;">
                <b style="color: #2980b9; font-size: 14px;">🌙 משמרת לילה:</b> <span style="font-size: 16px; font-weight: bold; color: #ffffff;">{', '.join(s2_names)}</span>
            </div>
        </div>
    """, unsafe_allow_html=True)

    date_str = st.session_state.log_date.strftime("%Y-%m-%d")
    units = [('טורבינה 1', 1), ('טורבינה 2', 2), ('טורבינה קיטורית', 3)]
    saved_inputs = {}
    
    for u_name, u_num in units:
        st.markdown(f'<div style="height:15px;"></div>', unsafe_allow_html=True)
        
        c_morn, c_space, c_night = st.columns([10, 0.5, 10])
        
        m_data = get_journal_data_list(date_str, u_name, 'Morning')
        n_data = get_journal_data_list(date_str, u_name, 'Night')
        
        with c_morn:
            st.markdown(f'<div class="header-orange"><p>{u_num}. {u_name} - משמרת בוקר</p></div>', unsafe_allow_html=True)
            for idx in range(6):
                c_d, c_h, c_b = st.columns([11.5, 2.5, 1.5])
                with c_d:
                    d_m = st.text_input(f"dm_{u_num}_{idx}", value=m_data[idx].get('Description',''), key=f"dm_{u_num}_{idx}_{date_str}", label_visibility="collapsed")
                with c_h:
                    h_m = st.text_input(f"hm_{u_num}_{idx}", value=m_data[idx].get('Hour',''), key=f"hm_{u_num}_{idx}_{date_str}", label_visibility="collapsed")
                with c_b:
                    if st.button("@", key=f"btn_wh_m_{u_num}_{idx}_{date_str}", type="secondary", use_container_width=True):
                        h_val_cur = st.session_state.get(f"hm_{u_num}_{idx}_{date_str}", "")
                        d_val_cur = st.session_state.get(f"dm_{u_num}_{idx}_{date_str}", "")
                        if h_val_cur.strip() or d_val_cur.strip():
                            success, msg = send_warehouse_email(warehouse_email_target, u_name, "בוקר", h_val_cur, d_val_cur, date_str)
                            if success: st.toast("נשלח למחסן בהצלחה!", icon="✅")
                            else: st.toast(f"שגיאה: {msg}", icon="❌")
                        else:
                            st.toast("השורה ריקה - אין מה לשלוח!", icon="⚠️")
                saved_inputs[(u_name, 'Morning', idx)] = (h_m, d_m)
                
        with c_night:
            st.markdown(f'<div class="header-blue"><p>{u_num}. {u_name} - משמרת לילה</p></div>', unsafe_allow_html=True)
            for idx in range(6):
                c_d, c_h, c_b = st.columns([11.5, 2.5, 1.5])
                with c_d:
                    d_n = st.text_input(f"dn_{u_num}_{idx}", value=n_data[idx].get('Description',''), key=f"dn_{u_num}_{idx}_{date_str}", label_visibility="collapsed")
                with c_h:
                    h_n = st.text_input(f"hn_{u_num}_{idx}", value=n_data[idx].get('Hour',''), key=f"hn_{u_num}_{idx}_{date_str}", label_visibility="collapsed")
                with c_b:
                    if st.button("@", key=f"btn_wh_n_{u_num}_{idx}_{date_str}", type="secondary", use_container_width=True):
                        h_val_cur = st.session_state.get(f"hn_{u_num}_{idx}_{date_str}", "")
                        d_val_cur = st.session_state.get(f"dn_{u_num}_{idx}_{date_str}", "")
                        if h_val_cur.strip() or d_val_cur.strip():
                            success, msg = send_warehouse_email(warehouse_email_target, u_name, "לילה", h_val_cur, d_val_cur, date_str)
                            if success: st.toast("נשלח למחסן בהצלחה!", icon="✅")
                            else: st.toast(f"שגיאה: {msg}", icon="❌")
                        else:
                            st.toast("השורה ריקה - אין מה לשלוח!", icon="⚠️")
                saved_inputs[(u_name, 'Night', idx)] = (h_n, d_n)

    st.markdown(f'<div style="height:20px;"></div>', unsafe_allow_html=True)
    
    col_save, col_send, col_addr, _ = st.columns([2, 2, 3, 5])
    with col_save:
        if st.button("💾 שמור יומן", type="primary", use_container_width=True):
            db = load_journal_db()
            db = db[db['Date'] != date_str]
            new_rows = []
            for (u_name, shift, idx), (h, d) in saved_inputs.items():
                if h.strip() or d.strip():
                    new_rows.append([date_str, u_name, shift, idx, h.strip(), d.strip()])
            if new_rows:
                new_df = pd.DataFrame(new_rows, columns=['Date', 'Unit', 'Shift', 'RowIdx', 'Hour', 'Description'])
                db = pd.concat([db, new_df])
            
            fh = io.BytesIO(); db.to_csv(fh, index=False, encoding='utf-8-sig'); fh.seek(0)
            media = MediaIoBaseUpload(fh, mimetype='text/csv', resumable=True)
            file_id = get_file_id(JOURNAL_DB)
            if file_id: 
                drive_service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
                st.cache_data.clear()
                st.success("היומן נשמר בענן בהצלחה!")
                st.rerun()
            else: 
                st.error("שגיאה 403: קובץ journal_db.csv לא נמצא.")
                
    with col_send:
        btn_send_log = st.button("✉️ שלח יומן", type="primary", use_container_width=True)
        
    with col_addr:
        target_log_email = st.selectbox("לשלוח יומן ל:", dropdown_emails_list, key="log_email", label_visibility="collapsed")
        
    if btn_send_log:
        with st.spinner("מכין ושולח דוח..."):
            journal_rows_for_email = []
            for (u_name, shift, idx), (h, d) in saved_inputs.items():
                if h.strip() or d.strip():
                    s_name = "בוקר" if shift == "Morning" else "לילה"
                    journal_rows_for_email.append({"יחידה": u_name, "משמרת": s_name, "שעה": h.strip(), "תיאור": d.strip()})
            success, msg = send_journal_email_func(target_log_email, journal_rows_for_email, date_str)
            if success: st.success(msg)
            else: st.error(msg)


# ==========================================
# ОКНО 2: РАСПИСАНИЕ (СИДУР)
# ==========================================
with tab_sch:
    st.markdown("<h3>טבלת סידור עבודה</h3>", unsafe_allow_html=True)
    if active_sch_id:
        try:
            request = drive_service.files().get_media(fileId=active_sch_id)
            fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: _, done = downloader.next_chunk()
            fh.seek(0); df_excel = pd.read_excel(fh, header=None).fillna("")
            raw_matrix = df_excel.values.tolist()
            cleaned_data = [[str(val).replace('.0', '') if val != "" else "" for val in row] for row in raw_matrix]
            
            df_clean = pd.DataFrame(cleaned_data)
            df_clean.columns = df_clean.columns.astype(str)
            
            rev_cols = list(df_clean.columns)[::-1]
            df_ui = df_clean[rev_cols]
            df_ui.rename(columns={'0': 'שם'}, inplace=True)
            
            target_day_str = str(st.session_state.log_date.day)
            target_col_name = None
            for idx, row in df_ui.iterrows():
                row_vals = [str(x).split('.')[0].strip() for x in row]
                if '1' in row_vals and '15' in row_vals:
                    for col in df_ui.columns:
                        if str(df_ui.at[idx, col]).split('.')[0].strip() == target_day_str:
                            target_col_name = str(col)
                    break
            
            styled_df = df_ui.style.apply(lambda df: generate_safe_styles(df, target_col_name), axis=None)
            styled_df = styled_df.set_properties(**{'text-align': 'center', 'font-weight': 'bold'})
            
            st.dataframe(styled_df, use_container_width=True, height=550)
        except Exception as e: st.error(f"שגיאה: {e}")


# ==========================================
# ОКНО 3: РАБОТЫ НА СЕГОДНЯ
# ==========================================
with tab_jobs:
    st.markdown("<h3>עבודות מתוכננות להיום</h3>", unsafe_allow_html=True)
    
    loaded_jobs = load_jobs_db(date_str)
    saved_jobs_inputs = []
    
    for i in range(15):
        c_task, c_num = st.columns([14, 1])
        with c_num:
            st.markdown(f'<div class="num-box"><p>{i+1}</p></div>', unsafe_allow_html=True)
        with c_task:
            t_val = st.text_input(f"משימה {i+1}", value=loaded_jobs[i], key=f"job_input_{i}_{date_str}", label_visibility="collapsed")
        saved_jobs_inputs.append({"מספר": str(i+1), "משימות ופעולות לביצוע": t_val})
        
    st.markdown(f'<div style="height:20px;"></div>', unsafe_allow_html=True)
    
    col_save_j, col_send_j, col_addr_j, _ = st.columns([2, 2, 3, 5])
    
    with col_save_j:
        if st.button("💾 שמור עבודות", type="primary", use_container_width=True, key="save_jobs_btn"):
            try:
                file_id = get_file_id(JOBS_FILE)
                if file_id:
                    request = drive_service.files().get_media(fileId=file_id)
                    fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while not done: _, done = downloader.next_chunk()
                    fh.seek(0)
                    df_all = pd.read_excel(fh).fillna("")
                    if "Date" not in df_all.columns:
                        df_all = pd.DataFrame(columns=['Date', 'RowIdx', 'Description'])
                else:
                    df_all = pd.DataFrame(columns=['Date', 'RowIdx', 'Description'])

                df_all = df_all[df_all['Date'] != date_str]

                new_job_rows = []
                for idx, job_data in enumerate(saved_jobs_inputs):
                    if job_data["משימות ופעולות לביצוע"].strip():
                        new_job_rows.append({'Date': date_str, 'RowIdx': idx, 'Description': job_data["משימות ופעולות לביצוע"].strip()})

                if new_job_rows:
                    df_new = pd.DataFrame(new_job_rows)
                    df_all = pd.concat([df_all, df_new], ignore_index=True)

                out_fh = io.BytesIO()
                df_all.to_excel(out_fh, index=False)
                out_fh.seek(0)
                media = MediaIoBaseUpload(out_fh, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
                
                if file_id: 
                    drive_service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
                    st.cache_data.clear()
                    st.success("העבודות נשמרו בהצלחה!")
                    st.rerun()
                else: 
                    st.error("שגיאה 403: קובץ jobs_internal.xlsx לא נמצא.")
            except Exception as e:
                st.error(f"שגיאה בשמירה: {e}")
            
    with col_send_j:
        btn_send_j = st.button("✉️ שלח עבודות", type="primary", use_container_width=True, key="send_jobs_btn")
        
    with col_addr_j:
        target_email_j = st.selectbox("לשלוח עבודות ל:", dropdown_emails_list, key="jobs_email_target", label_visibility="collapsed")
        
    if btn_send_j:
        with st.spinner("שולח..."):
            success, msg = send_jobs_email(target_email_j, saved_jobs_inputs, date_str)
            if success: st.success(msg)
            else: st.error(msg)

# ==========================================
# ОКНО 4: НАСТРОЙКИ (הגדרות)
# ==========================================
with tab_settings:
    st.markdown("<h3>הגדרות מערכת</h3>", unsafe_allow_html=True)
    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
    
    col_set_1, col_space, col_set_2 = st.columns([10, 1, 10])
    
    with col_set_1:
        st.markdown("<div style='color: #adbac7; font-weight: bold; font-size: 16px;'>רשימת אימיילים (לרשימה נפתחת - אחד בכל שורה):</div><div style='height: 15px;'></div>", unsafe_allow_html=True)
        emails_str = "\n".join(dropdown_emails_list)
        new_emails_str = st.text_area("Drop Emails", value=emails_str, height=150, label_visibility="collapsed")
        
    with col_set_2:
        st.markdown("<div style='color: #adbac7; font-weight: bold; font-size: 16px;'>אימייל מחסן (לכפתור @ בדוח משמרת):</div><div style='height: 15px;'></div>", unsafe_allow_html=True)
        new_wh_email = st.text_input("WH Email", value=warehouse_email_target, label_visibility="collapsed")
        
    st.markdown(f'<div style="height:20px;"></div>', unsafe_allow_html=True)
    
    col_save_set, _ = st.columns([2, 8])
    with col_save_set:
        if st.button("💾 שמור הגדרות", type="primary", use_container_width=True):
            new_settings = {
                "dropdown_emails": [e.strip() for e in new_emails_str.split("\n") if e.strip()],
                "warehouse_email": new_wh_email.strip()
            }
            
            settings_json_str = json.dumps(new_settings, ensure_ascii=False, indent=4)
            fh = io.BytesIO(settings_json_str.encode('utf-8'))
            media = MediaIoBaseUpload(fh, mimetype='application/json', resumable=True)
            
            file_id = get_file_id(SETTINGS_FILE)
            try:
                if file_id:
                    drive_service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
                else:
                    file_metadata = {'name': SETTINGS_FILE, 'parents': [FOLDER_ID]}
                    drive_service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True).execute()
                
                st.cache_data.clear()
                st.success("ההגדרות נשמרו בהצלחה בענן!")
                st.rerun()
            except Exception as e:
                st.error(f"שגיאה בשמירת הגדרות: {e}")
