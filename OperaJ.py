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
def send_jobs_email(to_email, df_jobs, date_str):
    from_email = "mashav.journal@gmail.com"
    password = st.secrets.get("EMAIL_PASS", "").replace(" ", "")
    if not password:
        return False, "Пароль приложений (EMAIL_PASS) не найден в Secrets!"
    
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = Header(f"עבודות מתוכננות להיום - {date_str}", 'utf-8')
    
    html_table = df_jobs.to_html(index=False, justify='right')
    html_body = f"""
    <html dir="rtl">
    <body style="font-family: Arial, sans-serif; text-align: right; direction: rtl;">
        <h2>עבודות מתוכננות להיום ({date_str})</h2>
        {html_table}
    </body>
    </html>
    """
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    
    excel_buffer = io.BytesIO()
    df_jobs.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)
    part = MIMEApplication(excel_buffer.read(), Name=f"Jobs_{date_str}.xlsx")
    part['Content-Disposition'] = f'attachment; filename="Jobs_{date_str}.xlsx"'
    msg.attach(part)
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(from_email, password)
        server.send_message(msg)
        server.quit()
        return True, "הודעה נשלחה בהצלחה למייל!"
    except Exception as e:
        return False, f"שגיאה בשרת הדואר: {str(e)}"

# --- 3. ГЛОБАЛЬНЫЙ CSS СТИЛЬ ---
st.markdown("""
    <style>
    .block-container { padding-top: 3.5rem !important; padding-bottom: 1rem !important; max-width: 98% !important; }
    .stApp { background-color: #9ba4b5; }
    * { direction: rtl !important; text-align: right !important; }
    
    .stTabs [data-baseweb="tab-list"] { background-color: #7a8594; border-radius: 5px; padding: 2px; margin-bottom: 0px;}
    .stTabs [data-baseweb="tab"] { font-size: 22px !important; font-weight: bold !important; color: white !important; padding: 10px 20px; }
    .stTabs [aria-selected="true"] { background-color: #2c3e50 !important; color: #fff !important; border-radius: 5px; }

    .stButton button[kind="primary"] { 
        background-color: #28a745 !important; color: white !important; 
        font-weight: bold; font-size: 18px; border: 2px solid #1e7e34 !important;
    }

    [data-testid="stDataEditor"] { font-size: 16px !important; font-weight: bold !important; direction: rtl !important; }
    [data-testid="stDataEditor"] input, [data-testid="stDataEditor"] textarea { direction: rtl !important; text-align: right !important; }
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

def get_journal_slice(date_str, unit, shift):
    df = load_journal_db()
    sub = df[(df['Date'] == date_str) & (df['Unit'] == unit) & (df['Shift'] == shift)].copy()
    if not sub.empty:
        sub['RowIdx'] = pd.to_numeric(sub['RowIdx'])
        sub = sub.sort_values('RowIdx')
    
    records = sub.to_dict('records')
    while len(records) < 6:
        records.append({'Hour': '', 'Description': ''})

    out = pd.DataFrame(records[:6])
    out = out[['Description', 'Hour']]
    out.columns = ['תיאור התקלה / עבודה', 'שעה']
    return out

def save_all_journal_grids(date_str, dfs_list):
    try:
        db = load_journal_db()
        db = db[db['Date'] != date_str] 
        rows = []
        for unit, shift, df in dfs_list:
            for idx, r in df.iterrows():
                h = str(r['שעה']).strip()
                d = str(r['תיאור התקלה / עבודה']).strip()
                if h or d:
                    rows.append([date_str, unit, shift, idx, h, d])

        new_df = pd.DataFrame(rows, columns=['Date', 'Unit', 'Shift', 'RowIdx', 'Hour', 'Description'])
        db = pd.concat([db, new_df])
        
        file_id = get_file_id(JOURNAL_DB)
        fh = io.BytesIO(); db.to_csv(fh, index=False, encoding='utf-8-sig'); fh.seek(0)
        media = MediaIoBaseUpload(fh, mimetype='text/csv', resumable=True)
        
        if file_id:
            drive_service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
        else:
            file_metadata = {'name': JOURNAL_DB, 'parents': [FOLDER_ID]}
            drive_service.files().create(body=file_metadata, media_body=media, supportsAllDrives=True).execute()
        
        st.cache_data.clear()
        return True, "היומן נשמר בענן בהצלחה!"
    except Exception as e:
        return False, str(e)

@st.cache_data(ttl=5)
def load_jobs_db():
    file_id = get_file_id(JOBS_FILE)
    if file_id:
        try:
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: _, done = downloader.next_chunk()
            fh.seek(0)
            df_jobs = pd.read_excel(fh).fillna("")
            df_jobs.columns = df_jobs.columns.astype(str)
            if "מספר" not in df_jobs.columns: df_jobs["מספר"] = [str(i) for i in range(1, 16)]
            if "משימות ופעולות לביצוע" not in df_jobs.columns: df_jobs["משימות ופעולות לביצוע"] = ["" for _ in range(15)]
            return df_jobs
        except: pass
    return pd.DataFrame({"מספר": [str(i) for i in range(1, 16)], "משימות ופעולות לביצוע": ["" for _ in range(15)]})

def draw_turbine_block(unit_name, section_num, date_str):
    c_morn, c_night = st.columns(2)
    
    df_m = get_journal_slice(date_str, unit_name, 'Morning')
    df_n = get_journal_slice(date_str, unit_name, 'Night')
    
    # Возвращаем рабочий и проверенный код стилей: только фон, цвет текста и жесткое выравнивание вправо
    styled_m = df_m.style.set_properties(**{'background-color': '#fdf5e6', 'color': 'black', 'text-align': 'right'})
    styled_n = df_n.style.set_properties(**{'background-color': '#fdf5e6', 'color': 'black', 'text-align': 'right'})
    
    # "small" - это минимально допустимая алгоритмом ширина для колонки
    config = {
        "שעה": st.column_config.TextColumn("שעה", width="small", max_chars=5),
        "תיאור התקלה / עבודה": st.column_config.TextColumn("תיאור התקלה / עבודה", width="large")
    }

    with c_morn:
        st.markdown(f"""
        <div style="background-color:#d35400; color:white; padding:4px; display:flex; border: 2px solid black; border-bottom: none; align-items:center;">
            <div style="flex:1; text-align:right; font-weight:bold; font-size:14px; padding-right:10px;">משמרת בוקר</div>
            <div style="flex:2; text-align:center; font-weight:bold; font-size:16px;">{section_num}. {unit_name}</div>
            <div style="flex:1;"></div>
        </div>
        """, unsafe_allow_html=True)
        ed_m = st.data_editor(styled_m, key=f"m_{section_num}_{date_str}", use_container_width=True, height=210, hide_index=True, column_config=config)

    with c_night:
        st.markdown(f"""
        <div style="background-color:#2980b9; color:white; padding:4px; display:flex; border: 2px solid black; border-bottom: none; align-items:center;">
            <div style="flex:1; text-align:right; font-weight:bold; font-size:14px; padding-right:10px;">משמרת לילה</div>
            <div style="flex:2; text-align:center; font-weight:bold; font-size:16px;">{section_num}. {unit_name}</div>
            <div style="flex:1;"></div>
        </div>
        """, unsafe_allow_html=True)
        ed_n = st.data_editor(styled_n, key=f"n_{section_num}_{date_str}", use_container_width=True, height=210, hide_index=True, column_config=config)

    return (unit_name, 'Morning', ed_m), (unit_name, 'Night', ed_n)

def colorize_schedule(val):
    v = str(val).split('.')[0].strip()
    if v == '1': return 'background-color: #a9dfbf; color: black; font-weight: bold; font-size: 16px;'
    elif v == '2': return 'background-color: #abb2b9; color: black; font-weight: bold; font-size: 16px;'
    elif v in ['8', '9']: return 'background-color: #f9e79f; color: black; font-size: 16px;'
    elif v in ['ח', 'מ']: return 'background-color: #f5b7b1; color: black; font-weight: bold; font-size: 16px;'
    return 'color: black; font-size: 16px;'

# --- 5. ВКЛАДКИ ОКОН ---
tab_log, tab_sch, tab_jobs = st.tabs(["דוח משמרת", "סידור", "עבודות היום"])

# ==========================================
# ОКНО 1: ОПЕРАТИВНЫЙ ЖУРНАЛ
# ==========================================
with tab_log:
    col_logo, col_title, col_cal_r, col_cal_m, col_cal_l = st.columns([1, 4, 1.2, 1.8, 1.2])
    with col_logo:
        if os.path.exists(LOGO_FILE): st.image(LOGO_FILE, width=120)
    with col_title:
        st.markdown("<h4 style='color: #2c3e50; margin-top: 2px; font-weight: bold; font-size: 20px;'>דוח משמרת תחנת כוח משאב</h4>", unsafe_allow_html=True)
    with col_cal_r:
        if st.button("▶ יום הבא", type="primary", use_container_width=True): st.session_state.log_date += timedelta(days=1); st.rerun()
    with col_cal_m:
        new_date = st.date_input("תאריך", value=st.session_state.log_date, label_visibility="collapsed")
        if new_date != st.session_state.log_date: st.session_state.log_date = new_date; st.rerun()
    with col_cal_l:
        if st.button("יום קודם ◀", type="primary", use_container_width=True): st.session_state.log_date -= timedelta(days=1); st.rerun()

    active_sch_id, sch_name = get_schedule_file_drive(st.session_state.log_date.month)
    s1_names, s2_names = get_operators(active_sch_id, st.session_state.log_date.day)

    st.markdown(f"""
        <div style="display: flex; gap: 10px; margin-top: 2px; margin-bottom: 5px;">
            <div style="flex: 1; border: 2px solid #2c3e50; padding: 4px 12px; background-color: #f8f9fa; border-radius: 4px; display: flex; justify-content: space-between; align-items: center;">
                <b style="color: #333; font-size: 14px;">🌞 משמרת בוקר:</b> <span style="font-size: 16px; font-weight: bold; color: black;">{', '.join(s1_names)}</span>
            </div>
            <div style="flex: 1; border: 2px solid black; padding: 4px 12px; background-color: #343a40; border-radius: 4px; display: flex; justify-content: space-between; align-items: center;">
                <b style="color: #ddd; font-size: 14px;">🌙 משמרת לילה:</b> <span style="font-size: 16px; font-weight: bold; color: white;">{', '.join(s2_names)}</span>
            </div>
        </div>
    """, unsafe_allow_html=True)

    date_str = st.session_state.log_date.strftime("%Y-%m-%d")
    
    grids = []
    grids.extend(draw_turbine_block('טורבינה 1', 1, date_str))
    grids.extend(draw_turbine_block('טורבינה 2', 2, date_str))
    grids.extend(draw_turbine_block('טורבינה קיטורית', 3, date_str))

    if st.button("💾 שמור כל השינויים ביומן", type="primary", use_container_width=True):
        success, msg = save_all_journal_grids(date_str, grids)
        if success: st.success(msg)
        else: st.error(f"שגיאה: {msg}")


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
            
            styled_df = df_ui.style.map(colorize_schedule).set_properties(**{'text-align': 'center', 'font-weight': 'bold'})
            st.dataframe(styled_df, use_container_width=True, height=550)
        except Exception as e: st.error(f"שגיאה: {e}")


# ==========================================
# ОКНО 3: РАБОТЫ НА СЕГОДНЯ
# ==========================================
with tab_jobs:
    st.markdown("<h3>עבודות מתוכננות להיום</h3>", unsafe_allow_html=True)
    
    df_jobs = load_jobs_db()
    df_ui_jobs = df_jobs[["משימות ופעולות לביצוע", "מספר"]]
    
    styled_jobs = df_ui_jobs.style.set_properties(**{'background-color': '#ebf5fb', 'color': 'black', 'text-align': 'right'})
    
    # "small" ширина и запрет на изменение текста номера
    config_jobs = {
        "מספר": st.column_config.TextColumn("מספר", width="small", disabled=True),
        "משימות ופעולות לביצוע": st.column_config.TextColumn("משימות ופעולות לביצוע", width="large")
    }
    
    edited_df = st.data_editor(styled_jobs, num_rows="dynamic", use_container_width=True, height=550, hide_index=True, column_config=config_jobs)
    
    col_save, col_send, col_addr = st.columns([2, 2, 6])
    
    with col_save:
        if st.button("💾 שמור עבודות (בענן)", type="primary", use_container_width=True):
            try:
                df_save_jobs = edited_df[["מספר", "משימות ופעולות לביצוע"]]
                
                file_id = get_file_id(JOBS_FILE)
                fh = io.BytesIO(); df_save_jobs.to_excel(fh, index=False); fh.seek(0)
                media = MediaIoBaseUpload(fh, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
                if file_id: drive_service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
                else: drive_service.files().create(body={'name': JOBS_FILE, 'parents': [FOLDER_ID]}, media_body=media, supportsAllDrives=True).execute()
                
                st.cache_data.clear()
                st.success("נשמר בהצלחה ב-Google Drive!")
            except Exception as e: st.error(f"שגיאה בשמירת עבודות: {e}")
            
    with col_send:
        btn_send = st.button("✉️ שלח למייל", use_container_width=True)
        
    with col_addr:
        target_email = st.selectbox("לשלוח עבודות ל:", ["wider71@gmail.com"], label_visibility="collapsed")
        
    if btn_send:
        with st.spinner("שולח..."):
            date_str_jobs = st.session_state.log_date.strftime("%Y-%m-%d")
            df_save_jobs = edited_df[["מספר", "משימות ופעולות לביצוע"]]
            success, msg = send_jobs_email(target_email, df_save_jobs, date_str_jobs)
            if success: st.success(msg)
            else: st.error(msg)
