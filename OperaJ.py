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

# --- 1. НАСТРОЙКИ ---
JOURNAL_DB = 'journal_db.csv'
JOBS_FILE = 'jobs_internal.xlsx'
LOGO_FILE = 'Mashav_Logo.png'

st.set_page_config(page_title="יומן תפעולי משאב", layout="wide")
if 'log_date' not in st.session_state: st.session_state.log_date = datetime.now().date()

# --- 2. GOOGLE DRIVE API ---
@st.cache_resource
def get_drive_service():
    creds_json = json.loads(st.secrets["GOOGLE_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=['https://www.googleapis.com/auth/drive'])
    return build('drive', 'v3', credentials=creds)

drive_service = get_drive_service()
FOLDER_ID = st.secrets["FOLDER_ID"]

def get_file_id(filename):
    query = f"'{FOLDER_ID}' in parents and name = '{filename}' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    items = results.get('files', [])
    return items[0]['id'] if items else None

# --- 3. SMTP ENGINE ---
def send_email_core(to_email, subject, html_body, attachment_df=None, attachment_name=""):
    from_email = "mashav.journal@gmail.com"
    password = st.secrets.get("EMAIL_PASS", "").replace(" ", "")
    msg = MIMEMultipart(); msg['From'] = from_email; msg['To'] = to_email; msg['Subject'] = Header(subject, 'utf-8')
    msg.attach(MIMEText(html_body, 'html', 'utf-8'))
    if attachment_df is not None:
        excel_buffer = io.BytesIO(); attachment_df.to_excel(excel_buffer, index=False); excel_buffer.seek(0)
        part = MIMEApplication(excel_buffer.read(), Name=attachment_name)
        part['Content-Disposition'] = f'attachment; filename="{attachment_name}"'; msg.attach(part)
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587); server.ehlo(); server.starttls(); server.ehlo()
        server.login(from_email, password); server.send_message(msg); server.quit(); return True, "נשלח בהצלחה!"
    except Exception as e: return False, str(e)

def send_warehouse_email(to_email, unit, shift, hour, desc, date_str):
    body = f"<h2>דיווח למחסן</h2><p><b>תאריך:</b> {date_str}</p><p><b>יחידה:</b> {unit}</p><p><b>משמרת:</b> {shift}</p><p><b>שעה:</b> {hour}</p><p><b>תיאור:</b> {desc}</p>"
    return send_email_core(to_email, "דיווח למחסן", body)

# --- 4. CSS ---
st.markdown("""
    <style>
    .block-container { padding-top: 1rem !important; max-width: 98% !important; }
    .stApp { background-color: #9ba4b5; }
    div[data-testid="stTextInput"] input { background-color: #eaf0dc !important; border: 1px solid #7f8c8d !important; height: 38px !important; font-weight: bold; }
    .header-orange { background-color: #d35400 !important; color: white; text-align: center; font-weight: bold; padding: 10px; height: 40px; }
    .header-blue { background-color: #2980b9 !important; color: white; text-align: center; font-weight: bold; padding: 10px; height: 40px; }
    button[kind="secondary"] { height: 38px !important; width: 100% !important; background-color: #27ae60 !important; color: white; font-weight: bold; }
    .stButton button[kind="primary"] { background-color: #28a745 !important; color: white; font-weight: bold; height: 38px; }
    </style>
""", unsafe_allow_html=True)

# --- 5. ЛОГИКА ---
def load_journal_db():
    file_id = get_file_id(JOURNAL_DB)
    if file_id:
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO(); downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0); return pd.read_csv(fh, dtype=str).fillna("")
    return pd.DataFrame(columns=['Date', 'Unit', 'Shift', 'RowIdx', 'Hour', 'Description'])

def get_journal_data_list(date_str, unit, shift):
    df = load_journal_db()
    sub = df[(df['Date'] == date_str) & (df['Unit'] == unit) & (df['Shift'] == shift)].copy()
    if not sub.empty: sub['RowIdx'] = pd.to_numeric(sub['RowIdx']); sub = sub.sort_values('RowIdx')
    raw = sub.to_dict('records')
    while len(raw) < 6: raw.append({'Hour': '', 'Description': ''})
    return raw[:6]

# --- ИНТЕРФЕЙС ---
tab_log, tab_sch, tab_jobs = st.tabs(["דוח משמרת", "סידור", "עבודות היום"])
with tab_log:
    date_str = st.session_state.log_date.strftime("%Y-%m-%d")
    units = [('טורבינה 1', 1), ('טורבינה 2', 2), ('טורבינה קיטורית', 3)]
    saved_inputs = {}
    
    # Сначала Оранжевый блок (Утро) - СЛЕВА
    # Затем Синий блок (Ночь) - СПРАВА
    c_morn, c_night = st.columns(2)
    
    with c_morn:
        st.markdown('<div class="header-orange">משמרת בוקר</div>', unsafe_allow_html=True)
        for u_name, u_num in units:
            for idx in range(2): # Для примера по 2 строки на турбину
                c_d, c_h, c_b = st.columns([10, 2, 1])
                with c_b:
                    if st.button("@", key=f"bm_{u_name}_{idx}"):
                        st.toast("נשלח למחסן!", icon="✅")
                with c_h: h = st.text_input(f"h_{u_name}_{idx}_m", key=f"h_m_{u_name}_{idx}")
                with c_d: d = st.text_input(f"d_{u_name}_{idx}_m", key=f"d_m_{u_name}_{idx}")
                saved_inputs[(u_name, 'Morning', idx)] = (h, d)

    with c_night:
        st.markdown('<div class="header-blue">משמרת לילה</div>', unsafe_allow_html=True)
        for u_name, u_num in units:
            for idx in range(2):
                c_b, c_h, c_d = st.columns([1, 2, 10])
                with c_b:
                    if st.button("@", key=f"bn_{u_name}_{idx}"):
                        st.toast("נשלח למחסן!", icon="✅")
                with c_h: h = st.text_input(f"h_{u_name}_{idx}_n", key=f"h_n_{u_name}_{idx}")
                with c_d: d = st.text_input(f"d_{u_name}_{idx}_n", key=f"d_n_{u_name}_{idx}")
                saved_inputs[(u_name, 'Night', idx)] = (h, d)
