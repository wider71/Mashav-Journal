import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

# --- 1. НАСТРОЙКИ ПУТЕЙ И КОНСТАНТЫ ---
JOURNAL_DB = 'journal_db.csv'
OLD_LOG = 'log.csv'
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

# --- 2. ПОДКЛЮЧЕНИЕ К GOOGLE DRIVE ---
@st.cache_resource
def get_drive_service():
    creds_json = json.loads(st.secrets["GOOGLE_JSON"])
    creds = Credentials.from_service_account_info(
        creds_json,
        scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=creds)

try:
    drive_service = get_drive_service()
    FOLDER_ID = st.secrets["FOLDER_ID"]
except Exception as e:
    st.error(f"שגיאה בהתחברות לענן. בדוק את Secrets. Error: {e}")
    st.stop()

def get_file_id(filename):
    query = f"'{FOLDER_ID}' in parents and name = '{filename}' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    return items[0]['id'] if items else None


# --- 3. ГЛОБАЛЬНЫЙ CSS СТИЛЬ ---
st.markdown("""
    <style>
    .block-container { padding-top: 3rem !important; padding-bottom: 1rem !important; max-width: 98% !important; }
    .stApp { background-color: #9ba4b5; }
    * { direction: rtl !important; text-align: right !important; }
    .stTextInput input, .stTextArea textarea, .stSelectbox > div > div { direction: rtl; text-align: right; }
    
    .stTabs [data-baseweb="tab-list"] { background-color: #7a8594; border-radius: 5px; padding: 2px; margin-bottom: 0px;}
    .stTabs [data-baseweb="tab"] { font-size: 22px !important; font-weight: bold !important; color: white !important; padding: 10px 20px; }
    .stTabs [aria-selected="true"] { background-color: #2c3e50 !important; color: #fff !important; border-radius: 5px; }

    .stButton button[kind="primary"] { 
        background-color: #28a745 !important; color: white !important; 
        font-weight: bold; font-size: 18px; border: 2px solid #1e7e34 !important;
    }

    [data-testid="stDataEditor"] { font-size: 16px !important; }
    .sidur-container { overflow-x: auto; border: 3px solid black; background-color: white; }
    </style>
""", unsafe_allow_html=True)


# --- 4. ФУНКЦИИ БЭКЕНДА (ОБЛАКО G-DRIVE) ---
def migrate_old_logs():
    if os.path.exists(OLD_LOG) and get_file_id(JOURNAL_DB) is None:
        try:
            old = pd.read_csv(OLD_LOG, encoding='utf-8-sig')
            rows = []
            for _, r in old.iterrows():
                d = str(r.get('DateObj', ''))
                h = str(r.get('Hour', str(r.get('שעה', ''))))
                s = str(r.get('Shift', str(r.get('משמרת', ''))))
                s_new = 'Morning' if 'בוקר' in s or '1' in s else 'Night'
                u = str(r.get('Unit', str(r.get('יחידה / ציוד', ''))))
                desc = str(r.get('Description', str(r.get('תיאור התקלה / עבודה', ''))))
                if h or desc:
                    rows.append({'Date': d, 'Unit': u, 'Shift': s_new, 'RowIdx': 0, 'Hour': h, 'Description': desc})
            if rows:
                df = pd.DataFrame(rows)
                fh = io.BytesIO()
                df.to_csv(fh, index=False, encoding='utf-8-sig')
                fh.seek(0)
                media = MediaIoBaseUpload(fh, mimetype='text/csv', resumable=True)
                file_metadata = {'name': JOURNAL_DB, 'parents': [FOLDER_ID]}
                drive_service.files().create(body=file_metadata, media_body=media).execute()
        except: pass

@st.cache_data(ttl=30)
def get_schedule_file_drive(target_month):
    markers = MONTH_MARKERS.get(target_month, [])
    query = f"'{FOLDER_ID}' in parents and trashed = false"
    try:
        results = drive_service.files().list(q=query, fields="files(id, name)").execute()
        for f in results.get('files', []):
            name = f['name']
            if name.endswith(('.xlsx', '.xls', '.csv')):
                if any(m in name.lower() for m in markers):
                    return f['id'], name
    except Exception: pass
    return None, None

@st.cache_data(ttl=60) 
def get_operators(file_id, target_day):
    if not file_id: return ["קובץ חסר"], ["קובץ חסר"]
    try:
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done: _, done = downloader.next_chunk()
        fh.seek(0)
        
        df = pd.read_excel(fh, header=None)
        raw = df.values.tolist()
        
        cal_row_idx = None
        cleaned_cal_row = []
        for i, row in enumerate(raw[:15]):
            c_row = [str(x).split('.')[0].strip() for x in row]
            if '1' in c_row and '2' in c_row and '15' in c_row:
                cal_row_idx = i; cleaned_cal_row = c_row; break
                
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
    except Exception:
        return ["שגיאה"], []

@st.cache_data(ttl=30)
def load_journal_db():
    file_id = get_file_id(JOURNAL_DB)
    if file_id:
        try:
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: _, done = downloader.next_chunk()
            fh.seek(0)
            return pd.read_csv(fh, dtype=str).fillna("")
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

    out = pd.DataFrame(records[:6])[['Hour', 'Description']]
    out.columns = ['שעה', 'תיאור התקלה / עבודה']
    return out

def save_all_journal_grids(date_str, dfs_list):
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
    fh = io.BytesIO()
    db.to_csv(fh, index=False, encoding='utf-8-sig')
    fh.seek(0)
    media = MediaIoBaseUpload(fh, mimetype='text/csv', resumable=True)
    
    if file_id:
        drive_service.files().update(fileId=file_id, media_body=media).execute()
    else:
        file_metadata = {'name': JOURNAL_DB, 'parents': [FOLDER_ID]}
        drive_service.files().create(body=file_metadata, media_body=media).execute()
    
    st.cache_data.clear()

@st.cache_data(ttl=30)
def load_jobs_db():
    file_id = get_file_id(JOBS_FILE)
    if file_id:
        try:
            request = drive_service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: _, done = downloader.next_chunk()
            fh.seek(0)
            return pd.read_excel(fh).fillna("")
        except: pass
    return pd.DataFrame({"מספר": [i for i in range(1, 16)], "משימות ופעולות לביצוע": ["" for _ in range(15)]})

def draw_turbine_block(unit_name, section_num, date_str):
    c_morn, c_night = st.columns(2)
    
    df_m = get_journal_slice(date_str, unit_name, 'Morning')
    df_n = get_journal_slice(date_str, unit_name, 'Night')
    
    config = {
        "שעה": st.column_config.TextColumn("שעה", width="small"),
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
        ed_m = st.data_editor(df_m, key=f"m_{section_num}_{date_str}", use_container_width=True, height=230, hide_index=True, column_config=config)

    with c_night:
        st.markdown(f"""
        <div style="background-color:#2980b9; color:white; padding:4px; display:flex; border: 2px solid black; border-bottom: none; align-items:center;">
            <div style="flex:1; text-align:right; font-weight:bold; font-size:14px; padding-right:10px;">משמרת לילה</div>
            <div style="flex:2; text-align:center; font-weight:bold; font-size:16px;">{section_num}. {unit_name}</div>
            <div style="flex:1;"></div>
        </div>
        """, unsafe_allow_html=True)
        ed_n = st.data_editor(df_n, key=f"n_{section_num}_{date_str}", use_container_width=True, height=230, hide_index=True, column_config=config)

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    return (unit_name, 'Morning', ed_m), (unit_name, 'Night', ed_n)

def colorize_schedule(val):
    v = str(val).split('.')[0].strip()
    if v == '1': return 'background-color: #a9dfbf; color: black; font-weight: bold;'
    elif v == '2': return 'background-color: #abb2b9; color: black; font-weight: bold;'
    elif v in ['8', '9']: return 'background-color: #f9e79f; color: black;'
    elif v in ['ח', 'מ']: return 'background-color: #f5b7b1; color: black; font-weight: bold;'
    return ''

# Запуск миграции старых данных
migrate_old_logs()

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
        <div style="display: flex; gap: 10px; margin-top: 2px; margin-bottom: 10px;">
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
        save_all_journal_grids(date_str, grids)
        st.success("היומן נשמר בענן בהצלחה!")

# ==========================================
# ОКНО 2: РАСПИСАНИЕ (СИДУР ИЗ G-DRIVE)
# ==========================================
with tab_sch:
    st.markdown("<h3>עריכת טבלת סידור עבודה</h3>", unsafe_allow_html=True)
    if active_sch_id:
        try:
            request = drive_service.files().get_media(fileId=active_sch_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done: _, done = downloader.next_chunk()
            fh.seek(0)
            
            df_excel = pd.read_excel(fh, header=None).fillna("")
            raw_matrix = df_excel.values.tolist()
            cleaned_data = [[str(val).replace('.0', '') if val != "" else "" for val in row] for row in raw_matrix]
            df_clean = pd.DataFrame(cleaned_data)
            
            try: styled_df = df_clean.style.map(colorize_schedule)
            except AttributeError: styled_df = df_clean.style.applymap(colorize_schedule)
            
            edited_schedule = st.data_editor(styled_df, use_container_width=True, height=600, hide_index=True)
            
            if st.button("💾 שמור שינויים בענן (Google Drive)", type="primary", use_container_width=True):
                out_fh = io.BytesIO()
                edited_schedule.to_excel(out_fh, index=False, header=False)
                out_fh.seek(0)
                media = MediaIoBaseUpload(out_fh, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
                drive_service.files().update(fileId=active_sch_id, media_body=media).execute()
                
                st.cache_data.clear()
                st.success("הסידור התעדכן ונשמר בענן בהצלחה!")
                st.rerun()
                
        except Exception as e:
            st.error(f"שגיאה קריאת קובץ: {e}")
    else:
        st.warning("קובץ סידור לא נמצא ב-Google Drive בתיקייה Mashav_DB.")

# ==========================================
# ОКНО 3: РАБОТЫ НА СЕГОДНЯ
# ==========================================
with tab_jobs:
    st.markdown("<h3>עבודות מתוכננות להיום</h3>", unsafe_allow_html=True)
    
    df_jobs = load_jobs_db()
    edited_df = st.data_editor(df_jobs, num_rows="dynamic", use_container_width=True, height=520, hide_index=True)
    
    col_save, col_dl, _ = st.columns([2, 2, 6])
    with col_save:
        if st.button("שמור עבודות (בענן)", type="primary", use_container_width=True):
            file_id = get_file_id(JOBS_FILE)
            fh = io.BytesIO()
            edited_df.to_excel(fh, index=False)
            fh.seek(0)
            media = MediaIoBaseUpload(fh, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
            if file_id:
                drive_service.files().update(fileId=file_id, media_body=media).execute()
            else:
                file_metadata = {'name': JOBS_FILE, 'parents': [FOLDER_ID]}
                drive_service.files().create(body=file_metadata, media_body=media).execute()
            
            st.cache_data.clear()
            st.success("נשמר בהצלחה ב-Google Drive!")
            
    with col_dl:
        buffer = io.BytesIO()
        edited_df.to_excel(buffer, index=False)
        st.download_button("📥 שמור בשם (הורד כ-Excel)", data=buffer.getvalue(), file_name=f"Jobs_{st.session_state.log_date.strftime('%Y_%m_%d')}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
