import streamlit as st
import pandas as pd
import numpy as np
from io import StringIO, BytesIO
import json
import os
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="Reconciliation Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# STYLES
# ============================================================
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F3864, #2E75B6);
        color: white;
        padding: 20px 30px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .metric-card {
        background: white;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        border-left: 4px solid;
        margin-bottom: 10px;
    }
    .metric-green  { border-left-color: #375623; }
    .metric-red    { border-left-color: #C00000; }
    .metric-orange { border-left-color: #C55A11; }
    .metric-purple { border-left-color: #4A148C; }
    .metric-teal   { border-left-color: #00695C; }
    .status-ok     { color: #375623; font-weight: bold; }
    .status-warn   { color: #C55A11; font-weight: bold; }
    .status-err    { color: #C00000; font-weight: bold; }
    .stDataFrame   { font-size: 12px; }
    div[data-testid="stMetricValue"] { font-size: 28px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

HISTORY_COLS = [
    'date', 'sup_cbd', 'our_eup', 'diff',
    'matched_count', 'sup_only_count', 'sup_only_cbd',
    'our_only_count', 'our_only_eup', 'real_gap',
    'pending_count', 'refunds_eup', 'net_billed'
]

# Detail columns saved per transaction for cross-day matching
DETAIL_COLS = [
    'date', 'category', 'phone', 'operator',
    'product', 'amount', 'supplier_date', 'our_date', 'reason'
]

@st.cache_resource
def get_gspread_client():
    if not GSPREAD_AVAILABLE:
        return None
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception:
        return None

def get_spreadsheet():
    gc = get_gspread_client()
    if gc is None:
        return None
    try:
        spreadsheet_id = st.secrets["google_sheets"]["spreadsheet_id"]
        return gc.open_by_key(spreadsheet_id)
    except Exception:
        return None

def get_or_create_sheet(sh, title, headers):
    """Get existing sheet or create new one with headers"""
    try:
        ws = sh.worksheet(title)
        return ws
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=len(headers))
        ws.append_row(headers)
        return ws

def load_history(month=None):
    """Load history from Google Sheets. If month given (YYYY-MM), load that month's sheet."""
    sh = get_spreadsheet()
    if sh is None:
        return _load_local_history()
    try:
        if month:
            sheet_title = datetime.strptime(month, '%Y-%m').strftime('%B %Y')
        else:
            # Load all months
            all_records = []
            for ws in sh.worksheets():
                if _is_month_sheet(ws.title):
                    try:
                        records = ws.get_all_records()
                        all_records.extend(records)
                    except Exception:
                        pass
            return sorted(all_records, key=lambda x: x.get('date', ''))

        ws = get_or_create_sheet(sh, sheet_title, HISTORY_COLS)
        return ws.get_all_records()
    except Exception:
        return _load_local_history()

def _is_month_sheet(title):
    months = ['January','February','March','April','May','June',
              'July','August','September','October','November','December']
    return any(m in title for m in months)

def save_to_sheets(record):
    """Save daily summary to the correct month sheet in Google Sheets"""
    sh = get_spreadsheet()
    if sh is None:
        _save_local_history(record)
        return False, "Google Sheets not connected — saved locally"
    try:
        # Month sheet name e.g. "May 2026"
        dt = datetime.strptime(record['date'], '%Y-%m-%d')
        sheet_title = dt.strftime('%B %Y')
        ws = get_or_create_sheet(sh, sheet_title, HISTORY_COLS)

        # Check if date already exists — update if so
        existing = ws.get_all_records()
        for i, row in enumerate(existing):
            if row.get('date') == record['date']:
                row_num = i + 2  # +1 header, +1 1-indexed
                ws.update(f'A{row_num}', [[record.get(c, '') for c in HISTORY_COLS]])
                return True, f"Updated existing record for {record['date']} in '{sheet_title}'"

        # Append new row
        ws.append_row([record.get(c, '') for c in HISTORY_COLS])
        return True, f"Saved to '{sheet_title}' sheet"
    except Exception as e:
        _save_local_history(record)
        return False, f"Sheets error: {e} — saved locally"

def save_details_to_sheets(report_date, result):
    """Save phone-level details for cross-day matching"""
    sh = get_spreadsheet()
    if sh is None:
        return False, "Not connected"
    try:
        ws = get_or_create_sheet(sh, 'Transaction Details', DETAIL_COLS)

        rows = []
        # Supplier only
        for _, r in result['sup_only'].iterrows():
            rows.append([report_date, 'Supplier Only', r.get('Phone_Display',''),
                        '', r.get('Package',''), r.get('CBD',0),
                        r.get('Sup_Date',''), '', r.get('Reason','')])
        # Our only
        for _, r in result['our_only'].iterrows():
            rows.append([report_date, 'Our Only', r.get('Phone_Display',''),
                        r.get('Operator',''), r.get('Product Name',''),
                        r.get('End User Price',0), '', r.get('Date & Time',''),
                        r.get('Reason','')])

        # Always update — remove existing rows for this date then re-add
        all_data = ws.get_all_records()
        keep = [r for r in all_data if str(r.get('date','')) != str(report_date)]
        ws.clear()
        ws.append_row(DETAIL_COLS)
        if keep:
            ws.append_rows([[r.get(c,'') for c in DETAIL_COLS] for r in keep])
        if rows:
            ws.append_rows(rows)
        return True, f"Saved {len(rows)} detail rows for {report_date}"
    except Exception as e:
        return False, f"Details save error: {e}"

def cross_day_match(result, report_date):
    """Find phones from 'Our Only' today that were 'Supplier Only' yesterday, or vice versa"""
    sh = get_spreadsheet()
    if sh is None:
        return [], []
    try:
        ws = sh.worksheet('Transaction Details')
        all_details = ws.get_all_records()
        df = pd.DataFrame(all_details)
        if df.empty:
            return [], []

        # Yesterday's supplier only phones
        dt = datetime.strptime(report_date, '%Y-%m-%d')
        yesterday = (dt - timedelta(days=1)).strftime('%Y-%m-%d')

        yest_sup_only = set(df[(df['date']==yesterday) & (df['category']=='Supplier Only')]['phone'])

        our_only_phones = set(result['our_only']['Phone_Display']) if len(result['our_only']) > 0 else set()
        sup_only_phones = set(result['sup_only']['Phone_Display']) if len(result['sup_only']) > 0 else set()

        # Our Only today that appeared as Supplier Only yesterday = date shift confirmed
        confirmed_shifts_our  = list(our_only_phones & yest_sup_only)
        # Supplier Only today that appeared as Our Only yesterday = date shift confirmed
        confirmed_shifts_sup  = list(sup_only_phones & set(
            df[(df['date']==yesterday) & (df['category']=='Our Only')]['phone']
        ))

        return confirmed_shifts_our, confirmed_shifts_sup
    except Exception:
        return [], []

# ---- Local fallback (when Sheets not available) ----
HISTORY_FILE = os.path.join(os.path.dirname(__file__), "history.json")

def _load_local_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def _save_local_history(record):
    history = _load_local_history()
    history = [h for h in history if h.get('date') != record.get('date')]
    history.append(record)
    history.sort(key=lambda x: x.get('date', ''))
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

def add_to_history(record):
    """Legacy wrapper — now saves to Sheets"""
    save_to_sheets(record)

# ============================================================
# PHONE NORMALIZATION
# ============================================================
def norm_phone(phone):
    if phone is None or (isinstance(phone, float) and np.isnan(phone)):
        return ''
    s = str(phone).strip().replace('.0', '').replace(' ', '').replace('+', '')
    if 'E' in s.upper():
        try:
            s = str(int(float(s)))
        except:
            pass
    s = s.replace('.0', '')
    if s.startswith('00972'):
        s = s[5:]
    elif s.startswith('972'):
        s = s[3:]
    if s.startswith('0'):
        s = s[1:]
    return s.strip()

def display_phone(norm):
    """Convert normalized phone back to 05X display format"""
    if norm and len(norm) >= 8:
        return '0' + norm
    return norm

# ============================================================
# LOAD FILES
# ============================================================
def load_supplier(file_bytes, filename):
    try:
        if filename.lower().endswith('.xls'):
            text = file_bytes.decode('cp1255', errors='replace')
        else:
            for enc in ['utf-8-sig', 'windows-1255', 'cp1255', 'latin1']:
                try:
                    text = file_bytes.decode(enc)
                    break
                except:
                    continue

        lines = text.replace('\r\n', '\n').replace('\r', '\n').strip().split('\n')
        start = 0
        first_line_cols = [c for c in lines[0].split('\t') if c.strip()]
        if len(first_line_cols) == 1:
            start = 1

        sep = '\t' if lines[start].count('\t') > lines[start].count(',') else ','
        df = pd.read_csv(StringIO('\n'.join(lines[start:])), sep=sep,
                         dtype={'MSISDN': str}, on_bad_lines='skip')

        col_map = {
            'שם לקוח': 'Cust_Name', 'מספר לקוח': 'Cust_Num',
            'מס טרנזקציה': 'Sup_TxID', 'תאריך ושעת טעינה': 'Sup_Date',
            'MSISDN': 'MSISDN', 'חיוב לפני הנחה כולל מעמ': 'CBD',
            'נטו כולל מעמ': 'Net_Total', 'שם כרטיס': 'Package',
            'סוג כרטיס': 'Card_Type', 'אחוז הנחה': 'Discount_Pct',
            'סכום הנחה כולל מעמ': 'Discount_Amt'
        }
        df.rename(columns=col_map, inplace=True)
        df['CBD'] = pd.to_numeric(df.get('CBD', 0), errors='coerce').fillna(0)
        df['Net_Total'] = pd.to_numeric(df.get('Net_Total', 0), errors='coerce').fillna(0)
        df['phone_norm'] = df['MSISDN'].apply(norm_phone)
        df['Phone_Display'] = df['phone_norm'].apply(display_phone)
        df = df[df['phone_norm'].str.len() >= 8]
        return df, None
    except Exception as e:
        return None, str(e)

def load_our(file_bytes, operator_name):
    try:
        for enc in ['utf-8-sig', 'utf-8', 'windows-1255', 'cp1255', 'latin1']:
            try:
                text = file_bytes.decode(enc)
                break
            except:
                continue

        df = pd.read_csv(StringIO(text), dtype={'Phone Number': str}, on_bad_lines='skip')
        df['Operator'] = operator_name
        df['End User Price'] = pd.to_numeric(df.get('End User Price', 0), errors='coerce').fillna(0)
        df['Customer price'] = pd.to_numeric(df.get('Customer price', 0), errors='coerce').fillna(0)

        df = df[~df['Action'].isin(['REWARD', 'REFUND_REWARD'])]

        df['Is_Refund'] = df['Action'] == 'REFUND'
        df['Eff_Status'] = df.apply(
            lambda r: 'CANCELLED' if r['Action'] == 'REFUND' else r['Status'], axis=1)
        df['phone_norm'] = df['Phone Number'].apply(norm_phone)
        df['Phone_Display'] = df['phone_norm'].apply(display_phone)

        # Detect late transactions (after 22:00 — likely date shift)
        try:
            df['_dt'] = pd.to_datetime(df['Date & Time'], dayfirst=True, errors='coerce')
            df['Is_Late'] = df['_dt'].dt.hour >= 22
        except:
            df['Is_Late'] = False

        df = df[df['phone_norm'].str.len() >= 7]
        return df, None
    except Exception as e:
        return None, str(e)

# ============================================================
# RECONCILIATION LOGIC
# ============================================================
def run_reconciliation(sup_df, partner_df, talk_df):
    our_all = pd.concat([partner_df, talk_df], ignore_index=True)

    our_dc      = our_all[(our_all['Eff_Status'].isin(['DONE', 'CANCELLED'])) & (~our_all['Is_Refund'])].copy()
    our_pending = our_all[our_all['Eff_Status'] == 'PENDING_CANCELLATION'].copy()
    our_refunds = our_all[our_all['Is_Refund']].copy()
    our_failed  = our_all[our_all['Eff_Status'] == 'FAILED'].copy()

    sup_phones     = set(sup_df['phone_norm'])
    our_dc_phones  = set(our_dc['phone_norm'])
    our_pnd_phones = set(our_pending['phone_norm'])

    matched_phones  = sup_phones & our_dc_phones
    sup_only_phones = sup_phones - our_dc_phones - our_pnd_phones
    sup_pending_phones = sup_phones & our_pnd_phones   # supplier has it, we have PENDING
    our_only_phones = our_dc_phones - sup_phones

    # ---- MATCHED detail ----
    matched_rows = []
    used_our = set()
    for _, sup_row in sup_df[sup_df['phone_norm'].isin(matched_phones)].iterrows():
        our_match = our_dc[
            (our_dc['phone_norm'] == sup_row['phone_norm']) &
            (~our_dc['Transaction ID'].isin(used_our))
        ]
        if len(our_match) > 0:
            our_row = our_match.iloc[0]
            used_our.add(our_row['Transaction ID'])
            diff = sup_row['CBD'] - our_row['End User Price']
            matched_rows.append({
                'Phone':              our_row['Phone_Display'],
                'Supplier Date':      sup_row.get('Sup_Date', ''),
                'Supplier Package':   sup_row.get('Package', ''),
                'Supplier Tx ID':     sup_row.get('Sup_TxID', ''),
                'Supplier CBD (NIS)': sup_row['CBD'],
                'Our Tx ID':          our_row['Transaction ID'],
                'Our Date':           our_row['Date & Time'],
                'Our Operator':       our_row['Operator'],
                'Our Status':         our_row['Eff_Status'],
                'Our Product':        our_row['Product Name'],
                'Our EUP (NIS)':      our_row['End User Price'],
                'Difference (NIS)':   diff,
            })

    matched_df = pd.DataFrame(matched_rows)

    # ---- SUPPLIER ONLY detail ----
    sup_only_df = sup_df[sup_df['phone_norm'].isin(sup_only_phones)].copy()
    sup_only_df['Reason'] = 'Not found in our system'

    # ---- SUPPLIER vs PENDING ----
    sup_pending_df = sup_df[sup_df['phone_norm'].isin(sup_pending_phones)].copy()

    # ---- OUR ONLY detail ----
    our_only_df = our_dc[our_dc['phone_norm'].isin(our_only_phones)].copy()
    our_only_df['Reason'] = our_only_df.apply(
        lambda r: 'Late transaction (after 22:00)' if r.get('Is_Late', False) else 'Not found at supplier',
        axis=1
    )

    # ---- TOTALS ----
    totals = {
        'sup_cbd':         matched_df['Supplier CBD (NIS)'].sum() if len(matched_df) else 0,
        'our_eup':         matched_df['Our EUP (NIS)'].sum() if len(matched_df) else 0,
        'diff':            matched_df['Difference (NIS)'].sum() if len(matched_df) else 0,
        'sup_only_cbd':    sup_only_df['CBD'].sum() if len(sup_only_df) else 0,
        'sup_pending_cbd': sup_pending_df['CBD'].sum() if len(sup_pending_df) else 0,
        'our_only_eup':    our_only_df['End User Price'].sum() if len(our_only_df) else 0,
        'pending_eup':     our_pending['End User Price'].sum() if len(our_pending) else 0,
        'refunds_eup':     our_refunds['End User Price'].sum() if len(our_refunds) else 0,
        'partner_eup':     partner_df[partner_df['Eff_Status'].isin(['DONE', 'CANCELLED']) & ~partner_df['Is_Refund']]['End User Price'].sum(),
        'talk012_eup':     talk_df[talk_df['Eff_Status'].isin(['DONE', 'CANCELLED']) & ~talk_df['Is_Refund']]['End User Price'].sum(),
        'partner_ref':     partner_df[partner_df['Is_Refund']]['End User Price'].sum(),
        'talk012_ref':     talk_df[talk_df['Is_Refund']]['End User Price'].sum(),
        'matched_count':        len(matched_phones),
        'sup_only_count':       len(sup_only_phones),
        'sup_pending_count':    len(sup_pending_phones),
        'our_only_count':       len(our_only_phones),
        'pending_count':        len(our_pending),
        'refunds_count':        len(our_refunds),
        'failed_count':         len(our_failed),
    }
    # Real gap: supplier charges for phones we don't see vs we charged for phones supplier doesn't see
    totals['real_gap'] = round(totals['sup_only_cbd'] - totals['our_only_eup'], 2)

    return {
        'matched':      matched_df,
        'sup_only':     sup_only_df,
        'sup_pending':  sup_pending_df,
        'our_only':     our_only_df,
        'pending':      our_pending,
        'refunds':      our_refunds,
        'failed':       our_failed,
        'totals':       totals,
    }

# ============================================================
# EXCEL EXPORT
# ============================================================
def create_excel_report(result, report_date):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def bd():
        s = Side(style='thin', color='CCCCCC')
        return Border(left=s, right=s, top=s, bottom=s)

    def H(cell, bg, fg='FFFFFF', sz=10, bold=True):
        cell.font = Font(bold=bold, color=fg, size=sz, name='Arial')
        cell.fill = PatternFill('solid', start_color=bg)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = bd()

    def D(cell, bg='FFFFFF', align='left', fg='000000', fmt=None, bold=False):
        cell.font = Font(color=fg, size=9, name='Arial', bold=bold)
        cell.fill = PatternFill('solid', start_color=bg)
        cell.alignment = Alignment(horizontal=align, vertical='center')
        cell.border = bd()
        if fmt:
            cell.number_format = fmt

    NAVY = '1F3864'; BLUE = '2E75B6'; LBLUE = 'DEEAF1'
    ORANGE = 'C55A11'; LORAN = 'FCE4D6'
    GREEN = '375623'; LGREEN = 'E2EFDA'
    RED = 'C00000'; LRED = 'FFE0E0'
    YELL = 'FFE699'; LYELL = 'FFFACD'
    TEAL = '00695C'; LTEAL = 'E0F2F1'
    PURPLE = '4A148C'; LPURP = 'EDE7F6'
    WHITE = 'FFFFFF'

    t = result['totals']

    def ttl(ws, txt, ncols, bg, sz=12):
        ws.merge_cells(f'A1:{get_column_letter(ncols)}1')
        ws['A1'].value = txt
        ws['A1'].font = Font(bold=True, color='FFFFFF', size=sz, name='Arial')
        ws['A1'].fill = PatternFill('solid', start_color=bg)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws['A1'].border = bd()
        ws.row_dimensions[1].height = 30

    def sec(ws, r, txt, ncols, bg):
        ws.merge_cells(f'A{r}:{get_column_letter(ncols)}{r}')
        ws[f'A{r}'].value = txt
        H(ws[f'A{r}'], bg, sz=11)
        ws.row_dimensions[r].height = 22

    def write_df(ws, df, start_row, hdr_bg, alt_bg, num_cols=None):
        if len(df) == 0:
            return start_row
        for ci, col in enumerate(df.columns, 1):
            H(ws.cell(row=start_row, column=ci, value=col), hdr_bg, sz=9)
        ws.row_dimensions[start_row].height = 35
        for ri, (_, row) in enumerate(df.iterrows()):
            r = start_row + 1 + ri
            bg = alt_bg if ri % 2 == 0 else WHITE
            for ci, col in enumerate(df.columns, 1):
                val = row[col]
                if pd.isna(val): val = ''
                cell = ws.cell(row=r, column=ci, value=val)
                is_num = num_cols and col in num_cols
                fmt = '#,##0.00' if is_num else None
                D(cell, bg, align='right' if is_num else ('center' if ci > 1 else 'left'), fmt=fmt)
            ws.row_dimensions[r].height = 15
        return start_row + 1 + len(df)

    # ---- SHEET 1: SUMMARY ----
    wss = wb.create_sheet("Summary")
    ttl(wss, f"📊  RECONCILIATION SUMMARY  —  {report_date}", 7, NAVY, 14)

    sec(wss, 3, "A.  TRANSACTION COUNT DISCREPANCY", 7, BLUE)
    for ci, h in enumerate(["Category", "# Phones / Tx", "Amount (NIS)", "Description", "Action Required", "", ""], 1):
        H(wss.cell(row=4, column=ci, value=h), NAVY)
    wss.row_dimensions[4].height = 28

    gap_text  = f"Supplier charges MORE by {t['real_gap']:,.2f} NIS" if t['real_gap'] > 0 else \
                (f"We charge MORE by {abs(t['real_gap']):,.2f} NIS" if t['real_gap'] < 0 else "Balanced")

    rows_a = [
        ("✅  MATCHED",              t['matched_count'],        round(t['sup_cbd'], 2),        "Phone in BOTH — DONE/CANCELLED",                         "—",                              LGREEN),
        ("❌  SUPPLIER ONLY",        t['sup_only_count'],        round(t['sup_only_cbd'], 2),   "Supplier charges — NOT in our system",                   "Investigate each phone",         LRED),
        ("⚠️  OUR SYSTEM ONLY",      t['our_only_count'],        round(t['our_only_eup'], 2),   "We charged — supplier does NOT list",                    "Check if date shift",            LYELL),
        ("🔄  SUP vs PENDING",       t['sup_pending_count'],     round(t['sup_pending_cbd'], 2),"Supplier charged — we have PENDING",                     "Verify cancellation next day",   LPURP),
        ("🕐  PENDING (ours)",        t['pending_count'],         round(t['pending_eup'], 2),    "Awaiting supplier decision",                             "Check next day",                 LPURP),
        ("↩️  REFUNDS",              t['refunds_count'],         round(t['refunds_eup'], 2),    "Credits from previous period",                           "Deduct from month total",        LTEAL),
        ("📊  REAL GAP",             t['sup_only_count'] - t['our_only_count'], round(t['real_gap'], 2), gap_text, "= Supplier Only CBD − Our Only EUP", LORAN),
    ]
    for i, (cat, cnt, amt, desc, action, bg) in enumerate(rows_a):
        r = 5 + i
        for ci, val in enumerate([cat, cnt, amt, desc, action, '', ''], 1):
            cell = wss.cell(row=r, column=ci, value=val)
            D(cell, bg, align='left' if ci in [1, 4, 5] else 'center', bold=(ci == 1))
            if ci == 3 and isinstance(val, float):
                cell.number_format = '#,##0.00'
        wss.row_dimensions[r].height = 20

    sec(wss, 13, "B.  NET BILLING SUMMARY", 7, GREEN)
    for ci, h in enumerate(["Item", "Partner (NIS)", "012Talk (NIS)", "TOTAL (NIS)", "", "", ""], 1):
        H(wss.cell(row=14, column=ci, value=h), GREEN)
    net = [
        ("Our EUP — DONE+CANCELLED",      round(t['partner_eup'], 2),  round(t['talk012_eup'], 2),  round(t['partner_eup'] + t['talk012_eup'], 2),                                           LBLUE),
        ("Refunds — credit back",          round(t['partner_ref'], 2),  round(t['talk012_ref'], 2),  round(t['partner_ref'] + t['talk012_ref'], 2),                                           LTEAL),
        ("PENDING (unconfirmed)",          round(t['pending_eup'], 2),  0.0,                         round(t['pending_eup'], 2),                                                               LPURP),
        ("NET Our System Total",           round(t['partner_eup'] + t['partner_ref'], 2), round(t['talk012_eup'] + t['talk012_ref'], 2), round(t['partner_eup'] + t['talk012_eup'] + t['partner_ref'] + t['talk012_ref'], 2), LYELL),
        ("Supplier CBD — matched phones",  '—', '—', round(t['sup_cbd'], 2),                                                                                                                  LORAN),
        ("Supplier CBD — supplier only",   '—', '—', round(t['sup_only_cbd'], 2),                                                                                                             LRED),
        ("REAL GAP (Sup Only − Our Only)", '—', '—', round(t['real_gap'], 2),                                                                                                                 LORAN),
    ]
    for i, row in enumerate(net):
        r = 15 + i
        bg = row[-1]
        is_bold = i in [3, 6]
        for ci, val in enumerate(row[:-1], 1):
            cell = wss.cell(row=r, column=ci, value=val)
            D(cell, bg, align='left' if ci == 1 else 'center', bold=is_bold or ci == 1)
            if ci in [2, 3, 4] and isinstance(val, float):
                cell.number_format = '#,##0.00'
        wss.row_dimensions[r].height = 18

    wss.column_dimensions['A'].width = 34
    wss.column_dimensions['B'].width = 18
    wss.column_dimensions['C'].width = 20
    wss.column_dimensions['D'].width = 40
    wss.column_dimensions['E'].width = 30
    wss.freeze_panes = 'A3'

    # ---- SHEET 2: SUPPLIER ONLY ----
    ws_so = wb.create_sheet("Supplier Only")
    so_cols = ['Phone_Display', 'Sup_Date', 'Package', 'Sup_TxID', 'CBD', 'Net_Total', 'Cust_Name', 'Reason']
    so_display = result['sup_only'][[c for c in so_cols if c in result['sup_only'].columns]].copy() if len(result['sup_only']) > 0 else pd.DataFrame()
    if len(so_display) > 0:
        rename = {'Phone_Display': 'Phone', 'Sup_Date': 'Supplier Date', 'Sup_TxID': 'Supplier Tx ID',
                  'CBD': 'CBD (NIS)', 'Net_Total': 'Net Total (NIS)', 'Cust_Name': 'Customer Name'}
        so_display.rename(columns=rename, inplace=True)

    ncols_so = max(len(so_display.columns) + 1, 9) if len(so_display) else 9
    ttl(ws_so, f"❌  SUPPLIER ONLY — {t['sup_only_count']} phones  |  CBD: {t['sup_only_cbd']:,.2f} NIS", ncols_so, RED)
    ws_so.merge_cells(f'A2:{get_column_letter(ncols_so)}2')
    ws_so['A2'].value = "ℹ️  Phones billed by supplier but NOT found in our system (DONE/CANCELLED). Possible date shift or missing transaction."
    ws_so['A2'].font = Font(color=RED, size=9, name='Arial')
    ws_so['A2'].fill = PatternFill('solid', start_color=LRED)
    ws_so['A2'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws_so['A2'].border = bd()
    ws_so.row_dimensions[2].height = 24

    if len(so_display) > 0:
        last_r_so = write_df(ws_so, so_display, 3, RED, LRED, {'CBD (NIS)', 'Net Total (NIS)'})
        # Total row
        tr = last_r_so
        for ci, col in enumerate(so_display.columns, 1):
            cell = ws_so.cell(row=tr, column=ci)
            if col == 'Phone': cell.value = 'TOTAL'
            elif col == 'CBD (NIS)': cell.value = round(so_display['CBD (NIS)'].sum(), 2)
            H(cell, GREEN)
            if col == 'CBD (NIS)': cell.number_format = '#,##0.00'
        # Verified column
        ver_col = len(so_display.columns) + 1
        H(ws_so.cell(row=3, column=ver_col, value='Verified ✓'), YELL, fg='000000')
        for ri in range(len(so_display)):
            r = ri + 4
            vc = ws_so.cell(row=r, column=ver_col, value='⬜ Not Verified')
            vc.font = Font(bold=True, color='7F6000', size=9, name='Arial')
            vc.fill = PatternFill('solid', start_color=YELL)
            vc.alignment = Alignment(horizontal='center', vertical='center')
            vc.border = bd()
        vcl = get_column_letter(ver_col)
        dv = DataValidation(type="list",
             formula1='"⬜ Not Verified,✅ Found — date shift confirmed,❌ Not found — investigate"',
             allow_blank=False, showDropDown=False)
        dv.sqref = f"{vcl}4:{vcl}{len(so_display)+3}"
        ws_so.add_data_validation(dv)

    ws_so.column_dimensions['A'].width = 16
    ws_so.column_dimensions['B'].width = 18
    ws_so.column_dimensions['C'].width = 22
    ws_so.freeze_panes = 'A4'

    # ---- SHEET 3: OUR SYSTEM ONLY ----
    ws_oo = wb.create_sheet("Our System Only")
    oo_cols = ['Phone_Display', 'Date & Time', 'Transaction ID', 'Operator', 'Eff_Status', 'Product Name', 'End User Price', 'Customer name', 'Reason']
    oo_display = result['our_only'][[c for c in oo_cols if c in result['our_only'].columns]].copy() if len(result['our_only']) > 0 else pd.DataFrame()
    if len(oo_display) > 0:
        rename = {'Phone_Display': 'Phone', 'Transaction ID': 'Our Tx ID', 'Eff_Status': 'Status',
                  'End User Price': 'End User Price (NIS)', 'Customer name': 'Customer Name'}
        oo_display.rename(columns=rename, inplace=True)

    ncols_oo = max(len(oo_display.columns) + 1, 9) if len(oo_display) else 9
    ttl(ws_oo, f"⚠️  OUR SYSTEM ONLY — {t['our_only_count']} phones  |  EUP: {t['our_only_eup']:,.2f} NIS", ncols_oo, ORANGE)
    ws_oo.merge_cells(f'A2:{get_column_letter(ncols_oo)}2')
    ws_oo['A2'].value = "ℹ️  DONE in our system but missing from supplier. Late transactions (after 22:00) typically appear in next day's supplier file."
    ws_oo['A2'].font = Font(color=ORANGE, size=9, name='Arial')
    ws_oo['A2'].fill = PatternFill('solid', start_color=LORAN)
    ws_oo['A2'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws_oo['A2'].border = bd()
    ws_oo.row_dimensions[2].height = 24

    if len(oo_display) > 0:
        last_r_oo = write_df(ws_oo, oo_display, 3, ORANGE, LORAN, {'End User Price (NIS)'})
        tr = last_r_oo
        for ci, col in enumerate(oo_display.columns, 1):
            cell = ws_oo.cell(row=tr, column=ci)
            if col == 'Phone': cell.value = 'TOTAL'
            elif col == 'End User Price (NIS)': cell.value = round(oo_display['End User Price (NIS)'].sum(), 2)
            H(cell, GREEN)
            if col == 'End User Price (NIS)': cell.number_format = '#,##0.00'
        # Color late rows orange
        if 'Reason' in oo_display.columns:
            for ri, (_, row) in enumerate(oo_display.iterrows()):
                r = ri + 4
                if 'Late' in str(row.get('Reason', '')):
                    for ci in range(1, len(oo_display.columns) + 1):
                        cell = ws_oo.cell(row=r, column=ci)
                        cell.fill = PatternFill('solid', start_color='FFF2CC')
        ver_col = len(oo_display.columns) + 1
        H(ws_oo.cell(row=3, column=ver_col, value='Verified ✓'), YELL, fg='000000')
        for ri in range(len(oo_display)):
            r = ri + 4
            vc = ws_oo.cell(row=r, column=ver_col, value='⬜ Not Verified')
            vc.font = Font(bold=True, color='7F6000', size=9, name='Arial')
            vc.fill = PatternFill('solid', start_color=YELL)
            vc.alignment = Alignment(horizontal='center', vertical='center')
            vc.border = bd()
        vcl = get_column_letter(ver_col)
        dv = DataValidation(type="list",
             formula1='"⬜ Not Verified,✅ Found in next day supplier report,❌ Not found — investigate"',
             allow_blank=False, showDropDown=False)
        dv.sqref = f"{vcl}4:{vcl}{len(oo_display)+3}"
        ws_oo.add_data_validation(dv)

    ws_oo.column_dimensions['A'].width = 16
    ws_oo.column_dimensions['B'].width = 18
    ws_oo.freeze_panes = 'A4'

    # ---- SHEET 4: MATCHED ----
    ws_m = wb.create_sheet("Matched")
    num_cols_m = {'Supplier CBD (NIS)', 'Our EUP (NIS)', 'Difference (NIS)'}
    if len(result['matched']) > 0:
        ttl(ws_m, f"✅  MATCHED — {len(result['matched'])} records  |  CBD: {t['sup_cbd']:,.2f} NIS  |  EUP: {t['our_eup']:,.2f} NIS  |  Diff: {t['diff']:,.2f} NIS",
            len(result['matched'].columns), NAVY)
        last_r_m = write_df(ws_m, result['matched'], 2, NAVY, LBLUE, num_cols_m)
        tr = last_r_m
        for ci, col in enumerate(result['matched'].columns, 1):
            cell = ws_m.cell(row=tr, column=ci)
            if col == 'Supplier Tx ID': cell.value = 'TOTAL'
            elif col in num_cols_m: cell.value = round(result['matched'][col].sum(), 2)
            H(cell, GREEN)
            if col in num_cols_m: cell.number_format = '#,##0.00'
        diff_col = list(result['matched'].columns).index('Difference (NIS)') + 1
        for ri in range(len(result['matched'])):
            r = ri + 3
            diff_val = result['matched'].iloc[ri]['Difference (NIS)']
            if abs(diff_val) > 0.01:
                cell = ws_m.cell(row=r, column=diff_col)
                cell.font = Font(bold=True, color=RED if diff_val > 0 else GREEN, size=9, name='Arial')
    else:
        ttl(ws_m, "✅  MATCHED — No data", 12, NAVY)

    for ci, w in enumerate([16, 14, 24, 14, 18, 16, 16, 10, 12, 26, 18, 16], 1):
        ws_m.column_dimensions[get_column_letter(ci)].width = w
    ws_m.freeze_panes = 'A3'

    # ---- SHEET 5: SUPPLIER vs PENDING ----
    ws_sp = wb.create_sheet("Supplier vs Pending")
    sp_display = result['sup_pending'][['Phone_Display', 'Sup_Date', 'Package', 'Sup_TxID', 'CBD']].copy() if len(result['sup_pending']) > 0 else pd.DataFrame()
    if len(sp_display) > 0:
        sp_display.rename(columns={'Phone_Display': 'Phone', 'Sup_Date': 'Supplier Date',
                                    'Sup_TxID': 'Supplier Tx ID', 'CBD': 'CBD (NIS)'}, inplace=True)
    ttl(ws_sp, f"🔄  SUPPLIER BILLED / WE HAVE PENDING — {t['sup_pending_count']} records  |  {t['sup_pending_cbd']:,.2f} NIS", 7, PURPLE)
    ws_sp.merge_cells('A2:G2')
    ws_sp['A2'].value = "ℹ️  Supplier already charged these phones but our system shows PENDING_CANCELLATION. Verify: did cancellation go through?"
    ws_sp['A2'].font = Font(color=PURPLE, size=9, name='Arial')
    ws_sp['A2'].fill = PatternFill('solid', start_color=LPURP)
    ws_sp['A2'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws_sp['A2'].border = bd()
    ws_sp.row_dimensions[2].height = 24
    if len(sp_display) > 0:
        write_df(ws_sp, sp_display, 3, PURPLE, LPURP, {'CBD (NIS)'})
    ws_sp.freeze_panes = 'A4'

    # ---- SHEET 6: PENDING ----
    ws_pnd = wb.create_sheet("Pending Cancellation")
    pnd_cols = ['Phone_Display', 'Date & Time', 'Transaction ID', 'Operator', 'Eff_Status', 'Product Name', 'End User Price', 'Customer price', 'Customer name']
    pnd_display = result['pending'][[c for c in pnd_cols if c in result['pending'].columns]].copy() if len(result['pending']) > 0 else pd.DataFrame()
    if len(pnd_display) > 0:
        pnd_display.rename(columns={'Phone_Display': 'Phone', 'Transaction ID': 'Our Tx ID',
                                    'Eff_Status': 'Status', 'End User Price': 'End User Price (NIS)',
                                    'Customer price': 'Customer Price (NIS)', 'Customer name': 'Customer Name'}, inplace=True)
    ttl(ws_pnd, f"🕐  PENDING CANCELLATION — {t['pending_count']} record(s)  |  EUP: {t['pending_eup']:,.2f} NIS",
        max(len(pnd_display.columns) + 1, 10) if len(pnd_display) else 10, PURPLE)
    ws_pnd.merge_cells(f'A2:{get_column_letter(10)}2')
    ws_pnd['A2'].value = "ℹ️  Check next day: did supplier send REFUND (approved) or back to DONE (rejected)? Update Verified column."
    ws_pnd['A2'].font = Font(color=PURPLE, size=9, name='Arial')
    ws_pnd['A2'].fill = PatternFill('solid', start_color=LPURP)
    ws_pnd['A2'].alignment = Alignment(horizontal='left', vertical='center')
    ws_pnd['A2'].border = bd()
    ws_pnd.row_dimensions[2].height = 20
    if len(pnd_display) > 0:
        ncols_pnd = len(pnd_display.columns) + 1
        write_df(ws_pnd, pnd_display, 3, PURPLE, LPURP, {'End User Price (NIS)', 'Customer Price (NIS)'})
        H(ws_pnd.cell(row=3, column=ncols_pnd, value='Verified ✓'), YELL, fg='000000')
        for ri in range(len(pnd_display)):
            r = ri + 4
            vc = ws_pnd.cell(row=r, column=ncols_pnd, value='⬜ Not Verified')
            vc.font = Font(bold=True, color='7F6000', size=9, name='Arial')
            vc.fill = PatternFill('solid', start_color=YELL)
            vc.alignment = Alignment(horizontal='center', vertical='center')
            vc.border = bd()
        vcl = get_column_letter(ncols_pnd)
        dv = DataValidation(type="list",
             formula1='"⬜ Not Verified,✅ Confirmed REFUND — cancellation approved,🔄 Back to DONE — cancellation rejected"',
             allow_blank=False, showDropDown=False)
        dv.sqref = f"{vcl}4:{vcl}{len(pnd_display)+3}"
        ws_pnd.add_data_validation(dv)
    ws_pnd.freeze_panes = 'A4'

    # ---- SHEET 7: REFUNDS ----
    ws_ref = wb.create_sheet("Refunds")
    ref_cols = ['Operator', 'Phone_Display', 'Date & Time', 'Transaction ID', 'Product Name', 'End User Price', 'Customer price']
    ref_display = result['refunds'][[c for c in ref_cols if c in result['refunds'].columns]].copy() if len(result['refunds']) > 0 else pd.DataFrame()
    if len(ref_display) > 0:
        ref_display.rename(columns={'Phone_Display': 'Phone', 'Transaction ID': 'Our Tx ID',
                                    'End User Price': 'End User Price (NIS)', 'Customer price': 'Customer Price (NIS)'}, inplace=True)
    ttl(ws_ref, f"↩️  REFUNDS — Previous Period  |  {t['refunds_count']} records  |  EUP: {t['refunds_eup']:,.2f} NIS",
        max(len(ref_display.columns), 7) if len(ref_display) else 7, TEAL)
    if len(ref_display) > 0:
        write_df(ws_ref, ref_display, 2, TEAL, LTEAL, {'End User Price (NIS)', 'Customer Price (NIS)'})
    ws_ref.freeze_panes = 'A3'

    # ---- SHEET 8: FAILED ----
    ws_f = wb.create_sheet("Failed")
    fail_cols = ['Operator', 'Phone_Display', 'Date & Time', 'Transaction ID', 'Product Name', 'End User Price', 'Error description']
    fail_display = result['failed'][[c for c in fail_cols if c in result['failed'].columns]].copy() if len(result['failed']) > 0 else pd.DataFrame()
    if len(fail_display) > 0:
        fail_display.rename(columns={'Phone_Display': 'Phone', 'Transaction ID': 'Our Tx ID',
                                    'End User Price': 'End User Price (NIS)', 'Error description': 'Error Description'}, inplace=True)
    ttl(ws_f, f"🔴  FAILED TRANSACTIONS — {t['failed_count']} records",
        max(len(fail_display.columns), 7) if len(fail_display) else 7, '7F0000')
    if len(fail_display) > 0:
        write_df(ws_f, fail_display, 2, '7F0000', LRED)
    ws_f.freeze_panes = 'A3'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# MONTHLY EXCEL EXPORT
# ============================================================
def create_monthly_excel(history, month_label):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monthly Summary"

    def bd():
        s = Side(style='thin', color='CCCCCC')
        return Border(left=s, right=s, top=s, bottom=s)
    def H(cell, bg, fg='FFFFFF', sz=10):
        cell.font = Font(bold=True, color=fg, size=sz, name='Arial')
        cell.fill = PatternFill('solid', start_color=bg)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = bd()
    def D(cell, bg='FFFFFF', align='center', fmt=None):
        cell.font = Font(size=9, name='Arial')
        cell.fill = PatternFill('solid', start_color=bg)
        cell.alignment = Alignment(horizontal=align, vertical='center')
        cell.border = bd()
        if fmt: cell.number_format = fmt

    NAVY = '1F3864'; LBLUE = 'DEEAF1'; GREEN = '375623'; WHITE = 'FFFFFF'; RED = 'C00000'

    ws.merge_cells('A1:K1')
    ws['A1'].value = f"📅  MONTHLY SUMMARY — {month_label}"
    ws['A1'].font = Font(bold=True, color='FFFFFF', size=14, name='Arial')
    ws['A1'].fill = PatternFill('solid', start_color=NAVY)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].border = bd()
    ws.row_dimensions[1].height = 36

    headers = ['Date', 'Supplier CBD (NIS)', 'Our EUP (NIS)', 'Difference (NIS)',
               'Matched', 'Supplier Only', 'Supplier Only CBD', 'Our Only', 'Our Only EUP',
               'Real Gap (NIS)', 'Net Billed (NIS)']
    for ci, h in enumerate(headers, 1):
        H(ws.cell(row=2, column=ci, value=h), NAVY)
    ws.row_dimensions[2].height = 35

    for ri, rec in enumerate(history):
        r = ri + 3
        bg = LBLUE if ri % 2 == 0 else WHITE
        vals = [
            rec.get('date', ''),
            rec.get('sup_cbd', 0),
            rec.get('our_eup', 0),
            rec.get('diff', 0),
            rec.get('matched_count', 0),
            rec.get('sup_only_count', 0),
            rec.get('sup_only_cbd', 0),
            rec.get('our_only_count', 0),
            rec.get('our_only_eup', 0),
            rec.get('real_gap', 0),
            rec.get('net_billed', 0),
        ]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=ci, value=val)
            D(cell, bg, fmt='#,##0.00' if ci in [2, 3, 4, 7, 9, 10, 11] else None)
            if ci == 10 and isinstance(val, (int, float)) and abs(val) > 0.01:
                cell.font = Font(bold=True, color=RED if val > 0 else GREEN, size=9, name='Arial')
        ws.row_dimensions[r].height = 18

    tr = len(history) + 3
    for ci, val in enumerate(
        ['TOTAL MONTH'] + [f'=SUM({get_column_letter(c)}3:{get_column_letter(c)}{tr-1})' for c in range(2, 12)], 1):
        cell = ws.cell(row=tr, column=ci, value=val)
        H(cell, GREEN, sz=11)
        if ci > 1: cell.number_format = '#,##0.00'
    ws.row_dimensions[tr].height = 24

    for ci, w in enumerate([14, 18, 16, 16, 12, 14, 18, 12, 16, 18, 18], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'A3'

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# MAIN APP
# ============================================================
def main():
    st.markdown("""
    <div class="main-header">
        <h1 style="margin:0; font-size:28px;">📊 Reconciliation Dashboard</h1>
        <p style="margin:5px 0 0 0; opacity:0.85; font-size:14px;">Supplier vs Our System (Partner + 012Talk)</p>
    </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### 📋 Navigation")
        page = st.radio("Navigation Menu", ["🔄 Daily Reconciliation", "📅 Monthly Summary", "ℹ️ Instructions"],
                        label_visibility="collapsed")
        st.markdown("---")
        st.markdown("### 📊 History")
        sh = get_spreadsheet()
        if sh is not None:
            st.success("✅ Google Sheets connected")
        else:
            st.warning("⚠️ Sheets not connected")
        history = load_history()
        if history:
            st.success(f"✅ {len(history)} days recorded")
            last = history[-1]
            st.info(f"Last: {last.get('date', 'N/A')}")
        else:
            st.info("No history yet")

    # ============================================================
    # PAGE: DAILY RECONCILIATION
    # ============================================================
    if page == "🔄 Daily Reconciliation":
        st.markdown("## 🔄 Daily Reconciliation")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("**1️⃣ Supplier File (.xls)**")
            sup_file = st.file_uploader("Supplier file", type=['xls', 'xlsx', 'csv'],
                                        label_visibility="collapsed", key="sup")
            st.caption("Supplier report — Charge Before Discount (col M)")
        with col2:
            st.markdown("**2️⃣ Partner File (.csv)**")
            part_file = st.file_uploader("Partner file", type=['csv', 'xlsx'],
                                         label_visibility="collapsed", key="part")
            st.caption("Our system export — Partner operator")
        with col3:
            st.markdown("**3️⃣ 012Talk File (.csv)**")
            talk_file = st.file_uploader("012Talk file", type=['csv', 'xlsx'],
                                         label_visibility="collapsed", key="talk")
            st.caption("Our system export — 012Talk operator")

        report_date = st.date_input("📅 Report Date", value=date.today())

        if sup_file and part_file and talk_file:
            if st.button("▶ Run Reconciliation", type="primary", use_container_width=True):
                with st.spinner("Processing reconciliation..."):
                    sup_df,  err1 = load_supplier(sup_file.read(), sup_file.name)
                    part_df, err2 = load_our(part_file.read(), 'Partner')
                    talk_df, err3 = load_our(talk_file.read(), '012Talk')

                    if err1: st.error(f"❌ Supplier file error: {err1}"); return
                    if err2: st.error(f"❌ Partner file error: {err2}"); return
                    if err3: st.error(f"❌ 012Talk file error: {err3}"); return

                    auto_date = str(report_date)
                    if sup_df is not None and 'Sup_Date' in sup_df.columns and len(sup_df) > 0:
                        try:
                            parsed = pd.to_datetime(sup_df['Sup_Date'].iloc[0], dayfirst=True)
                            auto_date = parsed.strftime('%Y-%m-%d')
                            st.info(f"📅 Date auto-detected from supplier file: **{auto_date}**")
                        except:
                            pass

                    result = run_reconciliation(sup_df, part_df, talk_df)
                    st.session_state['result'] = result
                    st.session_state['report_date'] = auto_date

                    # Cross-day matching
                    shifts_our, shifts_sup = cross_day_match(result, auto_date)
                    st.session_state['shifts_our'] = shifts_our
                    st.session_state['shifts_sup'] = shifts_sup

                    st.success("✅ Reconciliation complete!")

        # ---- RESULTS ----
        if 'result' in st.session_state:
            result = st.session_state['result']
            t = result['totals']
            report_date_str = st.session_state.get('report_date', str(date.today()))

            st.markdown("---")
            st.markdown("### 📊 Transaction Count Discrepancy")

            # Cross-day shift banner
            shifts_our = st.session_state.get('shifts_our', [])
            shifts_sup = st.session_state.get('shifts_sup', [])
            if shifts_our:
                st.success(f"✅ Date Shift Confirmed: {len(shifts_our)} phone(s) from 'Our Only' found in yesterday's 'Supplier Only' — these are normal date shifts: {', '.join(shifts_our)}")
            if shifts_sup:
                st.success(f"✅ Date Shift Confirmed: {len(shifts_sup)} phone(s) from 'Supplier Only' found in yesterday's 'Our Only' — normal date shifts: {', '.join(shifts_sup)}")

            # Top metrics — focus on counts and real gap
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("✅ Matched", f"{t['matched_count']:,}",
                          delta="OK" if t['matched_count'] > 0 else None)
            with col2:
                st.metric("❌ Supplier Only", t['sup_only_count'],
                          delta=f"−{t['sup_only_cbd']:,.2f} NIS" if t['sup_only_count'] > 0 else "0",
                          delta_color="inverse")
            with col3:
                st.metric("⚠️ Our Only", t['our_only_count'],
                          delta=f"−{t['our_only_eup']:,.2f} NIS" if t['our_only_count'] > 0 else "0",
                          delta_color="inverse")
            with col4:
                st.metric("🔄 Sup vs Pending", t['sup_pending_count'],
                          delta=f"{t['sup_pending_cbd']:,.2f} NIS" if t['sup_pending_count'] > 0 else "0",
                          delta_color="off")
            with col5:
                gap = t['real_gap']
                gap_label = f"+{gap:,.2f} NIS (sup higher)" if gap > 0 else (f"{gap:,.2f} NIS (we higher)" if gap < 0 else "0.00 NIS")
                st.metric("📊 Real Gap", gap_label,
                          delta="Supplier Only CBD − Our Only EUP",
                          delta_color="inverse" if gap > 0 else ("normal" if gap < 0 else "off"))

            st.markdown("---")

            # ---- TABS ----
            tabs = st.tabs([
                f"❌ Supplier Only ({t['sup_only_count']})",
                f"⚠️ Our System Only ({t['our_only_count']})",
                f"🔄 Sup vs Pending ({t['sup_pending_count']})",
                f"✅ Matched ({t['matched_count']})",
                f"🕐 Pending ({t['pending_count']})",
                f"↩️ Refunds ({t['refunds_count']})",
                f"🔴 Failed ({t['failed_count']})",
            ])

            # TAB 1 — SUPPLIER ONLY
            with tabs[0]:
                st.markdown("**Phones billed by supplier — NOT found in our system**")
                st.caption(f"Total CBD: **{t['sup_only_cbd']:,.2f} NIS** — these are charges you may owe but can't verify")
                if len(result['sup_only']) > 0:
                    display_cols = ['Phone_Display', 'Sup_Date', 'Package', 'Sup_TxID', 'CBD', 'Cust_Name', 'Reason']
                    show = result['sup_only'][[c for c in display_cols if c in result['sup_only'].columns]].copy()
                    show.rename(columns={'Phone_Display': 'Phone', 'Sup_Date': 'Date',
                                         'Sup_TxID': 'Supplier Tx ID', 'CBD': 'CBD (NIS)',
                                         'Cust_Name': 'Customer Name'}, inplace=True)
                    st.dataframe(show, use_container_width=True, hide_index=True)
                    st.info(f"📋 Total: {len(show)} phones | {t['sup_only_cbd']:,.2f} NIS")
                else:
                    st.success("✅ No supplier-only records — perfect match on transaction count!")

            # TAB 2 — OUR ONLY
            with tabs[1]:
                st.markdown("**Phones in our system — NOT found in supplier report**")
                st.caption(f"Total EUP: **{t['our_only_eup']:,.2f} NIS** — transactions supplier didn't list (often late-night date shifts)")
                if len(result['our_only']) > 0:
                    display_cols = ['Phone_Display', 'Date & Time', 'Operator', 'Product Name', 'End User Price', 'Eff_Status', 'Reason']
                    show = result['our_only'][[c for c in display_cols if c in result['our_only'].columns]].copy()
                    show.rename(columns={'Phone_Display': 'Phone', 'End User Price': 'EUP (NIS)',
                                         'Eff_Status': 'Status'}, inplace=True)

                    # Highlight late transactions
                    late_count = result['our_only'].get('Is_Late', pd.Series(dtype=bool)).sum() if 'Is_Late' in result['our_only'].columns else 0
                    if late_count > 0:
                        st.warning(f"⏰ {late_count} transaction(s) after 22:00 — likely to appear in tomorrow's supplier report")

                    st.dataframe(show, use_container_width=True, hide_index=True)
                    st.info(f"📋 Total: {len(show)} phones | {t['our_only_eup']:,.2f} NIS")
                else:
                    st.success("✅ No our-only records — supplier listed everything!")

            # TAB 3 — SUPPLIER vs PENDING
            with tabs[2]:
                st.markdown("**Supplier already charged — we show PENDING_CANCELLATION**")
                st.caption("These phones: supplier billed them, but our system thinks cancellation is pending. Needs verification.")
                if len(result['sup_pending']) > 0:
                    display_cols = ['Phone_Display', 'Sup_Date', 'Package', 'Sup_TxID', 'CBD']
                    show = result['sup_pending'][[c for c in display_cols if c in result['sup_pending'].columns]].copy()
                    show.rename(columns={'Phone_Display': 'Phone', 'Sup_Date': 'Supplier Date',
                                         'Sup_TxID': 'Supplier Tx ID', 'CBD': 'CBD (NIS)'}, inplace=True)
                    st.dataframe(show, use_container_width=True, hide_index=True)
                    st.warning(f"⚠️ {t['sup_pending_count']} phone(s) | {t['sup_pending_cbd']:,.2f} NIS — check if cancellation was processed")
                else:
                    st.success("✅ No supplier vs pending conflicts!")

            # TAB 4 — MATCHED
            with tabs[3]:
                st.markdown("**Phones matched in both systems (DONE/CANCELLED)**")
                if len(result['matched']) > 0:
                    st.dataframe(result['matched'], use_container_width=True, hide_index=True)
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Supplier CBD", f"{t['sup_cbd']:,.2f} NIS")
                    col2.metric("Our EUP", f"{t['our_eup']:,.2f} NIS")
                    col3.metric("Difference", f"{t['diff']:,.2f} NIS",
                                delta="OK" if abs(t['diff']) < 0.01 else "Check amounts")
                else:
                    st.info("No matched records")

            # TAB 5 — PENDING
            with tabs[4]:
                if len(result['pending']) > 0:
                    st.warning(f"⚠️ {t['pending_count']} transaction(s) pending. Check next day!")
                    display_cols = ['Phone_Display', 'Date & Time', 'Transaction ID', 'Operator', 'Product Name', 'End User Price']
                    show = result['pending'][[c for c in display_cols if c in result['pending'].columns]].copy()
                    show.rename(columns={'Phone_Display': 'Phone', 'End User Price': 'EUP (NIS)'}, inplace=True)
                    st.dataframe(show, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ No pending transactions!")

            # TAB 6 — REFUNDS
            with tabs[5]:
                if len(result['refunds']) > 0:
                    st.info(f"↩️ {t['refunds_count']} refunds — {t['refunds_eup']:,.2f} NIS credit (arrive end of month)")
                    display_cols = ['Operator', 'Phone_Display', 'Date & Time', 'Product Name', 'End User Price']
                    show = result['refunds'][[c for c in display_cols if c in result['refunds'].columns]].copy()
                    show.rename(columns={'Phone_Display': 'Phone', 'End User Price': 'EUP (NIS)'}, inplace=True)
                    st.dataframe(show, use_container_width=True, hide_index=True)
                else:
                    st.info("No refunds today")

            # TAB 7 — FAILED
            with tabs[6]:
                if len(result['failed']) > 0:
                    st.error(f"🔴 {t['failed_count']} failed transactions")
                    display_cols = ['Operator', 'Phone_Display', 'Date & Time', 'Product Name', 'Error description']
                    show = result['failed'][[c for c in display_cols if c in result['failed'].columns]].copy()
                    show.rename(columns={'Phone_Display': 'Phone'}, inplace=True)
                    st.dataframe(show, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ No failed transactions!")

            # ---- NET BILLING SUMMARY ----
            st.markdown("---")
            st.markdown("### 💰 Net Billing Summary")
            net_data = {
                'Item': [
                    'Our EUP — DONE+CANCELLED',
                    'Refunds — credit back (prev. period)',
                    'PENDING (unconfirmed)',
                    'NET Our System Total (excl. PENDING)',
                    'Supplier CBD — matched phones',
                    'Supplier CBD — supplier only phones',
                    '📊 REAL GAP (Supplier Only − Our Only)',
                ],
                'Partner (NIS)': [
                    round(t['partner_eup'], 2), round(t['partner_ref'], 2),
                    round(t['pending_eup'], 2), round(t['partner_eup'] + t['partner_ref'], 2),
                    '—', '—', '—'
                ],
                '012Talk (NIS)': [
                    round(t['talk012_eup'], 2), round(t['talk012_ref'], 2),
                    0, round(t['talk012_eup'] + t['talk012_ref'], 2),
                    '—', '—', '—'
                ],
                'TOTAL (NIS)': [
                    round(t['partner_eup'] + t['talk012_eup'], 2),
                    round(t['partner_ref'] + t['talk012_ref'], 2),
                    round(t['pending_eup'], 2),
                    round(t['partner_eup'] + t['talk012_eup'] + t['partner_ref'] + t['talk012_ref'], 2),
                    round(t['sup_cbd'], 2),
                    round(t['sup_only_cbd'], 2),
                    round(t['real_gap'], 2),
                ],
            }
            st.dataframe(pd.DataFrame(net_data), use_container_width=True, hide_index=True)

            # ---- DOWNLOAD & SAVE ----
            st.markdown("---")
            col1, col2 = st.columns(2)
            with col1:
                excel_buf = create_excel_report(result, report_date_str)
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_buf,
                    file_name=f"Reconciliation_{report_date_str.replace('-', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            with col2:
                if st.button("💾 Save to Monthly History", use_container_width=True):
                    record = {
                        'date':             report_date_str,
                        'sup_cbd':          round(t['sup_cbd'], 2),
                        'our_eup':          round(t['our_eup'], 2),
                        'diff':             round(t['diff'], 2),
                        'matched_count':    t['matched_count'],
                        'sup_only_count':   t['sup_only_count'],
                        'sup_only_cbd':     round(t['sup_only_cbd'], 2),
                        'our_only_count':   t['our_only_count'],
                        'our_only_eup':     round(t['our_only_eup'], 2),
                        'real_gap':         round(t['real_gap'], 2),
                        'pending_count':    t['pending_count'],
                        'refunds_eup':      round(t['refunds_eup'], 2),
                        'net_billed':       round(t['partner_eup'] + t['talk012_eup'] + t['partner_ref'] + t['talk012_ref'], 2),
                    }
                    ok, msg = save_to_sheets(record)
                    ok2, msg2 = save_details_to_sheets(report_date_str, result)
                    if ok:
                        st.success(f"✅ {msg}")
                    else:
                        st.warning(f"⚠️ {msg}")
                    if ok2:
                        st.info(f"📋 {msg2}")
                    else:
                        st.warning(f"⚠️ Details: {msg2}")

    # ============================================================
    # PAGE: MONTHLY SUMMARY
    # ============================================================
    elif page == "📅 Monthly Summary":
        st.markdown("## 📅 Monthly Summary")

        # Get available months from Sheets
        sh = get_spreadsheet()
        available_months = []
        if sh is not None:
            try:
                for ws in sh.worksheets():
                    if _is_month_sheet(ws.title):
                        try:
                            dt = datetime.strptime(ws.title, '%B %Y')
                            available_months.append(dt.strftime('%Y-%m'))
                        except Exception:
                            pass
                available_months = sorted(set(available_months), reverse=True)
            except Exception:
                pass

        if not available_months:
            # Fallback to local
            history = _load_local_history()
            available_months = sorted(set(h['date'][:7] for h in history), reverse=True)

        if not available_months:
            st.info("No history yet. Run daily reconciliations and click 'Save to Monthly History'.")
            return

        selected_month = st.selectbox("Select Month", available_months,
                                      format_func=lambda m: datetime.strptime(m, '%Y-%m').strftime('%B %Y'))
        month_history = load_history(month=selected_month)

        if not month_history:
            st.warning("No data for selected month")
            return

        total_sup      = sum(h.get('sup_cbd', 0) for h in month_history)
        total_eup      = sum(h.get('our_eup', 0) for h in month_history)
        total_gap      = sum(h.get('real_gap', 0) for h in month_history)
        total_refunds  = sum(h.get('refunds_eup', 0) for h in month_history)

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("📅 Days Recorded", len(month_history))
        col2.metric("Supplier CBD (matched)", f"{total_sup:,.2f} NIS")
        col3.metric("Our EUP (matched)", f"{total_eup:,.2f} NIS")
        col4.metric("↩️ Total Refunds", f"{total_refunds:,.2f} NIS")
        col5.metric("📊 Monthly Real Gap", f"{total_gap:,.2f} NIS",
                    delta="Sup Only − Our Only" ,
                    delta_color="inverse" if total_gap > 0 else "normal")

        st.markdown("---")
        df_month = pd.DataFrame(month_history)
        st.dataframe(df_month, use_container_width=True, hide_index=True)

        month_label = datetime.strptime(selected_month, '%Y-%m').strftime('%B %Y')
        monthly_buf = create_monthly_excel(month_history, month_label)
        st.download_button(
            label=f"📥 Download Monthly Report — {month_label}",
            data=monthly_buf,
            file_name=f"Monthly_Summary_{selected_month.replace('-', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

    # ============================================================
    # PAGE: INSTRUCTIONS
    # ============================================================
    elif page == "ℹ️ Instructions":
        st.markdown("## ℹ️ How to Use")
        st.markdown("""
        ### 📋 Daily Process

        **Step 1 — Upload Files**
        - Supplier `.xls` file (keep original format — CSV truncates phone numbers!)
        - Partner `.csv` from our system
        - 012Talk `.csv` from our system

        **Step 2 — Run Reconciliation**
        - Select the report date
        - Click **▶ Run Reconciliation**

        **Step 3 — Review the Gap**
        - **❌ Supplier Only** — supplier billed these, we have no record. Investigate!
        - **⚠️ Our Only** — we charged these, supplier didn't list. Often late-night (after 22:00) date shifts — check next day's supplier file.
        - **🔄 Sup vs Pending** — supplier already billed but we show PENDING. Did cancellation go through?
        - **📊 Real Gap** = Supplier Only CBD − Our Only EUP. This is your actual financial exposure for the day.

        **Step 4 — Download & Save**
        - **📥 Download Excel Report** — save to SharePoint. Contains all sheets with phone details.
        - **💾 Save to Monthly History** — adds daily totals to the monthly tracker.

        ---

        ### 📊 Real Gap Explained
        | Situation | Meaning |
        |---|---|
        | Real Gap = 0 | Every missing phone on one side is balanced by a missing phone on the other |
        | Real Gap > 0 | Supplier charges more than we see — you may owe money |
        | Real Gap < 0 | We have more transactions than supplier billed — supplier may owe a correction |

        ---

        ### ⚠️ Important Rules
        - Always use original `.xls` supplier file — CSV format truncates phone numbers to scientific notation
        - PENDING_CANCELLATION = check next day if supplier approved (REFUND) or rejected (stays DONE)
        - REFUND rows = credits from previous period, arrive end of month
        - REWARD and REFUND_REWARD rows are automatically excluded
        - Late transactions (after 22:00) often appear in the NEXT day's supplier file — normal
        """)


if __name__ == "__main__":
    main()
    
