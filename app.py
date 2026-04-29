import streamlit as st
import pandas as pd
import numpy as np
from io import StringIO, BytesIO
import json
import os
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

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
# HISTORY FILE
# ============================================================
HISTORY_FILE = os.path.join(os.path.dirname(__file__), "history.json")

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_history(history):
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

def add_to_history(record):
    history = load_history()
    # Remove duplicate for same date
    history = [h for h in history if h.get('date') != record.get('date')]
    history.append(record)
    history.sort(key=lambda x: x.get('date',''))
    save_history(history)

# ============================================================
# PHONE NORMALIZATION
# ============================================================
def norm_phone(phone):
    if phone is None or (isinstance(phone, float) and np.isnan(phone)):
        return ''
    s = str(phone).strip().replace('.0','').replace(' ','').replace('+','')
    if 'E' in s.upper() or 'e' in s:
        try: s = str(int(float(s)))
        except: pass
    s = s.replace('.0','')
    if s.startswith('00972'): s = s[5:]
    elif s.startswith('972'): s = s[3:]
    if s.startswith('0'): s = s[1:]
    return s.strip()

# ============================================================
# LOAD FILES
# ============================================================
def load_supplier(file_bytes, filename):
    """Load supplier XLS/CSV file"""
    try:
        if filename.endswith('.xls') or filename.endswith('.XLS'):
            text = file_bytes.decode('cp1255', errors='replace')
        else:
            for enc in ['utf-8-sig','windows-1255','cp1255','latin1']:
                try:
                    text = file_bytes.decode(enc)
                    break
                except: continue
        
        lines = text.replace('\r\n','\n').replace('\r','\n').strip().split('\n')
        # Skip title row if only 1 non-empty cell
        start = 0
        first_line_cols = [c for c in lines[0].split('\t') if c.strip()]
        if len(first_line_cols) == 1:
            start = 1
        
        # Detect separator
        sep = '\t' if lines[start].count('\t') > lines[start].count(',') else ','
        df = pd.read_csv(StringIO('\n'.join(lines[start:])), sep=sep,
                         dtype={'MSISDN': str}, on_bad_lines='skip')
        
        # Rename Hebrew columns
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
        df = df[df['phone_norm'].str.len() >= 8]
        return df, None
    except Exception as e:
        return None, str(e)

def load_our(file_bytes, operator_name):
    """Load our system CSV (Partner or 012Talk)"""
    try:
        for enc in ['utf-8-sig','utf-8','windows-1255','cp1255','latin1']:
            try:
                text = file_bytes.decode(enc)
                break
            except: continue
        
        df = pd.read_csv(StringIO(text), dtype={'Phone Number': str}, on_bad_lines='skip')
        df['Operator'] = operator_name
        df['End User Price'] = pd.to_numeric(df.get('End User Price', 0), errors='coerce').fillna(0)
        df['Customer price'] = pd.to_numeric(df.get('Customer price', 0), errors='coerce').fillna(0)
        
        # Filter out REWARD and REFUND_REWARD
        df = df[~df['Action'].isin(['REWARD', 'REFUND_REWARD'])]
        
        # Add effective status
        df['Is_Refund'] = df['Action'] == 'REFUND'
        df['Eff_Status'] = df.apply(
            lambda r: 'CANCELLED' if r['Action'] == 'REFUND' else r['Status'], axis=1)
        df['phone_norm'] = df['Phone Number'].apply(norm_phone)
        df = df[df['phone_norm'].str.len() >= 7]
        return df, None
    except Exception as e:
        return None, str(e)

# ============================================================
# RECONCILIATION LOGIC
# ============================================================
def run_reconciliation(sup_df, partner_df, talk_df):
    our_all = pd.concat([partner_df, talk_df], ignore_index=True)
    
    our_dc      = our_all[(our_all['Eff_Status'].isin(['DONE','CANCELLED'])) & (~our_all['Is_Refund'])].copy()
    our_pending = our_all[our_all['Eff_Status'] == 'PENDING_CANCELLATION'].copy()
    our_refunds = our_all[our_all['Is_Refund']].copy()
    our_failed  = our_all[our_all['Eff_Status'] == 'FAILED'].copy()
    
    sup_phones = set(sup_df['phone_norm'])
    our_dc_phones = set(our_dc['phone_norm'])
    our_pnd_phones = set(our_pending['phone_norm'])
    
    matched_phones  = sup_phones & our_dc_phones
    sup_only_phones = sup_phones - our_dc_phones - our_pnd_phones
    our_only_phones = our_dc_phones - sup_phones
    
    # Build matched detail
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
            matched_rows.append({
                'MSISDN': sup_row['MSISDN'],
                'Supplier Date': sup_row.get('Sup_Date',''),
                'Supplier Package': sup_row.get('Package',''),
                'Supplier Tx ID': sup_row.get('Sup_TxID',''),
                'Supplier CBD (NIS)': sup_row['CBD'],
                'Our Tx ID': our_row['Transaction ID'],
                'Our Date': our_row['Date & Time'],
                'Our Operator': our_row['Operator'],
                'Our Status': our_row['Eff_Status'],
                'Our Product': our_row['Product Name'],
                'Our EUP (NIS)': our_row['End User Price'],
                'Difference (NIS)': sup_row['CBD'] - our_row['End User Price'],
            })
    
    matched_df = pd.DataFrame(matched_rows)
    sup_only_df = sup_df[sup_df['phone_norm'].isin(sup_only_phones)].copy()
    our_only_df = our_dc[our_dc['phone_norm'].isin(our_only_phones)].copy()
    
    # Totals
    totals = {
        'sup_cbd': matched_df['Supplier CBD (NIS)'].sum() if len(matched_df) else 0,
        'our_eup': matched_df['Our EUP (NIS)'].sum() if len(matched_df) else 0,
        'diff': matched_df['Difference (NIS)'].sum() if len(matched_df) else 0,
        'sup_only_cbd': sup_only_df['CBD'].sum() if len(sup_only_df) else 0,
        'our_only_eup': our_only_df['End User Price'].sum() if len(our_only_df) else 0,
        'pending_eup': our_pending['End User Price'].sum() if len(our_pending) else 0,
        'refunds_eup': our_refunds['End User Price'].sum() if len(our_refunds) else 0,
        'partner_eup': partner_df[partner_df['Eff_Status'].isin(['DONE','CANCELLED']) & ~partner_df['Is_Refund']]['End User Price'].sum(),
        'talk012_eup': talk_df[talk_df['Eff_Status'].isin(['DONE','CANCELLED']) & ~talk_df['Is_Refund']]['End User Price'].sum(),
        'partner_ref': partner_df[partner_df['Is_Refund']]['End User Price'].sum(),
        'talk012_ref': talk_df[talk_df['Is_Refund']]['End User Price'].sum(),
        'matched_count': len(matched_phones),
        'sup_only_count': len(sup_only_phones),
        'our_only_count': len(our_only_phones),
        'pending_count': len(our_pending),
        'refunds_count': len(our_refunds),
        'failed_count': len(our_failed),
    }
    
    return {
        'matched': matched_df,
        'sup_only': sup_only_df,
        'our_only': our_only_df,
        'pending': our_pending,
        'refunds': our_refunds,
        'failed': our_failed,
        'totals': totals,
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
        if fmt: cell.number_format = fmt
    
    NAVY='1F3864'; BLUE='2E75B6'; LBLUE='DEEAF1'
    ORANGE='C55A11'; LORAN='FCE4D6'
    GREEN='375623'; LGREEN='E2EFDA'
    RED='C00000'; LRED='FFE0E0'
    YELL='FFE699'; LYELL='FFFACD'
    TEAL='00695C'; LTEAL='E0F2F1'
    PURPLE='4A148C'; LPURP='EDE7F6'
    WHITE='FFFFFF'
    
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
    
    def write_df_to_sheet(ws, df, start_row, hdr_bg, alt_bg, num_cols=None):
        if len(df) == 0:
            return start_row
        # Headers
        for ci, col in enumerate(df.columns, 1):
            H(ws.cell(row=start_row, column=ci, value=col), hdr_bg, sz=9)
        ws.row_dimensions[start_row].height = 35
        # Data
        for ri, (_, row) in enumerate(df.iterrows()):
            r = start_row + 1 + ri
            bg = alt_bg if ri % 2 == 0 else WHITE
            for ci, col in enumerate(df.columns, 1):
                val = row[col]
                if pd.isna(val): val = ''
                cell = ws.cell(row=r, column=ci, value=val)
                fmt = '#,##0.00' if num_cols and col in num_cols else None
                D(cell, bg, align='right' if num_cols and col in num_cols else ('center' if ci > 1 else 'left'), fmt=fmt)
            ws.row_dimensions[r].height = 15
        return start_row + 1 + len(df)
    
    # ---- SHEET 1: SUMMARY ----
    wss = wb.create_sheet("Summary")
    ttl(wss, f"📊  RECONCILIATION SUMMARY  —  {report_date}", 7, NAVY, 14)
    
    sec(wss, 3, "A.  FILE OVERVIEW", 7, BLUE)
    for ci, h in enumerate(["File","Role","Total Rows","DONE+CANCELLED","PENDING","FAILED","Notes"], 1):
        H(wss.cell(row=4, column=ci, value=h), NAVY)
    ov = [
        ("Supplier file (.xls)", "SUPPLIER REPORT", len(result['matched'])+len(result['sup_only']), 
         len(result['matched'])+len(result['sup_only']), "—", "—",
         "Amount: Charge Before Discount (col M)", LBLUE),
        ("Partner file (.csv)", "OUR SYSTEM — Partner",
         t['matched_count']+t['pending_count']+t['failed_count']+t['refunds_count'],
         "✓", t['pending_count'], t['failed_count'],
         f"incl. {len(result['refunds'][result['refunds']['Operator']=='Partner'])} REFUND rows", LORAN),
        ("012Talk file (.csv)", "OUR SYSTEM — 012Talk",
         t['matched_count']+t['failed_count'],
         "✓", 0, t['failed_count'],
         f"incl. {len(result['refunds'][result['refunds']['Operator']=='012Talk'])} REFUND rows", LYELL),
    ]
    for i, row in enumerate(ov):
        r = 5+i
        bg = row[-1]
        for ci, val in enumerate(row[:-1], 1):
            cell = wss.cell(row=r, column=ci, value=val)
            D(cell, bg, align='left' if ci in [1,2,7] else 'center', bold=(ci<=2))
        wss.row_dimensions[r].height = 18
    
    sec(wss, 9, "B.  RECONCILIATION RESULTS  —  Supplier: CBD  vs  Ours: End User Price", 7, BLUE)
    for ci, h in enumerate(["Category","Description","# Phones","Supplier CBD (NIS)","Our EUP (NIS)","Difference (NIS)","Verification Status"], 1):
        H(wss.cell(row=10, column=ci, value=h), NAVY)
    wss.row_dimensions[10].height = 30
    
    b_data = [
        ("✅  MATCHED", "Phone in supplier AND our system (DONE/CANCELLED)",
         t['matched_count'], round(t['sup_cbd'],2), round(t['our_eup'],2), round(t['diff'],2),
         "✅ Perfect match — 0.00 NIS" if abs(t['diff']) < 0.01 else f"⚠️ Diff: {t['diff']:.2f} NIS",
         LGREEN, GREEN if abs(t['diff']) < 0.01 else RED),
        ("❌  SUPPLIER ONLY", "In supplier — NOT in our system",
         t['sup_only_count'], round(t['sup_only_cbd'],2), "—", "—",
         "N/A — 0 phones" if t['sup_only_count'] == 0 else "⬜ Check Supplier Only sheet",
         LRED if t['sup_only_count'] > 0 else 'F5F5F5', RED),
        ("⚠️  OUR SYSTEM ONLY", "In our system — NOT in supplier",
         t['our_only_count'], "—", round(t['our_only_eup'],2), "—",
         "N/A — 0 phones" if t['our_only_count'] == 0 else "⬜ Check Our System Only sheet",
         LYELL if t['our_only_count'] > 0 else 'F5F5F5', '7F6000'),
        ("🕐  PENDING_CANCELLATION", "Awaiting supplier decision",
         t['pending_count'], "—", round(t['pending_eup'],2), "—",
         "None today" if t['pending_count'] == 0 else "⬜ Check next day — update Pending sheet",
         LPURP if t['pending_count'] > 0 else 'F5F5F5', PURPLE),
    ]
    for i, (cat, desc, cnt, s_a, o_a, diff, ver, bg, ver_fg) in enumerate(b_data):
        r = 11+i
        for ci, val in enumerate([cat, desc, cnt, s_a, o_a, diff, ver], 1):
            cell = wss.cell(row=r, column=ci, value=val)
            if ci == 1:
                D(cell, bg, bold=True)
            elif ci == 7:
                cell.font = Font(bold=True, color=ver_fg, size=9, name='Arial')
                cell.fill = PatternFill('solid', start_color=LGREEN if '✅' in ver else (YELL if '⬜' in ver else 'F5F5F5'))
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.border = bd()
            else:
                D(cell, bg, align='center')
                if ci in [4,5,6] and isinstance(val, float): cell.number_format = '#,##0.00'
        wss.row_dimensions[r].height = 22
    
    sec(wss, 16, "C.  REFUNDS — Previous Period (REFUND action, arrive end of month)", 7, TEAL)
    for ci, h in enumerate(["Operator","# Refund Tx","Total EUP (NIS)","Effect","","",""], 1):
        H(wss.cell(row=17, column=ci, value=h), TEAL)
    p_ref = result['refunds'][result['refunds']['Operator']=='Partner']
    t_ref = result['refunds'][result['refunds']['Operator']=='012Talk']
    for i, (op, sub) in enumerate([('Partner', p_ref), ('012Talk', t_ref), ('TOTAL', result['refunds'])]):
        r = 18+i
        bg = LTEAL if i < 2 else GREEN
        is_tot = (i == 2)
        for ci, val in enumerate([op, len(sub), round(sub['End User Price'].sum(),2), "Credit — reduces amount owed","","",""], 1):
            cell = wss.cell(row=r, column=ci, value=val)
            if is_tot:
                H(cell, GREEN)
                if ci == 3: cell.number_format = '#,##0.00'
            else:
                D(cell, bg, align='left' if ci in [1,4] else 'center', bold=(ci==1))
                if ci == 3: cell.number_format = '#,##0.00'
        wss.row_dimensions[r].height = 18
    
    sec(wss, 22, "D.  NET BILLING SUMMARY", 7, GREEN)
    for ci, h in enumerate(["Item","Partner (NIS)","012Talk (NIS)","TOTAL (NIS)","","",""], 1):
        H(wss.cell(row=23, column=ci, value=h), GREEN)
    net = [
        ("Our EUP — DONE+CANCELLED", round(t['partner_eup'],2), round(t['talk012_eup'],2), round(t['partner_eup']+t['talk012_eup'],2), LBLUE),
        ("Refunds — credit back (prev. period)", round(t['partner_ref'],2), round(t['talk012_ref'],2), round(t['partner_ref']+t['talk012_ref'],2), LTEAL),
        ("PENDING_CANCELLATION (unconfirmed)", round(t['pending_eup'],2), 0.0, round(t['pending_eup'],2), LPURP),
        ("NET Our System Total (excl. PENDING)", round(t['partner_eup']+t['partner_ref'],2), round(t['talk012_eup']+t['talk012_ref'],2), round(t['partner_eup']+t['talk012_eup']+t['partner_ref']+t['talk012_ref'],2), LYELL),
        ("Supplier CBD — matched phones", "—", "—", round(t['sup_cbd'],2), LORAN),
    ]
    for i, row in enumerate(net):
        r = 24+i
        bg = row[-1]
        is_bold = (i == 3)
        for ci, val in enumerate(row[:-1], 1):
            cell = wss.cell(row=r, column=ci, value=val)
            D(cell, bg, align='left' if ci==1 else 'center', bold=is_bold or ci==1)
            if ci in [2,3,4] and isinstance(val, float): cell.number_format = '#,##0.00'
        for ci in range(len(row)-1+1, 8):
            wss.cell(row=r, column=ci).fill = PatternFill('solid', start_color=bg)
            wss.cell(row=r, column=ci).border = bd()
        wss.row_dimensions[r].height = 18
    
    wss.column_dimensions['A'].width = 36
    wss.column_dimensions['B'].width = 44
    for ci in range(3, 8):
        wss.column_dimensions[get_column_letter(ci)].width = 18
    wss.column_dimensions['G'].width = 40
    wss.freeze_panes = 'A3'
    
    # ---- SHEET 2: MATCHED ----
    ws_m = wb.create_sheet("Matched")
    num_cols_m = {'Supplier CBD (NIS)', 'Our EUP (NIS)', 'Difference (NIS)'}
    if len(result['matched']) > 0:
        ttl(ws_m, f"✅  MATCHED — {len(result['matched'])} records  |  CBD: {t['sup_cbd']:,.2f} NIS  |  EUP: {t['our_eup']:,.2f} NIS  |  Diff: {t['diff']:,.2f} NIS", len(result['matched'].columns), NAVY)
        last_r = write_df_to_sheet(ws_m, result['matched'], 2, NAVY, LBLUE, num_cols_m)
        # Total row
        tr = last_r
        for ci, col in enumerate(result['matched'].columns, 1):
            cell = ws_m.cell(row=tr, column=ci)
            if col == 'Supplier Tx ID': cell.value = 'TOTAL'
            elif col in num_cols_m: cell.value = round(result['matched'][col].sum(),2)
            H(cell, GREEN)
            if col in num_cols_m: cell.number_format = '#,##0.00'
        # Color diff cells
        diff_col = list(result['matched'].columns).index('Difference (NIS)') + 1
        for ri in range(len(result['matched'])):
            r = ri + 3
            diff_val = result['matched'].iloc[ri]['Difference (NIS)']
            if abs(diff_val) > 0.01:
                cell = ws_m.cell(row=r, column=diff_col)
                cell.font = Font(bold=True, color=RED if diff_val < 0 else GREEN, size=9, name='Arial')
    else:
        ttl(ws_m, "✅  MATCHED — No data", 12, NAVY)
    
    col_widths_m = [16,14,24,14,20,16,16,10,12,26,18,16]
    for ci, w in enumerate(col_widths_m, 1):
        ws_m.column_dimensions[get_column_letter(ci)].width = w
    ws_m.freeze_panes = 'A3'
    
    # ---- SHEET 3: PENDING ----
    ws_pnd = wb.create_sheet("Pending Cancellation")
    pnd_display = result['pending'][['phone_norm','Date & Time','Transaction ID','Operator','Eff_Status','Product Name','End User Price','Customer price','Customer name']].copy() if len(result['pending']) > 0 else pd.DataFrame()
    if len(pnd_display) > 0:
        pnd_display.columns = ['Phone','Date & Time','Our Tx ID','Operator','Status','Product Name','End User Price (NIS)','Customer Price (NIS)','Customer Name']
    ttl(ws_pnd, f"🕐  PENDING CANCELLATION — {t['pending_count']} record(s)  |  EUP: {t['pending_eup']:,.2f} NIS", max(len(pnd_display.columns)+1, 10) if len(pnd_display) else 10, PURPLE)
    # Info
    ws_pnd.merge_cells(f'A2:{get_column_letter(10)}2')
    ws_pnd['A2'].value = "ℹ️  Check next day: did supplier send REFUND (approved) or back to DONE (rejected)? Update Verified ✓ column."
    ws_pnd['A2'].font = Font(color=PURPLE, size=9, name='Arial')
    ws_pnd['A2'].fill = PatternFill('solid', start_color=LPURP)
    ws_pnd['A2'].alignment = Alignment(horizontal='left', vertical='center')
    ws_pnd['A2'].border = bd()
    ws_pnd.row_dimensions[2].height = 20
    
    if len(pnd_display) > 0:
        ncols = len(pnd_display.columns) + 1
        last_r_pnd = write_df_to_sheet(ws_pnd, pnd_display, 3, PURPLE, LPURP, {'End User Price (NIS)','Customer Price (NIS)'})
        # Add Verified column header
        H(ws_pnd.cell(row=3, column=ncols, value='Verified ✓'), YELL, fg='000000')
        # Add verified dropdown cells
        for ri in range(len(pnd_display)):
            r = ri + 4
            ver_cell = ws_pnd.cell(row=r, column=ncols, value='⬜ Not Verified')
            ver_cell.font = Font(bold=True, color='7F6000', size=9, name='Arial')
            ver_cell.fill = PatternFill('solid', start_color=YELL)
            ver_cell.alignment = Alignment(horizontal='center', vertical='center')
            ver_cell.border = bd()
        ver_col_letter = get_column_letter(ncols)
        dv = DataValidation(type="list",
            formula1='"⬜ Not Verified,✅ Confirmed REFUND — cancellation approved,🔄 Back to DONE — cancellation rejected"',
            allow_blank=False, showDropDown=False)
        dv.sqref = f"{ver_col_letter}4:{ver_col_letter}{len(pnd_display)+3}"
        ws_pnd.add_data_validation(dv)
    ws_pnd.freeze_panes = 'A4'
    
    # ---- SHEET 4: SUPPLIER ONLY ----
    ws_so = wb.create_sheet("Supplier Only")
    so_display = result['sup_only'][['phone_norm','Sup_Date','Package','Sup_TxID','CBD','Net_Total','Cust_Name']].copy() if len(result['sup_only']) > 0 else pd.DataFrame()
    if len(so_display) > 0:
        so_display.columns = ['Phone (Supplier)','Supplier Date','Package','Supplier Tx ID','CBD (NIS)','Net Total (NIS)','Customer Name']
    ttl(ws_so, f"❌  SUPPLIER ONLY — {t['sup_only_count']} phones  |  CBD: {t['sup_only_cbd']:,.2f} NIS", max(len(so_display.columns)+1,9) if len(so_display) else 9, RED)
    ws_so.merge_cells('A2:I2')
    ws_so['A2'].value = "ℹ️  These phones are in supplier report but NOT in our system. Usually from previous date."
    ws_so['A2'].font = Font(color=RED, size=9, name='Arial')
    ws_so['A2'].fill = PatternFill('solid', start_color=LRED)
    ws_so['A2'].alignment = Alignment(horizontal='left', vertical='center')
    ws_so['A2'].border = bd()
    ws_so.row_dimensions[2].height = 20
    
    if len(so_display) > 0:
        ncols_so = len(so_display.columns) + 1
        write_df_to_sheet(ws_so, so_display, 3, RED, LRED, {'CBD (NIS)','Net Total (NIS)'})
        H(ws_so.cell(row=3, column=ncols_so, value='Verified ✓'), YELL, fg='000000')
        for ri in range(len(so_display)):
            r = ri + 4
            ver_cell = ws_so.cell(row=r, column=ncols_so, value='⬜ Not Verified')
            ver_cell.font = Font(bold=True, color='7F6000', size=9, name='Arial')
            ver_cell.fill = PatternFill('solid', start_color=YELL)
            ver_cell.alignment = Alignment(horizontal='center', vertical='center')
            ver_cell.border = bd()
        ver_col_letter2 = get_column_letter(ncols_so)
        dv2 = DataValidation(type="list",
            formula1='"⬜ Not Verified,✅ Verified — matches prev. date,❌ Not found — investigate"',
            allow_blank=False, showDropDown=False)
        dv2.sqref = f"{ver_col_letter2}4:{ver_col_letter2}{len(so_display)+3}"
        ws_so.add_data_validation(dv2)
    ws_so.freeze_panes = 'A4'
    
    # ---- SHEET 5: OUR SYSTEM ONLY ----
    ws_oo = wb.create_sheet("Our System Only")
    oo_display = result['our_only'][['phone_norm','Date & Time','Transaction ID','Operator','Eff_Status','Product Name','End User Price','Customer name']].copy() if len(result['our_only']) > 0 else pd.DataFrame()
    if len(oo_display) > 0:
        oo_display.columns = ['Phone','Date & Time','Our Tx ID','Operator','Status','Product Name','End User Price (NIS)','Customer Name']
    ttl(ws_oo, f"⚠️  OUR SYSTEM ONLY — {t['our_only_count']} phones  |  EUP: {t['our_only_eup']:,.2f} NIS", max(len(oo_display.columns)+1,9) if len(oo_display) else 9, ORANGE)
    ws_oo.merge_cells('A2:I2')
    ws_oo['A2'].value = "ℹ️  DONE in our system but missing from supplier. Usually late end-of-day transactions."
    ws_oo['A2'].font = Font(color=ORANGE, size=9, name='Arial')
    ws_oo['A2'].fill = PatternFill('solid', start_color=LORAN)
    ws_oo['A2'].alignment = Alignment(horizontal='left', vertical='center')
    ws_oo['A2'].border = bd()
    ws_oo.row_dimensions[2].height = 20
    
    if len(oo_display) > 0:
        ncols_oo = len(oo_display.columns) + 1
        write_df_to_sheet(ws_oo, oo_display, 3, ORANGE, LORAN, {'End User Price (NIS)'})
        H(ws_oo.cell(row=3, column=ncols_oo, value='Verified ✓'), YELL, fg='000000')
        for ri in range(len(oo_display)):
            r = ri + 4
            ver_cell = ws_oo.cell(row=r, column=ncols_oo, value='⬜ Not Verified')
            ver_cell.font = Font(bold=True, color='7F6000', size=9, name='Arial')
            ver_cell.fill = PatternFill('solid', start_color=YELL)
            ver_cell.alignment = Alignment(horizontal='center', vertical='center')
            ver_cell.border = bd()
        ver_col_letter3 = get_column_letter(ncols_oo)
        dv3 = DataValidation(type="list",
            formula1='"⬜ Not Verified,✅ Verified — found in next day supplier report,❌ Not found — investigate"',
            allow_blank=False, showDropDown=False)
        dv3.sqref = f"{ver_col_letter3}4:{ver_col_letter3}{len(oo_display)+3}"
        ws_oo.add_data_validation(dv3)
    ws_oo.freeze_panes = 'A4'
    
    # ---- SHEET 6: REFUNDS ----
    ws_ref = wb.create_sheet("Refunds")
    ref_display = result['refunds'][['Operator','phone_norm','Date & Time','Transaction ID','Product Name','End User Price','Customer price']].copy() if len(result['refunds']) > 0 else pd.DataFrame()
    if len(ref_display) > 0:
        ref_display.columns = ['Operator','Phone','Date & Time','Our Tx ID','Product Name','End User Price (NIS)','Customer Price (NIS)']
    ttl(ws_ref, f"↩️  REFUNDS — Previous Period  |  {t['refunds_count']} records  |  EUP: {t['refunds_eup']:,.2f} NIS", max(len(ref_display.columns),7) if len(ref_display) else 7, TEAL)
    if len(ref_display) > 0:
        write_df_to_sheet(ws_ref, ref_display, 2, TEAL, LTEAL, {'End User Price (NIS)','Customer Price (NIS)'})
    ws_ref.freeze_panes = 'A3'
    
    # ---- SHEET 7: FAILED ----
    ws_f = wb.create_sheet("Failed")
    fail_display = result['failed'][['Operator','phone_norm','Date & Time','Transaction ID','Product Name','End User Price','Error description']].copy() if len(result['failed']) > 0 else pd.DataFrame()
    if len(fail_display) > 0:
        fail_display.columns = ['Operator','Phone','Date & Time','Our Tx ID','Product Name','End User Price','Error Description']
    ttl(ws_f, f"🔴  FAILED TRANSACTIONS — {t['failed_count']} records", max(len(fail_display.columns),7) if len(fail_display) else 7, '7F0000')
    if len(fail_display) > 0:
        write_df_to_sheet(ws_f, fail_display, 2, '7F0000', LRED)
    ws_f.freeze_panes = 'A3'
    
    # Save to buffer
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
    
    NAVY='1F3864'; LBLUE='DEEAF1'; GREEN='375623'; LGREEN='E2EFDA'
    RED='C00000'; LRED='FFE0E0'; WHITE='FFFFFF'
    
    # Title
    ws.merge_cells('A1:J1')
    ws['A1'].value = f"📅  MONTHLY SUMMARY — {month_label}"
    ws['A1'].font = Font(bold=True, color='FFFFFF', size=14, name='Arial')
    ws['A1'].fill = PatternFill('solid', start_color=NAVY)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].border = bd()
    ws.row_dimensions[1].height = 36
    
    headers = ['Date','Supplier CBD (NIS)','Our EUP (NIS)','Difference (NIS)',
               'Matched Phones','Supplier Only','Our Only','PENDING',
               'Refunds (NIS)','Net Billed (NIS)']
    for ci, h in enumerate(headers, 1):
        H(ws.cell(row=2, column=ci, value=h), NAVY, sz=10)
    ws.row_dimensions[2].height = 35
    
    for ri, rec in enumerate(history):
        r = ri + 3
        bg = LBLUE if ri % 2 == 0 else WHITE
        vals = [
            rec.get('date',''), rec.get('sup_cbd',0), rec.get('our_eup',0),
            rec.get('diff',0), rec.get('matched_count',0), rec.get('sup_only_count',0),
            rec.get('our_only_count',0), rec.get('pending_count',0),
            rec.get('refunds_eup',0), rec.get('net_billed',0)
        ]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=ci, value=val)
            D(cell, bg, fmt='#,##0.00' if ci in [2,3,4,9,10] else None)
            if ci == 4 and isinstance(val, (int,float)):
                if abs(val) > 0.01:
                    cell.font = Font(bold=True, color=RED if val < 0 else '375623', size=9, name='Arial')
        ws.row_dimensions[r].height = 18
    
    # Total row
    tr = len(history) + 3
    for ci, val in enumerate(['TOTAL MONTH']+[f'=SUM({get_column_letter(c)}3:{get_column_letter(c)}{tr-1})' for c in range(2,11)], 1):
        cell = ws.cell(row=tr, column=ci, value=val)
        H(cell, GREEN, sz=11)
        if ci > 1: cell.number_format = '#,##0.00'
    ws.row_dimensions[tr].height = 24
    
    widths = [14,20,18,18,14,14,12,12,16,18]
    for ci, w in enumerate(widths, 1):
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
    # Header
    st.markdown("""
    <div class="main-header">
        <h1 style="margin:0; font-size:28px;">📊 Reconciliation Dashboard</h1>
        <p style="margin:5px 0 0 0; opacity:0.85; font-size:14px;">Supplier vs Our System (Partner + 012Talk)</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown("### 📋 Navigation")
        page = st.radio("Navigation Menu", ["🔄 Daily Reconciliation", "📅 Monthly Summary", "ℹ️ Instructions"], 
                        label_visibility="collapsed")
        
        st.markdown("---")
        st.markdown("### 📊 History")
        history = load_history()
        if history:
            st.success(f"✅ {len(history)} days recorded")
            last = history[-1]
            st.info(f"Last: {last.get('date','N/A')}")
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
            sup_file = st.file_uploader("Supplier file", type=['xls','xlsx','csv'], 
                                         label_visibility="collapsed", key="sup")
            st.caption("File from supplier — Charge Before Discount (col M)")
        with col2:
            st.markdown("**2️⃣ Partner File (.csv)**")
            part_file = st.file_uploader("Partner file", type=['csv','xlsx'],
                                          label_visibility="collapsed", key="part")
            st.caption("Our system export — Partner operator")
        with col3:
            st.markdown("**3️⃣ 012Talk File (.csv)**")
            talk_file = st.file_uploader("012Talk file", type=['csv','xlsx'],
                                          label_visibility="collapsed", key="talk")
            st.caption("Our system export — 012Talk operator")
        
        report_date = st.date_input("📅 Report Date", value=date.today(),
                                    help="Date is auto-detected from supplier file. You can override manually.")
        
        if sup_file and part_file and talk_file:
            if st.button("▶ Run Reconciliation", type="primary", use_container_width=True):
                with st.spinner("Processing reconciliation..."):
                    # Load files
                    sup_bytes = sup_file.read()
                    sup_df, err1 = load_supplier(sup_bytes, sup_file.name)
                    part_df, err2 = load_our(part_file.read(), 'Partner')
                    talk_df, err3 = load_our(talk_file.read(), '012Talk')
                    
                    
                    if err1: st.error(f"❌ Supplier file error: {err1}"); return
                    if err2: st.error(f"❌ Partner file error: {err2}"); return
                    if err3: st.error(f"❌ 012Talk file error: {err3}"); return
                    
                    # Auto-detect date from supplier file
                    auto_date = str(report_date)
                    if sup_df is not None and 'Sup_Date' in sup_df.columns and len(sup_df) > 0:
                        try:
                            raw_date = sup_df['Sup_Date'].iloc[0]
                            parsed = pd.to_datetime(raw_date, dayfirst=True)
                            auto_date = parsed.strftime('%Y-%m-%d')
                            st.info(f"📅 Date auto-detected from supplier file: **{auto_date}**")
                        except:
                            pass
                    
                    result = run_reconciliation(sup_df, part_df, talk_df)
                    t = result['totals']
                    
                    st.session_state['result'] = result
                    st.session_state['report_date'] = auto_date
                    st.success("✅ Reconciliation complete!")
        
        # Show results if available
        if 'result' in st.session_state:
            result = st.session_state['result']
            t = result['totals']
            report_date_str = st.session_state.get('report_date', str(date.today()))
            
            st.markdown("---")
            st.markdown("### 📊 Results")
            
            # Key metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("✅ Matched", f"{t['matched_count']:,}", 
                          delta=f"Diff: {t['diff']:.2f} NIS" if abs(t['diff']) > 0.01 else "0.00 NIS ✓")
            with col2:
                st.metric("❌ Supplier Only", t['sup_only_count'],
                          delta=f"{t['sup_only_cbd']:.2f} NIS" if t['sup_only_count'] > 0 else None,
                          delta_color="inverse")
            with col3:
                st.metric("⚠️ Our Only", t['our_only_count'],
                          delta=f"{t['our_only_eup']:.2f} NIS" if t['our_only_count'] > 0 else None,
                          delta_color="inverse")
            with col4:
                st.metric("🕐 Pending", t['pending_count'],
                          delta=f"{t['pending_eup']:.2f} NIS" if t['pending_count'] > 0 else None,
                          delta_color="off")
            
            # Summary table
            st.markdown("#### B. Reconciliation Results")
            summary_data = {
                'Category': ['✅ MATCHED', '❌ SUPPLIER ONLY', '⚠️ OUR SYSTEM ONLY', '🕐 PENDING'],
                '# Phones': [t['matched_count'], t['sup_only_count'], t['our_only_count'], t['pending_count']],
                'Supplier CBD (NIS)': [round(t['sup_cbd'],2), round(t['sup_only_cbd'],2), '—', '—'],
                'Our EUP (NIS)': [round(t['our_eup'],2), '—', round(t['our_only_eup'],2), round(t['pending_eup'],2)],
                'Difference (NIS)': [round(t['diff'],2), '—', '—', '—'],
                'Status': [
                    '✅ Perfect match' if abs(t['diff']) < 0.01 else f'⚠️ Diff: {t["diff"]:.2f}',
                    'N/A' if t['sup_only_count'] == 0 else '⬜ Verify in Excel',
                    'N/A' if t['our_only_count'] == 0 else '⬜ Verify in Excel',
                    'None' if t['pending_count'] == 0 else '⬜ Check next day'
                ]
            }
            st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)
            
            # Net billing
            st.markdown("#### D. Net Billing Summary")
            net_data = {
                'Item': ['Our EUP (DONE+CANCELLED)', 'Refunds (credit back)', 'PENDING (unconfirmed)', 'NET Our System Total', 'Supplier CBD (matched)'],
                'Partner (NIS)': [round(t['partner_eup'],2), round(t['partner_ref'],2), round(t['pending_eup'],2), round(t['partner_eup']+t['partner_ref'],2), '—'],
                '012Talk (NIS)': [round(t['talk012_eup'],2), round(t['talk012_ref'],2), 0, round(t['talk012_eup']+t['talk012_ref'],2), '—'],
                'TOTAL (NIS)': [round(t['partner_eup']+t['talk012_eup'],2), round(t['partner_ref']+t['talk012_ref'],2), round(t['pending_eup'],2), round(t['partner_eup']+t['talk012_eup']+t['partner_ref']+t['talk012_ref'],2), round(t['sup_cbd'],2)],
            }
            st.dataframe(pd.DataFrame(net_data), use_container_width=True, hide_index=True)
            
            # Tabs for details
            tabs = st.tabs(["✅ Matched", "❌ Supplier Only", "⚠️ Our Only", "🕐 Pending", "↩️ Refunds", "🔴 Failed"])
            with tabs[0]:
                if len(result['matched']) > 0:
                    st.dataframe(result['matched'], use_container_width=True, hide_index=True)
                else:
                    st.info("No matched records")
            with tabs[1]:
                if len(result['sup_only']) > 0:
                    st.dataframe(result['sup_only'][['phone_norm','Sup_Date','Package','Sup_TxID','CBD']], use_container_width=True, hide_index=True)
                else:
                    st.success("✅ No supplier-only records!")
            with tabs[2]:
                if len(result['our_only']) > 0:
                    st.dataframe(result['our_only'][['phone_norm','Date & Time','Operator','Product Name','End User Price']], use_container_width=True, hide_index=True)
                else:
                    st.success("✅ No our-only records!")
            with tabs[3]:
                if len(result['pending']) > 0:
                    st.warning(f"⚠️ {t['pending_count']} transaction(s) pending supplier decision. Check next day!")
                    st.dataframe(result['pending'][['phone_norm','Date & Time','Transaction ID','Product Name','End User Price']], use_container_width=True, hide_index=True)
                else:
                    st.success("✅ No pending transactions!")
            with tabs[4]:
                if len(result['refunds']) > 0:
                    st.info(f"↩️ {t['refunds_count']} refunds — {t['refunds_eup']:.2f} NIS credit (arrive end of month)")
                    st.dataframe(result['refunds'][['Operator','phone_norm','Date & Time','Product Name','End User Price']], use_container_width=True, hide_index=True)
                else:
                    st.info("No refunds today")
            with tabs[5]:
                if len(result['failed']) > 0:
                    st.error(f"🔴 {t['failed_count']} failed transactions")
                    st.dataframe(result['failed'][['Operator','phone_norm','Date & Time','Product Name','Error description']], use_container_width=True, hide_index=True)
                else:
                    st.success("✅ No failed transactions!")
            
            st.markdown("---")
            col1, col2 = st.columns(2)
            
            with col1:
                # Download Excel
                excel_buf = create_excel_report(result, report_date_str)
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_buf,
                    file_name=f"Reconciliation_{report_date_str.replace('-','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            
            with col2:
                # Save to history
                if st.button("💾 Save to Monthly History", use_container_width=True):
                    record = {
                        'date': report_date_str,
                        'sup_cbd': round(t['sup_cbd'],2),
                        'our_eup': round(t['our_eup'],2),
                        'diff': round(t['diff'],2),
                        'matched_count': t['matched_count'],
                        'sup_only_count': t['sup_only_count'],
                        'our_only_count': t['our_only_count'],
                        'pending_count': t['pending_count'],
                        'refunds_eup': round(t['refunds_eup'],2),
                        'net_billed': round(t['partner_eup']+t['talk012_eup']+t['partner_ref']+t['talk012_ref'],2),
                    }
                    add_to_history(record)
                    st.success(f"✅ Saved to monthly history!")
    
    # ============================================================
    # PAGE: MONTHLY SUMMARY
    # ============================================================
    elif page == "📅 Monthly Summary":
        st.markdown("## 📅 Monthly Summary")
        
        history = load_history()
        
        if not history:
            st.info("No history yet. Run daily reconciliations and click 'Save to Monthly History' to build history.")
            return
        
        # Filter by month
        months = sorted(set(h['date'][:7] for h in history), reverse=True)
        selected_month = st.selectbox("Select Month", months)
        
        month_history = [h for h in history if h['date'].startswith(selected_month)]
        
        if not month_history:
            st.warning("No data for selected month")
            return
        
        # Summary metrics
        total_sup = sum(h.get('sup_cbd',0) for h in month_history)
        total_eup = sum(h.get('our_eup',0) for h in month_history)
        total_diff = sum(h.get('diff',0) for h in month_history)
        total_refunds = sum(h.get('refunds_eup',0) for h in month_history)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("📅 Days Recorded", len(month_history))
        with col2: st.metric("💰 Supplier CBD Total", f"{total_sup:,.2f} NIS")
        with col3: st.metric("💰 Our EUP Total", f"{total_eup:,.2f} NIS")
        with col4: st.metric("↩️ Total Refunds", f"{total_refunds:,.2f} NIS")
        
        # Monthly table
        df_month = pd.DataFrame(month_history)
        st.dataframe(df_month, use_container_width=True, hide_index=True)
        
        # Download monthly Excel
        month_label = datetime.strptime(selected_month, '%Y-%m').strftime('%B %Y')
        monthly_buf = create_monthly_excel(month_history, month_label)
        st.download_button(
            label=f"📥 Download Monthly Report — {month_label}",
            data=monthly_buf,
            file_name=f"Monthly_Summary_{selected_month.replace('-','_')}.xlsx",
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
        - Upload the supplier `.xls` file (NOT .csv — keep original format)
        - Upload Partner `.csv` from our system
        - Upload 012Talk `.csv` from our system
        
        **Step 2 — Run**
        - Select the report date
        - Click **▶ Run Reconciliation**
        
        **Step 3 — Review Results**
        - Check Summary table — if Difference = 0.00 NIS ✅ all good
        - Check **Pending** tab — update verification next day
        - Check **Supplier Only** tab — usually previous date transactions
        - Check **Our Only** tab — usually late end-of-day transactions
        
        **Step 4 — Download & Save**
        - Click **📥 Download Excel Report** — save to SharePoint folder
        - Click **💾 Save to Monthly History** — adds to monthly tracker
        
        ---
        
        ### 📅 Monthly Report
        - Go to **📅 Monthly Summary** in sidebar
        - Select the month
        - Click **📥 Download Monthly Report**
        
        ---
        
        ### ⚠️ Important Rules
        - Always use original `.xls` supplier file (not CSV — phones get truncated!)
        - PENDING_CANCELLATION = check next day if supplier approved or rejected
        - REFUND rows = credits from previous period, arrive end of month
        - REWARD and REFUND_REWARD rows are automatically excluded
        """)

if __name__ == "__main__":
    main()
