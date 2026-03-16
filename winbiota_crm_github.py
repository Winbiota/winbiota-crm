import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime, os, json, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

creds_json = json.loads(os.environ['GOOGLE_CREDENTIALS'])
SHEET_ID   = os.environ['SHEET_ID']
EMAIL_USER = os.environ['EMAIL_USER']
EMAIL_PASS = os.environ['EMAIL_PASS']
OUTPUT     = '/tmp/Winbiota_CRM_Report.xlsx'

scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
creds  = Credentials.from_service_account_info(creds_json, scopes=scopes)
sheets = build('sheets', 'v4', credentials=creds)
result = sheets.spreadsheets().values().get(
    spreadsheetId=SHEET_ID,
    range='CRM Winbiota',
    valueRenderOption='FORMATTED_VALUE',
    dateTimeRenderOption='FORMATTED_STRING'
).execute()
all_values = result.get('values', [])
max_cols = max(len(r) for r in all_values)
all_values = [r + [''] * (max_cols - len(r)) for r in all_values]
df_raw = pd.DataFrame(all_values)

data = df_raw.iloc[3:].copy()
data.columns = range(len(data.columns))
data = data.replace('', pd.NA)
data = data[~data[2].astype(str).str.contains('dummy data', na=False)]
data = data[data[2].astype(str).str.strip().replace('nan','') != ''].reset_index(drop=True)
N = len(data)

sel = data[[1,2,4,8,9,10,11,12,14,15,16,20]].copy()
sel.columns = ['Fecha','Nombre','Tel','Contestaron Whats','Comunicacion',
               'Fecha 1era Llamada','Contesto Llamada 1','Estatus 1era Llamada',
               'Fecha 2nda Llamada','Contesto Llamada 2','Estatus 2nda Llamada','Nota']
sel['Tel'] = sel['Tel'].astype(str).str.replace('p:','',regex=False).str.strip().replace('nan','')
sel = sel.fillna('')

fecha_lead = pd.to_datetime(data[1], errors='coerce', dayfirst=True).dt.date
estatus1   = data[12].astype(str).str.strip().str.upper()
estatus2   = data[16].astype(str).str.strip().str.upper()

buenos_vals1 = ['CONTESTO','CONTESTA','INTERES','LLAMADA','NO INT','SI']
buenos_vals2 = ['SI CONTESTA','INTERESADA']
buenos1 = estatus1.isin(buenos_vals1)
buenos2 = estatus2.isin(buenos_vals2)

leads_day    = fecha_lead.dropna().value_counts().sort_index()
data['semana'] = pd.to_datetime(data[1], errors='coerce', dayfirst=True).dt.to_period('W')
leads_semana   = data.groupby('semana').size()

print(f"N={N}, Buenos1={buenos1.sum()}, Buenos2={buenos2.sum()}")

G_D='1B5E20'; G_M='2E7D32'; G_L='E8F5E9'; G_LL='F1F8E9'
WHITE='FFFFFF'; GRAY='F5F5F5'; AMBER='FFF8E1'; BLUE_D='1565C0'

def st(cell, bold=False, sz=9, col="000000", bg=None, ha='left', va='center', wrap=False):
    cell.font = Font(name='Arial', bold=bold, size=sz, color=col)
    if bg: cell.fill = PatternFill('solid', start_color=bg)
    cell.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)

thin  = Side(style='thin', color='C8E6C9')
bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)
thin2 = Side(style='thin', color='B0BEC5')
bdr2  = Border(left=thin2, right=thin2, top=thin2, bottom=thin2)

def hrow(ws, row, labels, widths, bg=G_M):
    for i,(h,w) in enumerate(zip(labels,widths),1):
        c = ws.cell(row=row, column=i, value=h)
        st(c, bold=True, sz=9, col=WHITE, bg=bg, ha='center', wrap=True)
        c.border = bdr
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30

def trow(ws, row, text, ncols, bg=G_D):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    st(c, bold=True, sz=12, col=WHITE, bg=bg, ha='center')
    ws.row_dimensions[row].height = 26

def section(ws, row, text, ncols, bg=BLUE_D):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    st(c, bold=True, sz=10, col=WHITE, bg=bg, ha='left')
    ws.row_dimensions[row].height = 20

def kpi(ws, row, value, label, fmt='number', bg=WHITE, ncols=10):
    c = ws.cell(row=row, column=1, value=value)
    c.number_format = '0.0%' if fmt=='pct' else '0'
    st(c, bold=True, sz=13, col=G_D, bg=bg, ha='center')
    c.border = bdr2
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=ncols)
    st(ws.cell(row=row, column=2, value=label), sz=10, bg=bg)
    ws.row_dimensions[row].height = 22

wb = openpyxl.Workbook()

ws1 = wb.active
ws1.title = "Datos"
NCOLS1 = 12
trow(ws1, 1, 'WINBIOTA - Seguimiento Leads CRM', NCOLS1)
h1 = ['Fecha','Nombre','Telefono','Contestaron Whats','Comunicacion',
      'Fecha 1era Llamada','Contesto Llamada 1','Estatus 1era Llamada',
      'Fecha 2nda Llamada','Contesto Llamada 2','Estatus 2nda Llamada','Nota']
w1 = [12,26,14,12,22,13,10,20,13,10,20,26]
hrow(ws1, 2, h1, w1)
for ri,(_, row) in enumerate(sel.iterrows(), 3):
    bg = G_L if ri%2==0 else WHITE
    for ci, val in enumerate(row.values, 1):
        v = str(val) if str(val) not in ['','nan','NaT','<NA>'] else ''
        c = ws1.cell(row=ri, column=ci, value=v)
        st(c, sz=9, bg=bg, wrap=(ci in [4,5,8,12]))
        c.border = bdr
    ws1.row_dimensions[ri].height = 16
ws1.freeze_panes = 'A3'

ws2 = wb.create_sheet("Estadisticas")
NCOLS2 = 10
DR = 2 + N
s1 = "'Datos'"
def ref(col, r1=3, r2=DR): return f"{s1}!{col}{r1}:{col}{r2}"

trow(ws2, 1, 'WINBIOTA - Estadisticas & KPIs', NCOLS2)
for i,w in enumerate([14,30,12,12,12,11,11,13,14,13],1):
    ws2.column_dimensions[get_column_letter(i)].width = w

r = 3
section(ws2, r, 'TOTALES', NCOLS2); r+=1

buenos_f  = '+'.join([f'COUNTIF({ref("H")},"{v}")' for v in buenos_vals1])
buenos2_f = '+'.join([f'COUNTIF({ref("K")},"{v}")' for v in buenos_vals2])
rng = ref("E")
precio_f = (
    f'=SUMPRODUCT(((ISNUMBER(SEARCH("PRECIO",{rng}))'
    f'+ISNUMBER(SEARCH("CARO",{rng}))'
    f'+ISNUMBER(SEARCH("VISTO",{rng}))'
    f'+ISNUMBER(SEARCH("DINERO",{rng}))'
    f'+ISNUMBER(SEARCH("PAGA",{rng})))>0)'
    f'*(ISNUMBER(SEARCH("SIN PRECIO",{rng}))=FALSE))'
)

rows_totales = [
    (f'=COUNTA({ref("B")})',  'Total leads',                         'number', G_LL),
    (f'=COUNTA({ref("D")})',  'Contestaron WhatsApp',                'number', WHITE),
    (f'=COUNTA({ref("G")})',  'Llamadas 1 realizadas',               'number', G_LL),
    (f'={buenos_f}',          'Llamada 1 efectiva (buenos estatus)', 'number', WHITE),
    (f'=COUNTA({ref("J")})',  'Llamadas 2 realizadas',               'number', G_LL),
    (f'={buenos2_f}',         'Llamada 2 efectiva (buenos estatus)', 'number', WHITE),
    (precio_f,                'Piden precio / bloqueo economico',    'number', AMBER),
]
total_rows = {}
keys = ['total_leads','contestaron_whats','ll1_hechas','buenos1','ll2_hechas','buenos2','precio']
for i,key in enumerate(keys): total_rows[key] = r + i
for val, label, fmt, bg in rows_totales:
    kpi(ws2, r, val, label, fmt, bg, NCOLS2); r+=1
r+=1

section(ws2, r, 'PORCENTAJES', NCOLS2); r+=1
tr = total_rows
pct_rows = [
    (f'=IFERROR(A{tr["contestaron_whats"]}/A{tr["total_leads"]},0)',                              'Contestaron Whats / Total leads', GRAY),
    (f'=IFERROR(A{tr["ll1_hechas"]}/A{tr["contestaron_whats"]},0)',                               'Llamadas 1 hechas / Contestaron Whats', WHITE),
    (f'=IFERROR(A{tr["buenos1"]}/A{tr["ll1_hechas"]},0)',                                         'Llamada 1 efectiva (buenos) / Llamadas 1 hechas', GRAY),
    (f'=IFERROR(A{tr["ll2_hechas"]}/A{tr["ll1_hechas"]},0)',                                      'Llamadas 2 hechas / Llamadas 1 hechas', WHITE),
    (f'=IFERROR(A{tr["buenos2"]}/A{tr["ll2_hechas"]},0)',                                         'Llamada 2 efectiva (buenos) / Llamadas 2 hechas', GRAY),
    (f'=IFERROR((A{tr["buenos1"]}+A{tr["buenos2"]})/(A{tr["ll1_hechas"]}+A{tr["ll2_hechas"]}),0)','% Efectividad total', WHITE),
    (f'=IFERROR(A{tr["precio"]}/A{tr["total_leads"]},0)',                                         '% leads bloqueo economico', AMBER),
]
for val, label, bg in pct_rows:
    c = ws2.cell(row=r, column=1, value=val)
    c.number_format = '0.0%'
    st(c, bold=True, sz=13, col=G_D, bg=bg, ha='center')
    c.border = bdr2
    ws2.merge_cells(start_row=r, start_column=2, end_row=r, end_column=NCOLS2)
    st(ws2.cell(row=r, column=2, value=label), sz=10, bg=bg)
    ws2.row_dimensions[r].height = 22; r+=1
r+=1

section(ws2, r, 'LEADS NUEVOS POR DIA', NCOLS2); r+=1
hrow(ws2, r, ['Fecha','Leads nuevos']+['']*8, [13,12]+[12]*8); r+=1
total_dia=0
for date in sorted(leads_day.index):
    bg = G_L if r%2==0 else WHITE
    count = leads_day[date]; total_dia += count
    try: ds = datetime.datetime.strptime(str(date),'%Y-%m-%d').strftime('%d/%m/%Y')
    except: ds = str(date)
    c1 = ws2.cell(row=r, column=1, value=ds)
    st(c1, sz=9, bg=bg); c1.border = bdr2
    c2 = ws2.cell(row=r, column=2, value=int(count))
    st(c2, sz=10, bg=bg, ha='center'); c2.border = bdr2
    for ci in range(3, NCOLS2+1):
        ws2.cell(row=r, column=ci).fill = PatternFill('solid', start_color=bg)
    ws2.row_dimensions[r].height = 16; r+=1
for ci,val in enumerate(['TOTAL',int(total_dia)]+['']*8,1):
    c = ws2.cell(row=r, column=ci, value=val)
    st(c, bold=True, sz=10, col=WHITE, bg=G_M, ha='center'); c.border = bdr2
ws2.row_dimensions[r].height = 20; r+=2

section(ws2, r, 'LEADS NUEVOS POR SEMANA', NCOLS2); r+=1
hrow(ws2, r, ['Semana','Leads nuevos']+['']*8, [20,12]+[12]*8); r+=1
total_sem=0
for period, count in leads_semana.items():
    bg = G_L if r%2==0 else WHITE
    try: label = f"{period.start_time.strftime('%d/%m/%Y')} - {period.end_time.strftime('%d/%m/%Y')}"
    except: label = str(period)
    c1 = ws2.cell(row=r, column=1, value=label)
    st(c1, sz=9, bg=bg); c1.border = bdr2
    c2 = ws2.cell(row=r, column=2, value=int(count))
    st(c2, sz=10, bg=bg, ha='center'); c2.border = bdr2
    for ci in range(3, NCOLS2+1):
        ws2.cell(row=r, column=ci).fill = PatternFill('solid', start_color=bg)
    total_sem += count
    ws2.row_dimensions[r].height = 16; r+=1
for ci,val in enumerate(['TOTAL',int(total_sem)]+['']*8,1):
    c = ws2.cell(row=r, column=ci, value=val)
    st(c, bold=True, sz=10, col=WHITE, bg=G_M, ha='center'); c.border = bdr2
ws2.row_dimensions[r].height = 20

wb.save(OUTPUT)
print(f"Listo. N={N}")

now = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
msg = MIMEMultipart()
msg['From'] = EMAIL_USER
msg['To']   = EMAIL_USER
msg['Subject'] = f'Winbiota CRM - Reporte {now}'
msg.attach(MIMEText(f"Reporte CRM generado el {now}.\nLeads: {N} | Buenos1: {buenos1.sum()} | Buenos2: {buenos2.sum()}", 'plain'))
with open(OUTPUT, 'rb') as f:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename="Winbiota_CRM_{datetime.datetime.now().strftime("%Y%m%d_%H%M")}.xlsx"')
msg.attach(part)
with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
    server.login(EMAIL_USER, EMAIL_PASS)
    server.sendmail(EMAIL_USER, EMAIL_USER, msg.as_string())
print("Email enviado.")
