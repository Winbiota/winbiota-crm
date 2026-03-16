import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
import gspread
from google.oauth2.service_account import Credentials
import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# ── Credentials & config ──────────────────────────────────────────
creds_json = json.loads(os.environ['GOOGLE_CREDENTIALS'])
SHEET_ID   = os.environ['SHEET_ID']
EMAIL_USER = os.environ['EMAIL_USER']
EMAIL_PASS = os.environ['EMAIL_PASS']
EMAIL_TO   = EMAIL_USER
OUTPUT     = '/tmp/Winbiota_CRM_Report.xlsx'

# ── Read Google Sheet ─────────────────────────────────────────────
scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
creds  = Credentials.from_service_account_info(creds_json, scopes=scopes)
gc     = gspread.authorize(creds)
ws_src = gc.open_by_key(SHEET_ID).worksheet('CRM Winbiota')

all_values = ws_src.get_all_values()
df_raw = pd.DataFrame(all_values)

data = df_raw.iloc[3:].copy()
data.columns = range(len(data.columns))
data = data[~data[2].astype(str).str.contains('dummy data', na=False)]
data = data[data[2].astype(str).str.strip() != ''].reset_index(drop=True)
N = len(data)

sel = data[[1,2,4,8,9,10,11,12,14,15,16,20]].copy()
sel.columns = ['Fecha','Nombre','Tel','Contestaron Whats','Comunicación',
               'Fecha 1era Llamada','Contesto Llamada 1','Estatus 1era Llamada',
               'Fecha 2nda Llamada','Contesto Llamada 2','Estatus 2nda Llamada','Nota']
sel['Tel'] = sel['Tel'].astype(str).str.replace('p:','',regex=False).str.strip().replace('nan','')
for col in ['Fecha','Fecha 1era Llamada','Fecha 2nda Llamada']:
    sel[col] = pd.to_datetime(sel[col], errors='coerce').dt.strftime('%d/%m/%Y').replace('NaT','')
sel = sel.fillna('')

fecha_lead = pd.to_datetime(data[1],  errors='coerce').dt.date
fecha_ll1  = pd.to_datetime(data[10], errors='coerce').dt.date
fecha_ll2  = pd.to_datetime(data[14], errors='coerce').dt.date
estatus1   = data[12].astype(str).str.strip().str.upper()
estatus2   = data[16].astype(str).str.strip().str.upper()

buenos_vals1 = ['CONTESTO','CONTESTA','INTERES','LLAMADA','NO INT','SI']
buenos_vals2 = ['SI CONTESTA','INTERESADA']
buenos1 = estatus1.isin(buenos_vals1)
buenos2 = estatus2.isin(buenos_vals2)

ll1_day        = fecha_ll1.value_counts().sort_index()
ll2_day        = fecha_ll2.value_counts().sort_index()
ll1_buenos_day = fecha_ll1[buenos1].value_counts().sort_index()
ll2_buenos_day = fecha_ll2[buenos2].value_counts().sort_index()
leads_day      = fecha_lead.value_counts().sort_index()
data['semana'] = pd.to_datetime(data[1], errors='coerce').dt.to_period('W')
leads_semana   = data.groupby('semana').size()

print(f"N={N}, Buenos1={buenos1.sum()}, Buenos2={buenos2.sum()}")

# ── Styles ────────────────────────────────────────────────────────
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

# ── SHEET 1 ───────────────────────────────────────────────────────
ws1 = wb.active
ws1.title = "Datos"
NCOLS1 = 12
trow(ws1, 1, 'WINBIOTA — Seguimiento Leads CRM', NCOLS1)
h1 = ['Fecha','Nombre','Teléfono','Contestaron\nWhats','Comunicación',
      'Fecha\n1era Llamada','Contesto\nLlamada 1','Estatus\n1era Llamada',
      'Fecha\n2nda Llamada','Contesto\nLlamada 2','Estatus\n2nda Llamada','Nota']
w1 = [12,26,14,12,22,13,10,20,13,10,20,26]
hrow(ws1, 2, h1, w1)
for ri,(_, row) in enumerate(sel.iterrows(), 3):
    bg = G_L if ri%2==0 else WHITE
    for ci, val in enumerate(row.values, 1):
        v = str(val) if str(val) not in ['','nan','NaT'] else ''
        c = ws1.cell(row=ri, column=ci, value=v)
        st(c, sz=9, bg=bg, wrap=(ci in [4,5,8,12]))
        c.border = bdr
    ws1.row_dimensions[ri].height = 16
ws1.freeze_panes = 'A3'

# ── SHEET 2 ───────────────────────────────────────────────────────
ws2 = wb.create_sheet("Estadisticas")
NCOLS2 = 10
DR = 2 + N
s1 = "'Datos'"
def ref(col, r1=3, r2=DR): return f"{s1}!{col}{r1}:{col}{r2}"

trow(ws2, 1, 'WINBIOTA — Estadísticas & KPIs', NCOLS2)
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
    (f'=COUNTA({ref("B")})',   'Total leads',                         'number', G_LL),
    (f'=COUNTA({ref("D")})',   'Contestaron WhatsApp',                'number', WHITE),
    (f'=COUNTA({ref("F")})',   'Llamadas 1 realizadas',               'number', G_LL),
    (f'={buenos_f}',           'Llamada 1 efectiva (buenos estatus)', 'number', WHITE),
    (f'=COUNTA({ref("I")})',   'Llamadas 2 realizadas',               'number', G_LL),
    (f'={buenos2_f}',          'Llamada 2 efectiva (buenos estatus)', 'number', WHITE),
    (precio_f,                 'Piden precio / bloqueo económico',    'number', AMBER),
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
    (f'=IFERROR(A{tr["contestaron_whats"]}/A{tr["total_leads"]},0)',           'Contestaron Whats / Total leads', GRAY),
    (f'=IFERROR(A{tr["ll1_hechas"]}/A{tr["contestaron_whats"]},0)',            'Llamadas 1 hechas / Contestaron Whats', WHITE),
    (f'=IFERROR(A{tr["buenos1"]}/A{tr["ll1_hechas"]},0)',                      'Llamada 1 efectiva (buenos) / Llamadas 1 hechas', GRAY),
    (f'=IFERROR(A{tr["ll2_hechas"]}/A{tr["ll1_hechas"]},0)',                   'Llamadas 2 hechas / Llamadas 1 hechas', WHITE),
    (f'=IFERROR(A{tr["buenos2"]}/A{tr["ll2_hechas"]},0)',                      'Llamada 2 efectiva (buenos) / Llamadas 2 hechas', GRAY),
    (f'=IFERROR((A{tr["buenos1"]}+A{tr["buenos2"]})/(A{tr["ll1_hechas"]}+A{tr["ll2_hechas"]}),0)', '% Efectividad total (buenos ll1+ll2) / (ll1+ll2 hechas)', WHITE),
    (f'=IFERROR(A{tr["precio"]}/A{tr["total_leads"]},0)',                      '% leads bloqueo económico — excl. "sin precio"', AMBER),
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

section(ws2, r, 'LLAMADAS POR DIA', NCOLS2); r+=1
hrow(ws2, r,
     ['Fecha','LL1\nhechas','LL1\nbuenos','% Efect.\n1','LL2\nhechas','LL2\nbuenos','% Efect.\n2',
      'LL tot.\nhechas','Efectivas\ntotales','% Efect.\ntotal'],
     [13,10,10,10,10,10,10,12,13,12]); r+=1

all_d = sorted(set(list(ll1_day.index)+list(ll2_day.index)))
tot_ll1=tot_ll2=tot_b1=tot_b2=0
for date in all_d:
    bg  = G_L if r%2==0 else WHITE
    ll1 = ll1_day.get(date,0); ll2 = ll2_day.get(date,0)
    b1  = ll1_buenos_day.get(date,0); b2 = ll2_buenos_day.get(date,0)
    tot_ll1+=ll1; tot_ll2+=ll2; tot_b1+=b1; tot_b2+=b2
    try: ds = datetime.datetime.strptime(str(date),'%Y-%m-%d').strftime('%d/%m/%Y')
    except: ds = str(date)
    vals = [ds, ll1, b1, b1/ll1 if ll1>0 else 0, ll2, b2, b2/ll2 if ll2>0 else 0]
    for ci,val in enumerate(vals,1):
        c = ws2.cell(row=r, column=ci, value=val)
        if ci in [4,7]: c.number_format='0%'
        st(c, sz=9, bg=bg, ha='center' if ci>1 else 'left'); c.border = bdr2
    tot_d=ll1+ll2; btot_d=b1+b2
    for ci,val in enumerate([tot_d, btot_d, btot_d/tot_d if tot_d>0 else 0], 8):
        c = ws2.cell(row=r, column=ci, value=val)
        if ci==10: c.number_format='0%'
        st(c, bold=True, sz=9, bg=bg, ha='center'); c.border = bdr2
    ws2.row_dimensions[r].height = 16; r+=1

tot_d=tot_ll1+tot_ll2; btot=tot_b1+tot_b2
vals_tot = ['TOTAL', tot_ll1, tot_b1, tot_b1/tot_ll1 if tot_ll1>0 else 0,
            tot_ll2, tot_b2, tot_b2/tot_ll2 if tot_ll2>0 else 0,
            tot_d, btot, btot/tot_d if tot_d>0 else 0]
for ci,val in enumerate(vals_tot,1):
    c = ws2.cell(row=r, column=ci, value=val)
    if ci in [4,7,10]: c.number_format='0%'
    st(c, bold=True, sz=10, col=WHITE, bg=G_M, ha='center'); c.border = bdr2
ws2.row_dimensions[r].height = 20; r+=2

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
    try: label = f"{period.start_time.strftime('%d/%m/%Y')} → {period.end_time.strftime('%d/%m/%Y')}"
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
print(f"Excel guardado en {OUTPUT}")

# ── Send email ────────────────────────────────────────────────────
now = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
msg = MIMEMultipart()
msg['From']    = EMAIL_USER
msg['To']      = EMAIL_TO
msg['Subject'] = f'Winbiota CRM — Reporte {now}'

body = f"""Hola,

Adjunto el reporte CRM de Winbiota generado automáticamente el {now}.

Resumen:
• Total leads: {N}
• Llamadas 1 efectivas (buenos estatus): {buenos1.sum()}
• Llamadas 2 efectivas (buenos estatus): {buenos2.sum()}

Saludos,
Bot Winbiota CRM
"""
msg.attach(MIMEText(body, 'plain'))

with open(OUTPUT, 'rb') as f:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename="Winbiota_CRM_{datetime.datetime.now().strftime("%Y%m%d_%H%M")}.xlsx"')
msg.attach(part)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
    server.login(EMAIL_USER, EMAIL_PASS)
    server.sendmail(EMAIL_USER, EMAIL_TO, msg.as_string())

print(f"Email enviado a {EMAIL_TO}")
