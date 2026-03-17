import os, json, datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

creds_json = json.loads(os.environ['GOOGLE_CREDENTIALS'])
SHEET_ID   = os.environ['SHEET_ID']

scopes = ['https://www.googleapis.com/auth/spreadsheets']
creds  = Credentials.from_service_account_info(creds_json, scopes=scopes)
sheets = build('sheets', 'v4', credentials=creds)

# Read CRM data
result = sheets.spreadsheets().values().get(
    spreadsheetId=SHEET_ID,
    range='CRM Winbiota',
    valueRenderOption='FORMATTED_VALUE',
    dateTimeRenderOption='FORMATTED_STRING'
).execute()
all_values = result.get('values', [])
max_cols = max(len(r) for r in all_values)
all_values = [r + [''] * (max_cols - len(r)) for r in all_values]

buenos_vals1 = ['CONTESTO','CONTESTA','INTERES','LLAMADA','NO INT','SI']
buenos_vals2 = ['SI CONTESTA','INTERESADA']

# Count from row 4 (index 3), skip dummy data
ll1_total = 0
ll1_buenos = 0
ll2_total = 0
ll2_buenos = 0

for row in all_values[3:]:
    name = str(row[2]).strip() if len(row) > 2 else ''
    if 'dummy data' in name.lower() or name == '':
        continue
    e1 = str(row[12]).strip().upper() if len(row) > 12 else ''
    e2 = str(row[16]).strip().upper() if len(row) > 16 else ''
    c1 = str(row[11]).strip() if len(row) > 11 else ''
    c2 = str(row[15]).strip() if len(row) > 15 else ''
    if c1 and c1 not in ['', 'nan', 'None']:
        ll1_total += 1
        if e1 in buenos_vals1:
            ll1_buenos += 1
    if c2 and c2 not in ['', 'nan', 'None']:
        ll2_total += 1
        if e2 in buenos_vals2:
            ll2_buenos += 1

now = datetime.datetime.utcnow()
timestamp = now.strftime('%Y-%m-%d %H:%M')
date_str  = now.strftime('%Y-%m-%d')
hour_str  = now.strftime('%H:00')

snapshot_row = [timestamp, date_str, hour_str, ll1_total, ll1_buenos, ll2_total, ll2_buenos]
print(f"Snapshot: {snapshot_row}")

# Ensure Snapshots sheet exists
try:
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=SHEET_ID,
        body={'requests': [{'addSheet': {'properties': {'title': 'Snapshots', 'hidden': True}}}]}
    ).execute()
    print("Hoja Snapshots creada")
    # Add header
    sheets.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range='Snapshots!A1',
        valueInputOption='RAW',
        body={'values': [['Timestamp','Fecha','Hora','LL1 total','LL1 buenos','LL2 total','LL2 buenos']]}
    ).execute()
except Exception as e:
    if 'already exists' in str(e).lower():
        print("Hoja Snapshots ya existe")
    else:
        print(f"Warning: {e}")

# Append snapshot row
sheets.spreadsheets().values().append(
    spreadsheetId=SHEET_ID,
    range='Snapshots!A1',
    valueInputOption='RAW',
    insertDataOption='INSERT_ROWS',
    body={'values': [snapshot_row]}
).execute()
print(f"Snapshot guardado: LL1={ll1_total}({ll1_buenos}) LL2={ll2_total}({ll2_buenos})")
