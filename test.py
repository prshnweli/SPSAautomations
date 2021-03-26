from Google import Create_Service
import win32com.client as win32

xlApp = win32.Dispatch('Excel.Application')
wb = xlApp.Workbooks.Open(r"FP/ElToro.xlsx")
ws = wb.Worksheets('FP')
rngData = ws.Range('A1').CurrentRegion()

# Google Sheet Id
gsheet_id = '1cY3CgX7T1SPV-277sh8BvC8d2_FezJDF5cZzk_Ii63A'
CLIENT_SECRET_FILE = 'keys/creds.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

response = service.spreadsheets().values().append(
    spreadsheetId=gsheet_id,
    valueInputOption='RAW',
    range='Sheet1!A1',
    body=dict(
        majorDimension='ROWS',
        values=rngData
    )
).execute()
