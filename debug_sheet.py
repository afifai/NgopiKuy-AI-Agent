import os
import urllib.parse
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession
from dotenv import load_dotenv

load_dotenv()

HERMES_HOME = os.environ.get("HERMES_HOME", os.path.expanduser("~/.hermes_kopi"))
SHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "KAS")

cred_path = os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials.json")
if not os.path.isabs(cred_path):
    CREDENTIALS_FILE = os.path.join(HERMES_HOME, cred_path)
else:
    CREDENTIALS_FILE = cred_path

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def debug_headers():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    authed_session = AuthorizedSession(creds)
    
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}?includeGridData=true&ranges={urllib.parse.quote(SHEET_NAME)}"
    res = authed_session.get(url)
    
    if res.status_code != 200:
        print(f"Error: {res.text}")
        return
        
    res_json = res.json()
    print("=== ISI TEKS 5 BARIS PERTAMA (DARI API) ===")
    
    # Looping yang aman menembus array JSON
    for sheet in res_json.get('sheets', []):
        for data in sheet.get('data', []):
            row_data = data.get('rowData', [])
            for r_idx, row in enumerate(row_data[:5]):
                vals = []
                for cell in row.get('values', []):
                    val = cell.get('formattedValue', '')
                    vals.append(val)
                print(f"Baris {r_idx}: {vals}")

if __name__ == "__main__":
    debug_headers()
