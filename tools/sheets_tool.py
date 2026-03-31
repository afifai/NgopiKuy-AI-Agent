import os
import re
import urllib.parse
from datetime import datetime
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession
from pydantic import BaseModel, Field
from dotenv import load_dotenv

load_dotenv()

# ==================== KONFIGURASI ====================
HERMES_HOME = os.environ.get("HERMES_HOME", os.path.expanduser("~/.hermes_kopi"))
SHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "KAS")

cred_path = os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials.json")
if not os.path.isabs(cred_path):
    CREDENTIALS_FILE = os.path.join(HERMES_HOME, cred_path)
else:
    CREDENTIALS_FILE = cred_path

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Batas iuran untuk validasi nominal (iuran kas per bulan)
BATAS_IURAN_WAJAR = 100000 

if not SHEET_ID:
    raise ValueError("SPREADSHEET_ID tidak ditemukan di file .env!")

# ==================== CORE FUNCTIONS ====================
def _parse_month_year(mm_yyyy):
    try:
        m, y = mm_yyyy.strip().split('/')
        return (int(y), int(m))
    except:
        return (9999, 12)

def _parse_rupiah(nominal_str):
    try:
        if not nominal_str: return 0
        s = str(nominal_str).strip()
        
        # Kebal terhadap baris persentase
        if '%' in s: return 0
        
        if s.endswith(',00') or s.endswith('.00'): s = s[:-3]
        elif s.endswith('.0'): s = s[:-2]
        cleaned = ''.join(c for c in s if c.isdigit())
        return int(cleaned) if cleaned else 0
    except:
        return 0

def _fetch_sheet_data_with_colors():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    authed_session = AuthorizedSession(creds)
    encoded_sheet = urllib.parse.quote(SHEET_NAME)
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}?includeGridData=true&ranges={encoded_sheet}"
    res = authed_session.get(url)
    
    if res.status_code != 200:
        raise Exception(f"API Error {res.status_code}: {res.text}")
        
    parsed_data = []
    for sheet in res.json().get('sheets', []):
        for data in sheet.get('data', []):
            for row in data.get('rowData', []):
                row_data = []
                for cell in row.get('values', []):
                    val = cell.get('formattedValue', '')
                    if not val:
                        uv = cell.get('userEnteredValue', {})
                        if 'stringValue' in uv: val = uv['stringValue']
                        elif 'numberValue' in uv: val = str(uv['numberValue'])
                            
                    color = "white"
                    eff_fmt = cell.get('effectiveFormat')
                    if isinstance(eff_fmt, dict):
                        bg = eff_fmt.get('backgroundColor', {})
                        if isinstance(bg, dict):
                            r, g, b = bg.get('red', 0), bg.get('green', 0), bg.get('blue', 0)
                            if r < 0.1 and g < 0.1 and b < 0.1: color = "black"
                            elif g > r and g > b: color = "green"
                    row_data.append({"value": str(val).strip(), "color": color})
                parsed_data.append(row_data)
            break 
        break 
    return parsed_data

def _find_header_and_data_v2():
    data = _fetch_sheet_data_with_colors()
    if not data: return [], -1, -1, []
        
    header_row_idx, nama_col_idx = -1, -1
    # 1. Cari Header sebagai Start Marker
    for r_idx, row in enumerate(data[:15]): 
        for c_idx, cell in enumerate(row):
            if cell['value'].lower() in ['nama', 'nama anggota']:
                header_row_idx, nama_col_idx = r_idx, c_idx
                break
        if header_row_idx != -1: break
            
    if header_row_idx == -1: return data, -1, -1, []
    headers = data[header_row_idx]
    
    # 2. Ambil data HANYA sampai ketemu Stop Marker "TOTAL per bulan"
    actual_data = []
    for row in data[header_row_idx + 1:]:
        if len(row) <= nama_col_idx: continue
        
        nama_raw = row[nama_col_idx]['value'].strip()
        
        # STOP MARKER: Berhenti total jika menyentuh baris totalan
        if "TOTAL PER BULAN" in nama_raw.upper():
            break
            
        # Ambil hanya baris yang punya nama (bukan baris kosong di tengah)
        if nama_raw:
            actual_data.append(row)
        
    return headers, header_row_idx, nama_col_idx, actual_data

# ==================== TOOLS ====================
class CekKasInput(BaseModel):
    nama_anggota: str = Field(..., description="Nama anggota yang dicari status kasnya")

def cek_status_kas(input_data: CekKasInput) -> str:
    try:
        headers, _, nama_idx, actual_data = _find_header_and_data_v2()
        if nama_idx == -1: return "Error: Kolom 'Nama Anggota' tidak ditemukan."
            
        for row in actual_data:
            nama_di_sheet = row[nama_idx]['value']
            if input_data.nama_anggota.lower() in nama_di_sheet.lower():
                bulan_bayar, tunggakan, bulan_libur = [], [], []
                
                for i, h_cell in enumerate(headers):
                    h_val = h_cell['value']
                    if re.match(r"^\d{2}/\d{4}$", h_val):
                        cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                        nominal = _parse_rupiah(cell['value'])
                        if cell['color'] == "black": bulan_libur.append(h_val)
                        elif cell['color'] == "green" or 0 < nominal <= BATAS_IURAN_WAJAR:
                            bulan_bayar.append(h_val)
                        elif nominal == 0: tunggakan.append(h_val)
                            
                terakhir_bayar = bulan_bayar[-1] if bulan_bayar else "Belum pernah bayar"
                bulan_ini = datetime.now().strftime("%m/%Y")
                tunggakan_berjalan = sorted([b for b in tunggakan if _parse_month_year(b) <= _parse_month_year(bulan_ini)], key=lambda x: _parse_month_year(x))
                
                pesan = f"Status Kas untuk *{nama_di_sheet}*:\n"
                pesan += f"Terakhir bayar/dianggap lunas: {terakhir_bayar}\n"
                if tunggakan_berjalan:
                    pesan += f"Tunggakan aktif ({len(tunggakan_berjalan)} bulan): {', '.join(tunggakan_berjalan)}\n"
                else:
                    pesan += "Status: AMAN (Tidak ada tunggakan sampai bulan ini).\n"
                if bulan_libur:
                    pesan += f"*(Catatan: Bebas tagihan/libur sebanyak {len(bulan_libur)} bulan karena cell hitam)*"
                return pesan
        return f"Anggota bernama '{input_data.nama_anggota}' tidak ditemukan."
    except Exception as e: return f"Error: {str(e)}"

class RingkasanBulanInput(BaseModel):
    bulan_tahun: str = Field(..., description="Bulan target. Contoh: '03/2026'")

def ringkasan_kas_bulan(input_data: RingkasanBulanInput) -> str:
    try:
        headers, _, nama_idx, actual_data = _find_header_and_data_v2()
        target_col_idx = next((i for i, h in enumerate(headers) if input_data.bulan_tahun in h['value']), -1)
        if target_col_idx == -1: return f"Kolom bulan '{input_data.bulan_tahun}' tidak ditemukan."
        
        sudah_bayar, belum_bayar, diliburkan, total_uang = [], [], [], 0
        for row in actual_data:
            nama = row[nama_idx]['value'].strip()
            cell = row[target_col_idx] if target_col_idx < len(row) else {"value": "", "color": "white"}
            nominal = _parse_rupiah(cell['value'])
            
            if cell['color'] == "black":
                diliburkan.append(nama)
            elif cell['color'] == "green" or 0 < nominal <= BATAS_IURAN_WAJAR:
                sudah_bayar.append(nama)
                total_uang += nominal
            elif nominal == 0:
                belum_bayar.append(nama)
        
        total_anggota = len(sudah_bayar) + len(belum_bayar) + len(diliburkan)
        wajib_bayar = total_anggota - len(diliburkan)
        target_uang = wajib_bayar * 50000
        persentase = (total_uang / target_uang) * 100 if target_uang > 0 else 0
        
        pesan = f"📊 *Ringkasan Kas Kuota Bulan {input_data.bulan_tahun}*\n"
        pesan += f"• Total Anggota: {total_anggota} orang\n"
        pesan += f"• Lunas: {len(sudah_bayar)} orang\n"
        pesan += f"• Belum Bayar: {len(belum_bayar)} orang\n"
        if diliburkan: pesan += f"• Bebas Kas (Cell Hitam): {len(diliburkan)} orang\n"
        
        pesan += f"\n💰 *Finansial:*\n"
        pesan += f"• Target: Rp{target_uang:,} ({wajib_bayar} org x 50k)\n"
        pesan += f"• Terkumpul: Rp{total_uang:,}\n"
        pesan += f"• Capaian: {persentase:.1f}%\n"
        
        if target_uang > total_uang:
            pesan += f"• Kurang: Rp{target_uang - total_uang:,}\n"
            pesan += f"\nDaftar Belum Bayar: {', '.join(belum_bayar[:10])}"
            if len(belum_bayar) > 10: pesan += f" ... (+{len(belum_bayar)-10} orang)"
        return pesan
    except Exception as e: return f"Error: {str(e)}"
