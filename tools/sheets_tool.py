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
SHEET_NAME_LOG = "TRANSAKSI" 

cred_path = os.getenv("GOOGLE_CREDENTIALS_PATH", "credentials.json")
if not os.path.isabs(cred_path):
    CREDENTIALS_FILE = os.path.join(HERMES_HOME, cred_path)
else:
    CREDENTIALS_FILE = cred_path

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

BLOCK_LIST_NAMA = ["TOTAL", "JUMLAH", "ESTIMATED", "NAMA ANGGOTA", "KETERANGAN", "GRANDE TOTAL", "SALDO", "PCN"]
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
    sheet_id = 0
    for sheet in res.json().get('sheets', []):
        sheet_id = sheet.get('properties', {}).get('sheetId', 0)
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
    return parsed_data, sheet_id

def _find_header_and_data_v2():
    data, sheet_id = _fetch_sheet_data_with_colors()
    if not data: return [], -1, -1, [], sheet_id, data
    header_row_idx, nama_col_idx = -1, -1
    for r_idx, row in enumerate(data[:15]): 
        for c_idx, cell in enumerate(row):
            if cell['value'].lower() in ['nama', 'nama anggota']:
                header_row_idx, nama_col_idx = r_idx, c_idx
                break
        if header_row_idx != -1: break
    if header_row_idx == -1: return data, -1, -1, [], sheet_id, data
    headers = data[header_row_idx]
    actual_data = []
    for row in data[header_row_idx + 1:]:
        if len(row) <= nama_col_idx: continue
        nama_raw = row[nama_col_idx]['value'].strip()
        if "TOTAL PER BULAN" in nama_raw.upper(): break
        if nama_raw: actual_data.append(row)
    return headers, header_row_idx, nama_col_idx, actual_data, sheet_id, data

def _execute_batch_update(requests):
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    authed_session = AuthorizedSession(creds)
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}:batchUpdate"
    res = authed_session.post(url, json={"requests": requests})
    if res.status_code != 200:
        raise Exception(f"BatchUpdate Error {res.status_code}: {res.text}")
    return res.json()

def _append_transaction_log(nama, nominal, ket_bulan, tipe="CASH"):
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    authed_session = AuthorizedSession(creds)
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    values = [[timestamp, nama, nominal, f"Iuran {ket_bulan} ({tipe})"]]
    
    encoded_log_sheet = urllib.parse.quote(f"{SHEET_NAME_LOG}!A1")
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{encoded_log_sheet}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS"
    
    res = authed_session.post(url, json={"values": values})
    if res.status_code != 200:
        raise Exception(f"Gagal mencatat log TRANSAKSI: {res.text}")

# ==================== TOOLS BACA (READ) ====================
# [KEMBALI KE VERSI ORIGINAL YANG SUDAH DI-TEST STABIL]

class CekKasInput(BaseModel):
    nama_anggota: str = Field(..., description="Nama anggota yang dicari status kasnya")

def cek_status_kas(input_data: CekKasInput) -> str:
    try:
        headers, _, nama_idx, actual_data, _, _ = _find_header_and_data_v2()
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
                
                # INI DIA LOGIC YANG SEMPAT HILANG (FILTER BULAN)
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
        headers, _, nama_idx, actual_data, _, _ = _find_header_and_data_v2()
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

class CekTunggakanInput(BaseModel): dummy: str = Field("dummy")
def rekap_tunggakan(input_data: CekTunggakanInput) -> str:
    try:
        headers, _, nama_idx, actual_data, _, _ = _find_header_and_data_v2()
        bulan_ini = datetime.now().strftime("%m/%Y")
        
        kolom_bulan_idx = []
        for i, h_cell in enumerate(headers):
            h_val = h_cell['value']
            if re.match(r"^\d{2}/\d{4}$", h_val) and _parse_month_year(h_val) <= _parse_month_year(bulan_ini):
                kolom_bulan_idx.append(i)
                
        rekap = []
        for row in actual_data:
            nama = row[nama_idx]['value'].strip()
            jumlah_tunggakan = 0
            
            for idx in kolom_bulan_idx:
                cell = row[idx] if idx < len(row) else {"value": "", "color": "white"}
                nominal = _parse_rupiah(cell['value'])
                color = cell['color']
                
                if color != "black" and color != "green" and nominal == 0:
                    jumlah_tunggakan += 1
                    
            if jumlah_tunggakan > 0:
                rekap.append({"nama": nama, "tunggakan": jumlah_tunggakan})
                
        rekap_sorted = sorted(rekap, key=lambda x: x['tunggakan'], reverse=True)
        
        pesan = "⚠️ *Rekap Anggota Menunggak Terbanyak*\n"
        for idx, data in enumerate(rekap_sorted[:15]): 
            pesan += f"{idx+1}. {data['nama']} ({data['tunggakan']} bulan)\n"
            
        return pesan
    except Exception as e: return f"Error: {str(e)}"

class TotalPiutangInput(BaseModel): dummy: str = Field("dummy")
def total_piutang_global(input_data: TotalPiutangInput) -> str:
    try:
        headers, _, _, actual_data, _, _ = _find_header_and_data_v2()
        bulan_ini = datetime.now().strftime("%m/%Y")
        kol = [i for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(bulan_ini)]
        kas, piutang = 0, 0
        for row in actual_data:
            for i in kol:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                n = _parse_rupiah(cell['value'])
                if cell['color'] == "black": continue
                elif cell['color'] == "green" or 0 < n <= BATAS_IURAN_WAJAR: kas += n
                elif n == 0: piutang += 50000
        potensi = kas + piutang
        sehat = (kas / potensi * 100) if potensi > 0 else 0
        pesan = "💰 *Laporan Piutang & Kesehatan Kas*\n"
        pesan += f"• Kas Terkumpul: Rp{kas:,}\n• Piutang: Rp{piutang:,}\n"
        pesan += f"• Potensi: Rp{potensi:,}\n• Kesehatan: {sehat:.1f}%\n"
        return pesan
    except Exception as e: return f"Error: {str(e)}"

class HallOfFameInput(BaseModel): dummy: str = Field("dummy")
def hall_of_fame(input_data: HallOfFameInput) -> str:
    try:
        headers, _, nama_idx, actual_data, _, _ = _find_header_and_data_v2()
        bulan_ini = datetime.now().strftime("%m/%Y")
        kol = [i for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(bulan_ini)]
        donatur = []
        for row in actual_data:
            bayar, wajib = 0, 0
            for i in kol:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                if cell['color'] != "black":
                    wajib += 50000
                    n = _parse_rupiah(cell['value'])
                    if cell['color'] == "green" or 0 < n <= BATAS_IURAN_WAJAR: bayar += n
            if bayar > wajib: donatur.append({"nama": row[nama_idx]['value'], "lebih": bayar - wajib})
        if not donatur: return "Belum ada donatur ekstra saat ini."
        pesan = "🏆 *Pahlawan Kas*\n"
        for d in sorted(donatur, key=lambda x: x['lebih'], reverse=True)[:5]: pesan += f"• {d['nama']} (+Rp{d['lebih']:,})\n"
        return pesan
    except Exception as e: return f"Error: {str(e)}"

class TrenBulanInput(BaseModel): dummy: str = Field("dummy")
def tren_bulan_kritis(input_data: TrenBulanInput) -> str:
    try:
        headers, _, _, actual_data, _, _ = _find_header_and_data_v2()
        bulan_ini = datetime.now().strftime("%m/%Y")
        kol = [(i, h['value']) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(bulan_ini)]
        tren = []
        for i, h in kol:
            k, t = 0, 0
            for row in actual_data:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                if cell['color'] != "black":
                    t += 50000
                    n = _parse_rupiah(cell['value'])
                    if cell['color'] == "green" or 0 < n <= BATAS_IURAN_WAJAR: k += n
            if t > 0: tren.append({"b": h, "p": (k/t*100)})
        pesan = "📉 *Bulan Kritis*\n"
        for d in sorted(tren, key=lambda x: x['p'])[:3]: pesan += f"• {d['b']} ({d['p']:.1f}%)\n"
        return pesan
    except Exception as e: return f"Error: {str(e)}"

class GhostingAlertInput(BaseModel): dummy: str = Field("dummy")
def ghosting_alert(input_data: GhostingAlertInput) -> str:
    try:
        headers, _, nama_idx, actual_data, _, _ = _find_header_and_data_v2()
        bulan_ini = datetime.now().strftime("%m/%Y")
        kol = sorted([i for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(bulan_ini)])
        ghosting = []
        for row in actual_data:
            streak = 0
            for i in kol:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                if cell['color'] == "black": continue
                if cell['color'] == "green" or 0 < _parse_rupiah(cell['value']) <= BATAS_IURAN_WAJAR: streak = 0
                else: streak += 1
            if streak >= 6: ghosting.append({"nama": row[nama_idx]['value'], "streak": streak})
        pesan = "👻 *Ghosting Alert*\n"
        for d in sorted(ghosting, key=lambda x: x['streak'], reverse=True)[:10]: pesan += f"• {d['nama']} ({d['streak']} bln)\n"
        return pesan
    except Exception as e: return f"Error: {str(e)}"

class RekapPemasukanAktualInput(BaseModel):
    bulan_tahun: str = Field(..., description="MM/YYYY")

def rekap_pemasukan_aktual(input_data: RekapPemasukanAktualInput) -> str:
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        res = AuthorizedSession(creds).get(f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{urllib.parse.quote(SHEET_NAME_LOG)}")
        rows = res.json().get('values', [])
        cash, barang, detail = 0, 0, []
        for r in rows[1:]:
            if len(r) >= 4 and r.split(' ')[3:] == input_data.bulan_tahun:
                nom = _parse_rupiah(r)
                if "(BARANG)" in r.upper(): barang += nom
                else: cash += nom
                detail.append(f"• {r}: Rp{nom:,}")
        return f"📈 *Pemasukan {input_data.bulan_tahun}*\nCash: Rp{cash:,}\nBarang: Rp{barang:,}\nTotal: Rp{cash+barang:,}\n\nDetail:\n" + "\n".join(detail[:10])
    except Exception as e: return f"Error: {str(e)}"


# ==================== TOOLS TULIS (WRITE / UPDATE) ====================

class BayarKasInput(BaseModel):
    nama_anggota: str = Field(..., description="Nama anggota")
    nominal_per_bulan: int = Field(50000)
    jumlah_bulan: int = Field(1)
    bulan_mulai: str = Field(None)

def bayar_kas(input_data: BayarKasInput) -> str:
    try:
        headers, header_idx, nama_idx, actual_data, sheet_id, raw_data = _find_header_and_data_v2()
        abs_row, nama_found = -1, ""
        for i, row in enumerate(raw_data[header_idx+1:]):
            if len(row) > nama_idx and input_data.nama_anggota.lower() in row[nama_idx]['value'].lower():
                abs_row, nama_found = header_idx + 1 + i, row[nama_idx]['value']
                break
        if abs_row == -1: return f"Anggota '{input_data.nama_anggota}' tidak ditemukan."

        b_start = input_data.bulan_mulai or datetime.now().strftime("%m/%Y")
        kol_bulan = sorted([(i, h['value']) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) >= _parse_month_year(b_start)])
        
        to_update, list_b = [], []
        for i, h in kol_bulan:
            cell = raw_data[abs_row][i] if i < len(raw_data[abs_row]) else {"value": "", "color": "white"}
            if cell['color'] not in ["green", "black"] and _parse_rupiah(cell['value']) == 0:
                to_update.append(i)
                list_b.append(h)
                if len(to_update) == input_data.jumlah_bulan: break
        
        if len(to_update) < input_data.jumlah_bulan: return "Kolom bulan tidak cukup."

        reqs = [{
            "updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": abs_row, "endRowIndex": abs_row + 1, "startColumnIndex": i, "endColumnIndex": i + 1},
                "rows": [{"values": [{"userEnteredValue": {"numberValue": input_data.nominal_per_bulan}, "userEnteredFormat": {"numberFormat": {"type": "CURRENCY", "pattern": "Rp#,##0"}}}]}],
                "fields": "userEnteredValue,userEnteredFormat(numberFormat)"
            }
        } for i in to_update]
        _execute_batch_update(reqs)
        
        total = input_data.nominal_per_bulan * input_data.jumlah_bulan
        _append_transaction_log(nama_found, total, ', '.join(list_b), "CASH")
        return f"✅ Berhasil catat Rp{total:,} dari *{nama_found}*."
    except Exception as e: return f"Error: {str(e)}"

class KonversiBarangInput(BaseModel):
    nama_anggota: str = Field(..., description="Nama anggota")
    harga_barang: int = Field(...)
    bulan_mulai: str = Field(None)

def konversi_barang(input_data: KonversiBarangInput) -> str:
    try:
        headers, header_idx, nama_idx, actual_data, sheet_id, raw_data = _find_header_and_data_v2()
        abs_row, nama_found = -1, ""
        for i, row in enumerate(raw_data[header_idx+1:]):
            if len(row) > nama_idx and input_data.nama_anggota.lower() in row[nama_idx]['value'].lower():
                abs_row, nama_found = header_idx + 1 + i, row[nama_idx]['value']
                break
        if abs_row == -1: return f"Anggota '{input_data.nama_anggota}' tidak ditemukan."

        jml_bln = input_data.harga_barang // 50000
        b_start = input_data.bulan_mulai or datetime.now().strftime("%m/%Y")
        kol_bulan = sorted([(i, h['value']) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) >= _parse_month_year(b_start)])
        
        to_update, list_b = [], []
        for i, h in kol_bulan:
            cell = raw_data[abs_row][i] if i < len(raw_data[abs_row]) else {"value": "", "color": "white"}
            if cell['color'] not in ["green", "black"] and _parse_rupiah(cell['value']) == 0:
                to_update.append(i)
                list_b.append(h)
                if len(to_update) == jml_bln: break

        reqs = [{
            "updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": abs_row, "endRowIndex": abs_row + 1, "startColumnIndex": i, "endColumnIndex": i + 1},
                "rows": [{"values": [{"userEnteredValue": {"stringValue": "Barang"}, "userEnteredFormat": {"backgroundColor": {"red": 0.4, "green": 0.8, "blue": 0.4}}}]}],
                "fields": "userEnteredValue,userEnteredFormat(backgroundColor)"
            }
        } for i in to_update]
        _execute_batch_update(reqs)
        
        _append_transaction_log(nama_found, jml_bln * 50000, ', '.join(list_b), "BARANG")
        return f"✅ Berhasil konversi barang Rp{input_data.harga_barang:,} dari *{nama_found}*."
    except Exception as e: return f"Error: {str(e)}"

class TambahAnggotaInput(BaseModel):
    nama_baru: str = Field(..., description="Nama lengkap anggota baru")

def tambah_anggota(input_data: TambahAnggotaInput) -> str:
    try:
        headers, header_row_idx, nama_idx, actual_data, sheet_id, raw_data = _find_header_and_data_v2()
        
        nama_baru_clean = input_data.nama_baru.strip()
        members = []
        start_row = int(header_row_idx) + 1
        max_row_idx = start_row
        
        for i, row in enumerate(raw_data[start_row:]):
            abs_idx = start_row + i
            if len(row) > nama_idx:
                nama_exist = row[nama_idx]['value'].strip()
                if "TOTAL PER BULAN" in nama_exist.upper(): break
                if nama_exist and not any(x in nama_exist.upper() for x in BLOCK_LIST_NAMA):
                    if nama_exist.lower() == nama_baru_clean.lower():
                        return f"⚠️ Anggota *{nama_exist}* sudah ada."
                    members.append((nama_exist, abs_idx))
                    max_row_idx = max(max_row_idx, abs_idx)
                    
        insert_row_idx = -1
        for nama, abs_idx in members:
            if nama.lower() > nama_baru_clean.lower():
                insert_row_idx = int(abs_idx)
                break
                
        if insert_row_idx == -1:
            insert_row_idx = max_row_idx + 1 if members else start_row

        requests = [{
            "insertDimension": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": insert_row_idx, "endIndex": insert_row_idx + 1},
                "inheritFromBefore": True
            }
        }]

        current_m_tuple = _parse_month_year(datetime.now().strftime("%m/%Y"))
        row_values = []
        max_col = len(headers)
        
        for col_idx in range(max_col):
            if col_idx == nama_idx:
                row_values.append({"userEnteredValue": {"stringValue": nama_baru_clean}})
            else:
                h_val = headers[col_idx]['value']
                if re.match(r"^\d{2}/\d{4}$", h_val):
                    m_tuple = _parse_month_year(h_val)
                    if m_tuple < current_m_tuple:
                        row_values.append({"userEnteredFormat": {"backgroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}}})
                    else:
                        row_values.append({"userEnteredFormat": {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}})
                else:
                    row_values.append({})

        requests.append({
            "updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": insert_row_idx, "endRowIndex": insert_row_idx + 1, "startColumnIndex": 0, "endColumnIndex": max_col},
                "rows": [{"values": row_values}],
                "fields": "userEnteredValue,userEnteredFormat.backgroundColor"
            }
        })

        _execute_batch_update(requests)
        return f"✅ Anggota baru *{nama_baru_clean}* berhasil ditambah sesuai abjad!"
    except Exception as e: return f"Error: {str(e)}"
