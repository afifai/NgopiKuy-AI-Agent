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
        
        if "TOTAL PER BULAN" in nama_raw.upper():
            break
            
        if nama_raw:
            actual_data.append(row)
        
    return headers, header_row_idx, nama_col_idx, actual_data, sheet_id, data

def _execute_batch_update(requests):
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    authed_session = AuthorizedSession(creds)
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}:batchUpdate"
    res = authed_session.post(url, json={"requests": requests})
    if res.status_code != 200:
        raise Exception(f"BatchUpdate Error {res.status_code}: {res.text}")
    return res.json()


# ==================== TOOLS BACA (READ) ====================

class CekKasInput(BaseModel):
    nama_anggota: str = Field(..., description="Nama anggota")

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
                        elif cell['color'] == "green" or 0 < nominal <= BATAS_IURAN_WAJAR: bulan_bayar.append(h_val)
                        elif nominal == 0: tunggakan.append(h_val)
                            
                terakhir_bayar = bulan_bayar[-1] if bulan_bayar else "Belum pernah bayar"
                bulan_ini = datetime.now().strftime("%m/%Y")
                tunggakan_berjalan = sorted([b for b in tunggakan if _parse_month_year(b) <= _parse_month_year(bulan_ini)], key=lambda x: _parse_month_year(x))
                
                pesan = f"Status Kas untuk *{nama_di_sheet}*:\n"
                pesan += f"Terakhir bayar: {terakhir_bayar}\n"
                if tunggakan_berjalan: pesan += f"Tunggakan ({len(tunggakan_berjalan)} bln): {', '.join(tunggakan_berjalan)}\n"
                else: pesan += "Status: AMAN.\n"
                if bulan_libur: pesan += f"*(Libur {len(bulan_libur)} bulan)*"
                return pesan
        return f"Anggota '{input_data.nama_anggota}' tidak ditemukan."
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
            
            if cell['color'] == "black": diliburkan.append(nama)
            elif cell['color'] == "green" or 0 < nominal <= BATAS_IURAN_WAJAR:
                sudah_bayar.append(nama)
                total_uang += nominal
            elif nominal == 0: belum_bayar.append(nama)
        
        total_anggota = len(sudah_bayar) + len(belum_bayar) + len(diliburkan)
        wajib_bayar = total_anggota - len(diliburkan)
        target_uang = wajib_bayar * 50000
        persentase = (total_uang / target_uang) * 100 if target_uang > 0 else 0
        
        pesan = f"📊 *Ringkasan Kas Kuota Bulan {input_data.bulan_tahun}*\n"
        pesan += f"• Total Anggota: {total_anggota} orang\n"
        pesan += f"• Lunas: {len(sudah_bayar)} orang\n"
        pesan += f"• Belum Bayar: {len(belum_bayar)} orang\n"
        if diliburkan: pesan += f"• Bebas Kas: {len(diliburkan)} orang\n"
        pesan += f"\n💰 *Finansial:*\n"
        pesan += f"• Target: Rp{target_uang:,}\n• Terkumpul: Rp{total_uang:,} ({persentase:.1f}%)\n"
        if target_uang > total_uang:
            pesan += f"• Kurang: Rp{target_uang - total_uang:,}\n"
            pesan += f"\nBelum Bayar: {', '.join(belum_bayar[:10])}"
            if len(belum_bayar) > 10: pesan += f" ... (+{len(belum_bayar)-10} orang)"
        return pesan
    except Exception as e: return f"Error: {str(e)}"

class CekTunggakanInput(BaseModel): dummy: str = Field("dummy")
def rekap_tunggakan(input_data: CekTunggakanInput) -> str:
    try:
        headers, _, nama_idx, actual_data, _, _ = _find_header_and_data_v2()
        bulan_ini = datetime.now().strftime("%m/%Y")
        kolom_bulan_idx = [i for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(bulan_ini)]
        rekap = []
        for row in actual_data:
            nama, nunggak = row[nama_idx]['value'].strip(), 0
            for idx in kolom_bulan_idx:
                cell = row[idx] if idx < len(row) else {"value": "", "color": "white"}
                if cell['color'] not in ["black", "green"] and _parse_rupiah(cell['value']) == 0: nunggak += 1
            if nunggak > 0: rekap.append({"nama": nama, "tunggakan": nunggak})
        pesan = "⚠️ *Rekap Menunggak Terbanyak*\n"
        for idx, data in enumerate(sorted(rekap, key=lambda x: x['tunggakan'], reverse=True)[:15]): 
            pesan += f"{idx+1}. {data['nama']} ({data['tunggakan']} bln)\n"
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
            nama, bayar, wajib = row[nama_idx]['value'].strip(), 0, 0
            for i in kol:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                if cell['color'] != "black":
                    wajib += 50000
                    n = _parse_rupiah(cell['value'])
                    if cell['color'] == "green" or 0 < n <= BATAS_IURAN_WAJAR: bayar += n
            if bayar > wajib: donatur.append({"nama": nama, "lebih": bayar - wajib})
        if not donatur: return "Belum ada donatur ekstra saat ini."
        pesan = "🏆 *Hall of Fame (Pahlawan Kas)*\n"
        for idx, d in enumerate(sorted(donatur, key=lambda x: x['lebih'], reverse=True)[:10]):
            pesan += f"{idx+1}. {d['nama']} (Extra Rp{d['lebih']:,})\n"
        return pesan
    except Exception as e: return f"Error: {str(e)}"

class TrenBulanInput(BaseModel): dummy: str = Field("dummy")
def tren_bulan_kritis(input_data: TrenBulanInput) -> str:
    try:
        headers, _, _, actual_data, _, _ = _find_header_and_data_v2()
        bulan_ini = datetime.now().strftime("%m/%Y")
        kol = [(i, h['value']) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(bulan_ini)]
        tren = {h: {"t": 0, "k": 0} for _, h in kol}
        for row in actual_data:
            for i, h in kol:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                if cell['color'] != "black":
                    tren[h]['t'] += 50000
                    n = _parse_rupiah(cell['value'])
                    if cell['color'] == "green" or 0 < n <= BATAS_IURAN_WAJAR: tren[h]['k'] += n
        arr = [{"b": b, "p": (d['k']/d['t']*100), "k": d['k'], "t": d['t']} for b, d in tren.items() if d['t'] > 0]
        pesan = "📉 *Bulan Kritis (Pemasukan Terseret)*\n"
        for idx, t in enumerate(sorted(arr, key=lambda x: x['p'])[:3]):
            pesan += f"{idx+1}. Bulan {t['b']}: {t['p']:.1f}% (Rp{t['k']:,})\n"
        return pesan
    except Exception as e: return f"Error: {str(e)}"

class GhostingAlertInput(BaseModel): dummy: str = Field("dummy")
def ghosting_alert(input_data: GhostingAlertInput) -> str:
    try:
        headers, _, nama_idx, actual_data, _, _ = _find_header_and_data_v2()
        bulan_ini = datetime.now().strftime("%m/%Y")
        kol = sorted([(i, h['value']) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(bulan_ini)], key=lambda x: _parse_month_year(x))
        ghosting = []
        for row in actual_data:
            nama, streak = row[nama_idx]['value'].strip(), 0
            for i, h in kol:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                if cell['color'] == "black": continue
                elif cell['color'] == "green" or 0 < _parse_rupiah(cell['value']) <= BATAS_IURAN_WAJAR: streak = 0
                elif _parse_rupiah(cell['value']) == 0: streak += 1
            if streak >= 6: ghosting.append({"nama": nama, "streak": streak})
        if not ghosting: return "Aman! Tidak ada ghosting >= 6 bulan."
        pesan = "👻 *Ghosting Alert (>= 6 Bulan)*\n"
        for idx, g in enumerate(sorted(ghosting, key=lambda x: x['streak'], reverse=True)[:15]):
            pesan += f"{idx+1}. {g['nama']} ({g['streak']} bln)\n"
        return pesan
    except Exception as e: return f"Error: {str(e)}"


# ==================== TOOLS TULIS (WRITE / UPDATE) ====================

# 1. BAYAR KAS MURNI (TANPA WARNA HIJAU)
class BayarKasInput(BaseModel):
    nama_anggota: str = Field(..., description="Nama anggota")
    nominal_per_bulan: int = Field(50000, description="Nominal yang dibayar per bulan (misal 50000 atau 100000)")
    jumlah_bulan: int = Field(1, description="Berapa bulan yang mau dilunasi")
    bulan_mulai: str = Field(None, description="Bulan start pembayaran (MM/YYYY). Kosongkan untuk bulan saat ini.")

def bayar_kas(input_data: BayarKasInput) -> str:
    try:
        headers, header_row_idx, nama_idx, actual_data, sheet_id, raw_data = _find_header_and_data_v2()
        
        abs_row_idx, nama_ditemukan = -1, ""
        for i, row in enumerate(raw_data[header_row_idx+1:]):
            if len(row) > nama_idx:
                n = row[nama_idx]['value'].strip()
                if input_data.nama_anggota.lower() in n.lower():
                    abs_row_idx = header_row_idx + 1 + i
                    nama_ditemukan = n
                    break
                    
        if abs_row_idx == -1: return f"Anggota '{input_data.nama_anggota}' tidak ditemukan."

        bulan_mulai = input_data.bulan_mulai if input_data.bulan_mulai else datetime.now().strftime("%m/%Y")
        start_tuple = _parse_month_year(bulan_mulai)
        
        month_cols = [(i, h['value'], _parse_month_year(h['value'])) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value'])]
        month_cols.sort(key=lambda x: x)

        months_to_update = []
        for col_idx, mm_yyyy, m_tuple in month_cols:
            if m_tuple >= start_tuple:
                cell = raw_data[abs_row_idx][col_idx] if col_idx < len(raw_data[abs_row_idx]) else {"value": "", "color": "white"}
                if cell['color'] not in ['green', 'black'] and _parse_rupiah(cell['value']) == 0:
                    months_to_update.append((col_idx, mm_yyyy))
                    if len(months_to_update) == input_data.jumlah_bulan: break
                        
        if len(months_to_update) < input_data.jumlah_bulan:
            return f"Hanya ada {len(months_to_update)} kolom bulan tersisa di Sheets!"

        requests = []
        for col_idx, mm_yyyy in months_to_update:
            requests.append({
                "updateCells": {
                    "range": {"sheetId": sheet_id, "startRowIndex": abs_row_idx, "endRowIndex": abs_row_idx + 1, "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1},
                    "rows": [{"values": [{
                        "userEnteredValue": {"numberValue": input_data.nominal_per_bulan},
                        "userEnteredFormat": {
                            "numberFormat": {"type": "CURRENCY", "pattern": "Rp#,##0"}
                            # Hapus bagian backgroundColor di sini
                        }
                    }]}],
                    "fields": "userEnteredValue,userEnteredFormat(numberFormat)" # Hapus backgroundColor dari fields
                }
            })
            
        _execute_batch_update(requests)
        return f"✅ Kas Rp{input_data.nominal_per_bulan:,} masuk buat *{nama_ditemukan}*!\nLunas {input_data.jumlah_bulan} bulan: {', '.join([m for m in months_to_update])}."
    except Exception as e: return f"Error: {str(e)}"

# 2. KONVERSI BELI BARANG
class KonversiBarangInput(BaseModel):
    nama_anggota: str = Field(..., description="Nama anggota")
    harga_barang: int = Field(..., description="Total harga barang (misal 160000)")
    bulan_mulai: str = Field(None, description="Bulan start pembayaran (MM/YYYY).")

def konversi_barang(input_data: KonversiBarangInput) -> str:
    try:
        headers, header_row_idx, nama_idx, actual_data, sheet_id, raw_data = _find_header_and_data_v2()
        
        jumlah_bulan = input_data.harga_barang // 50000
        if jumlah_bulan < 1: return f"Harga Rp{input_data.harga_barang:,} tidak cukup (Min. 50k)."

        abs_row_idx, nama_ditemukan = -1, ""
        for i, row in enumerate(raw_data[header_row_idx+1:]):
            if len(row) > nama_idx:
                n = row[nama_idx]['value'].strip()
                if input_data.nama_anggota.lower() in n.lower():
                    abs_row_idx = header_row_idx + 1 + i
                    nama_ditemukan = n
                    break
                    
        if abs_row_idx == -1: return f"Anggota '{input_data.nama_anggota}' tidak ditemukan."

        bulan_mulai = input_data.bulan_mulai if input_data.bulan_mulai else datetime.now().strftime("%m/%Y")
        start_tuple = _parse_month_year(bulan_mulai)
        
        month_cols = [(i, h['value'], _parse_month_year(h['value'])) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value'])]
        month_cols.sort(key=lambda x: x)

        months_to_update = []
        for col_idx, mm_yyyy, m_tuple in month_cols:
            if m_tuple >= start_tuple:
                cell = raw_data[abs_row_idx][col_idx] if col_idx < len(raw_data[abs_row_idx]) else {"value": "", "color": "white"}
                if cell['color'] not in ['green', 'black'] and _parse_rupiah(cell['value']) == 0:
                    months_to_update.append((col_idx, mm_yyyy))
                    if len(months_to_update) == jumlah_bulan: break
                        
        if len(months_to_update) < jumlah_bulan:
            return f"Hanya ada {len(months_to_update)} kolom bulan tersisa."

        requests = []
        for col_idx, mm_yyyy in months_to_update:
            requests.append({
                "updateCells": {
                    "range": {"sheetId": sheet_id, "startRowIndex": abs_row_idx, "endRowIndex": abs_row_idx + 1, "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1},
                    "rows": [{"values": [{
                        "userEnteredValue": {"stringValue": "Barang"},
                        "userEnteredFormat": {"backgroundColor": {"red": 0.4, "green": 0.8, "blue": 0.4}}
                    }]}],
                    "fields": "userEnteredValue,userEnteredFormat(backgroundColor)"
                }
            })
            
        _execute_batch_update(requests)
        return f"✅ Barang senilai Rp{input_data.harga_barang:,} dari *{nama_ditemukan}* dikonversi!\nLunas {jumlah_bulan} bulan: {', '.join([m for m in months_to_update])}."
    except Exception as e: return f"Error: {str(e)}"

# 3. TAMBAH ANGGOTA BARU
class TambahAnggotaInput(BaseModel):
    nama_baru: str = Field(..., description="Nama lengkap anggota baru")

def tambah_anggota(input_data: TambahAnggotaInput) -> str:
    try:
        headers, header_row_idx, nama_idx, actual_data, sheet_id, raw_data = _find_header_and_data_v2()
        
        members = []
        for i, row in enumerate(raw_data[header_row_idx+1:]):
            abs_idx = header_row_idx + 1 + i
            if len(row) > nama_idx:
                nama = row[nama_idx]['value'].strip()
                if "TOTAL PER BULAN" in nama.upper(): break
                if nama and not any(x in nama.upper() for x in BLOCK_LIST_NAMA): members.append((nama, abs_idx))
                    
        insert_row_idx = -1
        for nama, abs_idx in members:
            if nama.lower() > input_data.nama_baru.lower():
                insert_row_idx = abs_idx
                break
                
        if insert_row_idx == -1:
            insert_row_idx = members[-1] + 1 if members else header_row_idx + 1

        requests = []
        requests.append({
            "insertDimension": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": insert_row_idx, "endIndex": insert_row_idx + 1},
                "inheritFromBefore": True
            }
        })

        current_m_tuple = _parse_month_year(datetime.now().strftime("%m/%Y"))
        row_values = []
        max_col = len(headers)
        
        for col_idx in range(max_col):
            if col_idx == nama_idx:
                row_values.append({"userEnteredValue": {"stringValue": input_data.nama_baru}})
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
        return f"✅ Anggota baru *{input_data.nama_baru}* berhasil ditambah sesuai abjad!\nCell bulan lalu otomatis dihitamkan."
    except Exception as e: return f"Error: {str(e)}"
