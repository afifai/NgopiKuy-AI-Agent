import os
import re
import ast
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
def _get_my(s):
    try:
        s_str = str(s).strip()
        match = re.search(r'(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})', s_str)
        if match:
            m = int(match.group(2))
            y = int(match.group(3))
            if y < 100: y += 2000
            return f"{m:02d}/{y}"
    except: pass
    return ""

def _parse_month_year(mm_yyyy):
    try:
        m, y = mm_yyyy.strip().split('/')
        return (int(y), int(m))
    except:
        return (9999, 12)

def _parse_rupiah(val):
    try:
        if not val: return 0
        s = str(val).strip()
        
        # 🔥 FILTER ANTI-ALIEN: Kebal dari data korup 19 Sekstiliun
        if '[' in s or ']' in s or len(s) > 35: return 0
        if '%' in s: return 0
        
        # 🔥 LOGIKA ORIGINAL YANG BENAR (Jangan diubah lagi)
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
    if res.status_code != 200: raise Exception(f"API Error {res.status_code}: {res.text}")
    
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
    h_idx, n_idx = -1, -1
    for r_idx, row in enumerate(data[:15]): 
        for c_idx, cell in enumerate(row):
            if cell['value'].lower() in ['nama', 'nama anggota']:
                h_idx, n_idx = r_idx, c_idx
                break
        if h_idx != -1: break
    if h_idx == -1: return data, -1, -1, [], sheet_id, data
    headers = data[h_idx]
    actual = []
    for row in data[h_idx + 1:]:
        if len(row) <= n_idx: continue
        n_raw = row[n_idx]['value'].strip()
        if "TOTAL PER BULAN" in n_raw.upper(): break
        if n_raw: actual.append(row)
    return headers, h_idx, n_idx, actual, sheet_id, data

def _execute_batch_update(requests):
    if not requests: return {}
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    authed_session = AuthorizedSession(creds)
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}:batchUpdate"
    res = authed_session.post(url, json={"requests": requests})
    if res.status_code != 200: raise Exception(f"BatchUpdate Error: {res.text}")
    return res.json()

def _append_transaction_log(nama, nominal, ket_bulan, tipe="CASH"):
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    authed_session = AuthorizedSession(creds)
    ts = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    vals = [[ts, nama, nominal, f"Iuran {ket_bulan} ({tipe})"]]
    enc = urllib.parse.quote(f"{SHEET_NAME_LOG}!A1")
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{enc}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS"
    authed_session.post(url, json={"values": vals})

# ==================== TOOLS READ ====================
class CekKasInput(BaseModel): nama_anggota: str = Field(..., description="Nama anggota")
def cek_status_kas(input_data: CekKasInput) -> str:
    try:
        headers, _, n_idx, actual, _, _ = _find_header_and_data_v2()
        for row in actual:
            if input_data.nama_anggota.lower() in row[n_idx]['value'].lower():
                b_bayar, tunggak = [], []
                for i, h in enumerate(headers):
                    if re.match(r"^\d{2}/\d{4}$", h['value']):
                        cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                        nom = _parse_rupiah(cell['value'])
                        if cell['color'] == "green" or 0 < nom <= BATAS_IURAN_WAJAR: b_bayar.append(h['value'])
                        elif nom == 0 and cell['color'] != "black": tunggak.append(h['value'])
                
                b_ini = datetime.now().strftime("%m/%Y")
                t_aktif = sorted([b for b in tunggak if _parse_month_year(b) <= _parse_month_year(b_ini)], key=lambda x: _parse_month_year(x))
                res = f"Status *{row[n_idx]['value']}*:\nTerakhir bayar: {b_bayar[-1] if b_bayar else 'Belum'}\n"
                res += f"Tunggakan aktif: {', '.join(t_aktif) if t_aktif else 'AMAN'}"
                return res
        return "Anggota tidak ditemukan."
    except Exception as e: return f"Error: {str(e)}"

class RingkasanBulanInput(BaseModel): bulan_tahun: str = Field(..., description="MM/YYYY")
def ringkasan_kas_bulan(input_data: RingkasanBulanInput) -> str:
    try:
        headers, _, n_idx, actual, _, _ = _find_header_and_data_v2()
        col = next((i for i, h in enumerate(headers) if input_data.bulan_tahun in h['value']), -1)
        if col == -1: return f"Bulan {input_data.bulan_tahun} tidak ditemukan."
        lunas, total = [], 0
        for row in actual:
            cell = row[col] if col < len(row) else {"value": "", "color": "white"}
            nom = _parse_rupiah(cell['value'])
            if cell['color'] == "green" or 0 < nom <= BATAS_IURAN_WAJAR:
                lunas.append(row[n_idx]['value']); total += nom
        return f"📊 *Ringkasan {input_data.bulan_tahun}*\nLunas: {len(lunas)} orang\nTerkumpul: Rp{total:,}"
    except Exception as e: return f"Error: {str(e)}"

class CekTunggakanInput(BaseModel): dummy: str = Field("dummy")
def rekap_tunggakan(input_data: CekTunggakanInput) -> str:
    try:
        headers, _, n_idx, actual, _, _ = _find_header_and_data_v2()
        b_ini = datetime.now().strftime("%m/%Y")
        kol = [i for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(b_ini)]
        rekap = []
        for row in actual:
            t = 0
            for idx in kol:
                cell = row[idx] if idx < len(row) else {"value": "", "color": "white"}
                if cell['color'] not in ["black", "green"] and _parse_rupiah(cell['value']) == 0: t += 1
            if t > 0: rekap.append({"n": row[n_idx]['value'], "t": t})
        res = "⚠️ *Rekap Tunggakan*\n"
        for i, d in enumerate(sorted(rekap, key=lambda x: x['t'], reverse=True)[:15]):
            res += f"{i+1}. {d['n']} ({d['t']} bln)\n"
        return res
    except Exception as e: return f"Error: {str(e)}"

class TotalPiutangInput(BaseModel): dummy: str = Field("dummy")
def total_piutang_global(input_data: TotalPiutangInput) -> str:
    try:
        headers, _, _, actual, _, _ = _find_header_and_data_v2()
        b_ini = datetime.now().strftime("%m/%Y")
        kol = [i for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(b_ini)]
        kas, piutang = 0, 0
        for row in actual:
            for i in kol:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                n = _parse_rupiah(cell['value'])
                if cell['color'] == "black": continue
                if cell['color'] == "green" or 0 < n <= BATAS_IURAN_WAJAR: kas += n
                elif n == 0: piutang += 50000
        return f"💰 *Kesehatan Kas*\nKas: Rp{kas:,} | Piutang: Rp{piutang:,}"
    except Exception as e: return f"Error: {str(e)}"

class HallOfFameInput(BaseModel): dummy: str = Field("dummy")
def hall_of_fame(input_data: HallOfFameInput) -> str:
    try:
        headers, _, n_idx, actual, _, _ = _find_header_and_data_v2()
        b_ini = datetime.now().strftime("%m/%Y")
        kol = [i for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(b_ini)]
        dnt = []
        for row in actual:
            byr, wjb = 0, 0
            for i in kol:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                if cell['color'] != "black":
                    wjb += 50000
                    n = _parse_rupiah(cell['value'])
                    if cell['color'] == "green" or 0 < n <= BATAS_IURAN_WAJAR: byr += n
            if byr > wjb: dnt.append({"n": row[n_idx]['value'], "l": byr - wjb})
        res = "🏆 *Pahlawan Kas*\n"
        for d in sorted(dnt, key=lambda x: x['l'], reverse=True)[:5]: res += f"• {d['n']} (+Rp{d['l']:,})\n"
        return res
    except Exception as e: return f"Error: {str(e)}"

class TrenBulanInput(BaseModel): dummy: str = Field("dummy")
def tren_bulan_kritis(input_data: TrenBulanInput) -> str:
    try:
        headers, _, _, actual, _, _ = _find_header_and_data_v2()
        b_ini = datetime.now().strftime("%m/%Y")
        kol = [(i, h['value']) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(b_ini)]
        tren = []
        for i, h in kol:
            k, t = 0, 0
            for row in actual:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                if cell['color'] != "black":
                    t += 50000
                    n = _parse_rupiah(cell['value'])
                    if cell['color'] == "green" or 0 < n <= BATAS_IURAN_WAJAR: k += n
            if t > 0: tren.append({"b": h, "p": (k/t*100)})
        res = "📉 *Bulan Kritis*\n"
        for d in sorted(tren, key=lambda x: x['p'])[:3]: res += f"• {d['b']} ({d['p']:.1f}%)\n"
        return res
    except Exception as e: return f"Error: {str(e)}"

class GhostingAlertInput(BaseModel): dummy: str = Field("dummy")
def ghosting_alert(input_data: GhostingAlertInput) -> str:
    try:
        headers, _, n_idx, actual, _, _ = _find_header_and_data_v2()
        b_ini = datetime.now().strftime("%m/%Y")
        kol = sorted([i for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) <= _parse_month_year(b_ini)])
        gst = []
        for row in actual:
            s = 0
            for i in kol:
                cell = row[i] if i < len(row) else {"value": "", "color": "white"}
                if cell['color'] == "black": continue
                if cell['color'] == "green" or 0 < _parse_rupiah(cell['value']) <= BATAS_IURAN_WAJAR: s = 0
                else: s += 1
            if s >= 6: gst.append({"n": row[n_idx]['value'], "s": s})
        res = "👻 *Ghosting Alert*\n"
        for d in sorted(gst, key=lambda x: x['s'], reverse=True)[:10]: res += f"• {d['n']} ({d['s']} bln)\n"
        return res if gst else "Aman!"
    except Exception as e: return f"Error: {str(e)}"

class RekapPemasukanAktualInput(BaseModel): bulan_tahun: str = Field(..., description="MM/YYYY")
def rekap_pemasukan_aktual(input_data: RekapPemasukanAktualInput) -> str:
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        res = AuthorizedSession(creds).get(f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{urllib.parse.quote(SHEET_NAME_LOG)}")
        rows = res.json().get('values', [])
        c, b, dt = 0, 0, []
        for r in rows[1:]:
            try:
                if len(r) >= 4:
                    t_p = str(r).strip().split(' ')
                    if len(t_p) >= 10 and t_p[3:] == input_data.bulan_tahun:
                        n, k = _parse_rupiah(str(r)), str(r)
                        if "(BARANG)" in k.upper(): b += n
                        else: c += n
                        dt.append(f"• {str(r)}: Rp{n:,}")
            except: continue
        return f"📈 *Pemasukan {input_data.bulan_tahun}*\nCash: Rp{c:,} | Barang: Rp{b:,}\n" + "\n".join(dt[:10])
    except Exception as e: return f"Error: {str(e)}"

# ==================== TOOLS WRITE ====================
class BayarKasInput(BaseModel):
    nama_anggota: str = Field(..., description="Nama anggota")
    nominal_per_bulan: int = Field(50000)
    jumlah_bulan: int = Field(1)
    bulan_mulai: str = Field(None)

def bayar_kas(input_data: BayarKasInput) -> str:
    try:
        headers, h_idx, n_idx, actual, sheet_id, raw_data = _find_header_and_data_v2()
        abs_r, n_f = -1, ""
        for i, row in enumerate(raw_data[h_idx+1:]):
            if len(row) > n_idx and input_data.nama_anggota.lower() in row[n_idx]['value'].lower():
                abs_r, n_f = h_idx + 1 + i, row[n_idx]['value']; break
        if abs_r == -1: return "Anggota tidak ditemukan."
        b_s = input_data.bulan_mulai or datetime.now().strftime("%m/%Y")
        kol = sorted([(i, h['value']) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) >= _parse_month_year(b_s)])
        to_up, list_b = [], []
        for i, h in kol:
            cell = raw_data[abs_r][i] if i < len(raw_data[abs_r]) else {"value": "", "color": "white"}
            if cell['color'] not in ["green", "black"] and _parse_rupiah(cell['value']) == 0:
                to_up.append(i); list_b.append(h)
                if len(to_up) == input_data.jumlah_bulan: break
        if not to_up: return "⚠️ Tidak ada bulan kosong yang bisa diisi."
        reqs = [{"updateCells": {"range": {"sheetId": sheet_id, "startRowIndex": abs_r, "endRowIndex": abs_r + 1, "startColumnIndex": i, "endColumnIndex": i + 1}, "rows": [{"values": [{"userEnteredValue": {"numberValue": input_data.nominal_per_bulan}, "userEnteredFormat": {"numberFormat": {"type": "CURRENCY", "pattern": "Rp#,##0"}}}]}], "fields": "userEnteredValue,userEnteredFormat(numberFormat)"}} for i in to_up]
        _execute_batch_update(reqs); _append_transaction_log(n_f, input_data.nominal_per_bulan*len(to_up), ', '.join(list_b), "CASH")
        return f"✅ Berhasil catat Rp{input_data.nominal_per_bulan*len(to_up):,} dari *{n_f}*."
    except Exception as e: return f"Error: {str(e)}"

class KonversiBarangInput(BaseModel):
    nama_anggota: str = Field(..., description="Nama anggota")
    harga_barang: int = Field(...)
    bulan_mulai: str = Field(None)

def konversi_barang(input_data: KonversiBarangInput) -> str:
    try:
        headers, h_idx, n_idx, actual, sheet_id, raw_data = _find_header_and_data_v2()
        abs_r, n_f = -1, ""
        for i, row in enumerate(raw_data[h_idx+1:]):
            if len(row) > n_idx and input_data.nama_anggota.lower() in row[n_idx]['value'].lower():
                abs_r, n_f = h_idx + 1 + i, row[n_idx]['value']; break
        if abs_r == -1: return "Anggota tidak ditemukan."
        j_b = input_data.harga_barang // 50000
        if j_b < 1: return "Nominal kurang dari Rp50.000."
        b_s = input_data.bulan_mulai or datetime.now().strftime("%m/%Y")
        kol = sorted([(i, h['value']) for i, h in enumerate(headers) if re.match(r"^\d{2}/\d{4}$", h['value']) and _parse_month_year(h['value']) >= _parse_month_year(b_s)])
        to_up, list_b = [], []
        for i, h in kol:
            cell = raw_data[abs_r][i] if i < len(raw_data[abs_r]) else {"value": "", "color": "white"}
            if cell['color'] not in ["green", "black"] and _parse_rupiah(cell['value']) == 0:
                to_up.append(i); list_b.append(h)
                if len(to_up) == j_b: break
        if not to_up: return "⚠️ Tidak ada bulan kosong yang bisa diisi."
        reqs = [{"updateCells": {"range": {"sheetId": sheet_id, "startRowIndex": abs_r, "endRowIndex": abs_r + 1, "startColumnIndex": i, "endColumnIndex": i + 1}, "rows": [{"values": [{"userEnteredValue": {"stringValue": "Barang"}, "userEnteredFormat": {"backgroundColor": {"red": 0.4, "green": 0.8, "blue": 0.4}}}]}], "fields": "userEnteredValue,userEnteredFormat(backgroundColor)"}} for i in to_up]
        _execute_batch_update(reqs); _append_transaction_log(n_f, len(to_up)*50000, ', '.join(list_b), "BARANG")
        return f"✅ Berhasil konversi barang Rp{input_data.harga_barang:,} dari *{n_f}*."
    except Exception as e: return f"Error: {str(e)}"

class TambahAnggotaInput(BaseModel): nama_baru: str = Field(..., description="Nama lengkap")
def tambah_anggota(input_data: TambahAnggotaInput) -> str:
    try:
        headers, h_idx, n_idx, actual, sheet_id, raw_data = _find_header_and_data_v2()
        n_b = input_data.nama_baru.strip(); start_r = h_idx + 1; members = []
        for i, row in enumerate(raw_data[start_r:]):
            abs_i = start_r + i
            if len(row) > n_idx:
                n_e = row[n_idx]['value'].strip()
                if "TOTAL PER BULAN" in n_e.upper(): break
                if n_e and not any(x in n_e.upper() for x in BLOCK_LIST_NAMA):
                    if n_e.lower() == n_b.lower(): return f"⚠️ Anggota *{n_e}* sudah ada."
                    members.append((n_e, abs_i))
        ins_r, last_idx = -1, start_r
        for n, idx in members:
            last_idx = int(idx)
            if n.lower() > n_b.lower(): ins_r = int(idx); break
        if ins_r == -1: ins_r = last_idx + 1 if members else start_r
        cur_m = _parse_month_year(datetime.now().strftime("%m/%Y"))
        row_vals = []
        for c in range(len(headers)):
            if c == n_idx: row_vals.append({"userEnteredValue": {"stringValue": n_b}})
            else:
                h = headers[c]['value']
                if re.match(r"^\d{2}/\d{4}$", h):
                    m_t = _parse_month_year(h)
                    if m_t < cur_m: row_vals.append({"userEnteredFormat": {"backgroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}}})
                    else: row_vals.append({"userEnteredFormat": {"backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}})
                else: row_vals.append({})
        reqs = [{"insertDimension": {"range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": ins_r, "endIndex": ins_r + 1}, "inheritFromBefore": True}}, {"updateCells": {"range": {"sheetId": sheet_id, "startRowIndex": ins_r, "endRowIndex": ins_r + 1, "startColumnIndex": 0, "endColumnIndex": len(headers)}, "rows": [{"values": row_vals}], "fields": "userEnteredValue,userEnteredFormat.backgroundColor"}}]
        _execute_batch_update(reqs)
        return f"✅ Anggota *{n_b}* berhasil ditambah."
    except Exception as e: return f"Error: {str(e)}"

# ==================== TRANSACTION TOOL (FLEXIBLE FORMULA) ====================
class CatatPengeluaranInput(BaseModel):
    nama_item: str = Field(..., description="Keterangan pengeluaran")
    nominal: int = Field(None, description="Harga total atau harga satuan")
    harga_satuan: int = Field(None, description="Harga satuan barang")
    jumlah_item: int = Field(None, description="Jumlah / Qty barang")
    tanggal: str = Field(None, description="DD/MM/YYYY")

def catat_pengeluaran(input_data: CatatPengeluaranInput) -> str:
    try:
        sheet_name_trans = "Transaction"
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        authed_session = AuthorizedSession(creds)
        
        # 1. Normalisasi Input (Skenario Fleksibel)
        qty = input_data.jumlah_item if input_data.jumlah_item is not None else 1
        
        # Tentukan Harga Satuan
        if input_data.harga_satuan is not None:
            price_per_unit = input_data.harga_satuan
        elif input_data.nominal is not None:
            # Jika jumlah_item > 1, nominal dianggap harga satuan. Jika tidak, harga total.
            price_per_unit = input_data.nominal 
        else:
            return "Error: Nominal atau Harga Satuan harus diisi."

        tgl_input = input_data.tanggal if input_data.tanggal else datetime.now().strftime("%d/%m/%Y")
        m_y_target = _get_my(tgl_input)
        if not m_y_target: return f"Error: Format tanggal tidak valid ({tgl_input})"
        
        # 2. Ambil Meta & Data (A:I)
        url_m = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}"
        r_meta = authed_session.get(url_m).json()
        s_id = next((s.get('properties', {}).get('sheetId', 0) for s in r_meta.get('sheets', []) if s.get('properties', {}).get('title') == sheet_name_trans), 0)
        
        url_d = f"https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/{urllib.parse.quote(sheet_name_trans)}!A:I"
        rows = authed_session.get(url_d).json().get('values', [])
        
        max_id = 0
        last_physical_row = 1
        month_start_idx = -1
        
        # 3. Scan Fisik Baris (Cari ID & Titik Awal Bulan)
        for i, row in enumerate(rows):
            if i < 2: continue # Skip Header
            if not row or not any(str(x).strip() for x in row): continue
            
            last_physical_row = i
            
            raw_id = str(row) if len(row) > 0 else ""
            id_match = re.search(r'(\d+)', raw_id)
            if id_match:
                max_id = max(max_id, int(id_match.group(1)))
            
            if len(row) > 1 and _get_my(row) == m_y_target:
                if month_start_idx == -1: month_start_idx = i

        # 4. Tentukan Lokasi & Formula
        new_row_idx = last_physical_row + 1
        new_id = max_id + 1
        if month_start_idx == -1: month_start_idx = new_row_idx

        row_sheet_idx = new_row_idx + 1
        start_row_sheet = month_start_idx + 1 
        end_row_sheet = new_row_idx + 1
        
        formula_total_item = f"=D{row_sheet_idx}*E{row_sheet_idx}"
        formula_sum_monthly = f"=SUM(F{start_row_sheet}:F{end_row_sheet})"

        # 5. Build Cells & Format
        b_s = {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}}
        brd = {"top": b_s, "bottom": b_s, "left": b_s, "right": b_s}
        
        def mk_c(v, t="stringValue", a="CENTER", is_cur=False):
            fmt = {"borders": brd, "verticalAlignment": "MIDDLE", "horizontalAlignment": a}
            if is_cur: fmt["numberFormat"] = {"type": "CURRENCY", "pattern": "Rp#,##0"}
            return {"userEnteredValue": {t: v}, "userEnteredFormat": fmt}
            
        def mk_f(f_str, a="CENTER"):
            fmt = {"borders": brd, "verticalAlignment": "MIDDLE", "horizontalAlignment": a, "numberFormat": {"type": "CURRENCY", "pattern": "Rp#,##0"}}
            return {"userEnteredValue": {"formulaValue": f_str}, "userEnteredFormat": fmt}

        row_vals = [
            mk_c(new_id, "numberValue"),
            mk_c(tgl_input),
            mk_c(input_data.nama_item, "stringValue", "LEFT"),
            mk_c(qty, "numberValue"),
            mk_c(price_per_unit, "numberValue", "RIGHT", True),
            mk_f(formula_total_item, "RIGHT"),
            mk_c("Sistem"),
            mk_c("LUNAS")
        ]
        
        if month_start_idx == new_row_idx: row_vals.append(mk_f(formula_sum_monthly))
        else: row_vals.append(mk_c("", "stringValue"))

        # 6. Batch Update
        reqs = [
            {"updateCells": {"range": {"sheetId": s_id, "startRowIndex": new_row_idx, "endRowIndex": new_row_idx + 1, "startColumnIndex": 0, "endColumnIndex": len(row_vals)}, "rows": [{"values": row_vals}], "fields": "userEnteredValue,userEnteredFormat"}}
        ]
        
        if month_start_idx != new_row_idx:
            reqs.append({"updateCells": {"range": {"sheetId": s_id, "startRowIndex": month_start_idx, "endRowIndex": month_start_idx + 1, "startColumnIndex": 8, "endColumnIndex": 9}, "rows": [{"values": [mk_f(formula_sum_monthly)]}], "fields": "userEnteredValue,userEnteredFormat"}})
            reqs.append({"unmergeCells": {"range": {"sheetId": s_id, "startRowIndex": month_start_idx, "endRowIndex": new_row_idx, "startColumnIndex": 8, "endColumnIndex": 9}}})
        
        reqs.append({"mergeCells": {"range": {"sheetId": s_id, "startRowIndex": month_start_idx, "endRowIndex": new_row_idx + 1, "startColumnIndex": 8, "endColumnIndex": 9}, "mergeType": "MERGE_ALL"}})
        
        _execute_batch_update(reqs)
        return f"✅ Berhasil catat ID #{new_id}. {input_data.nama_item} (Qty: {qty}). Kalkulasi & Rekap otomatis via formula."

    except Exception as e:
        import traceback
        return f"Error: {str(e)}\n{traceback.format_exc()}"