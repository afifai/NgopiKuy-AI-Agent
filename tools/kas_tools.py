import json
from tools.registry import registry
from tools.sheets_tool import (
    cek_status_kas, CekKasInput,
    ringkasan_kas_bulan, RingkasanBulanInput,
    rekap_tunggakan, CekTunggakanInput,
    total_piutang_global, TotalPiutangInput,
    hall_of_fame, HallOfFameInput,
    tren_bulan_kritis, TrenBulanInput,
    ghosting_alert, GhostingAlertInput,
    rekap_pemasukan_aktual, RekapPemasukanAktualInput,
    bayar_kas, BayarKasInput,
    konversi_barang, KonversiBarangInput,
    tambah_anggota, TambahAnggotaInput,
    catat_pengeluaran, CatatPengeluaranInput
)

TOOLSET_NAME = "kas_management"

def _safe_execute(func, input_data):
    try:
        res = func(input_data)
        return json.dumps({"status": "success", "data": res})
    except Exception as e:
        return json.dumps({"status": "error", "message": str(e)})

# ==========================================
# KATEGORI A: READ TOOLS (BACA DATA)
# ==========================================

def _handler_cek_status(args, **kwargs):
    return _safe_execute(cek_status_kas, CekKasInput(nama_anggota=args.get("nama_anggota", "")))

registry.register(
    name="cek_status_kas", toolset=TOOLSET_NAME,
    schema={
        "name": "cek_status_kas",
        "description": "Cek status pembayaran dan tunggakan uang kas per individu.",
        "parameters": {
            "type": "object",
            "properties": {"nama_anggota": {"type": "string", "description": "Nama anggota"}},
            "required": ["nama_anggota"]
        }
    }, handler=_handler_cek_status
)

def _handler_ringkasan_bulan(args, **kwargs):
    return _safe_execute(ringkasan_kas_bulan, RingkasanBulanInput(bulan_tahun=args.get("bulan_tahun", "")))

registry.register(
    name="ringkasan_kas_bulan", toolset=TOOLSET_NAME,
    schema={
        "name": "ringkasan_kas_bulan",
        "description": "Cek total lunas dan uang terkumpul di bulan tertentu.",
        "parameters": {
            "type": "object",
            "properties": {"bulan_tahun": {"type": "string", "description": "Format MM/YYYY"}},
            "required": ["bulan_tahun"]
        }
    }, handler=_handler_ringkasan_bulan
)

def _handler_rekap_tunggakan(args, **kwargs):
    return _safe_execute(rekap_tunggakan, CekTunggakanInput(dummy="x"))

registry.register(
    name="rekap_tunggakan", toolset=TOOLSET_NAME,
    schema={"name": "rekap_tunggakan", "description": "Lihat daftar anggota yang paling banyak nunggak kas.", "parameters": {"type": "object", "properties": {}}},
    handler=_handler_rekap_tunggakan
)

def _handler_total_piutang(args, **kwargs):
    return _safe_execute(total_piutang_global, TotalPiutangInput(dummy="x"))

registry.register(
    name="total_piutang_global", toolset=TOOLSET_NAME,
    schema={"name": "total_piutang_global", "description": "Cek total uang kas aktual vs piutang (uang yang belum dibayar anggota).", "parameters": {"type": "object", "properties": {}}},
    handler=_handler_total_piutang
)

def _handler_hall_of_fame(args, **kwargs):
    return _safe_execute(hall_of_fame, HallOfFameInput(dummy="x"))

registry.register(
    name="hall_of_fame", toolset=TOOLSET_NAME,
    schema={"name": "hall_of_fame", "description": "Melihat pahlawan kas atau anggota yang bayar lebih dari kewajiban.", "parameters": {"type": "object", "properties": {}}},
    handler=_handler_hall_of_fame
)

def _handler_tren_kritis(args, **kwargs):
    return _safe_execute(tren_bulan_kritis, TrenBulanInput(dummy="x"))

registry.register(
    name="tren_bulan_kritis", toolset=TOOLSET_NAME,
    schema={"name": "tren_bulan_kritis", "description": "Lihat bulan-bulan di mana pembayaran kas paling seret/kritis.", "parameters": {"type": "object", "properties": {}}},
    handler=_handler_tren_kritis
)

def _handler_ghosting_alert(args, **kwargs):
    return _safe_execute(ghosting_alert, GhostingAlertInput(dummy="x"))

registry.register(
    name="ghosting_alert", toolset=TOOLSET_NAME,
    schema={"name": "ghosting_alert", "description": "Lihat daftar anggota yang nunggak lebih dari 6 bulan (ghosting).", "parameters": {"type": "object", "properties": {}}},
    handler=_handler_ghosting_alert
)

def _handler_rekap_pemasukan(args, **kwargs):
    return _safe_execute(rekap_pemasukan_aktual, RekapPemasukanAktualInput(bulan_tahun=args.get("bulan_tahun", "")))

registry.register(
    name="rekap_pemasukan_aktual", toolset=TOOLSET_NAME,
    schema={
        "name": "rekap_pemasukan_aktual",
        "description": "Lihat log rincian pemasukan kas (cash & barang) di bulan tertentu.",
        "parameters": {
            "type": "object",
            "properties": {"bulan_tahun": {"type": "string", "description": "Format MM/YYYY"}},
            "required": ["bulan_tahun"]
        }
    }, handler=_handler_rekap_pemasukan
)

# ==========================================
# KATEGORI B: WRITE TOOLS (UBAH DATA)
# ==========================================

def _handler_bayar_kas(args, **kwargs):
    data = BayarKasInput(
        nama_anggota=args.get("nama_anggota", ""),
        nominal_per_bulan=args.get("nominal_per_bulan", 50000),
        jumlah_bulan=args.get("jumlah_bulan", 1),
        bulan_mulai=args.get("bulan_mulai")
    )
    return _safe_execute(bayar_kas, data)

registry.register(
    name="bayar_kas", toolset=TOOLSET_NAME,
    schema={
        "name": "bayar_kas",
        "description": "Catat pembayaran iuran kas anggota.",
        "parameters": {
            "type": "object",
            "properties": {
                "nama_anggota": {"type": "string", "description": "Nama anggota"},
                "nominal_per_bulan": {"type": "integer", "description": "Nominal per bulan (default 50000)"},
                "jumlah_bulan": {"type": "integer", "description": "Bayar untuk berapa bulan? (default 1)"},
                "bulan_mulai": {"type": "string", "description": "Mulai bayar dari bulan apa? Format MM/YYYY (opsional)"}
            },
            "required": ["nama_anggota"]
        }
    }, handler=_handler_bayar_kas
)

def _handler_konversi_barang(args, **kwargs):
    data = KonversiBarangInput(
        nama_anggota=args.get("nama_anggota", ""),
        harga_barang=args.get("harga_barang", 0),
        bulan_mulai=args.get("bulan_mulai")
    )
    return _safe_execute(konversi_barang, data)

registry.register(
    name="konversi_barang", toolset=TOOLSET_NAME,
    schema={
        "name": "konversi_barang",
        "description": "Catat anggota yang menyumbang barang senilai uang kas.",
        "parameters": {
            "type": "object",
            "properties": {
                "nama_anggota": {"type": "string", "description": "Nama anggota"},
                "harga_barang": {"type": "integer", "description": "Total harga barang yang disumbangkan"},
                "bulan_mulai": {"type": "string", "description": "Format MM/YYYY (opsional)"}
            },
            "required": ["nama_anggota", "harga_barang"]
        }
    }, handler=_handler_konversi_barang
)

def _handler_tambah_anggota(args, **kwargs):
    return _safe_execute(tambah_anggota, TambahAnggotaInput(nama_baru=args.get("nama_baru", "")))

registry.register(
    name="tambah_anggota", toolset=TOOLSET_NAME,
    schema={
        "name": "tambah_anggota",
        "description": "Tambahkan anggota baru ke dalam sheet kas.",
        "parameters": {
            "type": "object",
            "properties": {"nama_baru": {"type": "string", "description": "Nama lengkap anggota baru"}},
            "required": ["nama_baru"]
        }
    }, handler=_handler_tambah_anggota
)

def _handler_catat_pengeluaran(args, **kwargs):
    data = CatatPengeluaranInput(
        nama_item=args.get("nama_item", ""),
        nominal=args.get("nominal"),
        harga_satuan=args.get("harga_satuan"),
        jumlah_item=args.get("jumlah_item"),
        tanggal=args.get("tanggal"),
        platform=args.get("platform", "Sistem")
    )
    return _safe_execute(catat_pengeluaran, data)

registry.register(
    name="catat_pengeluaran", toolset=TOOLSET_NAME,
    schema={
        "name": "catat_pengeluaran",
        "description": "Catat transaksi belanja/pengeluaran ke sheet Transaction.",
        "parameters": {
            "type": "object",
            "properties": {
                "nama_item": {"type": "string", "description": "Nama barang/pengeluaran"},
                "nominal": {"type": "integer", "description": "Harga total atau harga satuan"},
                "harga_satuan": {"type": "integer", "description": "Harga per item (opsional)"},
                "jumlah_item": {"type": "integer", "description": "Jumlah/Qty (opsional)"},
                "tanggal": {"type": "string", "description": "Format DD/MM/YYYY (opsional)"},
                "platform": {"type": "string", "description": "Tempat belanja (opsional, misal: Tokopedia)"}
            },
            "required": ["nama_item"]
        }
    }, handler=_handler_catat_pengeluaran
)
