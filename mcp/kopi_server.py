from __future__ import annotations

import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from tools.sheets_tool import (
    cek_status_kas,
    ringkasan_kas_bulan,
    rekap_tunggakan,
    total_piutang_global,
    hall_of_fame,
    tren_bulan_kritis,
    ghosting_alert,
    rekap_pemasukan_aktual,
    bayar_kas,
    konversi_barang,
    tambah_anggota,
    catat_pengeluaran,
    CekKasInput,
    RingkasanBulanInput,
    CekTunggakanInput,
    TotalPiutangInput,
    HallOfFameInput,
    TrenBulanInput,
    GhostingAlertInput,
    RekapPemasukanAktualInput,
    BayarKasInput,
    KonversiBarangInput,
    TambahAnggotaInput,
    CatatPengeluaranInput,
)
from tools.identity_guard import resolve_member, is_admin

mcp = FastMCP("ngopikuy-kas")


def _ok(value):
    return str(value)


@mcp.tool()
def status_kas(nama_anggota: str) -> str:
    """Cek status kas anggota."""
    payload = CekKasInput(nama_anggota=nama_anggota)
    return _ok(cek_status_kas(payload))


@mcp.tool()
def ringkasan_bulan(bulan_tahun: str) -> str:
    """Ringkasan kas pada bulan tertentu. Format MM/YYYY."""
    payload = RingkasanBulanInput(bulan_tahun=bulan_tahun)
    return _ok(ringkasan_kas_bulan(payload))


@mcp.tool()
def rekap_tunggakan_semua() -> str:
    """Rekap tunggakan seluruh anggota."""
    payload = CekTunggakanInput(dummy="x")
    return _ok(rekap_tunggakan(payload))


@mcp.tool()
def kesehatan_kas() -> str:
    """Lihat total kas dan total piutang."""
    payload = TotalPiutangInput(dummy="x")
    return _ok(total_piutang_global(payload))


@mcp.tool()
def hall_of_fame_kas() -> str:
    """Lihat anggota dengan kontribusi kas tertinggi."""
    payload = HallOfFameInput(dummy="x")
    return _ok(hall_of_fame(payload))


@mcp.tool()
def tren_bulan_kritis_kas() -> str:
    """Lihat bulan-bulan paling kritis dari sisi pembayaran."""
    payload = TrenBulanInput(dummy="x")
    return _ok(tren_bulan_kritis(payload))


@mcp.tool()
def ghosting_alert_kas() -> str:
    """Lihat anggota yang lama tidak membayar."""
    payload = GhostingAlertInput(dummy="x")
    return _ok(ghosting_alert(payload))


@mcp.tool()
def pemasukan_aktual_bulan(bulan_tahun: str) -> str:
    """Rekap pemasukan aktual cash dan barang. Format MM/YYYY."""
    payload = RekapPemasukanAktualInput(bulan_tahun=bulan_tahun)
    return _ok(rekap_pemasukan_aktual(payload))


@mcp.tool()
def anggota_tambah(nama_baru: str) -> str:
    """Tambah anggota baru."""
    payload = TambahAnggotaInput(nama_baru=nama_baru)
    return _ok(tambah_anggota(payload))


@mcp.tool()
def kas_bayar(
    nama_anggota: str,
    nominal_per_bulan: int = 50000,
    jumlah_bulan: int = 1,
    bulan_mulai: str | None = None,
) -> str:
    """Catat pembayaran kas."""
    payload = BayarKasInput(
        nama_anggota=nama_anggota,
        nominal_per_bulan=nominal_per_bulan,
        jumlah_bulan=jumlah_bulan,
        bulan_mulai=bulan_mulai,
    )
    return _ok(bayar_kas(payload))


@mcp.tool()
def kas_konversi_barang(
    nama_anggota: str,
    harga_barang: int,
    bulan_mulai: str | None = None,
) -> str:
    """Konversi pembelian barang menjadi pembayaran kas."""
    payload = KonversiBarangInput(
        nama_anggota=nama_anggota,
        harga_barang=harga_barang,
        bulan_mulai=bulan_mulai,
    )
    return _ok(konversi_barang(payload))


@mcp.tool()
def pengeluaran_catat(
    nama_item: str,
    nominal: int | None = None,
    harga_satuan: int | None = None,
    jumlah_item: int | None = None,
    tanggal: str | None = None,
    platform: str | None = None,
) -> str:
    """Catat pengeluaran barang atau operasional."""
    payload = CatatPengeluaranInput(
        nama_item=nama_item,
        nominal=nominal,
        harga_satuan=harga_satuan,
        jumlah_item=jumlah_item,
        tanggal=tanggal,
        platform=platform,
    )
    return _ok(catat_pengeluaran(payload))


@mcp.tool()
def identitas_cek_pelapor(
    telegram_user_id: int | None = None,
    username: str | None = None,
    full_name: str | None = None,
) -> str:
    """Validasi pelapor Telegram terhadap config pairing lokal."""
    ident = resolve_member(
        telegram_user_id=telegram_user_id,
        username=username,
        full_name=full_name,
    )
    if not ident.ok:
        return f"⛔ Pelapor belum terdaftar ({ident.reason}). Hubungi admin untuk pairing akun."

    return f"✅ Terverifikasi sebagai *{ident.sheet_name}* ({ident.reason})."


@mcp.tool()
def status_kas_pelapor(
    telegram_user_id: int | None = None,
    username: str | None = None,
    full_name: str | None = None,
) -> str:
    """Cek status kas berdasarkan identitas pelapor Telegram (tanpa input nama anggota)."""
    ident = resolve_member(
        telegram_user_id=telegram_user_id,
        username=username,
        full_name=full_name,
    )
    if not ident.ok or not ident.sheet_name:
        return f"⛔ Pelapor belum terdaftar ({ident.reason}). Hubungi admin untuk pairing akun."

    payload = CekKasInput(nama_anggota=ident.sheet_name)
    return _ok(cek_status_kas(payload))


@mcp.tool()
def kas_bayar_pelapor(
    telegram_user_id: int | None = None,
    username: str | None = None,
    full_name: str | None = None,
    nominal_per_bulan: int = 50000,
    jumlah_bulan: int = 1,
    bulan_mulai: str | None = None,
) -> str:
    """Catat pembayaran kas untuk pelapor sendiri via pairing Telegram -> nama di sheet."""
    ident = resolve_member(
        telegram_user_id=telegram_user_id,
        username=username,
        full_name=full_name,
    )
    if not ident.ok or not ident.sheet_name:
        return f"⛔ Pelapor belum terdaftar ({ident.reason}). Hubungi admin untuk pairing akun."

    payload = BayarKasInput(
        nama_anggota=ident.sheet_name,
        nominal_per_bulan=nominal_per_bulan,
        jumlah_bulan=jumlah_bulan,
        bulan_mulai=bulan_mulai,
    )
    return _ok(bayar_kas(payload))


@mcp.tool()
def kas_bayar_admin_untuk(
    nama_anggota: str,
    telegram_user_id: int | None = None,
    username: str | None = None,
    full_name: str | None = None,
    nominal_per_bulan: int = 50000,
    jumlah_bulan: int = 1,
    bulan_mulai: str | None = None,
) -> str:
    """Admin-only: catat pembayaran kas untuk anggota lain."""
    ident = resolve_member(
        telegram_user_id=telegram_user_id,
        username=username,
        full_name=full_name,
    )
    if not ident.ok:
        return f"⛔ Pelapor belum terdaftar ({ident.reason})."
    if not is_admin(ident.role):
        return "⛔ Akses ditolak. Fitur ini hanya untuk admin."

    payload = BayarKasInput(
        nama_anggota=nama_anggota,
        nominal_per_bulan=nominal_per_bulan,
        jumlah_bulan=jumlah_bulan,
        bulan_mulai=bulan_mulai,
    )
    return _ok(bayar_kas(payload))


if __name__ == "__main__":
    mcp.run()

