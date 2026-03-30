# Kas Kopi Automation

Sistem otomatis untuk mengelola kas komunitas menggunakan AI dan Google Sheets.

Dirancang untuk membantu pencatatan pembayaran, monitoring status anggota, dan rekap kas secara praktis tanpa proses manual.

## Badges
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Hermes](https://img.shields.io/badge/agent-hermes-orange)
![Status](https://img.shields.io/badge/status-active-success)
![License](https://img.shields.io/badge/license-private-lightgrey)

## Fitur
- Ekstraksi data dari screenshot bukti transfer
- Identifikasi nama dan nominal otomatis
- Alokasi pembayaran ke bulan berjalan dan tunggakan
- Tracking status pembayaran per anggota
- Ringkasan kas bulanan
- Rekap tunggakan anggota
- Integrasi langsung dengan Google Sheets

## Prasyarat
- Python 3.10+
- Hermes Agent
- Google Cloud Service Account
- Google Sheets API dan Google Drive API aktif

## Setup

### 1. Clone Repository
```bash
git clone <url-repo-anda> ~/.hermes_kopi
cd ~/.hermes_kopi
```

### 2. Setup Workspace
```bash
export HERMES_HOME=~/.hermes_kopi
hermes setup
```

### 3. Konfigurasi
Buat file `.env` untuk API key LLM yang mendukung vision

Tambahkan file berikut:
- credentials.json dari Google Cloud

Pastikan:
- Service account memiliki akses Editor ke Google Sheets
- SHEET_ID sudah diisi di tools/sheets_tool.py

### 4. Run
```bash
hermes run
```

## Struktur
- tools/sheets_tool.py  
  Modul untuk komunikasi dengan Google Sheets

- SOUL.md  
  Instruksi sistem agent

## Kontributor
Afif Akbar Iskandar  
https://github.com/afifai

