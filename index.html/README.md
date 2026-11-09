# Daftar Kelas – Layanan Psikologi Psychosoul

Halaman info dan pendaftaran kelas/event Psychosoul (frontend statis).

## Isi
- `index.html` : halaman utama (daftar kelas, pop up More info, form pendaftaran).
- `apps-script-psychosoul.gs` : backend Google Apps Script (simpan pendaftar ke Google Sheet).

## Sumber data
- Daftar kelas dibaca dari tab `infoclas` di Google Sheet (harus dibagikan "Siapa saja yang memiliki link" - Viewer).
- Pendaftar disimpan ke tab sesuai `KODE_KELAS`; bukti bayar ke folder Google Drive.

## Pengaturan (di bagian `EDIT DATA` pada `index.html`)
- `WA_NUMBER`, `REKENING`, `QR_IMAGE`
- `API_URL` : URL web app Apps Script (berakhiran `/exec`)
- `SHEET_ID`, `SHEET_NAME`

## Deploy
GitHub Pages: Settings > Pages > Deploy from a branch > `main` / root.

## Catatan keamanan
Jangan menyimpan PIN admin di repo. PIN diatur di Apps Script > Project Settings > Script properties (`ADMIN_PIN`).
