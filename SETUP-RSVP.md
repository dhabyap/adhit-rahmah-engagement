# Setup RSVP → Google Sheets

## 1. Buat Google Sheet
1. Buka https://sheets.new → beri nama sheet `RSVP-Engagement` (atau apa saja).
2. File > Project → atau langsung **Extensions > Apps Script**.
3. Hapus isi editor → paste seluruh isi file `AppsScript.gs` (di repo ini).

## 2. Deploy Web App
1. Klik ikon **Deploy** (kanan atas editor) > **New deployment**.
2. App type: **Web app**.
   - **Execute as:** `Me` (akun kamu).
   - **Who has access:** `Anyone` (WAJIB — karena form tamu tanpa login).
   - Description: apa saja.
3. Klik **Deploy** → otorisasi akses (drive + sheets) → selesai.
4. Copy URL yang di-generate, bentuk:
   `https://script.google.com/macros/s/AKf.../exec`

## 3. Tempel di kode
- Ganti `XXXXXXXX` di `admin.html` (baris `const API = ...`) dengan URL lengkap.
- Di `index.html`, isi konstanta `RSVP_API` dengan URL yang sama:
  ```js
  const RSVP_API = "https://script.google.com/macros/s/AKf.../exec";
  ```
  Selama `RSVP_API` kosong (`""`), form otomatis fallback ke WhatsApp — jadi aman kalau belum deploy.

## 4. Test
1. Buka `admin.html?key=RA-ADMIT-2026&auto=1` → harus tampil "Belum ada konfirmasi." (bukan error).
2. Isi form RSVP di `index.html` → submit → data muncul di sheet (tab `RSVP`) + dashboard (auto-refresh 30s).
3. Tes CSV: tombol **Export CSV** di dashboard.

## 5. (Opsional) Rubah token
- Token diApps Script: baris `const TOKEN = "RA-ADMIT-2026";` — ganti, lalu redeploy (Deploy > Manage deployments > Edit > New version).
- Ganti juga `KEY` default di `admin.html` kalau pakai token beda.

## Struktur sheet (bentang otomatis oleh `doPost`)
| Nama | Hadir? | Jumlah | Ucapan | HP | Waktu |
|------|--------|--------|--------|----|-------|
| (string) | (bool "TRUE"/"FALSE") | (angka) | (string) | (string) | (timestamp) |

## Catatan
- `setup()` dan `onOpen()` diApps Script membuat menu **Setup > Prepare sheet + styles** untuk styling awal.
- Kuota Google Apps Script Web App ~ 5k request/hari, 1M request/bulan — lebih dari cukup untuk RSVP satu event.
- CORS: fetch keApps Script pakai `Content-Type: text/plain` (bukan JSON) supaya tidak kena preflight OPTIONS (Apps Script tidak support). Ini yang dipakai di `index.html`.
