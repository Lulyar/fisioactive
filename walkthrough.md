# ✅ FisioActive — Walkthrough & Setup Guide

## File yang Dibuat

| File | Ukuran | Deskripsi |
|------|--------|-----------|
| [Kode.gs](file:///d:/a_PROJECT%20ASAL/fisioactive/Kode.gs) | ~250 baris | Backend Google Apps Script |
| [index.html](file:///d:/a_PROJECT%20ASAL/fisioactive/index.html) | ~3000 baris | Frontend SPA (HTML + CSS + JS) |

## Perbedaan UI dari Fisiomedika

| Aspek | Fisiomedika | FisioActive |
|-------|-------------|-------------|
| Layout | Top navbar horizontal | **Sidebar kiri** (collapsible) |
| Warna | Teal (`#0d9488`) | **Deep Indigo** (`#6366f1`) |
| Login | Form sederhana tengah | **Split-screen** (branding kiri + form kanan) |
| Modal | Slide panel dari kanan | **Center modal** + backdrop blur |
| Jadwal | Timeline vertikal | **Calendar grid cards** |
| Chart | CSS bar chart | **Donut chart** (conic-gradient) + progress bars |
| Font | Inter | **Outfit** |
| Terapis | Profile cards | **Profile cards** dengan avatar squared |
| Cards | Border-left accent | **Full rounded** + icon besar |
| Gradient | - | ❌ Tidak ada gradient (flat colors) |

## 🔧 Cara Setup

### Step 1: Buka Google Spreadsheet
Buka: https://docs.google.com/spreadsheets/d/1EbwC1bIp5iqIrT81-_9XC1aSKb-ZAs5Svy_LdRltXdc/edit

### Step 2: Buka Apps Script
Menu **Extensions > Apps Script**

### Step 3: Paste Kode.gs
1. Di editor Apps Script, file default `Code.gs` sudah ada
2. Hapus isinya, paste seluruh isi dari [Kode.gs](file:///d:/a_PROJECT%20ASAL/fisioactive/Kode.gs)

### Step 4: Buat file Index.html
1. Klik **+** di sebelah "Files" > **HTML**
2. Beri nama **Index** (tanpa .html)
3. Paste seluruh isi dari [index.html](file:///d:/a_PROJECT%20ASAL/fisioactive/index.html)

### Step 5: Run initSheets
1. Di dropdown fungsi (di toolbar), pilih **initSheets**
2. Klik ▶️ **Run**
3. Authorize jika diminta
4. Ini akan otomatis membuat 6 sheet + headers + 2 user default:
   - `admin` / `admin123` (role: admin)
   - `terapis` / `terapis123` (role: terapis)

### Step 6: Deploy
1. Klik **Deploy > New deployment**
2. Type: **Web app**
3. Execute as: **Me**
4. Who has access: **Anyone**
5. Klik **Deploy**
6. **Copy URL** deployment

### Step 7: Buka dan Mainkan!
Buka URL Web App yang Anda dapatkan di browser.
> [!NOTE]
> FisioActive kini secara otomatis mendeteksi lingkungan Google Apps Script Web App (`google.script.run`). Anda **tidak perlu lagi manual mengganti `API_URL`** di `Index.html` jika Anda menjalankannya langsung sebagai Web App!
>
> Jika Anda meng-host file `index.html` secara eksternal (misal: di Vercel atau dijalankan lokal), baru Anda perlu mengganti `API_URL` di `index.html` dengan URL Web App Anda.

### Step 8: Login
Login dengan kredensial default berikut:
- **Admin**: `admin` / `admin123`
- **Terapis**: `terapis` / `terapis123`

## Hosting di Vercel (Opsional)
Jika ingin deploy frontend di Vercel:
1. Push folder `fisioactive` ke GitHub
2. Import di Vercel
3. Pastikan `API_URL` di `index.html` sudah benar

## Fitur Lengkap (12 Fitur, 2 Role + Publik)

### Publik (Tanpa Login)
- **Cek Status Pasien**: Pencarian riwayat jadwal dan rekam medis menggunakan ID Pasien (e.g. `FA001`) atau ID Rekam Medis (e.g. `MR001`). Fitur ini dapat diakses langsung dari halaman Login tanpa perlu masuk ke sistem.

### Admin (11 menu)
Dashboard, Pendaftaran, Data Pasien, Jadwal, Rekam Medis, Terapis, Jenis Terapi, Laporan, Cek Pasien, Keluar

### Terapis (7 menu)
Dashboard (personal), Jadwal, Rekam Medis, Terapis, Jenis Terapi, Cek Pasien, Keluar
