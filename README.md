# 📷 Scan QR APAR & FSS — Pemeliharaan Bulanan

Aplikasi web sederhana (tanpa server/backend) untuk scan QR/barcode kode alat
APAR & FSS, mengisi checklist **KONDISI** dan **KETERANGAN** (sesuai form
*Pemeliharaan APAR Bulanan*), lalu mengekspor hasilnya kembali ke file Excel
**Data FSS & APAR**.

## ✨ Fitur

- **Scan QR** kode alat (pakai kamera HP) atau ketik kode manual.
- Data alat (Ruangan, Gedung, Lantai, Merk, Kapasitas, Jenis Gas, Kode)
  otomatis muncul dari `data.js` (sumber: `DATA FSS & APAR`).
- Checklist **KONDISI**: Kebersihan Tabung APAR, Pemeriksaan Indikator
  Tekanan, Pemeriksaan Kunci Pengaman, Pemeriksaan Selang Semprot,
  Pemeriksaan Nozzle, TAG Label Pemeliharaan, dan **EVIDEN** (otomatis
  terisi berdasarkan ada/tidaknya foto bukti).
- Kolom **KETERANGAN** + tombol **Upload Foto** (langsung buka kamera/galeri
  HP) sebagai bukti pemeriksaan.
- Tab **Riwayat & Rekap**: daftar semua alat yang sudah discan/diisi
  (ditandai ✅), pencarian, progres tugas (`x / total selesai`, dan
  **"Done All"** kalau semua sudah selesai), serta tombol **Download Excel**.
- Download Excel berisi seluruh data asli + kolom tambahan `QRkode`,
  `KONDISI` (7 sub-kolom), `KETERANGAN`, foto bukti (ter-embed di sel),
  tanggal periksa, dan status tugas.
- Generator QR (`generate.html`) untuk mencetak/print QR semua alat
  berdasarkan kolom `kode`.
- 100% berjalan di browser (client-side). Riwayat pemeriksaan tersimpan di
  `localStorage` HP/perangkat yang dipakai scan.

## 📁 Struktur Project

```
.
├── index.html       # Halaman utama: scan, checklist, riwayat, download
├── generate.html     # Halaman generate/print QR kode alat
├── data.js           # Data master FSS & APAR (hasil ekspor dari Excel)
├── style.css          # Style bersama (dark theme, mobile-friendly)
└── README.md
```

## 🚀 Cara Pakai

1. Buka `index.html` langsung di browser HP/laptop (atau host via GitHub
   Pages, lihat di bawah).
2. Tab **Scan** → tekan **Mulai Scan QR**, izinkan akses kamera, arahkan ke
   QR alat. Data alat muncul otomatis.
3. Isi checklist **KONDISI**, tulis **KETERANGAN**, upload foto bukti bila
   perlu, lalu tekan **Simpan Hasil Pemeriksaan**.
4. Tab **Riwayat & Rekap** → lihat semua alat yang sudah diperiksa (✅),
   pantau progres tugas, lalu **Download Excel** kapan saja untuk laporan.
5. Butuh cetak QR baru/ulang? Buka `generate.html`.

## 🔄 Memperbarui data alat (`data.js`)

`data.js` dibuat dari sheet **APAR** pada file
`DATA_FSS___APAR_-_Copy.xlsx`. Kalau data master di Excel berubah
(tambah/kurang alat, ganti merk, dsb), regenerate `data.js` dengan salah
satu cara:

- **Manual**: edit array `APAR_DATA` di `data.js` langsung (formatnya JSON
  biasa, tiap alat 1 objek dengan field `no, ruangan, gedung, lantai, merk,
  kapasitas, jenisGas, kode, id`).
- **Dari Excel** (perlu Python + openpyxl), lalu tempel ulang ke `data.js`:
  ```python
  import openpyxl, json
  wb = openpyxl.load_workbook("DATA_FSS___APAR_-_Copy.xlsx", data_only=True)
  ws = wb["APAR"]
  rows = []
  for row in ws.iter_rows(min_row=3, values_only=True):
      no, ruangan, gedung, lantai, merk, kap, jenis, kode, qr = row
      if ruangan is None and gedung is None:
          continue
      rows.append({
          "no": no, "ruangan": ruangan, "gedung": gedung, "lantai": lantai,
          "merk": merk, "kapasitas": kap, "jenisGas": jenis,
          "kode": str(int(kode)) if isinstance(kode, (int, float)) else (str(kode).strip() if kode else None),
      })
  for i, r in enumerate(rows):
      r["id"] = r["kode"] if r["kode"] else f"row{i+1}"
  print(json.dumps(rows, ensure_ascii=False, indent=2))
  ```
  Setiap baris **wajib** punya kolom `kode` supaya bisa discan — kalau
  kosong, alat tetap muncul di rekap/download tapi tidak bisa dicari lewat
  scan QR.

## ⬆️ Upload ke GitHub

Repo ini sudah disiapkan sebagai git repository lokal (`git init` + commit
pertama). Untuk mengunggahnya ke akun GitHub Anda:

```bash
# 1. Buat repo baru di GitHub (lewat web, tanpa README/gitignore bawaan)
# 2. Di folder project ini, jalankan:
git remote add origin https://github.com/USERNAME/NAMA-REPO.git
git branch -M main
git push -u origin main
```

Ganti `USERNAME/NAMA-REPO` sesuai repo GitHub Anda.

### Hosting gratis via GitHub Pages (opsional)

Setelah push ke GitHub:
1. Buka repo → **Settings** → **Pages**.
2. Source: pilih branch `main`, folder `/ (root)` → **Save**.
3. Tunggu 1–2 menit, aplikasi bisa diakses via
   `https://USERNAME.github.io/NAMA-REPO/` — bisa langsung dibuka dari HP
   untuk scan QR tanpa perlu install apa-apa (butuh izin kamera & koneksi
   internet untuk load library QR/Excel dari CDN).

## 🛠️ Library yang dipakai (via CDN)

- [html5-qrcode](https://github.com/mebjas/html5-qrcode) — scan QR pakai kamera.
- [qrcodejs](https://github.com/davidshimjs/qrcodejs) — generate QR di `generate.html`.
- [ExcelJS](https://github.com/exceljs/exceljs) — membuat file `.xlsx` (termasuk embed foto) langsung di browser.

Tidak ada dependency npm/build step — cukup buka file `.html`-nya.

## ⚠️ Catatan

- Riwayat pemeriksaan disimpan di `localStorage` **per perangkat/browser**.
  Kalau dipakai beberapa petugas dengan HP berbeda, masing-masing punya
  riwayat sendiri — gabungkan hasilnya dengan mengunduh Excel dari
  tiap perangkat lalu digabung manual, atau host di GitHub Pages dan
  sepakati satu HP sebagai perangkat pencatat utama.
- Kolom `QRkode` yang sebelumnya `#VALUE!` di file sumber sekarang diisi
  otomatis dengan nilai `kode` yang sama dipakai untuk membuat QR-nya.
