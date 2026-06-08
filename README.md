# Dashboard Compare SAP vs Prodsys

Dashboard web untuk membandingkan data goods receipt dari SAP dengan data dari sistem Prodsys, menggantikan proses manual via Excel macro.

---

## Latar Belakang

Rekonsiliasi dilakukan antara:
- **SAP**: data goods receipt movement type **531** (by-product), **888** (main product), dikurangi pembatalan **532** / **889**
- **Prodsys**: data dari sistem eksternal, bersumber dari table `ZPPCPFINT_GRAUTO`

Dashboard memungkinkan operator men-download data SAP secara otomatis via macro, atau mengupload file export secara manual, lalu sistem menghitung selisih qty per order / material secara otomatis.

Akses internal: `http://10.204.10.32:8000`

---

## Tech Stack

| Komponen  | Teknologi                         |
|-----------|-----------------------------------|
| Backend   | Python 3 + Flask 3.x              |
| Database  | SQLite (`dashboard.db`)           |
| Frontend  | HTML / CSS / JavaScript (vanilla) |
| File      | Excel (.xlsx/.xls) atau CSV (tab-separated) |
| Export    | openpyxl (styled Excel)           |
| Automasi  | SAP GUI Scripting (VBScript)      |

---

## Cara Menjalankan

```bash
# 1. Install dependencies (sekali saja)
pip install -r requirements.txt

# 2. Jalankan server
python app.py
# atau klik ganda run_program.bat

# 3. Akses di browser
http://10.204.10.32:8000
```

Agar bisa diakses dari komputer lain di jaringan, buka port 8000 di Windows Firewall (jalankan PowerShell sebagai Administrator):

```powershell
New-NetFirewallRule -DisplayName "Dashboard SAP port 8000" `
    -Direction Inbound -Protocol TCP -LocalPort 8000 -Action Allow
```

---

## Setup Kredensial SAP (`.env`)

Untuk fitur automasi SAP macro, buat file `.env` di root project dengan isi:

```
SAP_CLIENT=100
SAP_USER=username_sap_anda
SAP_PASS=password_sap_anda
SAP_SYSTEM=nama_sistem_di_sap_logon
```

Lihat `.env.example` sebagai template. File `.env` **tidak di-commit ke git** (ada di `.gitignore`), sehingga kredensial tidak tersimpan di repository.

> SAP Logon harus sudah terpasang dan terdaftar sistem SAP-nya. Program akan login otomatis saat automasi dijalankan.

---

## Deployment ke PC Server (Git Workflow)

Program berjalan di PC server. Setiap ada perubahan dari laptop developer, deploy ke server dengan:

```bash
# Di PC server (sekali saja — setup awal)
git clone https://github.com/panji-wicaksono/dashboard-compare-sap.git

# Di PC server (setiap ada update)
git pull origin main

# Restart Flask setelah pull
python app.py
```

Repository bersifat public, sehingga `git pull` tidak memerlukan kredensial GitHub.

---

## Alur Penggunaan

Ada dua cara mengisi data: **otomatis via SAP Macro** (direkomendasikan) atau **manual upload**.

### Mode 1 — Otomatis via SAP Macro

1. Buka dashboard → tab **Upload Data**
2. Di bagian **⚡ Upload via SAP Macro**, pilih tanggal download
3. Klik **▶ Jalankan Automasi**
4. Program akan otomatis:
   - Login ke SAP
   - Download CAUFV → ekstrak Production Order → download AUFM → download ZPPCPFINT_GRAUTO
   - Upload semua data ke database
5. Pantau progress real-time di log yang muncul di UI
6. Setelah selesai, pindah ke tab **Perbandingan**

> Jika data untuk tanggal yang sama sudah ada, data lama **akan diganti** dengan data terbaru (bukan ditambahkan).

### Mode 2 — Manual Upload

1. Export masing-masing tabel dari SAP/Prodsys ke file .xlsx atau .xls
2. Buka dashboard → tab **Upload Data**
3. Upload file AUFM, CAUFV, MAKT, dan Prodsys secara terpisah
4. Pindah ke tab **Perbandingan**

### Melihat Hasil Perbandingan

1. Di tab **Perbandingan**, tanggal terbaru yang ada datanya otomatis terpilih
2. Pilih tanggal lain jika perlu, lalu klik **Tampilkan**
3. Gunakan filter status / search untuk menyaring data
4. Klik **Export Excel** untuk mengunduh hasil

---

## Upload Data

Data diupload per tabel. Upload file manual bersifat **upsert**: jika data sudah ada (berdasarkan primary key SAP), baris baru diabaikan. Upload via macro untuk tanggal yang sama akan **mengganti** data lama.

### AUFM — Goods Receipt SAP

Export dari SAP (table AUFM via SE16). Kolom wajib:

| Kolom   | Keterangan                              |
|---------|-----------------------------------------|
| MANDT   | Client SAP                              |
| MBLNR   | Nomor material document                 |
| MJAHR   | Tahun material document                 |
| ZEILE   | Nomor baris (line item)                 |
| AUFNR   | Nomor production order                  |
| MATNR   | Nomor material                          |
| BWART   | Movement type (531/532/888/889)         |
| MENGE   | Quantity                                |
| MEINS   | Unit of measure                         |
| BLDAT   | Tanggal posting (format DD.MM.YYYY)     |
| WERKS   | Plant                                   |

Kolom opsional:

| Kolom | Keterangan                      |
|-------|---------------------------------|
| LGORT | Storage location (SHUV dikecualikan dari perhitungan) |

> Primary key: `MANDT + MBLNR + MJAHR + ZEILE`

---

### CAUFV — Production Order Header

Export dari SAP (table CAUFV via SE16). Kolom wajib:

| Kolom | Keterangan                          |
|-------|-------------------------------------|
| AUFNR | Nomor production order              |
| AUART | Order type (mis. FCHO)              |

Kolom opsional:

| Kolom | Keterangan                                    |
|-------|-----------------------------------------------|
| GSTRP | Tanggal mulai produksi (format DD.MM.YYYY) — digunakan sebagai filter tanggal di dashboard |
| WERKS | Plant                                         |

> Primary key: `AUFNR`

---

### MAKT — Deskripsi Material

Export dari SAP (table MAKT via SE16). Kolom wajib:

| Kolom | Keterangan           |
|-------|----------------------|
| MATNR | Nomor material       |
| MAKTX | Deskripsi material   |

> Primary key: `MATNR`

---

### Prodsys — Data ZPPCPFINT_GRAUTO

Export dari sistem Prodsys (table ZPPCPFINT_GRAUTO via SE16). Kolom wajib:

| Kolom       | Keterangan                                          |
|-------------|-----------------------------------------------------|
| SAPID       | ID unik record Prodsys                              |
| CAUFV_AUFN  | Nomor production order                              |
| CAUFV_WERK  | Plant                                               |
| MSEG_MATNR  | Nomor material                                      |
| MSEG_MENGE  | Quantity original (disimpan, tidak dipakai compare) |
| MSEG_CONVM  | Quantity terkonversi — **digunakan sebagai Qty Prodsys** |
| ZDATA1      | Data count — **ditampilkan sebagai kolom Count**    |

Kolom opsional:

| Kolom       | Keterangan                                          |
|-------------|-----------------------------------------------------|
| CAUFV_GSTRP | Tanggal mulai produksi (format DD.MM.YYYY) — sumber utama filter tanggal |
| AFRUD_ISDD  | Tanggal aktual (ISO) — fallback jika CAUFV_GSTRP tidak ada |
| MSEG_MEINS  | Unit of measure                                     |
| CAUFV_AUAR  | Order type                                          |
| KEY_STATUS  | Status record                                       |

> Primary key: `SAPID + CAUFV_WERK`

---

## Logika Perhitungan

### Qty SAP
Dihitung per `AUFNR + MATNR`, dengan cancellation diperhitungkan:

```
Qty SAP = SUM(MENGE jika BWART IN ('531','888'))
        - SUM(MENGE jika BWART IN ('532','889'))
```

Baris dengan `LGORT = 'SHUV'` dikecualikan.

### Matching Tanggal
SAP tidak difilter berdasarkan `BLDAT` (tanggal posting) karena tanggal posting SAP bisa berbeda dengan `GSTRP`. Sebagai gantinya, data SAP diambil untuk semua `AUFNR` yang ada di Prodsys pada tanggal yang dipilih.

### Qty Prodsys
```
Qty Prodsys = SUM(CONVM)  per AUFNR + MATNR
```

### Count
```
Count = SUM(ZDATA1)  per AUFNR + MATNR
```

---

## Logika Status & Warna

Selisih = **Qty SAP − Qty Prodsys**

| Kondisi                     | Warna Baris   | Status Label          |
|-----------------------------|:-------------:|-----------------------|
| Qty SAP = Qty Prodsys       | Hijau         | `MATCH`               |
| Qty SAP > Qty Prodsys       | Merah         | `SAP LEBIH BESAR`     |
| Qty Prodsys > Qty SAP       | Oranye        | `PRODSYS LEBIH BESAR` |
| Hanya ada di SAP            | Merah muda    | `HANYA DI SAP`        |
| Hanya ada di Prodsys        | Kuning        | `HANYA DI PRODSYS`    |

Toleransi MATCH: `|Qty SAP − Qty Prodsys| < 0.01`

---

## Kolom Output Dashboard

| Kolom       | Keterangan                                                          |
|-------------|---------------------------------------------------------------------|
| Order       | Nomor production order (`AUFNR`)                                    |
| Material    | Nomor material (`MATNR`)                                            |
| Deskripsi   | Nama material dari MAKT                                             |
| Order Type  | Tipe order dari CAUFV (`AUART`)                                     |
| Plant       | Kode plant (`WERKS`)                                                |
| UoM         | Satuan (`MEINS`)                                                    |
| Qty SAP     | Qty net SAP (3 desimal)                                             |
| Qty Prodsys | Qty dari `MSEG_CONVM` (3 desimal)                                   |
| Count       | Data count dari `ZDATA1` (integer)                                  |
| Selisih     | Qty SAP − Qty Prodsys (positif = SAP lebih besar)                  |
| Status      | Label kondisi perbandingan                                          |

Tabel menampilkan **subtotal per production order** (baris abu-abu) yang merangkum total qty dan selisih seluruh material dalam order tersebut.

---

## Struktur File Project

```
dashboard-compare-sap/
├── app.py                        # Backend Flask + SQLite
├── templates/
│   └── index.html                # Frontend (single page)
├── static/
│   └── favicon.svg               # Icon browser tab
├── macro sap/
│   ├── sap_login.vbs             # Login SAP otomatis (baca kredensial dari .env)
│   ├── caufv.vbs                 # Download table CAUFV dari SAP
│   ├── aufm.vbs                  # Download table AUFM dari SAP
│   └── zppcpfint_grauto.vbs      # Download table ZPPCPFINT_GRAUTO dari SAP
├── data upload/                  # File hasil download macro (tidak di-commit)
│   ├── caufv.xls
│   ├── aufm.xls
│   └── zppcpfint_grauto.xls
├── .env                          # Kredensial SAP (TIDAK di-commit, ada di .gitignore)
├── .env.example                  # Template .env (di-commit sebagai panduan)
├── .gitignore
├── dashboard.db                  # Database SQLite (auto-create, tidak di-commit)
├── requirements.txt              # Dependencies Python
└── run_program.bat               # Shortcut jalankan server (Windows)
```
