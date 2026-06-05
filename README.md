# Dashboard Compare SAP vs Prodsys

Dashboard web untuk membandingkan data goods receipt dari SAP dengan data dari sistem Prodsys, menggantikan proses manual via Excel macro.

---

## Latar Belakang

Rekonsiliasi dilakukan antara:
- **SAP**: data goods receipt movement type **531** (by-product), **888** (main product), dikurangi pembatalan **532** / **889**
- **Prodsys**: data dari sistem eksternal, bersumber dari table `ZPPCPFINT_GRAUTO`

Dashboard memungkinkan operator mengupload data export SAP dan Prodsys, lalu sistem menghitung selisih qty per order / material secara otomatis.

Akses internal: `http://10.204.10.32:8000`

---

## Tech Stack

| Komponen  | Teknologi                         |
|-----------|-----------------------------------|
| Backend   | Python 3 + Flask 3.x              |
| Database  | SQLite (`dashboard.db`)           |
| Frontend  | HTML / CSS / JavaScript (vanilla) |
| File      | Excel (.xlsx/.xls) atau CSV       |
| Export    | openpyxl (styled Excel)           |

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

## Alur Penggunaan

1. Buka dashboard → tab **Upload Data**
2. Upload 4 file export SAP/Prodsys (urutan bebas, data lama tidak dihapus)
3. Pindah ke tab **Perbandingan**
4. Pilih **Tanggal Mulai Produksi (GSTRP)** lalu klik **Tampilkan**
5. Gunakan filter status / search untuk menyaring data
6. Klik **Export Excel** untuk mengunduh hasil

---

## Upload Data

Data diupload per tabel. Upload bersifat **upsert**: jika data sudah ada (berdasarkan primary key SAP), baris baru diabaikan. Data lama tidak dihapus.

### AUFM — Goods Receipt SAP

Export dari SAP (tcode COOIS / MB51). Kolom wajib:

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

Kolom opsional (digunakan jika tersedia):

| Kolom | Keterangan                      |
|-------|---------------------------------|
| LGORT | Storage location (SHUV dikecualikan dari perhitungan) |

> Primary key: `MANDT + MBLNR + MJAHR + ZEILE`

---

### CAUFV — Production Order Header

Export dari SAP (table CAUFV). Kolom wajib:

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

Export dari SAP (table MAKT). Kolom wajib:

| Kolom | Keterangan           |
|-------|----------------------|
| MATNR | Nomor material       |
| MAKTX | Deskripsi material   |

> Primary key: `MATNR`

---

### Prodsys — Data ZPPCPFINT_GRAUTO

Export dari sistem Prodsys. Kolom wajib:

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
├── app.py                    # Backend Flask + SQLite
├── templates/
│   └── index.html            # Frontend (single page)
├── dashboard.db              # Database SQLite (auto-create)
├── requirements.txt          # Dependencies Python
├── run_program.bat           # Shortcut jalankan server (Windows)
├── Contoh AUFM.xlsx          # Contoh file export SAP AUFM
├── Contoh CAUFV.xlsx         # Contoh file export SAP CAUFV
├── Contoh MAKT.xlsx          # Contoh file export SAP MAKT
└── Contoh ZPPCPFINT_GRAUTO.xlsx  # Contoh file export Prodsys
```
