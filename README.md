# excel-sales-performance-dashboard
Interactive Excel sales dashboard: data cleansing with Power Query, data modeling with Power Pivot, KPI formulas, Refresh and Fullscreen feature by Macros VBA and PivotChart/Slicer visuals on 5,130 orders (2019–2021).

# Excel Sales Performance Dashboard

Dashboard interaktif di Excel untuk memantau kinerja penjualan 5,130 order (Feb 2019 – Jan 2021): data cleansing dengan **Power Query**, pemodelan dengan **Power Pivot**, perhitungan KPI dengan **formula**, dan visualisasi dengan **PivotChart, Slicer, dan Timeline**.

![Dashboard Preview](images/dashboard_preview.png)

---

## 1. Overview

**Tujuan:** mengubah data penjualan mentah yang berantakan menjadi dashboard satu halaman yang membantu manajemen menjawab:

- Berapa total penjualan, volume, dan rata-rata nilai order?
- Region, negara bagian, produk, dan sales rep mana yang paling berkontribusi?
- Bagaimana perbandingan penjualan Retail vs Wholesale dan Business vs Individual?
- Bagaimana tren penjualan dari bulan ke bulan?

**Ringkasan hasil:**

| Total Sales | Quantity Sold | Total Order | Avg Order Value | Customer | Sales Rep |
|---|---|---|---|---|---|
| **$1.71M** | **16,523** | **5,130** | **$333.08** | 2,417 | 44 |

---

## 2. Dataset

| Item | Keterangan |
|---|---|
| File mentah | `Raw_Sales_Data.csv` (5,161 baris, 16 kolom, pemisah `;`) |
| File bersih | `Sales_Data.xlsx` (5,130 baris, 17 kolom) |
| Periode | 1 Feb 2019 – 31 Jan 2021 |
| Bisnis | Penjualan produk olive oil, gift basket, dan bath products ke pelanggan ritel dan bisnis di Amerika Serikat |

**Kolom utama:** OrderNum, Employee Name, Employee Job Title, Sales Region, OrderDate, OrderType (Retail/Wholesale), CustomerType (Individual/Business), CustID, CustName, CustState, ProdCategory, ProdNumber, ProdName, Quantity, Price, dan **Order Total** (kolom hasil perhitungan = Quantity × Price).

---

## 3. Tools

| Tool | Kegunaan |
|---|---|
| Power Query | Import CSV, membersihkan dan mentransformasi data |
| Power Pivot (Data Model) | Menghubungkan tabel `Sales` ke data model untuk PivotTable |
| Formula Excel | `GETPIVOTDATA`, `MAX`, `MIN`, `AVERAGE`, dan kolom hitung Order Total |
| PivotTable & PivotChart | Agregasi dan grafik penjualan |
| Slicer & Timeline | Filter interaktif (Region, Customer Type, Job Title, Product Category, Order Type, tanggal) |
| Macro VBA | Tombol Refresh Dashboard dan mode layar penuh (full-screen) |
| Excel `.xlsm` | Format workbook macro-enabled |
| Git & GitHub | Version control dan dokumentasi |

---

## 4. Langkah Pengerjaan

### a. Data Cleansing (Power Query)
Perbandingan data mentah dan data bersih menunjukkan transformasi berikut:

| Masalah di data mentah | Perbaikan |
|---|---|
| CSV berpemisah titik koma (`;`) dengan BOM | Import dengan delimiter `;` dan encoding UTF-8 |
| `OrderDate` bertipe teks | Konversi ke tipe Date |
| `Price` bertipe teks (ada nilai spasi/kosong) | Konversi ke tipe angka desimal |
| Nama karyawan tidak seragam (132 variasi, mis. `alexandra kundt`, `ALEXANDRA KUNDT`) | Trim dan Proper Case, menjadi **44 karyawan unik** |
| Spasi di awal nama pelanggan (`  Lanni D'Ambrogi`) | Trim |
| Singkatan region (`S Central East`, `N Central East`) | Distandarkan menjadi nama lengkap |
| 31 baris tidak lengkap (ProdNumber, Quantity, atau Price kosong) | Dihapus: **5,161 → 5,130 baris** |
| Tidak ada kolom nilai order | Menambah kolom `Order Total` = Quantity × Price |

### b. Data Model (Power Pivot)
Tabel `Sales` dimuat ke Data Model sehingga seluruh PivotTable memakai satu sumber data yang konsisten.

### c. Perhitungan KPI (Formula)
Sheet pivot memuat KPI Total Sales, Quantity Sold, Average Order Value, Total Order, Total Customer, dan Total Sales Employee. Nilainya diambil dengan `GETPIVOTDATA`, dan tanggal data terbaru dengan `MAX`.

### d. Dashboard
Sheet **Dashboard** berisi kartu KPI (Total Sales, Average Order Value, Total Order, Total Customer, Quantity Sold) dan panel berikut:

- Total Order by Months (tren bulanan)
- Regional Sales Distribution
- Top 5 Customer States
- Top 5 Highest Value Customers
- Top 5 Salespeople Performance
- Top 10 Selling Products
- Product Category dan komposisi Retail vs Wholesale (berdasarkan quantity: 67% vs 33%)

Slicer (Sales Region, Customer Type, Employee Job Title, Product Category, Order Type) dan Timeline tanggal membuat seluruh grafik bisa difilter secara interaktif.

### e. Otomasi dengan Macro VBA
Dua modul VBA ditambahkan agar dashboard terasa seperti aplikasi, dan keduanya dipasang pada ikon tombol di sheet Dashboard.

| Macro | Modul | Fungsi | Tombol |
|---|---|---|---|
| `Refresh_All` | `Refresh.bas` | Menyegarkan seluruh data, PivotTable, dan grafik, lalu menampilkan pesan "Dashboard Updated" | Ikon Refresh |
| `Swtchprg` | `FullScreen.bas` | Toggle tampilan: menyembunyikan atau memunculkan ribbon, formula bar, status bar, scrollbar, heading, dan tab sheet agar dashboard tampil layar penuh | Ikon Full Screen |
| `HideAll` / `ShowAll` | `FullScreen.bas` | Sub pendukung untuk masuk dan keluar dari mode layar penuh | (dipanggil oleh `Swtchprg`) |

Cell `Dashboard!B1` dipakai sebagai penanda status (1 = tampilan normal, 2 = layar penuh).

```vba
' Refresh.bas
Sub Refresh_All()
    ActiveWorkbook.RefreshAll
    MsgBox "Dashboard Updated"
End Sub

' FullScreen.bas (ringkasan toggle)
Sub Swtchprg()
    If Range("B1").Value = 2 Then
        Range("B1").Value = 1
        Call ShowAll      ' kembali ke tampilan normal
    Else
        Range("B1").Value = 2
        Call HideAll      ' masuk mode layar penuh
    End If
End Sub
```

Kode lengkap tersedia di folder [`vba/`](vba/). Macro hanya mengatur tampilan Excel dan refresh data; tidak mengakses file lain maupun jaringan.

---

## 5. Hasil & Insight

| Insight | Temuan |
|---|---|
| **Produk** | Olive Oil menyumbang **89.7%** penjualan ($1.53M). Bath products hanya 2.3%. Produk teratas: *Extra Virgin Olive Oil – Bold 12pk* ($166K). |
| **Tipe order** | Wholesale hanya 14.7% dari jumlah order (755), tetapi menghasilkan **57%** penjualan. Rata-rata order Wholesale $1,289 vs Retail $168. |
| **Pelanggan** | Pelanggan Business menyumbang 58.6% penjualan. |
| **Region** | South Central East terbesar ($340K, 19.9%), disusul North Central East ($270K) dan Central East ($252K). |
| **Negara bagian** | Texas ($235K, 13.7%), California ($195K), Florida ($141K). |
| **Sales rep** | Alyssa Smoak ($77.8K), Paul Massimo ($73.8K), Brenda Albury ($69.8K). |
| **Tren** | Penjualan 2020 ($926K) naik 25.5% dari 2019 ($738K). Pada periode sebanding (Feb–Des) kenaikannya 15.4%. Puncak bulanan: Okt 2020 ($93K). |

> Catatan: data 2019 mulai Februari dan data 2021 hanya Januari, sehingga perbandingan tahunan perlu dibaca dengan hati-hati.

**Rekomendasi bisnis:**
1. Perkuat segmen **Wholesale/Business** yang bernilai order tinggi (program kemitraan, harga volume).
2. Fokus pada **Olive Oil kemasan 12pk**, produk dengan kontribusi penjualan terbesar.
3. Pertahankan Texas dan California, dan dorong region dengan kontribusi kecil (Northwest, North Central West).
4. Tinjau portofolio **Bath products** yang kontribusinya rendah: bundling dengan gift basket atau evaluasi ulang.
5. Replikasi praktik sales rep teratas ke tim lain lewat coaching.

---

## 6. Struktur Repository

```
excel-sales-performance-dashboard/
├── data/
│   ├── raw/Raw_Sales_Data.csv
│   └── clean/Sales_Data.xlsx
├── dashboard/
│   └── EXCEL_Sales_Performance_Dashboard.xlsm
├── vba/
│   ├── Refresh.bas
│   └── FullScreen.bas
├── images/
│   └── dashboard_preview.png
└── README.md
```

---

## 7. How to Run

**Prasyarat:** Microsoft Excel 2016 atau lebih baru (Windows) dengan Power Query dan Power Pivot aktif. Power Pivot tidak tersedia di Excel untuk Mac.

1. **Clone** repository:
   ```bash
   git clone https://github.com/<username>/excel-sales-performance-dashboard.git
   ```
2. Buka `dashboard/EXCEL_Sales_Performance_Dashboard.xlsm`, lalu klik **Enable Content** jika muncul peringatan macro.
3. Buka sheet **Dashboard** dan gunakan **Slicer** dan **Timeline** untuk memfilter data.
4. Gunakan tombol macro di dashboard:
   - **Ikon Refresh**: menyegarkan semua data (setara **Data → Refresh All**).
   - **Ikon Full Screen**: masuk/keluar mode layar penuh. Jika tampilan Excel tidak kembali normal, jalankan macro `ShowAll` lewat **Alt + F8**.
5. Untuk mengulang proses cleansing: **Data → Get Data → From Text/CSV**, pilih `data/raw/Raw_Sales_Data.csv`, atur delimiter ke *Semicolon*, lalu ikuti langkah pada bagian 4a di editor Power Query.

---


> Dataset bersifat sampel dan digunakan untuk tujuan pembelajaran dan portofolio.
