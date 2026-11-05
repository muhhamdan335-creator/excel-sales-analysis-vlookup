# Analisis Penjualan Perlengkapan Kantor Menggunakan Excel VLOOKUP

Proyek ini merupakan portofolio pengolahan data menggunakan **Microsoft Excel**
dengan fokus pada rumus **VLOOKUP**. Dataset yang digunakan mensimulasikan
transaksi penjualan perlengkapan kantor dari 12 pelanggan dan 15 produk.

Tujuan proyek ini adalah menggabungkan data transaksi penjualan dengan data
master pelanggan dan master produk, sehingga setiap transaksi memiliki
informasi lengkap untuk kebutuhan analisis penjualan.

## Gambaran Proyek

Pada kondisi awal, sheet **Data Penjualan** hanya berisi informasi dasar
seperti nomor order, tanggal, ID pelanggan, kode produk, dan jumlah pembelian.

Informasi seperti nama pelanggan, kota, tipe pelanggan, nama produk, kategori,
dan harga satuan belum tersedia pada data penjualan. Informasi tersebut diambil
secara otomatis dari sheet **Master Pelanggan** dan **Master Produk**
menggunakan VLOOKUP.

### Data Penjualan Sebelum VLOOKUP

![Data Penjualan Sebelum VLOOKUP](images/01-data-penjualan-sebelum-vlookup.png)

## Tujuan Proyek

- Menggabungkan data transaksi dengan master pelanggan menggunakan VLOOKUP.
- Menggabungkan data transaksi dengan master produk menggunakan VLOOKUP.
- Mengambil nama pelanggan, kota, dan tipe pelanggan secara otomatis.
- Mengambil nama produk, kategori, dan harga satuan secara otomatis.
- Menghitung total penjualan dari jumlah produk dan harga satuan.
- Menghasilkan dataset yang lebih lengkap, rapi, dan siap dianalisis.

## Tools dan Teknik

| Kategori | Tools / Teknik |
|---|---|
| Aplikasi | Microsoft Excel |
| Rumus utama | VLOOKUP |
| Rumus pendukung | Perkalian dasar Excel |
| Format data | XLSX dan CSV |
| Struktur data | Data transaksi dan data master |

## Struktur Dataset

Proyek ini menggunakan tiga tabel utama:

| File / Sheet | Keterangan |
|---|---|
| `data/data_penjualan.csv` | Berisi 40 transaksi penjualan, ID pelanggan, kode produk, dan jumlah pembelian |
| `data/master_pelanggan.csv` | Berisi ID pelanggan, nama pelanggan, kota, dan tipe pelanggan |
| `data/master_produk.csv` | Berisi kode produk, nama produk, kategori, dan harga satuan |
| `excel/dataset_portofolio_vlookup.xlsx` | File Excel utama berisi seluruh sheet dan hasil pengolahan VLOOKUP |

### Master Pelanggan

Sheet **Master Pelanggan** digunakan sebagai sumber data untuk mencari informasi
pelanggan berdasarkan **ID Pelanggan**.

![Master Pelanggan](images/02-master-pelanggan.png)

### Master Produk

Sheet **Master Produk** digunakan sebagai sumber data untuk mencari informasi
produk berdasarkan **Kode Produk**.

![Master Produk](images/03-master-produk.png)

## Rumus VLOOKUP yang Digunakan

VLOOKUP digunakan untuk mencari sebuah nilai pada kolom pertama tabel referensi,
kemudian mengembalikan nilai dari kolom lain pada baris yang sama.

Struktur rumus:

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, range_lookup)
```

| Bagian Rumus | Keterangan |
|---|---|
| `lookup_value` | Nilai yang dicari, misalnya ID Pelanggan atau Kode Produk |
| `table_array` | Range tabel sumber data |
| `col_index_num` | Nomor kolom hasil yang ingin diambil |
| `range_lookup` | `FALSE` untuk kecocokan persis |

Pada proyek ini, `FALSE` digunakan karena ID Pelanggan dan Kode Produk harus
cocok secara persis.

### 1. Mengambil Nama Pelanggan

```excel
=VLOOKUP(C2,Master_Pelanggan!$A:$D,2,FALSE)
```

Rumus ini mencari ID Pelanggan pada kolom C, kemudian mengambil Nama Pelanggan
dari kolom kedua pada sheet `Master_Pelanggan`.

![Rumus VLOOKUP Nama Pelanggan](images/04-rumus-vlookup-nama-pelanggan.png)

### 2. Mengambil Kota Pelanggan

```excel
=VLOOKUP(C2,Master_Pelanggan!$A:$D,3,FALSE)
```

Rumus ini mengambil Kota dari kolom ketiga pada sheet `Master_Pelanggan`.

![Rumus VLOOKUP Kota Pelanggan](images/05-rumus-vlookup-kota-pelanggan.png)

### 3. Mengambil Tipe Pelanggan

```excel
=VLOOKUP(C2,Master_Pelanggan!$A:$D,4,FALSE)
```

Rumus ini mengambil Tipe Pelanggan dari kolom keempat pada sheet
`Master_Pelanggan`.

![Rumus VLOOKUP Tipe Pelanggan](images/06-rumus-vlookup-tipe-pelanggan.png)

### 4. Mengambil Nama Produk

```excel
=VLOOKUP(D2,Master_Produk!$A:$D,2,FALSE)
```

Rumus ini mencari Kode Produk pada kolom D, kemudian mengambil Nama Produk
dari kolom kedua pada sheet `Master_Produk`.

![Rumus VLOOKUP Nama Produk](images/07-rumus-vlookup-nama-produk.png)

### 5. Mengambil Kategori Produk

```excel
=VLOOKUP(D2,Master_Produk!$A:$D,3,FALSE)
```

Rumus ini mengambil Kategori Produk dari kolom ketiga pada sheet
`Master_Produk`.

![Rumus VLOOKUP Kategori Produk](images/08-rumus-vlookup-kategori-produk.png)

### 6. Mengambil Harga Satuan

```excel
=VLOOKUP(D2,Master_Produk!$A:$D,4,FALSE)
```

Rumus ini mengambil Harga Satuan dari kolom keempat pada sheet
`Master_Produk`.

![Rumus VLOOKUP Harga Satuan](images/09-rumus-vlookup-harga-satuan.png)

### 7. Menghitung Total Penjualan

```excel
=E2*J2
```

`E2` berisi jumlah produk yang dibeli, sedangkan `J2` berisi harga satuan yang
telah diambil otomatis dari Master Produk menggunakan VLOOKUP.

![Rumus Total Penjualan](images/10-rumus-total-penjualan.png)

## Hasil Akhir

Setelah seluruh rumus VLOOKUP diterapkan, setiap baris transaksi memiliki
informasi lengkap, meliputi:

- Nama Pelanggan
- Kota Pelanggan
- Tipe Pelanggan
- Nama Produk
- Kategori Produk
- Harga Satuan
- Total Penjualan

![Hasil Akhir VLOOKUP](images/11-hasil-akhir-vlookup.png)

Data yang telah lengkap ini dapat digunakan untuk menganalisis penjualan
berdasarkan kota, tipe pelanggan, kategori produk, maupun produk tertentu.

## Ringkasan Analisis

Ringkasan analisis membantu menampilkan gambaran umum performa penjualan,
seperti total penjualan, jumlah transaksi, dan total unit terjual.

![Ringkasan Analisis](images/12-ringkasan-analisis.png)

## Cara Menggunakan File Excel

1. Unduh file `excel/dataset_portofolio_vlookup.xlsx`.
2. Buka sheet **Data Penjualan**.
3. Pastikan kolom ID Pelanggan dan Kode Produk sudah terisi.
4. Masukkan rumus VLOOKUP pada baris pertama data, yaitu baris 2.
5. Tarik rumus ke bawah menggunakan fill handle.
6. Gunakan rumus `=E2*J2` untuk menghitung Total Penjualan.
7. Periksa hasil VLOOKUP untuk memastikan tidak ada nilai `#N/A`.

## Pembelajaran yang Diperoleh

- Memahami cara kerja VLOOKUP untuk mencocokkan data antar tabel.
- Memahami perbedaan antara data transaksi dan data master.
- Menggunakan referensi absolut seperti `$A:$D` agar range tidak bergeser.
- Menggunakan `FALSE` untuk pencarian kecocokan persis.
- Menyusun dataset yang terstruktur dan siap digunakan untuk analisis.

## Dokumentasi Tambahan

Dokumentasi lebih rinci tersedia pada folder `documentation/`:

- [Panduan VLOOKUP](documentation/panduan_vlookup.md)
- [Struktur Dataset](documentation/struktur_dataset.md)
