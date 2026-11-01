# FOMO - Sistem Asesmen Holland Test (RIASEC)

Aplikasi Skoring dan Pelaporan Bakat & Minat berbasis Teori Karir RIASEC Holland. Dibangun menggunakan Google Apps Script (backend) dan Google Sheets (database cloud tanpa batas server fisik).

---

## Daftar Isi
1. [Fitur Utama](#fitur-utama)
2. [Peta Bidang RIASEC](#peta-bidang-riasec)
3. [Struktur File](#struktur-file)
4. [Petunjuk Setup & Deployment](#petunjuk-setup--deployment)
5. [Struktur Database Spreadsheet](#struktur-database-spreadsheet)
6. [Tautan Akses Aplikasi](#tautan-akses-aplikasi)

---

## Fitur Utama

- **Login & Otentikasi Token Unik**: Peserta masuk menggunakan token untuk mengisi biodata dan instrumen asesmen.
- **108 Butir Soal Terstandarisasi**: Soal standar Holland Test mencakup 6 dimensi kepribadian karir.
- **Perhitungan Skor & Ranking Otomatis**: Menghitung akumulasi skor Realistik, Investigatif, Artistik, Sosial, Enterprising, dan Konvensional (RIASEC).
- **Interpretasi & Rekomendasi Karir Komprehensif**: Menghasilkan 3 kode dominan (*Holland Code*), tingkat konsistensi, deskripsi profil, hingga rekomendasi profesi relevan.
- **Visualisasi Hasil Interaktif**: Menampilkan radar chart kepribadian, bar distribusi skor, dan podium 3 bidang minat teratas.
- **Cetak & Ekspor Laporan**: Fitur Cetak/Simpan PDF dan unduh laporan ke format Excel.
- **Panel Admin Terpadu**:
  - Dasbor statistik peserta dan distribusi bidang minat.
  - Pengelolaan token (generate massal, edit, aktivasi, reset tes, hapus permanen).
  - Pengelolaan bank soal (tambah, edit, aktif/nonaktifkan, hapus, impor soal Excel).
  - Laporan rekapitulasi data hasil tes seluruh peserta.

---

## Peta Bidang RIASEC

| Kode | Bidang Karir | Karakteristik & Minat |
| :---: | :--- | :--- |
| **R** | **Realistik** | Menyukai pekerjaan praktis, fisik, mekanik, alat kerja, dan aktivitas luar ruangan. |
| **I** | **Investigatif** | Menyukai pemikiran analitis, sains, penelitian, observasi, dan pemecahan masalah. |
| **A** | **Artistik** | Kreatif, ekspresif, intuitif, menyukai seni, desain, musik, dan lingkungan kerja fleksibel. |
| **S** | **Sosial** | Ramah, komunikatif, senang mengajar, melatih, menolong, dan bekerja sama dengan orang lain. |
| **E** | **Enterprising** | Memiliki jiwa kepemimpinan, persuasif, menyukai bisnis, negosiasi, dan pengambilan keputusan. |
| **C** | **Konvensional** | Terstruktur, teliti, prosedural, menyukai administrasi, pengarsipan data, dan perhitungan. |

---

## Struktur File

```text
fomo/
├── .gitignore                      # File ignore Git
├── README.md                       # Dokumentasi utama proyek
├── logo irada-1.png                # Aset logo Irada
├── LOGO FOMO-2026-S2.png           # Aset logo FOMO
├── Aplikasi Skoring RIASEC Holland.xlsx  # Lembar kerja acuan rumus & referensi
├── appscript/                      # Kode sumber Google Apps Script
│   ├── Code.gs                     # Backend server-side (Google Apps Script)
│   ├── CSS.html                    # Stylesheet dan variabel CSS terpusat
│   ├── Login.html                  # Halaman masuk siswa/wali murid
│   ├── LoginAdmin.html             # Halaman masuk administrator
│   ├── Biodata.html                # Form pengisian profil peserta
│   ├── Instrumen.html              # Antarmuka pengerjaan 108 soal asesmen
│   ├── Hasil.html                  # Halaman hasil & interpretasi siswa
│   ├── Admin.html                  # Panel dasbor, token, soal, & laporan admin
│   ├── Laporan.html                # Halaman rekapitulasi laporan lengkap
│   ├── LaporanSiswa.html           # Detail laporan individual per peserta
│   ├── referensi_riasec.json       # Master data deskripsi, profesi, & konsistensi
│   ├── soal_excel.json             # Master data 108 butir soal baku
│   └── README.md                   # Petunjuk teknis modul Apps Script
└── scratch/
    └── build_code_gs.py            # Skrip pembangun master Code.gs dari JSON
```

---

## Petunjuk Setup & Deployment

### 1. Persiapan Google Spreadsheet
1. Buka [Google Sheets](https://sheets.google.com) dan buat dokumen baru dengan nama **`FOMO Database`**.
2. Salin ID Spreadsheet dari tautan browser:
   ```text
   https://docs.google.com/spreadsheets/d/[SALIN_ID_DI_SINI]/edit
   ```

### 2. Pemasangan pada Google Apps Script
1. Buka [Google Apps Script](https://script.google.com) dan buat proyek baru dengan nama **`FOMO App`**.
2. Salin kode dari file `appscript/Code.gs` ke editor Apps Script.
3. Buat file HTML satu per satu sesuai nama yang ada pada folder `appscript/`:
   - `CSS`
   - `Login`
   - `LoginAdmin`
   - `Biodata`
   - `Instrumen`
   - `Hasil`
   - `Admin`
   - `Laporan`
   - `LaporanSiswa`
4. Tekan **Ctrl + S** untuk menyimpan seluruh file.

### 3. Inisialisasi Database
1. Pada menu toolbar Apps Script, pilih fungsi **`sinkronkanDatabase`** pada dropdown fungsi.
2. Klik tombol **Jalankan (Run)**.
3. Sistem secara otomatis akan menyusun baris header tabel yang baku di spreadsheet Anda.

### 4. Publikasi Web App
1. Klik tombol **Terapkan (Deploy)** di pojok kanan atas, lalu pilih **Penerapan baru (New deployment)**.
2. Pilih jenis: **Aplikasi Web (Web app)**.
3. Atur konfigurasi:
   - **Jalankan sebagai**: `Saya (akun Anda)`
   - **Yang memiliki akses**: `Siapa saja (Anyone)`
4. Klik **Terapkan (Deploy)**, lalu salin tautan URL yang diberikan.

---

## Struktur Database Spreadsheet

Sistem mengelola 6 lembar kerja (*sheet*) otomatis:

1. **`Token`**: Menyimpan token akses pengerjaan tes, status, dan waktu penggunaan.
   - Kolom: `Token`, `Nama_Wali`, `No_HP_Wali`, `Jenjang`, `Asal_Sekolah`, `Status`, `Tgl_Dibuat`, `Tgl_Dipakai`.
2. **`Biodata`**: Profil lengkap peserta tes.
   - Kolom: `Token`, `Nama_Siswa`, `Jenis_Kelamin`, `Tanggal_Lahir`, `Pendidikan`, `Asal_Sekolah`, `No_HP`, `Timestamp`.
3. **`Soal`**: Bank 108 soal asesmen RIASEC.
   - Kolom: `ID_Soal`, `Teks_Soal`, `Kode_RIASEC`, `Jenjang`, `Opsi_A`, `Opsi_B`, `Nilai_A`, `Nilai_B`, `Aktif`.
4. **`Jawaban`**: Riwayat respons butir soal per peserta.
   - Kolom: `Token`, `ID_Soal`, `Jawaban`, `Nilai`, `Timestamp`.
5. **`Hasil`**: Skor terhitung, ranking RIASEC, dan kode dominan peserta.
   - Kolom: `Token`, `Nama_Siswa`, `Jenis_Kelamin`, `Tanggal_Lahir`, `Pendidikan`, `Asal_Sekolah`, `Skor_R`, `Skor_I`, `Skor_A`, `Skor_S`, `Skor_E`, `Skor_C`, `Bidang_Utama`, `Bidang_2`, `Bidang_3`, `Kode_Holland`, `Konsistensi`, `Timestamp`.
6. **`Admin`**: Kredensial akun pengelola aplikasi.
   - Kolom: `Username`, `Password`, `Role`, `Tgl_Dibuat`.

---

## Tautan Akses Aplikasi

- **Tautan Siswa / Peserta Asesmen**:
  `https://script.google.com/macros/s/[DEPLOYMENT_ID]/exec?page=login`
- **Tautan Administrator & Guru**:
  `https://script.google.com/macros/s/[DEPLOYMENT_ID]/exec?page=admin-login`
