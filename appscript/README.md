# 🎯 APLIKASI FOMO — Find Out My Outstanding
### Aplikasi Skoring Bakat & Minat RIASEC Holland
**Irada · Google Apps Script + Google Sheets**

---

## 📋 DAFTAR ISI
1. [Tentang Aplikasi](#tentang)
2. [Struktur File](#struktur)
3. [Langkah Setup](#setup)
4. [Cara Deploy](#deploy)
5. [Cara Pakai Admin](#admin)
6. [Cara Pakai Siswa/Wali](#siswa)
7. [Format Import Soal Excel](#excel)
8. [Struktur Database Sheets](#database)
9. [FAQ & Troubleshooting](#faq)

---

## 🌟 TENTANG APLIKASI <a name="tentang"></a>

Aplikasi **FOMO** adalah sistem asesmen bakat & minat berbasis **RIASEC Holland** yang dirancang untuk siswa **SD dan TK**. Dibangun di atas Google Apps Script dengan Google Sheets sebagai database, aplikasi ini mampu digunakan oleh **ribuan pengguna secara bersamaan** dan menyimpan hingga **15.000+ data**.

### ✨ Fitur Utama
| Fitur | Keterangan |
|-------|-----------|
| 🔑 Login Token | Siswa/wali login pakai kode unik yang diberikan admin |
| 📝 Biodata | Input nama, usia, jenjang (SD/TK), sekolah, no HP |
| 📋 Instrumen Dinamis | Soal diambil langsung dari Sheets, bisa diubah kapan saja |
| 📊 Skoring Otomatis | 7 bidang: Olahraga, Seni, Matematika, Sains, Bahasa, Sosial, Religi |
| 🏆 Hasil Visual | Grafik, podium, deskripsi bidang, rekomendasi profesi |
| 👨‍💼 Panel Admin | Import soal Excel, generate token massal, laporan lengkap |
| 📥 Export Data | CSV & Excel untuk laporan |
| 📱 Responsive | Mobile-friendly, bisa dipakai di HP siswa |

### 🗺️ Peta Bidang RIASEC
| Kode | Bidang | Deskripsi Singkat |
|------|--------|-------------------|
| R | Olahraga | Aktivitas fisik, gerakan, ketangkasan |
| I | Sains | Ilmu pengetahuan, eksperimen, penelitian |
| A | Seni | Kreativitas, musik, gambar, tari |
| S | Sosial | Interaksi, membantu, kepemimpinan |
| E | Bahasa | Komunikasi, cerita, menulis, pidato |
| C | Matematika | Angka, logika, urutan, struktur |
| G | Religi | Agama, ibadah, hafalan, dakwah |

---

## 📁 STRUKTUR FILE <a name="struktur"></a>

```
📂 appscript/
├── Code.gs          → Backend utama (semua fungsi server-side)
├── CSS.html         → Stylesheet global tema biru
├── Login.html       → Halaman login (token siswa + admin)
├── Biodata.html     → Form data diri siswa
├── Instrumen.html   → Halaman pengerjaan soal
├── Hasil.html       → Halaman hasil & rekomendasi
├── Admin.html       → Panel admin lengkap
├── Laporan.html     → Laporan detail semua siswa
└── README.md        → Panduan ini
```

---

## ⚙️ LANGKAH SETUP <a name="setup"></a>

### LANGKAH 1 — Buat Google Spreadsheet Baru

1. Buka [sheets.google.com](https://sheets.google.com)
2. Klik **+ Blank spreadsheet** → beri nama: `FOMO Database`
3. Salin **ID Spreadsheet** dari URL:
   ```
   https://docs.google.com/spreadsheets/d/[INI_YANG_DISALIN]/edit
   ```
4. Simpan ID tersebut, akan dipakai di langkah berikutnya

### LANGKAH 2 — Buat Google Apps Script

1. Buka [script.google.com](https://script.google.com)
2. Klik **New Project**
3. Beri nama project: `FOMO App`

### LANGKAH 3 — Upload Semua File

Di Apps Script editor:

**A. File Code.gs**
- Hapus isi default `Code.gs`
- Copy-paste seluruh isi file `Code.gs`
- **Ganti baris ini** dengan ID Spreadsheet Anda:
  ```javascript
  var SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET_ANDA';
  //                    ↑ Ganti dengan ID dari Langkah 1
  ```

**B. File HTML**

Untuk setiap file HTML, klik **File → New → HTML file**, beri nama persis seperti berikut (tanpa .html):

| Nama File di Editor | Isi dari |
|--------------------|----------|
| `CSS` | `CSS.html` |
| `Login` | `Login.html` |
| `Biodata` | `Biodata.html` |
| `Instrumen` | `Instrumen.html` |
| `Hasil` | `Hasil.html` |
| `Admin` | `Admin.html` |
| `Laporan` | `Laporan.html` |

> ⚠️ **Penting:** Nama file harus **persis sama** (case-sensitive). Misalnya `Login` bukan `login`.

### LANGKAH 4 — Setup Awal Database

1. Di Apps Script editor, buka `Code.gs`
2. Klik tombol **Run** (▶) di toolbar
3. Pilih function: `setupAwal`
4. Klik **Run**
5. Izinkan akses Google Sheets saat diminta (klik **Review permissions → Allow**)

Ini akan membuat semua sheet otomatis:
- ✅ Sheet `Token`
- ✅ Sheet `Biodata`
- ✅ Sheet `Soal` (+ 35 soal contoh)
- ✅ Sheet `Jawaban`
- ✅ Sheet `Hasil`
- ✅ Sheet `Admin` (user: `admin`, password: `fomo2024`)

---

## 🚀 CARA DEPLOY <a name="deploy"></a>

### Deploy sebagai Web App

1. Di Apps Script editor, klik **Deploy → New Deployment**
2. Klik ikon ⚙️ di "Select type" → pilih **Web App**
3. Isi konfigurasi:
   ```
   Description    : FOMO App v1.0
   Execute as     : Me (akun Google Anda)
   Who has access : Anyone  ← PENTING! Agar siswa bisa akses
   ```
4. Klik **Deploy**
5. Salin **URL Web App** yang diberikan — ini adalah link aplikasi Anda!

### URL Format
```
https://script.google.com/macros/s/[ID_DEPLOYMENT]/exec
```

### Update Aplikasi (jika ada perubahan kode)
1. Klik **Deploy → Manage Deployments**
2. Pilih deployment yang ada → klik ✏️ Edit
3. Pilih "New version"
4. Klik **Deploy**

> ⚠️ **Setiap kali mengubah kode**, harus deploy ulang dengan "New version" agar perubahan aktif!

---

## 👨‍💼 CARA PAKAI ADMIN <a name="admin"></a>

### Login Admin
1. Buka URL Web App
2. Klik tab **🔧 Admin**
3. Username: `admin` | Password: `fomo2024`
4. *(Sangat disarankan ganti password setelah login pertama di sheet Admin)*

### Import Soal dari Excel
1. Login admin → klik menu **📥 Import Soal Excel**
2. Baca panduan format kolom
3. Klik **⬇️ Download Template Excel** untuk mendapat contoh format
4. Isi template dengan soal-soal Anda
5. Upload file Excel → Preview akan tampil
6. Centang "Hapus soal lama" jika ingin replace semua soal
7. Klik **✅ Import Soal**

### Generate Token untuk Siswa
1. Login admin → klik menu **🎫 Kelola Token**
2. Isi form:
   - **Jenjang**: SD atau TK
   - **Nama Sekolah**: nama sekolah siswa
   - **Jumlah Token**: berapa token yang ingin dibuat (1–500)
   - **Nama Wali**: opsional, nama kelas/wali
3. Klik **⚡ Generate Token**
4. Token akan tampil → klik **📋 Salin Semua** untuk menyalin
5. Bagikan token kepada siswa/wali melalui WhatsApp, kertas, dll.

### Lihat Laporan
1. Login admin → klik menu **📋 Laporan**
2. Atau buka langsung: `URL_APP?page=laporan`
3. Filter berdasarkan jenjang, bidang, atau cari nama
4. Klik baris siswa untuk melihat detail hasil
5. Export ke **CSV** atau **Excel** untuk arsip

---

## 👨‍👩‍👧 CARA PAKAI SISWA / WALI <a name="siswa"></a>

### Alur Pengerjaan

```
1. Buka link aplikasi
        ↓
2. Masukkan TOKEN yang diberikan sekolah
        ↓
3. Isi Biodata (nama, usia, jenjang, sekolah, no HP)
        ↓
4. Kerjakan soal instrumen (pilih A atau B)
        ↓
5. Klik "Selesai & Lihat Hasil"
        ↓
6. Lihat hasil + rekomendasi bidang bakat minat
        ↓
7. Cetak / simpan hasil sebagai PDF
```

### Tips untuk Siswa
- ✅ Jawab dengan **jujur** sesuai kebiasaan sehari-hari
- ✅ Tidak ada jawaban **benar atau salah**
- ✅ Bisa **kembali** ke soal sebelumnya
- ✅ Jawaban tersimpan otomatis, jika internet terputus bisa lanjut lagi
- ✅ Gunakan **keyboard** A/B untuk menjawab lebih cepat (di laptop)

---

## 📊 FORMAT IMPORT SOAL EXCEL <a name="excel"></a>

### Struktur Kolom (Sheet pertama, baris 1 = header)

| Kolom | Nama | Contoh | Keterangan |
|-------|------|--------|-----------|
| A | teks_soal | Saya suka bermain bola | Teks pernyataan soal |
| B | kode_riasec | R | Kode: R/I/A/S/E/C/G |
| C | jenjang | SEMUA | SD, TK, atau SEMUA |
| D | opsi_a | Ya, sangat suka | Pilihan A (biasanya "Ya") |
| E | opsi_b | Tidak terlalu | Pilihan B (biasanya "Tidak") |
| F | nilai_a | 1 | Skor jika memilih A |
| G | nilai_b | 0 | Skor jika memilih B |

### Contoh Soal yang Baik

| Bidang | Contoh Soal |
|--------|------------|
| Olahraga (R) | "Saya suka berlari dan bermain di luar rumah" |
| Sains (I) | "Saya suka bertanya kenapa sesuatu bisa terjadi" |
| Seni (A) | "Saya senang menggambar, mewarnai, atau membuat prakarya" |
| Sosial (S) | "Saya suka membantu teman yang kesusahan" |
| Bahasa (E) | "Saya senang bercerita di depan teman-teman" |
| Matematika (C) | "Saya suka berhitung dan mengerjakan soal angka" |
| Religi (G) | "Saya rajin mengaji dan belajar agama" |

### Tips Penulisan Soal
- ✅ Gunakan bahasa yang mudah dipahami anak SD/TK
- ✅ Satu soal = satu topik/tema
- ✅ Hindari kalimat negatif (jangan gunakan "tidak")
- ✅ Minimal **5 soal per kode RIASEC** untuk hasil yang akurat
- ✅ Rekomendasikan **28–42 soal total** (4–6 soal × 7 bidang)

---

## 🗄️ STRUKTUR DATABASE SHEETS <a name="database"></a>

### Sheet: Token
| Kolom | Keterangan |
|-------|-----------|
| Token | Kode unik (contoh: SD-AB3C2D) |
| Nama_Wali | Nama wali murid/kelas |
| No_HP_Wali | Nomor HP wali |
| Jenjang | SD atau TK |
| Asal_Sekolah | Nama sekolah |
| Status | aktif / dipakai / nonaktif |
| Tgl_Dibuat | Tanggal token dibuat |
| Tgl_Dipakai | Tanggal token pertama dipakai |

### Sheet: Biodata
| Kolom | Keterangan |
|-------|-----------|
| Token | Referensi ke Token |
| Nama_Siswa | Nama lengkap siswa |
| Usia | Usia dalam tahun |
| Jenjang | SD atau TK |
| Asal_Sekolah | Nama sekolah |
| No_HP | Nomor HP wali |
| Timestamp | Waktu pengisian |

### Sheet: Soal
| Kolom | Keterangan |
|-------|-----------|
| ID_Soal | Kode unik soal (S0001, dst) |
| Teks_Soal | Teks pernyataan soal |
| Kode_RIASEC | R/I/A/S/E/C/G |
| Jenjang | SD/TK/SEMUA |
| Opsi_A | Teks pilihan A |
| Opsi_B | Teks pilihan B |
| Nilai_A | Skor pilihan A |
| Nilai_B | Skor pilihan B |
| Aktif | aktif/tidak |

### Sheet: Jawaban
| Kolom | Keterangan |
|-------|-----------|
| Token | Referensi token |
| ID_Soal | Referensi soal |
| Jawaban | A atau B |
| Nilai | Skor yang diperoleh |
| Timestamp | Waktu submit |

### Sheet: Hasil
| Kolom | Keterangan |
|-------|-----------|
| Token | Referensi token |
| Nama_Siswa | Nama siswa |
| Jenjang | SD/TK |
| Asal_Sekolah | Nama sekolah |
| Skor_R...Skor_G | Skor per bidang |
| Bidang_Utama | Bidang dengan skor tertinggi |
| Bidang_2, Bidang_3 | Urutan ke-2 dan ke-3 |
| Timestamp | Waktu selesai tes |

---

## ❓ FAQ & TROUBLESHOOTING <a name="faq"></a>

### ❓ Apakah bisa digunakan bersamaan oleh banyak orang?
✅ Ya! Google Apps Script + Sheets mendukung akses bersamaan. Untuk skala besar (>500 orang bersamaan), pastikan:
- Gunakan akun Google Workspace (bukan personal) jika memungkinkan
- Quota harian Apps Script cukup untuk ribuan request

### ❓ Berapa kapasitas data Google Sheets?
✅ Google Sheets mendukung hingga **5 juta sel per spreadsheet**. Untuk 15.000 siswa dengan ~50 jawaban/siswa = 750.000 baris di sheet Jawaban — masih aman.

### ❓ Bagaimana jika siswa lupa atau kehilangan token?
✅ Admin bisa melihat daftar token di panel Admin → Kelola Token. Cari berdasarkan nama sekolah atau wali.

### ❓ Bisa kah soal diubah setelah beberapa siswa sudah mengerjakan?
✅ Bisa, tapi **hati-hati**: jika soal lama dihapus dan ada siswa yang sudah menjawab soal tersebut, skor mereka tetap tersimpan di sheet Hasil. Sebaiknya tandai soal sebagai "tidak" di kolom Aktif daripada menghapus.

### ❓ Kenapa hasil tidak muncul setelah submit?
Kemungkinan penyebab:
1. Koneksi internet terputus saat submit
2. Quota Apps Script harian habis (jarang terjadi)
3. Token sudah pernah disubmit → cek sheet Hasil

### ❓ Bagaimana cara ganti password admin?
Buka Google Sheets → sheet **Admin** → edit kolom B (Password) di baris admin.

### ❓ Bagaimana cara menambah bidang baru selain 7 bidang?
1. Di `Code.gs`, tambahkan kode baru di variabel `BIDANG_MAP`:
   ```javascript
   var BIDANG_MAP = {
     'R': 'Olahraga',
     // ... existing ...
     'X': 'Teknologi'  // ← tambahkan kode baru
   };
   ```
2. Di `Hasil.html`, tambahkan info bidang baru di variabel `BIDANG_INFO`
3. Deploy ulang aplikasi

### ❓ Aplikasi error "You do not have permission"
Saat deploy, pastikan setting **"Who has access"** diset ke **"Anyone"** (bukan "Anyone with Google Account").

### ❓ Bagaimana cara backup data?
1. Buka Google Spreadsheet
2. **File → Download → Microsoft Excel (.xlsx)**
3. Atau gunakan fitur Export CSV di panel Admin/Laporan

---

## 📞 INFORMASI KONTAK & KREDIT

```
Aplikasi    : FOMO — Find Out My Outstanding
Versi       : 1.0.0
Platform    : Google Apps Script + Google Sheets
Institusi   : Irada
Teknologi   : HTML5, CSS3, Vanilla JavaScript, SheetJS
```

### Struktur Pengembang
```
Admin Utama  → Akses penuh: import soal, kelola token, lihat laporan
Operator     → Akses terbatas: generate token, lihat laporan
Siswa/Wali   → Akses token: isi biodata, kerjakan soal, lihat hasil
```

---

## 📝 CATATAN PENTING

> 🔐 **Keamanan**: Ganti password admin default (`fomo2024`) segera setelah setup pertama!

> 📊 **Performa**: Untuk penggunaan massal (>1000 siswa/hari), pertimbangkan untuk memecah data ke beberapa Spreadsheet per periode/tahun ajaran.

> 💾 **Backup**: Lakukan backup data Spreadsheet secara berkala, minimal setiap akhir semester.

> 🔄 **Update Soal**: Setelah import soal baru, siswa yang **belum mengerjakan** akan mendapat soal terbaru. Siswa yang **sudah selesai** tidak terpengaruh.

---

*Dibuat dengan ❤️ untuk pendidikan Indonesia · FOMO App © 2026 Irada*
