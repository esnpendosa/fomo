import json
import os

with open('appscript/referensi_riasec.json', 'r', encoding='utf-8') as f:
    ref_data = json.load(f)

with open('appscript/soal_excel.json', 'r', encoding='utf-8') as f:
    soal_data = json.load(f)

ref_json_str = json.dumps(ref_data, ensure_ascii=False)
soal_json_str = json.dumps(soal_data, ensure_ascii=False)

code_template = '''// ============================================================
// APLIKASI FOMO - Skoring RIASEC Holland (Standar Excel Holland Test)
// Backend : Google Apps Script (Code.gs)
// Database: Google Sheets (dibuat OTOMATIS)
// ============================================================

// Nama sheet di Spreadsheet
var SHEET_TOKEN     = 'Token';
var SHEET_BIODATA   = 'Biodata';
var SHEET_SOAL      = 'Soal';
var SHEET_JAWABAN   = 'Jawaban';
var SHEET_HASIL     = 'Hasil';
var SHEET_ADMIN     = 'Admin';

// Nama file Spreadsheet yang akan dibuat otomatis
var SPREADSHEET_NAME = 'FOMO Database';

// ---- MAPPING BIDANG RIASEC Holland ----
var BIDANG_MAP = {
  'R': 'Realistik',    // Realistic    - Teknik, Mekanik, Praktis
  'I': 'Investigatif', // Investigative- Sains, Penelitian, Analitis
  'A': 'Artistik',     // Artistic     - Seni, Kreatif, Ekspresi
  'S': 'Sosial',       // Social       - Menolong, Mengajar, Relasi
  'E': 'Enterprising', // Enterprising - Kepemimpinan, Bisnis, Persuasi
  'C': 'Konvensional'  // Conventional - Administrasi, Prosedur, Data
};

var BIDANG_INGGRIS = {
  'R': 'Realistic',
  'I': 'Investigative',
  'A': 'Artistik',
  'S': 'Social',
  'E': 'Enterprising',
  'C': 'Conventional'
};

// ============================================================
// DATA REFERENSI HOLLAND (Sesuai Sheet Referensi Excel)
// ============================================================
var REFERENSI_RIASEC = ''' + ref_json_str + ''';

// ============================================================
// 108 SOAL STANDAR HOLLAND (Sesuai Sheet Soal Excel)
// ============================================================
var SOAL_STANDAR_HOLLAND = ''' + soal_json_str + ''';

// ============================================================
// ENTRY POINT - doGet
// ============================================================
function doGet(e) {
  e = e || { parameter: {} };
  var page = (e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : 'login';

  var pageMap = {
    'login'        : 'Login',
    'admin-login'  : 'LoginAdmin',
    'biodata'      : 'Biodata',
    'instrumen'    : 'Instrumen',
    'hasil'        : 'Hasil',
    'admin'        : 'Admin',
    'laporan'      : 'Laporan',
    'laporan-siswa': 'LaporanSiswa'
  };

  var viewName = pageMap[page] || 'Login';
  var template = HtmlService.createTemplateFromFile(viewName);
  template.page = page;

  return template.evaluate()
    .setTitle('FOMO - Report Holland Test (RIASEC)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// UTILITY: Include file HTML
// ============================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// UTILITY: Format Tanggal Aman (Anti-Crash)
// ============================================================
function safeFormatDate(val, formatStr) {
  if (!val) return '';
  formatStr = formatStr || 'dd/MM/yyyy HH:mm';
  try {
    var d;
    if (val instanceof Date) {
      d = val;
    } else {
      d = new Date(val);
    }
    if (isNaN(d.getTime())) {
      return String(val);
    }
    return Utilities.formatDate(d, 'Asia/Jakarta', formatStr);
  } catch(e) {
    return String(val || '');
  }
}

// ============================================================
// UTILITY: Ambil/Buat Spreadsheet OTOMATIS
// ============================================================
function getSpreadsheetId() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');

  if (id) {
    try {
      SpreadsheetApp.openById(id);
      return id;
    } catch(e) {
      props.deleteProperty('SPREADSHEET_ID');
    }
  }

  var ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  var newId = ss.getId();
  props.setProperty('SPREADSHEET_ID', newId);

  try {
    var file = DriveApp.getFileById(newId);
    var rootId = DriveApp.getRootFolder().getId();
    DriveApp.getFolderById(rootId).addFile(file);
  } catch(e) { }

  return newId;
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(getSpreadsheetId());
}

function getSheet(name) {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    initSheetHeader(sh, name);
  }
  return sh;
}

function getSpreadsheetUrl() {
  try {
    return { success: true, url: 'https://docs.google.com/spreadsheets/d/' + getSpreadsheetId() };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Map header kolom ke index (case-insensitive)
function getColumnMap(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var c = 0; c < headers.length; c++) {
    var key = headers[c] ? headers[c].toString().trim().toLowerCase().replace(/[^a-z0-9_]/g, '_') : '';
    if (key) map[key] = c;
  }
  return map;
}

// Inisialisasi header tiap sheet jika baru
function initSheetHeader(sh, name) {
  var headers = {
    'Token'   : ['Token', 'Nama_Wali', 'No_HP_Wali', 'Jenjang', 'Asal_Sekolah', 'Status', 'Tgl_Dibuat', 'Tgl_Dipakai'],
    'Biodata' : ['Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'No_HP', 'Timestamp'],
    'Soal'    : ['ID_Soal', 'Teks_Soal', 'Kode_RIASEC', 'Jenjang', 'Opsi_A', 'Opsi_B', 'Nilai_A', 'Nilai_B', 'Aktif'],
    'Jawaban' : ['Token', 'ID_Soal', 'Jawaban', 'Nilai', 'Timestamp'],
    'Hasil'   : ['Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'Skor_R', 'Skor_I', 'Skor_A', 'Skor_S', 'Skor_E', 'Skor_C', 'Bidang_Utama', 'Bidang_2', 'Bidang_3', 'Kode_Holland', 'Konsistensi', 'Timestamp'],
    'Admin'   : ['Username', 'Password', 'Role', 'Tgl_Dibuat']
  };

  if (headers[name] && sh.getLastRow() === 0) {
    sh.appendRow(headers[name]);
    sh.getRange(1, 1, 1, headers[name].length)
      .setBackground('#1565C0')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

// ============================================================
// FUNGSI INTERPRETASI HOLLAND (Persis Rumus Excel Hasil & Referensi)
// ============================================================
function getInterpretasiHolland(top3Kode) {
  top3Kode = (top3Kode || 'RIA').toString().trim().toUpperCase();
  var c1 = top3Kode.charAt(0) || 'R';
  var c2 = top3Kode.charAt(1) || 'I';
  var c3 = top3Kode.charAt(2) || 'A';
  var code2 = c1 + c2;
  var code3 = c1 + c2 + c3;

  var konsistensi = (REFERENSI_RIASEC.konsistensi && REFERENSI_RIASEC.konsistensi[code2]) ?
                    REFERENSI_RIASEC.konsistensi[code2] : 'Sedang';

  var combo2 = REFERENSI_RIASEC.combo2 && REFERENSI_RIASEC.combo2[code2];
  var single1 = (REFERENSI_RIASEC.single && REFERENSI_RIASEC.single[c1]) || { deskripsi: '', profesi: '' };

  var deskripsi = '';
  if (combo2 && combo2.deskripsi) {
    deskripsi = combo2.deskripsi;
  } else {
    deskripsi = single1.deskripsi || '';
  }

  var profesiTerkait = '';
  if (combo2 && combo2.profesi) {
    profesiTerkait = combo2.profesi;
  }
  if (single1.profesi) {
    profesiTerkait = (profesiTerkait ? profesiTerkait + ' ' : '') + single1.profesi;
  }

  var multi = REFERENSI_RIASEC.profesi_kombinasi || {};
  var profesi3 = multi[code3] || '';
  var profesi2 = multi[code2] || '';

  return {
    top3Kode: code3,
    code2: code2,
    c1: c1,
    c2: c2,
    c3: c3,
    nama1: BIDANG_MAP[c1] || c1,
    nama2: BIDANG_MAP[c2] || c2,
    nama3: BIDANG_MAP[c3] || c3,
    konsistensi: konsistensi,
    deskripsi: deskripsi,
    profesiTerkait: profesiTerkait,
    profesiKombinasi3: profesi3,
    profesiKombinasi2: profesi2
  };
}

// ============================================================
// SETUP AWAL
// ============================================================
function setupAwal() {
  var ss = getSpreadsheet();
  var sheets = [SHEET_TOKEN, SHEET_BIODATA, SHEET_SOAL, SHEET_JAWABAN, SHEET_HASIL, SHEET_ADMIN];

  sheets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      initSheetHeader(sh, name);
    }
  });

  // Admin default
  var adminSheet = getSheet(SHEET_ADMIN);
  if (adminSheet.getLastRow() <= 1) {
    adminSheet.appendRow(['admin', 'fomo2024', 'superadmin', new Date()]);
  }

  // Token contoh
  var tokenSheet = getSheet(SHEET_TOKEN);
  if (tokenSheet.getLastRow() <= 1) {
    tokenSheet.appendRow(['SD-SAMPLE', 'Admin Guru', '081234567890', 'SD', 'SDN Contoh 01', 'aktif', new Date(), '']);
    tokenSheet.appendRow(['TK-SAMPLE', 'Admin Guru', '081234567890', 'TK', 'TK Contoh 01', 'aktif', new Date(), '']);
    tokenSheet.appendRow(['UM-SAMPLE', 'Umum', '081234567890', 'SEMUA', 'Umum', 'aktif', new Date(), '']);
  }

  // Isi 108 soal standar Excel jika kosong
  var soalSheet = getSheet(SHEET_SOAL);
  if (soalSheet.getLastRow() <= 1) {
    insertSoalContoh(soalSheet);
  }

  return {
    success: true,
    message: 'Setup database berhasil dengan 108 soal standar Holland.',
    url: 'https://docs.google.com/spreadsheets/d/' + ss.getId()
  };
}

// ============================================================
// FUNGSI SINKRONISASI & PERBAIKAN DATABASE (Jalankan Sekali)
// ============================================================
function sinkronkanDatabase() {
  var ss = getSpreadsheet();
  
  // 1. Standarisasi Header Baku Semua Sheet
  var headers = {
    'Token'   : ['Token', 'Nama_Wali', 'No_HP_Wali', 'Jenjang', 'Asal_Sekolah', 'Status', 'Tgl_Dibuat', 'Tgl_Dipakai'],
    'Biodata' : ['Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'No_HP', 'Timestamp'],
    'Soal'    : ['ID_Soal', 'Teks_Soal', 'Kode_RIASEC', 'Jenjang', 'Opsi_A', 'Opsi_B', 'Nilai_A', 'Nilai_B', 'Aktif'],
    'Jawaban' : ['Token', 'ID_Soal', 'Jawaban', 'Nilai', 'Timestamp'],
    'Hasil'   : ['Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'Skor_R', 'Skor_I', 'Skor_A', 'Skor_S', 'Skor_E', 'Skor_C', 'Bidang_Utama', 'Bidang_2', 'Bidang_3', 'Kode_Holland', 'Konsistensi', 'Timestamp'],
    'Admin'   : ['Username', 'Password', 'Role', 'Tgl_Dibuat']
  };

  Object.keys(headers).forEach(function(shName) {
    var sh = ss.getSheetByName(shName);
    if (!sh) {
      sh = ss.insertSheet(shName);
    }
    var hList = headers[shName];
    sh.getRange(1, 1, 1, hList.length).setValues([hList]);
    sh.getRange(1, 1, 1, hList.length)
      .setBackground('#1565C0')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  });

  // 2. Rapikan sheet Hasil jika ada baris lama yang kolomnya belum lengkap
  var hasilSh = ss.getSheetByName('Hasil');
  if (hasilSh && hasilSh.getLastRow() > 1) {
    var lastRow = hasilSh.getLastRow();
    var lastCol = Math.max(hasilSh.getLastColumn(), 18);
    var data = hasilSh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var token = row[0] ? row[0].toString().trim().toUpperCase() : '';
      if (!token) continue;

      var kodeHolland = row[15];
      if (!kodeHolland || kodeHolland.toString().trim() === '') {
        var skor = {
          R: parseFloat(row[6]) || 0,
          I: parseFloat(row[7]) || 0,
          A: parseFloat(row[8]) || 0,
          S: parseFloat(row[9]) || 0,
          E: parseFloat(row[10]) || 0,
          C: parseFloat(row[11]) || 0
        };
        var ranked = ['R', 'I', 'A', 'S', 'E', 'C'].map(function(k) {
          return { kode: k, skor: skor[k], bidang: BIDANG_MAP[k] || k };
        }).sort(function(a, b) { return b.skor - a.skor; });

        var top3 = ranked.slice(0, 3).map(function(r) { return r.kode; }).join('');
        var interp = getInterpretasiHolland(top3);

        hasilSh.getRange(i + 2, 13).setValue(ranked[0].bidang);
        hasilSh.getRange(i + 2, 14).setValue(ranked[1] ? ranked[1].bidang : '');
        hasilSh.getRange(i + 2, 15).setValue(ranked[2] ? ranked[2].bidang : '');
        hasilSh.getRange(i + 2, 16).setValue(top3);
        hasilSh.getRange(i + 2, 17).setValue(interp.konsistensi);
      }
    }
  }

  return 'Database berhasil disinkronkan dan header diperbarui.';
}


function insertSoalContoh(sheet) {
  var rows = SOAL_STANDAR_HOLLAND.map(function(s) {
    var idStr = 'S' + ('000' + s.no).slice(-3);
    return [
      idStr,
      s.teks,
      s.kode,
      s.jenjang || 'SEMUA',
      s.opsiA || 'Ya',
      s.opsiB || 'Tidak',
      s.nilaiA !== undefined ? s.nilaiA : 1,
      s.nilaiB !== undefined ? s.nilaiB : 0,
      'aktif'
    ];
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

// ============================================================
// LOGIN TOKEN
// ============================================================
function loginToken(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var storedToken = row[0].toString().trim().toUpperCase();
      var status = (row[5] || '').toString().trim().toLowerCase();

      if (storedToken === token) {
        if (status === 'nonaktif') {
          return { success: false, message: 'Token sudah tidak aktif. Hubungi administrator.' };
        }

        // Ambil data hasil jika ada
        var hasilAda = cekHasilByToken(token);
        var bio = cekBiodata(token);

        if (hasilAda) {
          return {
            success    : true,
            status     : 'sudah_selesai',
            token      : token,
            jenjang    : row[3],
            sekolah    : row[4],
            namaPeserta: (bio.ada && bio.nama) ? bio.nama : (hasilAda.nama || ''),
            hasil      : hasilAda
          };
        }

        if (status === 'aktif') {
          sheet.getRange(i + 1, 6).setValue('dipakai');
          sheet.getRange(i + 1, 8).setValue(new Date());
        }

        return {
          success: true,
          status : bio.ada ? 'lanjut' : 'baru',
          token  : token,
          jenjang: row[3],
          sekolah: row[4],
          biodata: bio
        };
      }
    }
    return { success: false, message: 'Token "' + token + '" tidak ditemukan. Periksa kembali token Anda.' };
  } catch(err) {
    return { success: false, message: 'Terjadi kesalahan login: ' + err.message };
  }
}

// ============================================================
// RESET TOKEN UNTUK TES ULANG (Mulai dari 0)
// ============================================================
function resetToken(token) {
  try {
    token = token.toString().trim().toUpperCase();

    // 1. Update status di sheet Token
    var tokenSheet = getSheet(SHEET_TOKEN);
    var tData = tokenSheet.getDataRange().getValues();
    for (var i = 1; i < tData.length; i++) {
      if (tData[i][0].toString().trim().toUpperCase() === token) {
        tokenSheet.getRange(i + 1, 6).setValue('aktif');
        tokenSheet.getRange(i + 1, 8).setValue('');
        break;
      }
    }

    // 2. Hapus dari sheet Hasil
    var hasilSheet = getSheet(SHEET_HASIL);
    var hData = hasilSheet.getDataRange().getValues();
    for (var h = hData.length - 1; h >= 1; h--) {
      if (hData[h][0].toString().trim().toUpperCase() === token) {
        hasilSheet.deleteRow(h + 1);
      }
    }

    // 3. Hapus dari sheet Jawaban
    var jwbSheet = getSheet(SHEET_JAWABAN);
    var jData = jwbSheet.getDataRange().getValues();
    for (var j = jData.length - 1; j >= 1; j--) {
      if (jData[j][0].toString().trim().toUpperCase() === token) {
        jwbSheet.deleteRow(j + 1);
      }
    }

    // 4. Hapus dari sheet Biodata
    var bioSheet = getSheet(SHEET_BIODATA);
    var bData = bioSheet.getDataRange().getValues();
    for (var b = bData.length - 1; b >= 1; b--) {
      if (bData[b][0].toString().trim().toUpperCase() === token) {
        bioSheet.deleteRow(b + 1);
      }
    }

    return { success: true, message: 'Token ' + token + ' berhasil di-reset. Anda dapat mulai mengerjakan dari awal.' };
  } catch(err) {
    return { success: false, message: 'Gagal mereset token: ' + err.message };
  }
}

// ============================================================
// LOGIN ADMIN
// ============================================================
function loginAdmin(username, password) {
  try {
    var sheet = getSheet(SHEET_ADMIN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(username).trim() && String(data[i][1]).trim() === String(password).trim()) {
        return { success: true, role: data[i][2], username: username };
      }
    }
    return { success: false, message: 'Username atau password salah.' };
  } catch(err) {
    return { success: false, message: 'Terjadi kesalahan: ' + err.message };
  }
}

// ============================================================
// BIODATA SISWA
// ============================================================
function cekBiodata(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_BIODATA);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { ada: false };

    var col = getColumnMap(sheet);
    var cToken = col['token'] !== undefined ? col['token'] : 0;
    var cNama  = col['nama_siswa'] !== undefined ? col['nama_siswa'] : 1;
    var cJk    = col['jenis_kelamin'] !== undefined ? col['jenis_kelamin'] : -1;
    var cTgl   = col['tanggal_lahir'] !== undefined ? col['tanggal_lahir'] : -1;
    var cPend  = col['pendidikan'] !== undefined ? col['pendidikan'] : (col['jenjang'] !== undefined ? col['jenjang'] : 3);
    var cSek   = col['asal_sekolah'] !== undefined ? col['asal_sekolah'] : 4;
    var cHp    = col['no_hp'] !== undefined ? col['no_hp'] : 5;

    for (var i = 1; i < data.length; i++) {
      if (data[i][cToken] && data[i][cToken].toString().trim().toUpperCase() === token) {
        var row = data[i];
        return {
          ada         : true,
          token       : token,
          nama        : row[cNama] || '',
          jenisKelamin: cJk >= 0 ? (row[cJk] || 'Laki-laki') : (row[2] || 'Laki-laki'),
          tanggalLahir: cTgl >= 0 ? safeFormatDate(row[cTgl], 'yyyy-MM-dd') : '',
          tanggalLahirText: cTgl >= 0 ? safeFormatDate(row[cTgl], 'dd MMM yyyy') : '',
          pendidikan  : row[cPend] || 'SD',
          jenjang     : row[cPend] || 'SD',
          sekolah     : cSek >= 0 ? (row[cSek] || '') : '',
          noHp        : cHp >= 0 ? (row[cHp] || '') : ''
        };
      }
    }
    return { ada: false };
  } catch(err) {
    return { ada: false };
  }
}

function simpanBiodata(token, nama, jenisKelamin, tanggalLahir, pendidikan, sekolah, noHp) {
  try {
    token = token.toString().trim().toUpperCase();
    nama = String(nama || '').trim();
    jenisKelamin = String(jenisKelamin || 'Laki-laki').trim();
    tanggalLahir = tanggalLahir ? new Date(tanggalLahir) : '';
    pendidikan = String(pendidikan || 'SD').trim();
    sekolah = String(sekolah || '').trim();
    noHp = String(noHp || '').trim();

    var sheet = getSheet(SHEET_BIODATA);
    var data = sheet.getDataRange().getValues();
    var col = getColumnMap(sheet);
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        rowIndex = i + 1;
        break;
      }
    }

    // Jika header belum punya jenis kelamin, sesuaikan header
    if (col['jenis_kelamin'] === undefined && sheet.getLastRow() >= 1) {
      sheet.getRange(1, 1, 1, 8).setValues([[
        'Token', 'Nama_Siswa', 'Jenis_Kelamin', 'Tanggal_Lahir', 'Pendidikan', 'Asal_Sekolah', 'No_HP', 'Timestamp'
      ]]).setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold');
    }

    var rowValues = [token, nama, jenisKelamin, tanggalLahir, pendidikan, sekolah, noHp, new Date()];

    if (rowIndex > 1) {
      // Update data jika sudah ada
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
      return { success: true, message: 'Biodata berhasil diperbarui.', biodata: cekBiodata(token) };
    } else {
      sheet.appendRow(rowValues);
      return { success: true, message: 'Biodata berhasil disimpan.', biodata: cekBiodata(token) };
    }
  } catch(err) {
    return { success: false, message: 'Gagal menyimpan biodata: ' + err.message };
  }
}

// ============================================================
// SOAL
// ============================================================
function getSoal(jenjang) {
  try {
    var sheet = getSheet(SHEET_SOAL);
    var data = sheet.getDataRange().getValues();
    var soalList = [];

    // Jika soal masih kosong di sheet, isi otomatis dari 108 soal standar
    if (data.length <= 1) {
      insertSoalContoh(sheet);
      data = sheet.getDataRange().getValues();
    }

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var soalJenjang = (row[3] || '').toString().trim().toUpperCase();
      var aktif = (row[8] || '').toString().trim().toLowerCase();

      if (aktif !== 'tidak' && aktif !== 'false' && aktif !== '0') {
        if (!jenjang || jenjang === 'SEMUA' || soalJenjang === 'SEMUA' || soalJenjang === jenjang.toUpperCase()) {
          soalList.push({
            id       : row[0].toString(),
            no       : i,
            teks     : (row[1] || '').toString(),
            kode     : (row[2] || 'R').toString().trim().toUpperCase(),
            jenjang  : (row[3] || 'SEMUA').toString(),
            opsiA    : (row[4] || 'Ya').toString(),
            opsiB    : (row[5] || 'Tidak').toString(),
            nilaiA   : parseFloat(row[6]) !== undefined && !isNaN(parseFloat(row[6])) ? parseFloat(row[6]) : 1,
            nilaiB   : parseFloat(row[7]) !== undefined && !isNaN(parseFloat(row[7])) ? parseFloat(row[7]) : 0
          });
        }
      }
    }

    return { success: true, soal: soalList, total: soalList.length };
  } catch(err) {
    return { success: false, message: 'Gagal memuat soal: ' + err.message, soal: [] };
  }
}

// ============================================================
// SIMPAN JAWABAN & HITUNG SKOR RIASEC
// ============================================================
function simpanJawabanDanHitung(token, jawabanList) {
  try {
    token = token.toString().trim().toUpperCase();
    if (!jawabanList || !jawabanList.length) {
      return { success: false, message: 'Daftar jawaban kosong.' };
    }

    var ts = new Date();

    // 1. Simpan jawaban ke sheet Jawaban (bersihkan jawaban lama untuk token ini jika ada)
    var jawabanSheet = getSheet(SHEET_JAWABAN);
    var jData = jawabanSheet.getDataRange().getValues();
    for (var j = jData.length - 1; j >= 1; j--) {
      if (jData[j][0] && jData[j][0].toString().trim().toUpperCase() === token) {
        jawabanSheet.deleteRow(j + 1);
      }
    }

    var rowsJawaban = jawabanList.map(function(item) {
      return [token, item.idSoal, item.jawaban || '', item.nilai !== undefined ? item.nilai : 0, ts];
    });
    if (rowsJawaban.length) {
      jawabanSheet.getRange(jawabanSheet.getLastRow() + 1, 1, rowsJawaban.length, rowsJawaban[0].length).setValues(rowsJawaban);
    }

    // 2. Hitung skor per kode RIASEC
    var skor = { R: 0, I: 0, A: 0, S: 0, E: 0, C: 0 };

    var soalSheet = getSheet(SHEET_SOAL);
    var soalData = soalSheet.getDataRange().getValues();
    var soalMap = {};
    for (var s = 1; s < soalData.length; s++) {
      soalMap[soalData[s][0].toString()] = (soalData[s][2] || 'R').toString().trim().toUpperCase();
    }

    jawabanList.forEach(function(item) {
      var kode = soalMap[item.idSoal.toString()];
      if (kode && skor.hasOwnProperty(kode)) {
        skor[kode] += parseFloat(item.nilai) || 0;
      }
    });

    // 3. Urutkan ranking RIASEC
    var urutan = ['R', 'I', 'A', 'S', 'E', 'C'];
    var ranked = urutan.map(function(k) {
      return { kode: k, skor: skor[k], bidang: BIDANG_MAP[k] || k, namaInggris: BIDANG_INGGRIS[k] || k };
    }).sort(function(a, b) {
      return b.skor - a.skor;
    });

    var top3Kode = ranked.slice(0, 3).map(function(r) { return r.kode; }).join('');
    var interpretasi = getInterpretasiHolland(top3Kode);

    // 4. Ambil biodata terkini
    var biodata = cekBiodata(token);

    // 5. Simpan / Perbarui ke sheet Hasil
    var hasilSheet = getSheet(SHEET_HASIL);
    var hData = hasilSheet.getDataRange().getValues();
    var hasilRowIndex = -1;

    for (var h = 1; h < hData.length; h++) {
      if (hData[h][0] && hData[h][0].toString().trim().toUpperCase() === token) {
        hasilRowIndex = h + 1;
        break;
      }
    }

    var hasilValues = [
      token,
      biodata.nama || '',
      biodata.jenisKelamin || 'Laki-laki',
      biodata.tanggalLahirText || biodata.tanggalLahir || '',
      biodata.pendidikan || 'SD',
      biodata.sekolah || '',
      skor['R'], skor['I'], skor['A'], skor['S'], skor['E'], skor['C'],
      ranked[0].bidang,
      ranked[1] ? ranked[1].bidang : '',
      ranked[2] ? ranked[2].bidang : '',
      top3Kode,
      interpretasi.konsistensi,
      ts
    ];

    if (hasilRowIndex > 1) {
      hasilSheet.getRange(hasilRowIndex, 1, 1, hasilValues.length).setValues([hasilValues]);
    } else {
      hasilSheet.appendRow(hasilValues);
    }

    return {
      success      : true,
      token        : token,
      skor         : skor,
      ranked       : ranked,
      top3Kode     : top3Kode,
      biodata      : biodata,
      interpretasi : interpretasi,
      bidangUtama  : ranked[0].bidang,
      bidang2      : ranked[1] ? ranked[1].bidang : '',
      bidang3      : ranked[2] ? ranked[2].bidang : ''
    };
  } catch(err) {
    return { success: false, message: 'Gagal memproses jawaban: ' + err.message };
  }
}

// ============================================================
// CEK HASIL BY TOKEN
// ============================================================
function cekHasilByToken(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_HASIL);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;

    var col = getColumnMap(sheet);
    var cToken = col['token'] !== undefined ? col['token'] : 0;
    var cNama  = col['nama_siswa'] !== undefined ? col['nama_siswa'] : 1;
    var cJk    = col['jenis_kelamin'] !== undefined ? col['jenis_kelamin'] : -1;
    var cTgl   = col['tanggal_lahir'] !== undefined ? col['tanggal_lahir'] : -1;
    var cPend  = col['pendidikan'] !== undefined ? col['pendidikan'] : (col['jenjang'] !== undefined ? col['jenjang'] : 4);
    var cSek   = col['asal_sekolah'] !== undefined ? col['asal_sekolah'] : 5;
    var cR     = col['skor_r'] !== undefined ? col['skor_r'] : 6;
    var cI     = col['skor_i'] !== undefined ? col['skor_i'] : 7;
    var cA     = col['skor_a'] !== undefined ? col['skor_a'] : 8;
    var cS     = col['skor_s'] !== undefined ? col['skor_s'] : 9;
    var cE     = col['skor_e'] !== undefined ? col['skor_e'] : 10;
    var cC     = col['skor_c'] !== undefined ? col['skor_c'] : 11;
    var cUtama = col['bidang_utama'] !== undefined ? col['bidang_utama'] : 12;
    var cB2    = col['bidang_2'] !== undefined ? col['bidang_2'] : 13;
    var cB3    = col['bidang_3'] !== undefined ? col['bidang_3'] : 14;
    var cKod   = col['kode_holland'] !== undefined ? col['kode_holland'] : 15;
    var cKon   = col['konsistensi'] !== undefined ? col['konsistensi'] : 16;
    var cTs    = col['timestamp'] !== undefined ? col['timestamp'] : 17;

    for (var i = 1; i < data.length; i++) {
      if (data[i][cToken] && data[i][cToken].toString().trim().toUpperCase() === token) {
        var row = data[i];
        var skor = {
          R: parseFloat(row[cR]) || 0,
          I: parseFloat(row[cI]) || 0,
          A: parseFloat(row[cA]) || 0,
          S: parseFloat(row[cS]) || 0,
          E: parseFloat(row[cE]) || 0,
          C: parseFloat(row[cC]) || 0
        };

        var ranked = ['R', 'I', 'A', 'S', 'E', 'C'].map(function(k) {
          return { kode: k, skor: skor[k], bidang: BIDANG_MAP[k] || k, namaInggris: BIDANG_INGGRIS[k] || k };
        }).sort(function(a, b) { return b.skor - a.skor; });

        var top3Kode = (cKod >= 0 && row[cKod]) ? row[cKod].toString().trim() :
                       ranked.slice(0, 3).map(function(r) { return r.kode; }).join('');

        var interpretasi = getInterpretasiHolland(top3Kode);
        var bio = cekBiodata(token);

        return {
          token       : token,
          nama        : (bio.ada && bio.nama) ? bio.nama : (row[cNama] || token),
          jenisKelamin: (bio.ada && bio.jenisKelamin) ? bio.jenisKelamin : (cJk >= 0 ? (row[cJk] || 'Laki-laki') : 'Laki-laki'),
          tanggalLahir: (bio.ada && (bio.tanggalLahirText || bio.tanggalLahir)) ? (bio.tanggalLahirText || bio.tanggalLahir) : (cTgl >= 0 ? safeFormatDate(row[cTgl], 'dd MMM yyyy') : ''),
          pendidikan  : (bio.ada && bio.pendidikan) ? bio.pendidikan : (row[cPend] || 'SD'),
          jenjang     : (bio.ada && bio.pendidikan) ? bio.pendidikan : (row[cPend] || 'SD'),
          sekolah     : (bio.ada && bio.sekolah) ? bio.sekolah : (cSek >= 0 ? (row[cSek] || '') : ''),
          skor        : skor,
          ranked      : ranked,
          top3Kode    : top3Kode,
          interpretasi: interpretasi,
          bidangUtama : row[cUtama] || ranked[0].bidang,
          bidang2     : row[cB2] || (ranked[1] ? ranked[1].bidang : ''),
          bidang3     : row[cB3] || (ranked[2] ? ranked[2].bidang : ''),
          timestamp   : safeFormatDate(row[cTs], 'dd/MM/yyyy HH:mm')
        };
      }
    }
    return null;
  } catch(err) {
    return null;
  }
}

// ============================================================
// ADMIN: AMBIL SEMUA HASIL (LAPORAN) - ANTI CRASH
// ============================================================
function getAllHasil(filterJenjang, filterSekolah) {
  try {
    var sheet = getSheet(SHEET_HASIL);
    var data = sheet.getDataRange().getValues();
    var result = [];

    if (data.length <= 1) {
      return { success: true, data: [], total: 0 };
    }

    var col = getColumnMap(sheet);
    var cToken = col['token'] !== undefined ? col['token'] : 0;
    var cNama  = col['nama_siswa'] !== undefined ? col['nama_siswa'] : 1;
    var cJk    = col['jenis_kelamin'] !== undefined ? col['jenis_kelamin'] : -1;
    var cTgl   = col['tanggal_lahir'] !== undefined ? col['tanggal_lahir'] : -1;
    var cPend  = col['pendidikan'] !== undefined ? col['pendidikan'] : (col['jenjang'] !== undefined ? col['jenjang'] : 4);
    var cSek   = col['asal_sekolah'] !== undefined ? col['asal_sekolah'] : 5;
    var cR     = col['skor_r'] !== undefined ? col['skor_r'] : 6;
    var cI     = col['skor_i'] !== undefined ? col['skor_i'] : 7;
    var cA     = col['skor_a'] !== undefined ? col['skor_a'] : 8;
    var cS     = col['skor_s'] !== undefined ? col['skor_s'] : 9;
    var cE     = col['skor_e'] !== undefined ? col['skor_e'] : 10;
    var cC     = col['skor_c'] !== undefined ? col['skor_c'] : 11;
    var cUtama = col['bidang_utama'] !== undefined ? col['bidang_utama'] : 12;
    var cB2    = col['bidang_2'] !== undefined ? col['bidang_2'] : 13;
    var cB3    = col['bidang_3'] !== undefined ? col['bidang_3'] : 14;
    var cKod   = col['kode_holland'] !== undefined ? col['kode_holland'] : 15;
    var cKon   = col['konsistensi'] !== undefined ? col['konsistensi'] : 16;
    var cTs    = col['timestamp'] !== undefined ? col['timestamp'] : 17;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[cToken]) continue;

      var jenjang = (row[cPend] || '').toString().trim().toUpperCase();
      var sekolah = (cSek >= 0 && row[cSek]) ? row[cSek].toString().trim() : '';

      if (filterJenjang && filterJenjang !== 'semua' && jenjang !== filterJenjang.toUpperCase()) continue;
      if (filterSekolah && filterSekolah !== 'semua' && sekolah.toLowerCase().indexOf(filterSekolah.toLowerCase()) === -1) continue;

      var skor = {
        R: parseFloat(row[cR]) || 0,
        I: parseFloat(row[cI]) || 0,
        A: parseFloat(row[cA]) || 0,
        S: parseFloat(row[cS]) || 0,
        E: parseFloat(row[cE]) || 0,
        C: parseFloat(row[cC]) || 0
      };

      var ranked = ['R', 'I', 'A', 'S', 'E', 'C'].map(function(k) {
        return { kode: k, skor: skor[k], bidang: BIDANG_MAP[k] || k };
      }).sort(function(a, b) { return b.skor - a.skor; });

      var top3Kode = (cKod >= 0 && row[cKod]) ? row[cKod].toString().trim() :
                     ranked.slice(0, 3).map(function(r) { return r.kode; }).join('');

      var konsistensi = (cKon >= 0 && row[cKon]) ? row[cKon].toString().trim() :
                        ((REFERENSI_RIASEC.konsistensi && REFERENSI_RIASEC.konsistensi[top3Kode.slice(0, 2)]) || 'Sedang');

      result.push({
        no          : result.length + 1,
        token       : row[cToken],
        nama        : row[cNama] || row[cToken],
        jenisKelamin: cJk >= 0 ? (row[cJk] || 'Laki-laki') : 'Laki-laki',
        tanggalLahir: cTgl >= 0 ? safeFormatDate(row[cTgl], 'dd MMM yyyy') : '',
        jenjang     : row[cPend] || 'SD',
        pendidikan  : row[cPend] || 'SD',
        sekolah     : sekolah,
        skor        : skor,
        ranked      : ranked,
        top3Kode    : top3Kode,
        konsistensi : konsistensi,
        bidangUtama : row[cUtama] || ranked[0].bidang,
        bidang2     : row[cB2] || (ranked[1] ? ranked[1].bidang : ''),
        bidang3     : row[cB3] || (ranked[2] ? ranked[2].bidang : ''),
        timestamp   : safeFormatDate(row[cTs], 'dd/MM/yyyy HH:mm')
      });
    }

    return { success: true, data: result, total: result.length };
  } catch(err) {
    return { success: false, message: 'Gagal mengambil data laporan: ' + err.message, data: [] };
  }
}

// ============================================================
// LAPORAN LENGKAP PER SISWA
// ============================================================
function getLaporanSiswa(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var hasil = cekHasilByToken(token);
    if (!hasil) return { success: false, message: 'Data hasil tidak ditemukan untuk token ' + token };

    var biodata = cekBiodata(token);
    var MAKS_PER_DIMENSI = 18;
    var skor = hasil.skor || {};
    var persen = {};
    var urutan = ['R', 'I', 'A', 'S', 'E', 'C'];
    urutan.forEach(function(k) {
      var s = parseFloat(skor[k]) || 0;
      persen[k] = Math.round((s / MAKS_PER_DIMENSI) * 100);
    });

    // Detail jawaban per soal
    var jawabanSheet = getSheet(SHEET_JAWABAN);
    var jData = jawabanSheet.getDataRange().getValues();
    var detailJawaban = [];
    for (var j = 1; j < jData.length; j++) {
      if (jData[j][0] && jData[j][0].toString().trim().toUpperCase() === token) {
        detailJawaban.push({
          idSoal : jData[j][1],
          jawaban: jData[j][2],
          nilai  : jData[j][3]
        });
      }
    }

    // Map teks soal
    var soalSheet = getSheet(SHEET_SOAL);
    var sData = soalSheet.getDataRange().getValues();
    var soalMap = {};
    for (var s = 1; s < sData.length; s++) {
      soalMap[sData[s][0].toString()] = {
        teks  : sData[s][1],
        kode  : sData[s][2],
        bidang: BIDANG_MAP[sData[s][2]] || sData[s][2]
      };
    }

    var jawabanLengkap = detailJawaban.map(function(item) {
      var info = soalMap[item.idSoal.toString()] || {};
      return {
        idSoal : item.idSoal,
        teks   : info.teks || '',
        kode   : info.kode || '',
        bidang : info.bidang || '',
        jawaban: item.jawaban,
        nilai  : item.nilai
      };
    });

    var rekapDimensi = {};
    urutan.forEach(function(k) { rekapDimensi[k] = { ya: 0, tidak: 0, total: 0 }; });
    jawabanLengkap.forEach(function(item) {
      var k = (item.kode || '').toUpperCase();
      if (rekapDimensi[k]) {
        rekapDimensi[k].total++;
        if (parseFloat(item.nilai) > 0) rekapDimensi[k].ya++;
        else rekapDimensi[k].tidak++;
      }
    });

    return {
      success       : true,
      token         : token,
      biodata       : biodata.ada ? biodata : hasil,
      skor          : skor,
      persen        : persen,
      ranked        : hasil.ranked,
      top3Kode      : hasil.top3Kode,
      interpretasi  : hasil.interpretasi,
      bidangUtama   : hasil.bidangUtama,
      bidang2       : hasil.bidang2,
      bidang3       : hasil.bidang3,
      rekapDimensi  : rekapDimensi,
      jawabanLengkap: jawabanLengkap,
      timestamp     : hasil.timestamp
    };
  } catch(err) {
    return { success: false, message: 'Gagal mengambil laporan siswa: ' + err.message };
  }
}

// ============================================================
// ADMIN: STATISTIK DASHBOARD
// ============================================================
function getStatistik() {
  try {
    var tokenSheet = getSheet(SHEET_TOKEN);
    var hasilSheet = getSheet(SHEET_HASIL);
    var soalSheet  = getSheet(SHEET_SOAL);

    var tData = tokenSheet.getDataRange().getValues();
    var hData = hasilSheet.getDataRange().getValues();
    var sData = soalSheet.getDataRange().getValues();

    var totalToken = Math.max(tData.length - 1, 0);
    var tokenAktif = 0, tokenDipakai = 0;
    for (var i = 1; i < tData.length; i++) {
      var st = (tData[i][5] || '').toString().toLowerCase();
      if (st === 'aktif') tokenAktif++;
      else if (st === 'dipakai') tokenDipakai++;
    }

    var totalHasil = Math.max(hData.length - 1, 0);
    var totalSoal  = Math.max(sData.length - 1, 0);

    return {
      success     : true,
      totalToken  : totalToken,
      tokenAktif  : tokenAktif,
      tokenDipakai: tokenDipakai,
      totalHasil  : totalHasil,
      totalSoal   : totalSoal
    };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// ADMIN: TOKEN MANAGEMENT
// ============================================================
function generateToken(jumlah, jenjang, sekolah, namaWali) {
  try {
    var sheet = getSheet(SHEET_TOKEN);
    var tokens = [];
    var prefix = jenjang.toUpperCase() === 'SD' ? 'SD' : jenjang.toUpperCase() === 'TK' ? 'TK' : 'UM';

    var rows = [];
    for (var i = 0; i < jumlah; i++) {
      var token = prefix + '-' + generateRandomString(6);
      rows.push([token, namaWali || '', '', jenjang, sekolah, 'aktif', new Date(), '']);
      tokens.push(token);
    }
    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, tokens: tokens, count: tokens.length };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function generateRandomString(length) {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var res = '';
  for (var i = 0; i < length; i++) {
    res += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return res;
}

function getAllToken(filterJenjang, filterStatus) {
  try {
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var jenjang = (row[3] || '').toString().toUpperCase();
      var status  = (row[5] || '').toString().toLowerCase();

      if (filterJenjang && filterJenjang !== 'semua' && jenjang !== filterJenjang.toUpperCase()) continue;
      if (filterStatus && filterStatus !== 'semua' && status !== filterStatus.toLowerCase()) continue;

      result.push({
        token      : row[0],
        namaWali   : row[1] || '',
        noHpWali   : row[2] || '',
        jenjang    : row[3] || '',
        asalSekolah: row[4] || '',
        status     : row[5] || 'aktif',
        tglDibuat  : safeFormatDate(row[6], 'dd/MM/yyyy HH:mm'),
        tglDipakai : safeFormatDate(row[7], 'dd/MM/yyyy HH:mm')
      });
    }
    return { success: true, data: result, total: result.length };
  } catch(err) {
    return { success: false, message: err.message, data: [] };
  }
}

function updateToken(token, dataUpdate) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        var rowIdx = i + 1;
        if (dataUpdate.jenjang     !== undefined) sheet.getRange(rowIdx, 4).setValue(dataUpdate.jenjang);
        if (dataUpdate.asalSekolah !== undefined) sheet.getRange(rowIdx, 5).setValue(dataUpdate.asalSekolah);
        if (dataUpdate.namaWali    !== undefined) sheet.getRange(rowIdx, 2).setValue(dataUpdate.namaWali);
        if (dataUpdate.noHpWali    !== undefined) sheet.getRange(rowIdx, 3).setValue(dataUpdate.noHpWali);
        if (dataUpdate.status      !== undefined) sheet.getRange(rowIdx, 6).setValue(dataUpdate.status);
        return { success: true, message: 'Token ' + token + ' berhasil diperbarui.' };
      }
    }
    return { success: false, message: 'Token tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function hapusToken(token) {
  return nonaktifkanToken(token);
}

function nonaktifkanToken(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        sheet.getRange(i + 1, 6).setValue('nonaktif');
        return { success: true, message: 'Token ' + token + ' berhasil dinonaktifkan.' };
      }
    }
    return { success: false, message: 'Token tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function aktifkanToken(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        sheet.getRange(i + 1, 6).setValue('aktif');
        return { success: true, message: 'Token ' + token + ' berhasil diaktifkan kembali.' };
      }
    }
    return { success: false, message: 'Token tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

function hapusTokenPermanen(token) {
  try {
    token = token.toString().trim().toUpperCase();
    var sheet = getSheet(SHEET_TOKEN);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === token) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Token ' + token + ' berhasil dihapus permanen.' };
      }
    }
    return { success: false, message: 'Token tidak ditemukan.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// ADMIN: SOAL MANAGEMENT & IMPORT
// ============================================================
function getAllSoal() {
  try {
    var sheet = getSheet(SHEET_SOAL);
    var data = sheet.getDataRange().getValues();
    var result = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      result.push({
        id     : row[0],
        teks   : row[1],
        kode   : row[2],
        jenjang: row[3],
        opsiA  : row[4],
        opsiB  : row[5],
        nilaiA : row[6],
        nilaiB : row[7],
        aktif  : row[8]
      });
    }
    return { success: true, data: result, total: result.length };
  } catch(err) {
    return { success: false, message: err.message, data: [] };
  }
}

function importSoal(dataSoal, hapusLama) {
  try {
    var sheet = getSheet(SHEET_SOAL);
    if (hapusLama) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
    }

    var startId = sheet.getLastRow();
    var rows = dataSoal.map(function(s, idx) {
      var idStr = 'S' + ('000' + (startId + idx)).slice(-3);
      return [
        idStr,
        s.teks,
        s.kode ? s.kode.toUpperCase() : 'R',
        s.jenjang ? s.jenjang.toUpperCase() : 'SEMUA',
        s.opsiA || 'Ya',
        s.opsiB || 'Tidak',
        s.nilaiA !== undefined ? s.nilaiA : 1,
        s.nilaiB !== undefined ? s.nilaiB : 0,
        'aktif'
      ];
    });

    if (rows.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, message: 'Berhasil mengimpor ' + rows.length + ' soal.' };
  } catch(err) {
    return { success: false, message: 'Gagal impor soal: ' + err.message };
  }
}

function hapusSemuaSoal() {
  try {
    var sheet = getSheet(SHEET_SOAL);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return { success: true, message: 'Semua soal berhasil dihapus.' };
  } catch(err) {
    return { success: false, message: err.message };
  }
}
'''

with open('appscript/Code.gs', 'w', encoding='utf-8') as f:
    f.write(code_template)

print('appscript/Code.gs generated successfully!')
