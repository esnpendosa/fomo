with open('appscript/Admin.html', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Update table header in Laporan tab
old_th = '              <thead><tr><th>#</th><th>Token</th><th>Nama Siswa</th><th>Jenjang</th><th>Sekolah</th><th>Bidang Utama</th><th>Bidang 2</th><th>Bidang 3</th><th>Tanggal</th><th>Aksi</th></tr></thead>'
new_th = '              <thead><tr><th>#</th><th>Token</th><th>Nama Siswa</th><th>Pendidikan</th><th>Sekolah</th><th>R</th><th>I</th><th>A</th><th>S</th><th>E</th><th>C</th><th>Kode</th><th>Bidang Utama</th><th>Tanggal</th><th>Aksi</th></tr></thead>'
text = text.replace(old_th, new_th)

# 2. Add Reset button to token actions in renderTabelToken
old_aksi = '''      aksi += '<button class="btn btn-ghost btn-icon-sm" title="Hapus Permanen" onclick="hapusTokenPermanen(\\''+tokenEnc+'\\')"><svg class="icon icon-sm" style="color:var(--red-600);"><use href="#i-trash"/></svg></button>' +
        '</div>';'''
new_aksi = '''      aksi += '<button class="btn btn-ghost btn-icon-sm" title="Reset Tes / Ulangi" onclick="resetTokenAdmin(\\''+tokenEnc+'\\')"><svg class="icon icon-sm" style="color:var(--amber-600);"><use href="#i-refresh"/></svg></button>' +
        '<button class="btn btn-ghost btn-icon-sm" title="Hapus Permanen" onclick="hapusTokenPermanen(\\''+tokenEnc+'\\')"><svg class="icon icon-sm" style="color:var(--red-600);"><use href="#i-trash"/></svg></button>' +
        '</div>';'''
text = text.replace(old_aksi, new_aksi)

# 3. Add resetTokenAdmin function
old_aktif = '  // ---- AKTIFKAN TOKEN ----'
new_aktif = '''  // ---- RESET TOKEN (ADMIN) ----
  function resetTokenAdmin(token) {
    if (!confirm('Reset token ' + token + ' agar bisa mengerjakan tes ulang dari awal?\\n\\nHasil dan jawaban lama untuk token ini akan dihapus.')) return;
    setLoading(true, 'Mereset token...');
    google.script.run
      .withSuccessHandler(function(res) {
        setLoading(false);
        if (res.success) {
          showToast(res.message, 'success');
          muatDaftarToken();
          muatStatistik();
        } else {
          showToast('Gagal: ' + res.message, 'error');
        }
      })
      .withFailureHandler(function(err) {
        setLoading(false);
        showToast('Gagal: ' + (err.message || ''), 'error');
      })
      .resetToken(token);
  }

  // ---- AKTIFKAN TOKEN ----'''
text = text.replace(old_aktif, new_aktif)

# 4. Update renderTabelLaporan and implement eksporLaporanExcel
old_render_lap = '''  function renderTabelLaporan(data) {
    var body = document.getElementById('bodyLaporan');
    if (!data.length) { body.innerHTML = '<tr><td colspan="10" class="text-center" style="padding:28px;color:var(--gray-400);">Tidak ada data hasil.</td></tr>'; return; }
    var url = '<?= ScriptApp.getService().getUrl() ?>';
    body.innerHTML = data.map(function(r,i) {
      var jBadge = (r.jenjang||'').toUpperCase()==='SD' ? 'badge-primary' : 'badge-warning';
      return '<tr><td style="color:var(--gray-400);font-size:12px;">'+(i+1)+'</td>'+
        '<td><span class="mono" style="font-size:12px;color:var(--blue-700);">'+esc(r.token)+'</span></td>'+
        '<td style="font-weight:600;font-size:13px;">'+esc(r.nama||'—')+'</td>'+
        '<td><span class="badge '+jBadge+'">'+esc(r.jenjang||'—')+'</span></td>'+
        '<td style="font-size:12px;" class="truncate">'+esc(r.sekolah||'—')+'</td>'+
        '<td style="font-weight:600;color:var(--blue-700);font-size:12px;">'+esc(r.bidangUtama||'—')+'</td>'+
        '<td style="font-size:12px;color:var(--gray-500);">'+esc(r.bidang2||'—')+'</td>'+
        '<td style="font-size:12px;color:var(--gray-500);">'+esc(r.bidang3||'—')+'</td>'+
        '<td style="font-size:11px;white-space:nowrap;color:var(--gray-500);">'+esc(r.timestamp||'—')+'</td>'+
        '<td><a href="'+url+'?page=laporan-siswa&token='+encodeURIComponent(r.token)+'" target="_blank" class="btn btn-ghost btn-icon-sm" title="Lihat laporan per siswa"><svg class="icon icon-sm"><use href="#i-external"/></svg></a></td></tr>';
    }).join('');
  }

  function filterLaporan() {
    var q = document.getElementById('cariLaporan').value.toLowerCase();
    renderTabelLaporan(semuaLaporan.filter(function(r){ return (r.nama||'').toLowerCase().includes(q)||(r.token||'').toLowerCase().includes(q)||(r.sekolah||'').toLowerCase().includes(q); }));
  }

  function eksporLaporanCSV() {
    if (!semuaLaporan.length) { showToast('Tidak ada data.','warning'); return; }
    var h = 'Token,Nama_Siswa,Jenjang,Sekolah,Bidang_Utama,Bidang_2,Bidang_3,Skor_R,Skor_I,Skor_A,Skor_S,Skor_E,Skor_C,Skor_G,Tanggal\\n';
    var r = semuaLaporan.map(function(x){ var s=x.skor||{}; return [x.token,x.nama,x.jenjang,\'"\'+(x.sekolah||\'\')+\'"\',x.bidangUtama,x.bidang2,x.bidang3,s.R||0,s.I||0,s.A||0,s.S||0,s.E||0,s.C||0,s.G||0,x.timestamp].join(\',\'); }).join(\'\\n\');
    unduhCSV(\'laporan_fomo_\'+new Date().toISOString().slice(0,10)+\'.csv\', h+r);
    showToast(\'CSV laporan berhasil diunduh.\',\'success\');
  }'''

new_render_lap = '''  function renderTabelLaporan(data) {
    var body = document.getElementById('bodyLaporan');
    if (!data.length) { body.innerHTML = '<tr><td colspan="15" class="text-center" style="padding:28px;color:var(--gray-400);">Tidak ada data hasil.</td></tr>'; return; }
    var url = '<?= ScriptApp.getService().getUrl() ?>';
    body.innerHTML = data.map(function(r,i) {
      var s = r.skor || {};
      return '<tr><td style="color:var(--gray-400);font-size:12px;">'+(i+1)+'</td>'+
        '<td><span class="mono" style="font-size:11px;font-weight:700;color:var(--blue-700);">'+esc(r.token)+'</span></td>'+
        '<td style="font-weight:700;font-size:13px;">'+esc(r.nama||'—')+'</td>'+
        '<td style="font-size:12px;">'+esc(r.pendidikan||r.jenjang||'SD')+'</td>'+
        '<td style="font-size:12px;" class="truncate">'+esc(r.sekolah||'—')+'</td>'+
        '<td style="font-family:monospace;font-size:11px;font-weight:700;">'+(s.R||0)+'</td>'+
        '<td style="font-family:monospace;font-size:11px;font-weight:700;">'+(s.I||0)+'</td>'+
        '<td style="font-family:monospace;font-size:11px;font-weight:700;">'+(s.A||0)+'</td>'+
        '<td style="font-family:monospace;font-size:11px;font-weight:700;">'+(s.S||0)+'</td>'+
        '<td style="font-family:monospace;font-size:11px;font-weight:700;">'+(s.E||0)+'</td>'+
        '<td style="font-family:monospace;font-size:11px;font-weight:700;">'+(s.C||0)+'</td>'+
        '<td><span style="font-family:monospace;font-weight:800;color:#1D4ED8;background:#EFF6FF;padding:2px 6px;border-radius:4px;border:1px solid #BFDBFE;">'+esc(r.top3Kode||'—')+'</span></td>'+
        '<td style="font-weight:600;color:var(--blue-700);font-size:12px;">'+esc(r.bidangUtama||'—')+'</td>'+
        '<td style="font-size:11px;white-space:nowrap;color:var(--gray-500);">'+esc(r.timestamp||'—')+'</td>'+
        '<td><div class="d-flex gap-1">' +
          '<a href="'+url+'?page=hasil&token='+encodeURIComponent(r.token)+'" target="_blank" class="btn btn-ghost btn-icon-sm" title="Report Holland (Siswa)"><svg class="icon icon-sm"><use href="#i-external"/></svg></a>' +
          '<a href="'+url+'?page=laporan-siswa&token='+encodeURIComponent(r.token)+'" target="_blank" class="btn btn-ghost btn-icon-sm" title="Laporan Detail 108 Soal"><svg class="icon icon-sm"><use href="#i-file-text"/></svg></a>' +
        '</div></td></tr>';
    }).join('');
  }

  function filterLaporan() {
    var q = document.getElementById('cariLaporan').value.toLowerCase();
    renderTabelLaporan(semuaLaporan.filter(function(r){ return (r.nama||'').toLowerCase().includes(q)||(r.token||'').toLowerCase().includes(q)||(r.sekolah||'').toLowerCase().includes(q); }));
  }

  function eksporLaporanExcel() {
    if (!semuaLaporan.length) { showToast('Tidak ada data hasil untuk diekspor.','warning'); return; }
    var headers = ['No','Token','Nama Siswa','Jenis Kelamin','Tanggal Lahir','Pendidikan','Sekolah','Skor R','Skor I','Skor A','Skor S','Skor E','Skor C','Kode Holland','Bidang Utama','Bidang 2','Bidang 3','Konsistensi','Tanggal Selesai'];
    var rows = [headers];
    semuaLaporan.forEach(function(r, idx) {
      var s = r.skor || {};
      rows.push([
        idx + 1,
        r.token || '',
        r.nama || '',
        r.jenisKelamin || 'Laki-laki',
        r.tanggalLahir || '',
        r.pendidikan || r.jenjang || 'SD',
        r.sekolah || '',
        s.R !== undefined ? s.R : 0,
        s.I !== undefined ? s.I : 0,
        s.A !== undefined ? s.A : 0,
        s.S !== undefined ? s.S : 0,
        s.E !== undefined ? s.E : 0,
        s.C !== undefined ? s.C : 0,
        r.top3Kode || '',
        r.bidangUtama || '',
        r.bidang2 || '',
        r.bidang3 || '',
        r.konsistensi || 'Sedang',
        r.timestamp || ''
      ]);
    });
    unduhFileExcel('laporan_fomo_lengkap_' + new Date().toISOString().slice(0,10), rows);
    showToast('File Excel lengkap berhasil diunduh (' + semuaLaporan.length + ' data).','success');
  }

  function eksporLaporanCSV() {
    if (!semuaLaporan.length) { showToast('Tidak ada data.','warning'); return; }
    var headers = ['No','Token','Nama_Siswa','Jenis_Kelamin','Tanggal_Lahir','Pendidikan','Sekolah','Skor_R','Skor_I','Skor_A','Skor_S','Skor_E','Skor_C','Kode_Holland','Bidang_Utama','Bidang_2','Bidang_3','Konsistensi','Tanggal'];
    var lines = [headers.join(',')];
    semuaLaporan.forEach(function(r, idx) {
      var s = r.skor || {};
      lines.push([
        idx + 1,
        r.token,
        '"' + (r.nama || '').replace(/"/g, '""') + '"',
        r.jenisKelamin || 'Laki-laki',
        r.tanggalLahir || '',
        r.pendidikan || r.jenjang || 'SD',
        '"' + (r.sekolah || '').replace(/"/g, '""') + '"',
        s.R || 0, s.I || 0, s.A || 0, s.S || 0, s.E || 0, s.C || 0,
        r.top3Kode || '',
        '"' + (r.bidangUtama || '') + '"',
        '"' + (r.bidang2 || '') + '"',
        '"' + (r.bidang3 || '') + '"',
        r.konsistensi || 'Sedang',
        r.timestamp || ''
      ].join(','));
    });
    unduhCSV('laporan_fomo_' + new Date().toISOString().slice(0,10) + '.csv', lines.join('\\n'));
    showToast('CSV laporan berhasil diunduh.','success');
  }'''
text = text.replace(old_render_lap, new_render_lap)

# 5. Add unduhFileExcel
old_unduh = '''  function unduhCSV(nama, isi) {
    var blob = new Blob(['\\uFEFF'+isi],{type:'text/csv;charset=utf-8;'});
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a'); a.href=url; a.download=nama; a.click(); URL.revokeObjectURL(url);
  }'''

new_unduh = '''  function unduhCSV(nama, isi) {
    var blob = new Blob(['\\uFEFF'+isi],{type:'text/csv;charset=utf-8;'});
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a'); a.href=url; a.download=nama; a.click(); URL.revokeObjectURL(url);
  }

  function unduhFileExcel(filename, rows) {
    var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">' +
      '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
      '<x:Name>Laporan FOMO</x:Name>' +
      '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>' +
      '</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
      '<style>' +
      'table { border-collapse:collapse; font-family: Calibri, sans-serif; }' +
      'th { background:#F39C12; color:#FFFFFF; font-weight:bold; border:1px solid #D1D5DB; padding:8px 12px; text-align:center; }' +
      'td { border:1px solid #E5E7EB; padding:6px 10px; font-size:11pt; }' +
      '.mono { font-family: Consolas, monospace; font-weight:bold; }' +
      '</style>' +
      '</head><body><table>';

    rows.forEach(function(row, idx) {
      html += '<tr>';
      row.forEach(function(val) {
        var tag = idx === 0 ? 'th' : 'td';
        var txt = val !== undefined && val !== null ? String(val).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : '';
        html += '<' + tag + '>' + txt + '</' + tag + '>';
      });
      html += '</tr>';
    });

    html += '</table></body></html>';

    var blob = new Blob(['\\uFEFF' + html], { type: 'application/vnd.ms-excel;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename + '.xls';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }'''
text = text.replace(old_unduh, new_unduh)

with open('appscript/Admin.html', 'w', encoding='utf-8') as f:
    f.write(text)

print('Admin.html updated successfully!')
