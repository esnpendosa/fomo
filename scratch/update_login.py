with open('appscript/Login.html', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Add Modal HTML before </body>
modal_html = '''<!-- Modal Token Sudah Pernah Digunakan -->
<div class="modal-overlay hidden" id="modalTokenSelesai" style="position:fixed;inset:0;background:rgba(0,0,0,0.65);z-index:9999;display:none;align-items:center;justify-content:center;backdrop-filter:blur(3px);">
  <div style="background:#FFFFFF;border-radius:20px;max-width:420px;width:90%;padding:26px 22px;text-align:center;box-shadow:0 16px 40px rgba(0,0,0,0.25);animation:fadeIn .25s ease;">
    <div style="width:58px;height:58px;border-radius:50%;background:#FEF3C7;color:#D97706;display:flex;align-items:center;justify-content:center;margin:0 auto 14px;font-size:26px;">📋</div>
    <div style="font-size:18px;font-weight:800;color:#1F2937;margin-bottom:6px;">Token Sudah Pernah Digunakan</div>
    <div style="font-size:13.5px;color:#4B5563;line-height:1.55;margin-bottom:20px;">
      Token <strong id="selesaiTokenText" style="color:#1D4ED8;font-family:monospace;">—</strong> sudah memiliki hasil tes atas nama <strong id="selesaiNamaText" style="color:#111827;">—</strong>.
      <br/><span style="font-size:12px;color:#6B7280;margin-top:6px;display:block;">Pilih apakah Anda ingin melihat hasil sebelumnya atau mengulang tes dari awal.</span>
    </div>
    <div style="display:flex;flex-direction:column;gap:10px;">
      <button class="btn btn-primary btn-lg" onclick="bukaHasilSelesai()" style="background:#1565C0;color:white;border-radius:12px;padding:12px;font-weight:700;border:none;cursor:pointer;">
        👁️ Lihat Hasil Tes Sebelumnya
      </button>
      <button class="btn btn-secondary btn-lg" onclick="konfirmasiUlangiTes()" style="background:#F8FAFC;color:#D97706;border:1.5px solid #F59E0B;border-radius:12px;padding:12px;font-weight:700;cursor:pointer;">
        🔄 Ulangi Tes dari Awal (Reset)
      </button>
      <button class="btn btn-ghost" onclick="tutupModalSelesai()" style="background:transparent;border:none;padding:8px;font-size:12px;color:#64748B;cursor:pointer;">
        Batal / Gunakan Token Lain
      </button>
    </div>
  </div>
</div>
</body>'''

text = text.replace('</body>', modal_html)

# 2. Add handlers in script
old_status = '''        if (res.status === 'sudah_selesai') {
          setTimeout(function() {
            window.location.href = '<?= ScriptApp.getService().getUrl() ?>?page=hasil&token=' + res.token;
          }, 800);
        } else if (res.status === 'lanjut') {'''

new_status = '''        if (res.status === 'sudah_selesai') {
          setLoading(false);
          tampilModalSelesai(res);
          return;
        } else if (res.status === 'lanjut') {'''

text = text.replace(old_status, new_status)

# 3. Add JS helper functions before closing script
helper_funcs = '''  var tokenTerpilihSelesai = '';
  function tampilModalSelesai(res) {
    tokenTerpilihSelesai = res.token;
    document.getElementById('selesaiTokenText').textContent = res.token;
    document.getElementById('selesaiNamaText').textContent = res.namaPeserta || 'Peserta';
    var m = document.getElementById('modalTokenSelesai');
    m.style.display = 'flex';
    m.classList.remove('hidden');
  }

  function tutupModalSelesai() {
    var m = document.getElementById('modalTokenSelesai');
    m.style.display = 'none';
    m.classList.add('hidden');
  }

  function bukaHasilSelesai() {
    window.location.href = '<?= ScriptApp.getService().getUrl() ?>?page=hasil&token=' + encodeURIComponent(tokenTerpilihSelesai);
  }

  function konfirmasiUlangiTes() {
    if (!confirm('Yakin ingin mengulang tes dari awal?\\n\\nHasil dan jawaban lama untuk token ini akan di-reset (dihapus).')) return;
    setLoading(true, 'Mereset token untuk tes baru...');
    google.script.run
      .withSuccessHandler(function(res) {
        setLoading(false);
        if (res.success) {
          sessionStorage.removeItem('fomo_jwb_' + tokenTerpilihSelesai);
          window.location.href = '<?= ScriptApp.getService().getUrl() ?>?page=biodata&token=' + encodeURIComponent(tokenTerpilihSelesai);
        } else {
          alert('Gagal reset: ' + res.message);
        }
      })
      .withFailureHandler(function(err) {
        setLoading(false);
        alert('Gagal menghubungi server: ' + (err.message || ''));
      })
      .resetToken(tokenTerpilihSelesai);
  }

  window.onload = function() {'''

text = text.replace('  window.onload = function() {', helper_funcs)

with open('appscript/Login.html', 'w', encoding='utf-8') as f:
    f.write(text)

print('Login.html updated successfully!')
