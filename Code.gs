/* =========================================================
   SEIMBANG v1.0 — Media Pembelajaran Dasar-Dasar Akuntansi Kelas X SMK
   Backend Google Apps Script + Google Sheets
   Tempel SELURUH isi berkas ini ke Code.gs di proyek Apps Script.
   ========================================================= */

/* =========================================================
   SEIMBANG — PENGHUBUNG GOOGLE APPS SCRIPT
   Cara pasang ringkas (lihat panduan lengkap):
   1. Buat Google Sheet baru → Ekstensi → Apps Script.
   2. Tempel seluruh isi file ini ke Code.gs, lalu buat file HTML
      bernama "index" dan tempel isi index.html.
   3. Pilih fungsi setup lalu klik Jalankan (izinkan akses).
   4. Terapkan → Deployment baru → Aplikasi web →
      Jalankan sebagai: Saya, Akses: Siapa saja → salin URL.
   ========================================================= */

var SEIMBANG_VERSI = '1.1';

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'ping') {
    return ContentService.createTextOutput(JSON.stringify(handle_({ action: 'ping' }))).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SEIMBANG · Dasar-Dasar Akuntansi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* API untuk halaman yang di-host di tempat lain (mis. GitHub Pages) */
function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) { req = {}; }
  return ContentService.createTextOutput(JSON.stringify(handle_(req))).setMimeType(ContentService.MimeType.JSON);
}

/* API untuk halaman yang disajikan Apps Script (google.script.run) */
function api(reqText) {
  var req = {};
  try { req = JSON.parse(reqText); } catch (err) { req = {}; }
  return JSON.stringify(handle_(req));
}

function handle_(req) {
  return SBServer.create(gasEnv_()).handle(req);
}

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet belum terhubung. Jalankan fungsi setup() dari editor Apps Script.');
  return ss;
}

function gasEnv_() {
  var ss = ss_();
  var props = PropertiesService.getScriptProperties();
  var tables = {};
  var lockHeld = false;
  return {
    table: function (name) { return tables[name] = tables[name] || gasTable_(ss, name); },
    secret: function () {
      var s = props.getProperty('SECRET');
      if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('SECRET', s); }
      return s;
    },
    random: function () { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); },
    now: function () { return Date.now(); },
    lock: function (fn) {
      if (lockHeld) return fn();
      var lock = LockService.getScriptLock();
      lock.waitLock(25000);
      lockHeld = true;
      try {
        // muat ulang data di dalam kunci agar tidak menimpa tulisan siswa lain
        Object.keys(tables).forEach(function (k) { tables[k].reload(); });
        return fn();
      } finally { lockHeld = false; SpreadsheetApp.flush(); lock.releaseLock(); }
    },
    cache: {
      get: function (k) { return CacheService.getScriptCache().get(k); },
      put: function (k, v, ttl) { CacheService.getScriptCache().put(k, v, Math.min(21600, ttl || 600)); },
      remove: function (k) { CacheService.getScriptCache().remove(k); }
    },
    info: { mode: 'gas', versi: SEIMBANG_VERSI, sheetUrl: ss.getUrl() }
  };
}

function gasTable_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" belum ada. Jalankan fungsi setup() di editor Apps Script.');
  var cache = null;
  function toCell(v) {
    if (v === undefined || v === null) return '';
    if (typeof v === 'boolean' || typeof v === 'number') return v;
    if (v instanceof Date) return v;
    if (typeof v === 'object') v = JSON.stringify(v);
    v = String(v);
    if (/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(\.\d+)?Z$/.test(v)) return new Date(v);
    if (/^[=+\-@]/.test(v)) return "'" + v;
    return v;
  }
  function load() {
    if (cache) return cache;
    var lastCol = Math.max(1, sh.getLastColumn()), lastRow = sh.getLastRow();
    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    var vals = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
    var rows = [];
    vals.forEach(function (r, i) {
      if (r.join('') === '') return;
      var o = { _r: i + 2 };
      header.forEach(function (h, j) { var v = r[j]; if (v instanceof Date) v = v.toISOString(); o[h] = v; });
      rows.push(o);
    });
    cache = { header: header, rows: rows };
    return cache;
  }
  function rowValues(c, o) { return c.header.map(function (h) { return toCell(o[h]); }); }
  return {
    reload: function () { cache = null; },
    all: function () { return load().rows; },
    insert: function (o) {
      var c = load();
      sh.appendRow(rowValues(c, o));
      var r = { _r: sh.getLastRow() };
      c.header.forEach(function (h) { r[h] = o[h] === undefined ? '' : o[h]; });
      c.rows.push(r);
    },
    update: function (f, v, patch) {
      var c = load();
      for (var i = 0; i < c.rows.length; i++) {
        var r = c.rows[i];
        if (String(r[f]) === String(v)) {
          for (var k in patch) r[k] = patch[k];
          sh.getRange(r._r, 1, 1, c.header.length).setValues([rowValues(c, r)]);
          return true;
        }
      }
      return false;
    },
    remove: function (f, v) {
      var c = load();
      var idx = c.rows.filter(function (r) { return String(r[f]) === String(v); }).map(function (r) { return r._r; }).sort(function (a, b) { return b - a; });
      idx.forEach(function (r) { sh.deleteRow(r); });
      cache = null;
      return idx.length;
    }
  };
}

/* ---------- Jalankan sekali dari editor ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari menu Ekstensi di Google Sheet, lalu jalankan setup lagi.');
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SS_ID', ss.getId());
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());

  var warna = { Pengaturan: '#1d2b53', Akun: '#1f6f4a', Progres: '#2448a8', Kuis: '#8a5a00', Refleksi: '#7a2f6b' };
  Object.keys(SBServer.SCHEMA).forEach(function (name) {
    var cols = SBServer.SCHEMA[name];
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var head = sh.getRange(1, 1, 1, cols.length);
    var lama = sh.getLastRow() > 0 ? sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String).filter(function (x) { return x; }) : [];
    if (!lama.length) head.setValues([cols]);
    else {
      // versi baru menambah kolom: tambahkan di ujung tanpa mengganggu data lama
      cols.forEach(function (c) { if (lama.indexOf(c) < 0) { lama.push(c); sh.getRange(1, lama.length).setValue(c); } });
      head = sh.getRange(1, 1, 1, lama.length);
    }
    head.setFontWeight('bold').setFontColor('#ffffff').setBackground(warna[name] || '#333333');
    sh.setFrozenRows(1);
    var hdr = lama.length ? lama : cols;
    (SBServer.TEXT_COLS[name] || []).forEach(function (c) {
      var i = hdr.indexOf(c);
      if (i >= 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });
  var bawaan = ss.getSheetByName('Sheet1') || ss.getSheetByName('Lembar1');
  if (bawaan && bawaan.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(bawaan);

  var env = gasEnv_();
  var core = SBServer.create(env);
  var mentah = {};
  env.table('Pengaturan').all().forEach(function (r) { mentah[String(r.kunci)] = 1; });
  SETTINGS_DEF.forEach(function (d) { if (!mentah[d[0]]) env.table('Pengaturan').insert({ kunci: d[0], nilai: d[1], keterangan: d[2] }); });
  if (!core._findUser('guru')) core._buatAkun({ username: 'guru', nama: 'Dr. Herry Sofyandy P., M.Pd., Gr.', peran: 'guru', password: 'guru123' });
  SpreadsheetApp.flush();
  Logger.log('SEIMBANG siap. Akun guru: username "guru", sandi "guru123". Segera ganti sandi setelah masuk.');
  try { SpreadsheetApp.getUi().alert('SEIMBANG siap dipakai.\n\nAkun guru:\nusername: guru\nkata sandi: guru123\n\nLanjutkan dengan Terapkan → Deployment baru → Aplikasi web.'); } catch (e) {}
}

/* Lupa sandi guru? Jalankan fungsi ini dari editor: sandi kembali ke guru123 */
function resetSandiGuru() {
  var env = gasEnv_();
  env.table('Akun').remove('username', 'guru');
  SBServer.create(env)._buatAkun({ username: 'guru', nama: 'Dr. Herry Sofyandy P., M.Pd., Gr.', peran: 'guru', password: 'guru123' });
  Logger.log('Sandi akun guru dikembalikan ke guru123.');
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('SEIMBANG')
      .addItem('Siapkan / perbaiki sheet', 'setup')
      .addItem('Kembalikan sandi guru ke guru123', 'resetSandiGuru')
      .addToUi();
  } catch (e) {}
}


/* ---------- KONTEN (termasuk bank soal & kunci, hanya di server) ---------- */
/* =========================================================
   KONTEN PEMBELAJARAN — SEIMBANG
   Mapel  : Dasar-Dasar Akuntansi dan Keuangan Lembaga (Kelas X SMK, Fase E)
   Elemen : Prinsip-prinsip dan Konsep Akuntansi Dasar
   Konteks: usaha jasa "Laundry Kilat" milik Raka di Kota Jambi
   ========================================================= */

const CP = {
  mapel: 'Dasar-Dasar Akuntansi dan Keuangan Lembaga',
  fase: 'Fase E · Kelas X SMK',
  elemen: 'Prinsip-prinsip dan Konsep Akuntansi Dasar',
  teks: 'Pada akhir fase E, peserta didik mampu memahami pengertian akuntansi, tujuan pencatatan akuntansi, pihak-pihak yang membutuhkan informasi akuntansi, prinsip-prinsip akuntansi, serta konsep akuntansi dasar.',
  cakupan: 'Media ini mencakup bagian akuntansi dasar dari elemen tersebut dan menutupnya dengan persamaan dasar akuntansi sebagai jembatan menuju siklus akuntansi perusahaan jasa.',
  tp: [
    'Menjelaskan pengertian, kegiatan, bidang, dan tujuan akuntansi serta mengenali transaksi dan bukti transaksi.',
    'Mengidentifikasi pihak internal dan eksternal yang membutuhkan informasi akuntansi, laporan yang mereka baca, dan syarat informasi yang berguna.',
    'Mengidentifikasi asumsi dasar dan prinsip-prinsip akuntansi dalam kasus usaha sehari-hari.',
    'Menganalisis pengaruh transaksi keuangan terhadap persamaan dasar akuntansi perusahaan jasa.'
  ]
};

const MISI = [
  /* =============== MISI 1 =============== */
  {
    id: 'm1', no: 1, judul: 'Kenalan dengan Akuntansi',
    ringkas: 'Pengertian, transaksi, kegiatan, bidang, dan tujuan akuntansi',
    tp: 0,
    bagian: [
      {
        judul: 'Uang Raka ke mana?',
        html: `<p>Raka, lulusan SMK, membuka usaha <b>Laundry Kilat</b> di dekat kos-kosan mahasiswa di Kota Jambi. Pelanggannya ramai setiap hari. Anehnya, di akhir bulan uangnya hampir habis.</p>
<p>Raka bingung. Ia tidak tahu berapa yang sudah dibelanjakan untuk deterjen, pelanggan mana yang belum membayar, dan apakah usahanya untung atau rugi.</p>
<div class="analogi"><b>Bayangkan</b><p>Kamu main gim tanpa melihat skor dan sisa nyawa. Kamu tidak tahu sedang menang atau kalah. Usaha tanpa catatan keuangan rasanya seperti itu.</p></div>
<p>Masalah Raka sederhana: <b>ia tidak mencatat.</b> Di media ini kamu akan membantu Raka menata keuangan Laundry Kilat langkah demi langkah.</p>`,
        pemantik: {
          tanya: 'Menurutmu, catatan apa saja yang seharusnya dimiliki Raka?',
          jawab: 'Uang yang masuk dan keluar (kas), pelanggan yang belum membayar (piutang), utang ke toko deterjen, dan barang milik usaha seperti mesin cuci dan setrika. Semua catatan itu adalah bahan akuntansi.'
        }
      },
      {
        judul: 'Apa itu akuntansi?',
        html: `<div class="def"><span class="def-k">Definisi</span><p><b>Akuntansi</b> adalah proses mengidentifikasi, mencatat, menggolongkan, mengikhtisarkan, dan melaporkan transaksi keuangan suatu usaha, lalu menafsirkan hasilnya, sehingga menghasilkan informasi yang berguna untuk mengambil keputusan.</p></div>
<p>Definisi itu panjang. Kita bedah per bagian:</p>
<dl class="who">
<dt>"Proses"</dt><dd>Akuntansi adalah rangkaian kegiatan yang berurutan, bukan pekerjaan sekali jadi.</dd>
<dt>"Transaksi keuangan"</dt><dd>Inilah <b>bahan mentahnya</b>: kejadian yang mengubah keuangan usaha.</dd>
<dt>"Informasi"</dt><dd>Inilah <b>hasilnya</b>: laporan keuangan yang rapi dan mudah dibaca.</dd>
<dt>"Mengambil keputusan"</dt><dd>Inilah <b>tujuannya</b>: membantu orang memilih tindakan, misalnya menambah mesin cuci atau tidak.</dd>
</dl>
<div class="analogi"><b>Seperti dapur</b><p>Bahan mentah (transaksi) dicuci, dipotong, dan dimasak (dicatat, digolongkan, diringkas) sampai menjadi hidangan (laporan keuangan) yang siap dinikmati (dipakai untuk keputusan).</p></div>
<p>Karena laporannya dipahami pemilik, bank, pemerintah, dan banyak pihak lain, akuntansi sering disebut <b>bahasa bisnis</b>.</p>`
      },
      {
        judul: 'Transaksi dan buktinya',
        html: `<p>Tidak semua kejadian di usaha dicatat. Yang dicatat hanya <b>transaksi keuangan</b>, yaitu kejadian yang memenuhi tiga ciri:</p>
<ol class="steps">
<li><b>Mengubah keadaan keuangan</b><span>Ada harta, utang, atau modal yang bertambah atau berkurang.</span></li>
<li><b>Dapat diukur dengan uang</b><span>Nilainya jelas dalam rupiah.</span></li>
<li><b>Didukung bukti</b><span>Ada dokumen yang bisa diperiksa orang lain.</span></li>
</ol>
<p>Dokumen itu disebut <b>bukti transaksi</b>. Contohnya:</p>
<div class="tbl-wrap"><table><thead><tr><th>Bukti</th><th>Dipakai untuk</th></tr></thead><tbody>
<tr><td><b>Nota kontan</b></td><td>Pembelian atau penjualan tunai, misalnya beli deterjen di toko.</td></tr>
<tr><td><b>Faktur</b></td><td>Pembelian atau penjualan secara kredit (dibayar belakangan).</td></tr>
<tr><td><b>Kuitansi</b></td><td>Tanda terima uang yang ditandatangani penerima, misalnya bayar sewa ruko.</td></tr>
<tr><td><b>Bukti kas masuk / keluar</b></td><td>Catatan penerimaan atau pengeluaran uang yang dibuat usaha sendiri.</td></tr>
<tr><td><b>Memo internal</b></td><td>Perintah tertulis pimpinan ke bagian akuntansi, misalnya mencatat mesin yang rusak.</td></tr>
<tr><td><b>Struk digital</b></td><td>Pembayaran lewat QRIS, transfer bank, atau dompet digital.</td></tr>
</tbody></table></div>
<p class="note"><b>Bukti eksternal</b> berasal dari pihak luar (nota dari toko, faktur dari pemasok). <b>Bukti internal</b> dibuat sendiri oleh usaha (bukti kas keluar, memo).</p>`,
        cek: {
          judul: 'Cek cepat: transaksi atau bukan?',
          opsi: ['Transaksi', 'Bukan'],
          item: [
            { t: 'Raka membeli deterjen Rp150.000 dan menerima nota.', a: 0, why: 'Ada uang keluar, nilainya jelas, dan ada bukti nota.' },
            { t: 'Raka berkenalan dengan calon pelanggan baru.', a: 1, why: 'Tidak ada perubahan keuangan yang bisa diukur dengan uang.' },
            { t: 'Pelanggan membayar cucian 5 kg sebesar Rp35.000 lewat QRIS.', a: 0, why: 'Kas bertambah karena pendapatan jasa, buktinya struk digital.' },
            { t: 'Raka berencana mengecat ulang papan nama tokonya bulan depan.', a: 1, why: 'Baru rencana. Belum terjadi dan belum ada uang yang berpindah.' },
            { t: 'Raka menerima pesanan cucian lewat WhatsApp. Cucian belum diantar dan belum dibayar.', a: 1, why: 'Pesanan saja belum mengubah keuangan. Transaksi terjadi saat jasa selesai atau uang diterima.' },
            { t: 'Raka membayar sewa ruko lewat transfer bank dan menyimpan bukti transfernya.', a: 0, why: 'Kas berkurang untuk beban sewa dan ada bukti transfer.' }
          ]
        }
      },
      {
        judul: 'Enam kegiatan akuntansi',
        html: `<p>Kegiatan akuntansi berjalan berurutan. Setiap langkah menghasilkan sesuatu yang dipakai langkah berikutnya:</p>
<ol class="steps">
<li><b>Identifikasi</b><span>Memeriksa bukti: apakah kejadian ini transaksi? Contoh: memeriksa nota pembelian deterjen.<br><i>Hasil: bukti transaksi yang sah.</i></span></li>
<li><b>Pencatatan</b><span>Mencatat transaksi menurut urutan tanggal.<br><i>Hasil: jurnal.</i></span></li>
<li><b>Penggolongan</b><span>Mengelompokkan transaksi ke akun yang sama: Kas, Perlengkapan, Utang Usaha, dan lainnya.<br><i>Hasil: buku besar.</i></span></li>
<li><b>Pengikhtisaran</b><span>Meringkas saldo semua akun untuk memeriksa keseimbangan.<br><i>Hasil: neraca saldo.</i></span></li>
<li><b>Pelaporan</b><span>Menyusun laporan untuk pembaca.<br><i>Hasil: laporan laba rugi, perubahan modal, dan posisi keuangan.</i></span></li>
<li><b>Penafsiran</b><span>Membaca laporan untuk mengambil keputusan, misalnya Raka memutuskan menambah satu mesin cuci karena laba naik.</span></li>
</ol>
<div class="tahu"><b>Tahukah kamu?</b><p>Rangkaian ini diulang setiap periode, misalnya setiap bulan. Karena berulang seperti lingkaran, para akuntan menyebutnya <b>siklus akuntansi</b>. Kamu akan mempelajarinya lebih dalam setelah media ini.</p></div>
<p class="note"><b>Trik mengingat:</b> "<b>I-Pe-Pe-Pe-Pe-Pe</b>": Identifikasi, Pencatatan, Penggolongan, Pengikhtisaran, Pelaporan, Penafsiran.</p>`
      },
      {
        judul: 'Bidang akuntansi dan profesinya',
        html: `<p>Akuntansi punya banyak bidang. Setiap bidang fokus pada kebutuhan yang berbeda. Ketuk kartu untuk melihat penjelasannya.</p>`,
        flip: [
          { d: 'Akuntansi keuangan', b: 'Menyusun laporan keuangan untuk pihak luar seperti bank dan investor. Inilah bidang yang kamu pelajari sekarang.' },
          { d: 'Akuntansi manajemen', b: 'Menyediakan informasi untuk pimpinan di dalam usaha, misalnya menentukan tarif cuci per kilogram.' },
          { d: 'Akuntansi biaya', b: 'Menghitung biaya untuk menghasilkan barang atau jasa, misalnya biaya mencuci 1 kg pakaian.' },
          { d: 'Akuntansi perpajakan', b: 'Menghitung, membayar, dan melaporkan pajak usaha sesuai aturan.' },
          { d: 'Pemeriksaan (auditing)', b: 'Memeriksa apakah laporan keuangan sudah wajar dan bisa dipercaya. Dilakukan oleh auditor.' },
          { d: 'Akuntansi pemerintahan', b: 'Mencatat keuangan negara dan daerah, misalnya APBD Provinsi Jambi.' },
          { d: 'Sistem informasi akuntansi', b: 'Merancang sistem dan aplikasi pencatatan, termasuk aplikasi akuntansi di komputer dan HP.' },
          { d: 'Akuntansi anggaran', b: 'Menyusun rencana keuangan untuk periode mendatang lalu membandingkannya dengan kenyataan.' }
        ],
        html2: `<p><b>Profesi akuntan</b> antara lain akuntan publik (auditor independen), akuntan perusahaan, akuntan pemerintah, dan akuntan pendidik. Lulusan SMK Akuntansi bisa bekerja sebagai staf administrasi keuangan, kasir, staf pajak, operator aplikasi akuntansi, atau membuka usaha sendiri seperti Raka.</p>`
      },
      {
        judul: 'Tujuan dan peran akuntansi',
        html: `<p><b>Tujuan utama</b> akuntansi adalah menyediakan informasi keuangan yang berguna untuk pengambilan keputusan. Rinciannya:</p>
<ul class="dots">
<li>Mengetahui <b>posisi keuangan</b>: berapa harta, utang, dan modal usaha.</li>
<li>Mengetahui <b>kinerja</b>: apakah usaha memperoleh laba atau menderita rugi.</li>
<li>Bahan <b>pertanggungjawaban</b> pengelola kepada pemilik dan pihak lain.</li>
<li>Dasar <b>perencanaan dan pengendalian</b>, misalnya menghemat biaya listrik bulan depan.</li>
</ul>
<p><b>Peran akuntansi</b> ada di mana-mana: UMKM sampai perusahaan besar, pemerintah daerah, koperasi sekolah, masjid, bendahara OSIS, bahkan keuangan keluarga.</p>`,
        pemantik: {
          tanya: 'Di sekitarmu, siapa yang sebenarnya sudah "berakuntansi" tanpa menyebutnya akuntansi?',
          jawab: 'Contohnya bendahara kelas yang mencatat uang kas, ibu yang mencatat belanja bulanan, atau pemilik warung yang punya buku utang pelanggan. Semuanya mencatat transaksi untuk mengambil keputusan.'
        }
      },
      {
        judul: 'Rangkuman Misi 1',
        html: `<ul class="rangkum">
<li><b>Akuntansi</b> mengolah transaksi keuangan menjadi informasi untuk mengambil keputusan.</li>
<li><b>Transaksi</b> punya tiga ciri: mengubah keuangan, terukur dengan uang, dan ada buktinya.</li>
<li>Bukti transaksi: nota, faktur, kuitansi, bukti kas masuk/keluar, memo, dan struk digital.</li>
<li>Kegiatan akuntansi: identifikasi → pencatatan → penggolongan → pengikhtisaran → pelaporan → penafsiran.</li>
<li>Bidang akuntansi antara lain keuangan, manajemen, biaya, perpajakan, pemeriksaan, pemerintahan, sistem informasi, dan anggaran.</li>
<li>Tujuannya: mengetahui posisi keuangan dan kinerja, pertanggungjawaban, serta perencanaan.</li>
</ul>
<div class="istilah"><b>Istilah penting</b><span>Transaksi</span><span>Bukti transaksi</span><span>Faktur</span><span>Kuitansi</span><span>Jurnal</span><span>Buku besar</span><span>Neraca saldo</span><span>Siklus akuntansi</span></div>`
      }
    ],
    aktivitas: {
      tipe: 'urut',
      judul: 'Susun urutan kegiatan akuntansi',
      petunjuk: 'Ketuk kegiatan sesuai urutan yang benar, dari yang pertama sampai yang terakhir.',
      item: ['Identifikasi bukti transaksi', 'Pencatatan ke jurnal', 'Penggolongan ke buku besar', 'Pengikhtisaran (neraca saldo)', 'Pelaporan keuangan', 'Penafsiran untuk keputusan']
    }
  },

  /* =============== MISI 2 =============== */
  {
    id: 'm2', no: 2, judul: 'Siapa Butuh Info Akuntansi?',
    ringkas: 'Pengguna informasi, laporan keuangan, dan syarat informasi yang berguna',
    tp: 1,
    bagian: [
      {
        judul: 'Satu laporan, banyak pembaca',
        html: `<div class="analogi"><b>Seperti rapor</b><p>Rapormu dibaca olehmu, orang tua, wali kelas, dan sekolah. Isinya sama, tetapi setiap pembaca mencari hal yang berbeda: orang tua melihat nilai, wali kelas melihat kehadiran, sekolah melihat kenaikan kelas.</p></div>
<p>Laporan keuangan Laundry Kilat juga dibaca banyak pihak dengan kepentingan berbeda. Para pengguna informasi akuntansi dibagi dua kelompok:</p>
<div class="duo">
<div><b>Pihak internal</b><span>Berada di dalam usaha dan ikut mengelolanya: pemilik dan manajer atau pimpinan.</span></div>
<div><b>Pihak eksternal</b><span>Berada di luar usaha: investor, kreditur, pemasok, pelanggan, pemerintah, dan masyarakat.</span></div>
</div>
<p class="note"><b>Cara membedakan:</b> tanyakan "apakah ia ikut mengelola usaha sehari-hari?" Jika ya, ia pihak internal.</p>`,
        pemantik: {
          tanya: 'Kalau kamu pemilik Laundry Kilat, informasi apa yang paling ingin kamu ketahui setiap akhir bulan?',
          jawab: 'Biasanya pemilik ingin tahu untung atau rugi, berapa uang kas yang tersisa, pelanggan mana yang belum membayar, dan utang apa yang harus segera dilunasi. Semua itu ada di laporan keuangan.'
        }
      },
      {
        judul: 'Pihak internal',
        html: `<p>Pihak internal memakai informasi akuntansi untuk <b>mengelola</b> usaha.</p>
<dl class="who">
<dt>Pemilik</dt><dd><i>"Usahaku untung atau rugi? Perlu tambah modal? Berapa yang boleh kuambil untuk keperluan pribadi?"</i><br>Pemilik melihat laba dan perkembangan modal.</dd>
<dt>Manajer / pimpinan</dt><dd><i>"Biaya mana yang bisa dihemat? Berapa tarif cuci yang tepat? Perlu tambah karyawan?"</i><br>Manajer memakai informasi untuk merencanakan, mengendalikan biaya, dan menilai kinerja.</dd>
</dl>
<div class="tahu"><b>Tahukah kamu?</b><p>Di usaha kecil seperti Laundry Kilat, pemilik sering sekaligus menjadi manajer. Raka mengelola sendiri usahanya, jadi ia membaca laporan dengan dua "topi" sekaligus.</p></div>`
      },
      {
        judul: 'Pihak eksternal',
        html: `<p>Pihak eksternal memakai informasi akuntansi untuk <b>menilai</b> usaha dari luar.</p>
<div class="tbl-wrap"><table><thead><tr><th>Pihak</th><th>Pertanyaan mereka</th></tr></thead><tbody>
<tr><td><b>Investor</b></td><td>"Kalau aku menanam modal, untungnya berapa dan risikonya besar tidak?"</td></tr>
<tr><td><b>Kreditur</b> (bank, koperasi)</td><td>"Sanggupkah usaha ini mengembalikan pinjaman beserta bunganya?"</td></tr>
<tr><td><b>Pemasok</b></td><td>"Kalau dijual secara kredit, apakah dibayar tepat waktu?"</td></tr>
<tr><td><b>Pelanggan</b></td><td>"Apakah usaha ini akan terus buka dan layanannya berlanjut?"</td></tr>
<tr><td><b>Pemerintah</b></td><td>"Berapa pajak yang harus dibayar? Bagaimana perkembangan UMKM daerah?"</td></tr>
<tr><td><b>Masyarakat</b></td><td>"Apa sumbangan usaha ini bagi lapangan kerja dan lingkungan?"</td></tr>
</tbody></table></div>
<p class="note"><b>Catatan:</b> menurut kerangka konseptual IAI, <b>karyawan</b> juga pemakai laporan keuangan karena berkepentingan atas gaji dan kelangsungan kerja. Sebagian buku menggolongkannya sebagai pihak internal karena bekerja di dalam usaha.</p>`,
        cek: {
          judul: 'Cek cepat: internal atau eksternal?',
          opsi: ['Internal', 'Eksternal'],
          item: [
            { t: 'Bank menilai pengajuan kredit Laundry Kilat.', a: 1, why: 'Bank berada di luar usaha dan berkepentingan atas pinjaman.' },
            { t: 'Raka menghitung laba bulan ini.', a: 0, why: 'Raka adalah pemilik yang mengelola usahanya.' },
            { t: 'Kantor pajak memeriksa laporan usaha.', a: 1, why: 'Pemerintah adalah pihak luar.' }
          ]
        }
      },
      {
        judul: 'Laporan yang mereka baca',
        html: `<p>Hasil akhir akuntansi adalah <b>laporan keuangan</b>. Untuk usaha jasa kecil, ada tiga laporan utama. Ini laporan Laundry Kilat bulan Oktober 2026:</p>
<div class="lap"><div class="lap-h">1. Laporan laba rugi <small>Apakah usaha untung?</small></div>
<div class="lap-r"><span>Pendapatan jasa</span><span>Rp5.000.000</span></div>
<div class="lap-r in"><span>Beban listrik dan air</span><span>Rp600.000</span></div>
<div class="lap-r in"><span>Beban gaji</span><span>Rp1.200.000</span></div>
<div class="lap-r"><span>Total beban</span><span>(Rp1.800.000)</span></div>
<div class="lap-r tot"><span>Laba bersih</span><span>Rp3.200.000</span></div></div>
<div class="lap"><div class="lap-h">2. Laporan perubahan modal <small>Bagaimana modal berubah?</small></div>
<div class="lap-r"><span>Setoran modal awal (1 Oktober)</span><span>Rp20.000.000</span></div>
<div class="lap-r"><span>Ditambah laba bersih</span><span>Rp3.200.000</span></div>
<div class="lap-r"><span>Dikurangi prive</span><span>(Rp500.000)</span></div>
<div class="lap-r tot"><span>Modal akhir (31 Oktober)</span><span>Rp22.700.000</span></div></div>
<div class="lap"><div class="lap-h">3. Laporan posisi keuangan (neraca) <small>Apa yang dimiliki dan dari mana asalnya?</small></div>
<div class="lap-r"><span>Kas</span><span>Rp13.700.000</span></div>
<div class="lap-r"><span>Perlengkapan</span><span>Rp1.500.000</span></div>
<div class="lap-r"><span>Peralatan</span><span>Rp8.000.000</span></div>
<div class="lap-r tot"><span>Total harta</span><span>Rp23.200.000</span></div>
<div class="lap-r"><span>Utang usaha</span><span>Rp500.000</span></div>
<div class="lap-r"><span>Modal Raka</span><span>Rp22.700.000</span></div>
<div class="lap-r tot"><span>Total utang dan modal</span><span>Rp23.200.000</span></div></div>
<p>Perhatikan: laba dari laporan pertama masuk ke laporan kedua, dan modal akhir dari laporan kedua masuk ke neraca. Ketiga laporan saling terhubung.</p>`,
        cek: {
          judul: 'Cek cepat: laporan mana yang paling dibutuhkan?',
          opsi: ['Laba rugi', 'Perubahan modal', 'Neraca'],
          item: [
            { t: 'Raka ingin tahu apakah bulan ini usahanya untung.', a: 0, why: 'Laba atau rugi terlihat di laporan laba rugi.' },
            { t: 'Bank ingin membandingkan harta dengan utang Laundry Kilat.', a: 2, why: 'Neraca menyajikan harta, utang, dan modal pada tanggal tertentu.' },
            { t: 'Raka ingin tahu berapa modalnya berkurang karena prive.', a: 1, why: 'Prive dan laba yang mengubah modal disajikan di laporan perubahan modal.' }
          ]
        }
      },
      {
        judul: 'Syarat informasi yang berguna',
        html: `<p>Laporan baru berguna kalau informasinya berkualitas. Menurut kerangka konseptual Ikatan Akuntan Indonesia (IAI), informasi harus <b>relevan</b> dan <b>menyajikan keadaan dengan jujur</b>. Kegunaannya makin tinggi bila juga dapat dibandingkan, dapat diverifikasi, tepat waktu, dan dapat dipahami. Ketuk kartunya.</p>`,
        flip: [
          { d: 'Relevan', b: 'Membantu pembaca mengambil keputusan. Contoh: data biaya listrik per bulan membantu Raka memutuskan untuk berhemat.' },
          { d: 'Penyajian jujur (andal)', b: 'Sesuai kenyataan, lengkap, netral, dan bebas dari kesalahan penting. Kas di laporan harus sama dengan uang yang benar-benar ada.' },
          { d: 'Dapat dibandingkan', b: 'Bisa dibandingkan antarbulan atau antarusaha karena metodenya sama. Laporan Oktober bisa disandingkan dengan November.' },
          { d: 'Dapat diverifikasi', b: 'Orang lain yang memeriksa bukti akan sampai pada angka yang sama. Karena itu setiap catatan butuh nota atau kuitansi.' },
          { d: 'Tepat waktu', b: 'Tersedia saat dibutuhkan. Laporan Oktober yang baru selesai bulan Maret sudah terlambat untuk keputusan.' },
          { d: 'Dapat dipahami', b: 'Disajikan jelas dan ringkas sehingga mudah dimengerti pembaca yang punya pengetahuan wajar tentang bisnis.' }
        ]
      },
      {
        judul: 'Rangkuman Misi 2',
        html: `<ul class="rangkum">
<li>Pengguna <b>internal</b> mengelola usaha: pemilik dan manajer.</li>
<li>Pengguna <b>eksternal</b> menilai dari luar: investor, kreditur, pemasok, pelanggan, pemerintah, dan masyarakat.</li>
<li>Tiga laporan utama usaha jasa: <b>laba rugi</b> (untung atau rugi), <b>perubahan modal</b> (perubahan modal), dan <b>neraca</b> (harta, utang, modal).</li>
<li>Laba mengalir ke laporan perubahan modal, lalu modal akhir mengalir ke neraca.</li>
<li>Informasi berguna jika relevan dan disajikan jujur, serta dapat dibandingkan, diverifikasi, tepat waktu, dan dapat dipahami.</li>
</ul>
<div class="istilah"><b>Istilah penting</b><span>Pihak internal</span><span>Pihak eksternal</span><span>Kreditur</span><span>Investor</span><span>Laporan laba rugi</span><span>Neraca</span><span>Relevan</span><span>Dapat dibandingkan</span></div>`
      }
    ],
    aktivitas: {
      tipe: 'pilah',
      judul: 'Pilah para pengguna informasi',
      petunjuk: 'Tentukan apakah setiap pihak termasuk pengguna internal atau eksternal.',
      wadah: ['Internal', 'Eksternal'],
      item: [
        { t: 'Raka, pemilik Laundry Kilat', a: 0, why: 'Pemilik butuh info laba dan perkembangan usaha.' },
        { t: 'Bank pemberi pinjaman', a: 1, why: 'Kreditur menilai kemampuan membayar pinjaman dan bunga.' },
        { t: 'Manajer operasional', a: 0, why: 'Manajer memakai info untuk merencanakan dan mengendalikan biaya.' },
        { t: 'Toko pemasok deterjen', a: 1, why: 'Pemasok menilai kemampuan membayar utang dagang.' },
        { t: 'Kantor pajak', a: 1, why: 'Pemerintah memakai info untuk menetapkan pajak.' },
        { t: 'Calon investor', a: 1, why: 'Investor menilai keuntungan dan risiko.' },
        { t: 'Direktur perusahaan', a: 0, why: 'Pimpinan mengambil keputusan dari dalam perusahaan.' },
        { t: 'Pelanggan tetap', a: 1, why: 'Pelanggan ingin layanan terus berlanjut.' },
        { t: 'Warga sekitar', a: 1, why: 'Masyarakat melihat kontribusi usaha bagi lingkungan.' }
      ]
    }
  },

  /* =============== MISI 3 =============== */
  {
    id: 'm3', no: 3, judul: 'Prinsip & Konsep Dasar',
    ringkas: 'Asumsi dasar, prinsip akuntansi, dan basis akrual',
    tp: 2,
    bagian: [
      {
        judul: 'Mengapa perlu aturan main?',
        html: `<div class="analogi"><b>Seperti sepak bola</b><p>Bayangkan setiap tim memakai aturan sendiri: ada yang boleh memegang bola, ada yang gawangnya lebih lebar. Skornya tidak bisa dibandingkan dan pertandingan jadi kacau.</p></div>
<p>Akuntansi juga begitu. Kalau setiap usaha mencatat dengan caranya sendiri, laporannya tidak bisa dibandingkan dan sulit dipercaya. Karena itu akuntansi memakai <b>asumsi dasar</b> dan <b>prinsip</b> yang sama.</p>
<p>Di Indonesia, Standar Akuntansi Keuangan (SAK) disusun oleh <b>Ikatan Akuntan Indonesia (IAI)</b>. Ada empat pilar:</p>
<ul class="dots">
<li><b>SAK Internasional</b>, mengadopsi standar internasional (IFRS) secara penuh.</li>
<li><b>SAK Indonesia</b>, terutama untuk perusahaan yang sahamnya diperdagangkan di bursa.</li>
<li><b>SAK Indonesia untuk Entitas Privat</b>, menggantikan SAK ETAP sejak 1 Januari 2025.</li>
<li><b>SAK EMKM</b>, untuk usaha mikro, kecil, dan menengah seperti Laundry Kilat.</li>
</ul>
<p class="note">Instansi pemerintah memakai standar tersendiri, yaitu Standar Akuntansi Pemerintahan (SAP).</p>`,
        pemantik: {
          tanya: 'Apa yang terjadi jika Laundry Kilat mencatat mesin cuci dengan harga beli, sedangkan laundry lain mencatat dengan harga jual hari ini?',
          jawab: 'Laporan keduanya tidak bisa dibandingkan dengan adil. Pembaca bisa salah menilai usaha mana yang lebih besar. Karena itu semua usaha perlu prinsip yang sama, misalnya prinsip biaya historis.'
        }
      },
      {
        judul: 'Empat asumsi dasar',
        html: `<p><b>Asumsi</b> adalah anggapan dasar yang diterima begitu saja ketika mencatat, tanpa perlu dibuktikan setiap kali. Ketuk kartu untuk melihat contohnya di Laundry Kilat.</p>`,
        flip: [
          { d: 'Kesatuan usaha (entitas ekonomi)', b: 'Keuangan usaha terpisah dari keuangan pribadi pemilik. Uang belanja rumah Raka tidak boleh dicatat sebagai beban laundry.' },
          { d: 'Kelangsungan usaha (going concern)', b: 'Usaha dianggap akan terus beroperasi. Karena itu mesin cuci disusutkan selama umur pakainya, tidak langsung dianggap habis.' },
          { d: 'Periode akuntansi', b: 'Umur usaha dibagi ke periode tertentu supaya kinerja bisa diukur. Laundry Kilat menyusun laporan setiap akhir bulan.' },
          { d: 'Unit moneter', b: 'Transaksi dicatat dalam satuan uang (rupiah) dan nilai rupiah dianggap stabil. Keramahan pelayan tidak dicatat karena tidak bisa diukur dengan uang.' }
        ]
      },
      {
        judul: 'Prinsip saat mencatat',
        html: `<p>Tiga prinsip ini menjawab pertanyaan <b>"berapa yang dicatat dan atas dasar apa?"</b> Ketuk setiap kartu.</p>`,
        flip: [
          { d: 'Biaya historis (harga perolehan)', b: 'Aset dicatat sebesar harga saat diperoleh, termasuk ongkos sampai siap pakai. Mesin cuci dibeli Rp4.000.000, dicatat Rp4.000.000 walau harga pasar naik.' },
          { d: 'Objektivitas', b: 'Pencatatan harus didukung bukti yang dapat diperiksa, seperti nota dan kuitansi. Pengeluaran tanpa bukti tidak dicatat.' },
          { d: 'Materialitas', b: 'Jumlah kecil yang tidak memengaruhi keputusan boleh disederhanakan. Gunting Rp15.000 langsung dicatat sebagai beban, tidak sebagai peralatan.' }
        ]
      },
      {
        judul: 'Prinsip saat melaporkan',
        html: `<p>Empat prinsip ini menjawab pertanyaan <b>"kapan dicatat dan bagaimana menyajikannya?"</b></p>`,
        flip: [
          { d: 'Pengakuan pendapatan', b: 'Pendapatan diakui saat jasa selesai diberikan, bukan saat uang diterima. Cucian selesai 30 Oktober menjadi pendapatan Oktober.' },
          { d: 'Penandingan (matching)', b: 'Beban dicatat pada periode yang sama dengan pendapatan yang dihasilkannya. Gaji Oktober menjadi beban Oktober walau dibayar 1 November.' },
          { d: 'Konsistensi', b: 'Metode yang sama dipakai dari periode ke periode supaya laporan bisa dibandingkan. Jika berganti metode, alasannya harus diungkapkan.' },
          { d: 'Pengungkapan penuh', b: 'Semua informasi penting disajikan, termasuk lewat catatan atas laporan keuangan, misalnya perkara hukum yang belum selesai.' }
        ]
      },
      {
        judul: 'Basis kas atau basis akrual?',
        html: `<p>Ada dua cara menentukan <b>kapan</b> pendapatan dan beban dicatat:</p>
<div class="duo">
<div><b>Basis kas</b><span>Dicatat saat uang <b>diterima atau dibayar</b>.</span></div>
<div><b>Basis akrual</b><span>Dicatat saat <b>terjadi</b> (jasa selesai, beban terpakai), walau uangnya belum berpindah.</span></div>
</div>
<div class="contoh"><b>Contoh soal</b><p>Cucian asrama mahasiswa senilai Rp1.800.000 selesai 30 Oktober dan baru dibayar 5 November. Gaji karyawan Oktober Rp1.200.000 dibayar 1 November. Berapa pendapatan dan beban bulan Oktober?</p>
<div class="tbl-wrap"><table><thead><tr><th>Bulan Oktober</th><th class="r">Basis kas</th><th class="r">Basis akrual</th></tr></thead><tbody>
<tr><td>Pendapatan dari asrama</td><td class="r num">Rp0</td><td class="r num">Rp1.800.000</td></tr>
<tr><td>Beban gaji</td><td class="r num">Rp0</td><td class="r num">Rp1.200.000</td></tr>
</tbody></table></div>
<p style="margin:0">Basis akrual menunjukkan kinerja Oktober yang sebenarnya. Inilah inti prinsip <b>pengakuan pendapatan</b> dan <b>penandingan</b>.</p></div>
<p class="note">SAK, termasuk SAK EMKM, mewajibkan basis akrual untuk menyusun laporan keuangan.</p>`,
        cek: {
          judul: 'Cek cepat: basis kas atau akrual?',
          opsi: ['Basis kas', 'Basis akrual'],
          item: [
            { t: 'Pendapatan dicatat saat cucian selesai diserahkan, walau pelanggan bayar minggu depan.', a: 1, why: 'Dicatat saat jasa selesai, bukan saat uang diterima.' },
            { t: 'Beban listrik baru dicatat ketika tagihannya dibayar.', a: 0, why: 'Dicatat saat uang keluar.' },
            { t: 'Gaji Oktober yang dibayar 1 November dicatat sebagai beban Oktober.', a: 1, why: 'Beban dicatat di periode terjadinya, sesuai prinsip penandingan.' }
          ]
        }
      },
      {
        judul: 'Jurus kata kunci',
        html: `<p>Soal kasus prinsip akuntansi lebih mudah dikerjakan jika kamu mencari <b>kata kuncinya</b>:</p>
<div class="tbl-wrap"><table><thead><tr><th>Jika soal menyebut…</th><th>Kemungkinan jawabannya</th></tr></thead><tbody>
<tr><td>uang pribadi, keperluan rumah, prive</td><td><b>Kesatuan usaha</b></td></tr>
<tr><td>terus beroperasi, disusutkan selama umur pakai</td><td><b>Kelangsungan usaha</b></td></tr>
<tr><td>setiap bulan, setiap tahun, akhir periode</td><td><b>Periode akuntansi</b></td></tr>
<tr><td>rupiah, tidak bisa diukur dengan uang</td><td><b>Unit moneter</b></td></tr>
<tr><td>harga beli, harga perolehan, walau harga pasar naik</td><td><b>Biaya historis</b></td></tr>
<tr><td>nota, kuitansi, bukti, bisa diperiksa</td><td><b>Objektivitas</b></td></tr>
<tr><td>jumlah kecil, langsung dibebankan</td><td><b>Materialitas</b></td></tr>
<tr><td>jasa selesai, kapan pendapatan dicatat</td><td><b>Pengakuan pendapatan</b></td></tr>
<tr><td>beban dan pendapatan di periode yang sama</td><td><b>Penandingan</b></td></tr>
<tr><td>metode sama setiap tahun, bisa dibandingkan</td><td><b>Konsistensi</b></td></tr>
<tr><td>catatan atas laporan, diungkapkan</td><td><b>Pengungkapan penuh</b></td></tr>
</tbody></table></div>`,
        cek: {
          judul: 'Latihan jurus kata kunci',
          opsi: ['Kesatuan usaha', 'Biaya historis', 'Materialitas'],
          item: [
            { t: 'Raka membayar SPP adiknya dari kas laundry. Akuntan mencatatnya sebagai prive.', a: 0, why: 'Kata kunci "keperluan pribadi" → kesatuan usaha.' },
            { t: 'Setrika dicatat sebesar harga belinya, Rp1.200.000, walau kini harganya Rp1.500.000.', a: 1, why: 'Kata kunci "harga beli walau harga pasar naik" → biaya historis.' },
            { t: 'Stapler seharga Rp20.000 langsung dicatat sebagai beban.', a: 2, why: 'Kata kunci "jumlah kecil, langsung dibebankan" → materialitas.' }
          ]
        }
      },
      {
        judul: 'Rangkuman Misi 3',
        html: `<ul class="rangkum">
<li>Standar akuntansi di Indonesia disusun IAI. Usaha mikro, kecil, dan menengah memakai <b>SAK EMKM</b>.</li>
<li><b>Asumsi dasar:</b> kesatuan usaha, kelangsungan usaha, periode akuntansi, unit moneter.</li>
<li><b>Prinsip mencatat:</b> biaya historis, objektivitas, materialitas.</li>
<li><b>Prinsip melaporkan:</b> pengakuan pendapatan, penandingan, konsistensi, pengungkapan penuh.</li>
<li>Laporan keuangan disusun dengan <b>basis akrual</b>: dicatat saat terjadi, bukan saat uang berpindah.</li>
<li>Kerjakan soal kasus dengan mencari kata kuncinya.</li>
</ul>
<div class="istilah"><b>Istilah penting</b><span>SAK EMKM</span><span>Asumsi</span><span>Going concern</span><span>Biaya historis</span><span>Materialitas</span><span>Penandingan</span><span>Basis akrual</span><span>Prive</span></div>`
      }
    ],
    aktivitas: {
      tipe: 'kasus',
      judul: 'Kasus di Laundry Kilat',
      petunjuk: 'Baca setiap kasus, lalu pilih asumsi atau prinsip yang berlaku. Gunakan jurus kata kunci.',
      item: [
        { t: 'Raka membayar tagihan listrik rumahnya memakai uang kas laundry. Akuntan menolak mencatatnya sebagai beban usaha.', opsi: ['Kesatuan usaha', 'Konsistensi', 'Materialitas', 'Periode akuntansi'], a: 0, why: 'Keuangan pribadi harus dipisahkan dari keuangan usaha. Pengeluaran itu dicatat sebagai prive.' },
        { t: 'Setrika uap dibeli Rp1.200.000. Setahun kemudian harga pasarnya Rp1.500.000, tetapi catatan tidak diubah.', opsi: ['Unit moneter', 'Biaya historis', 'Pengungkapan penuh', 'Objektivitas'], a: 1, why: 'Aset dicatat sebesar harga perolehannya.' },
        { t: 'Cucian asrama selesai tanggal 30 Oktober, dibayar tanggal 5 November. Pendapatan dicatat bulan Oktober.', opsi: ['Penandingan', 'Kelangsungan usaha', 'Pengakuan pendapatan', 'Konsistensi'], a: 2, why: 'Pendapatan diakui saat jasa selesai diberikan.' },
        { t: 'Laundry Kilat selalu memakai metode garis lurus untuk menyusutkan mesin cuci setiap tahun.', opsi: ['Konsistensi', 'Materialitas', 'Objektivitas', 'Unit moneter'], a: 0, why: 'Metode yang sama dipakai terus supaya laporan dapat dibandingkan.' },
        { t: 'Pembelian gunting Rp15.000 langsung dicatat sebagai beban, tidak dimasukkan ke akun peralatan.', opsi: ['Biaya historis', 'Materialitas', 'Kesatuan usaha', 'Pengungkapan penuh'], a: 1, why: 'Jumlahnya kecil dan tidak memengaruhi keputusan.' },
        { t: 'Setiap transaksi baru dicatat jika ada nota, kuitansi, atau bukti transfer.', opsi: ['Periode akuntansi', 'Penandingan', 'Objektivitas', 'Kelangsungan usaha'], a: 2, why: 'Pencatatan harus didukung bukti yang dapat diperiksa.' }
      ]
    }
  },

  /* =============== MISI 4 =============== */
  {
    id: 'm4', no: 4, judul: 'Persamaan Dasar Akuntansi',
    ringkas: 'Harta = Utang + Modal, selalu seimbang',
    tp: 3,
    bagian: [
      {
        judul: 'Tiga unsur utama',
        html: `<dl class="who">
<dt>Harta (aset)</dt><dd>Sumber daya milik usaha yang bermanfaat di masa depan: kas, piutang usaha, perlengkapan, peralatan, kendaraan, gedung.</dd>
<dt>Utang (kewajiban)</dt><dd>Kewajiban kepada pihak lain yang harus dibayar: utang usaha, utang bank, utang gaji.</dd>
<dt>Modal (ekuitas)</dt><dd>Hak pemilik atas harta usaha setelah dikurangi utang.</dd>
</dl>
<div class="analogi"><b>Contoh sehari-hari</b><p>Dina membeli motor seharga <b>Rp20 juta</b>. Uangnya <b>Rp15 juta</b> dari tabungan sendiri dan <b>Rp5 juta</b> pinjaman dari bank.<br>Motor (harta) Rp20 juta = pinjaman (utang) Rp5 juta + uang sendiri (modal) Rp15 juta.</p></div>`,
        pemantik: {
          tanya: 'Kamu membeli HP Rp3.000.000: Rp2.000.000 dari tabunganmu dan Rp1.000.000 pinjam dari kakak. Berapa harta, utang, dan modalmu?',
          jawab: 'Harta (HP) Rp3.000.000. Utang (ke kakak) Rp1.000.000. Modal (uang sendiri) Rp2.000.000. Harta Rp3.000.000 = Utang Rp1.000.000 + Modal Rp2.000.000.'
        }
      },
      {
        judul: 'Rumusnya selalu seimbang',
        html: `<div class="eq" aria-label="Harta sama dengan Utang ditambah Modal"><span>Harta</span><i>=</i><span>Utang</span><i>+</i><span>Modal</span></div>
<p>Sisi kiri menunjukkan <b>apa yang dimiliki</b> usaha. Sisi kanan menunjukkan <b>dari mana asalnya</b>: dari pinjaman pihak lain (utang) atau dari pemilik (modal).</p>
<p>Karena setiap harta pasti punya asal, kedua sisi <b>selalu sama besar</b>. Setiap transaksi memengaruhi minimal dua hal sehingga persamaan tetap seimbang. Ini disebut <b>prinsip dualitas</b>.</p>
<p>Rumusnya bisa dibalik sesuai yang dicari:</p>
<ul class="dots"><li>Modal = Harta − Utang</li><li>Utang = Harta − Modal</li></ul>`,
        pemantik: {
          tanya: 'Harta Laundry Kilat Rp30.000.000 dan utangnya Rp6.000.000. Berapa modal Raka?',
          jawab: 'Modal = Harta − Utang = Rp30.000.000 − Rp6.000.000 = Rp24.000.000.'
        }
      },
      {
        judul: 'Pendapatan, beban, dan prive',
        html: `<div class="eq eq-sm"><span>Harta</span><i>=</i><span>Utang</span><i>+</i><span>Modal + Pendapatan − Beban − Prive</span></div>
<ul class="dots">
<li><b>Pendapatan</b> menambah modal. Contoh: menerima uang jasa cuci.</li>
<li><b>Beban</b> mengurangi modal. Contoh: membayar listrik, gaji, sewa.</li>
<li><b>Prive</b> adalah pengambilan harta usaha oleh pemilik untuk keperluan pribadi. Prive mengurangi modal, tetapi <b>bukan beban</b>, karena tidak dipakai untuk menjalankan usaha.</li>
</ul>
<div class="tahu"><b>Jebakan yang sering muncul</b><p>Uang masuk belum tentu pendapatan, dan uang keluar belum tentu beban:</p>
<ul class="dots"><li>Menerima <b>pinjaman bank</b>: kas naik, tetapi utang juga naik. Bukan pendapatan.</li><li>Pemilik <b>menyetor modal</b>: modal naik, tetapi bukan pendapatan.</li><li><b>Membeli peralatan</b>: kas turun, tetapi peralatan naik. Bukan beban.</li><li><b>Membayar utang</b>: kas turun, utang turun. Bukan beban.</li></ul></div>`,
        cek: {
          judul: 'Cek cepat: pendapatan, beban, prive, atau bukan?',
          opsi: ['Pendapatan', 'Beban', 'Prive', 'Bukan ketiganya'],
          item: [
            { t: 'Laundry Kilat menerima pinjaman dari bank Rp5.000.000.', a: 3, why: 'Kas naik dan utang naik. Pinjaman bukan pendapatan.' },
            { t: 'Raka mengambil kas usaha untuk membayar SPP adiknya.', a: 2, why: 'Untuk keperluan pribadi pemilik, jadi prive.' },
            { t: 'Membayar gaji karyawan bulan ini.', a: 1, why: 'Biaya untuk menjalankan usaha, jadi beban.' },
            { t: 'Menerima uang jasa cuci kiloan dari pelanggan.', a: 0, why: 'Hasil dari menjual jasa, jadi pendapatan.' }
          ]
        }
      },
      {
        judul: 'Empat langkah menganalisis transaksi',
        html: `<p>Gunakan empat langkah ini untuk setiap transaksi:</p>
<ol class="steps">
<li><b>Akun apa yang terpengaruh?</b><span>Cari minimal dua akun.</span></li>
<li><b>Masuk kelompok apa?</b><span>Harta, utang, atau modal (pendapatan, beban, dan prive masuk ke modal).</span></li>
<li><b>Bertambah atau berkurang, berapa?</b><span>Tulis tanda + atau − dan jumlahnya.</span></li>
<li><b>Cek keseimbangan</b><span>Perubahan sisi kiri harus sama dengan perubahan sisi kanan.</span></li>
</ol>
<div class="contoh"><b>Contoh 1</b><p><i>4 Oktober: membeli deterjen dan plastik Rp1.500.000 secara kredit.</i></p>
<ol class="langkah"><li>Akun: Perlengkapan dan Utang Usaha.</li><li>Perlengkapan = harta. Utang Usaha = utang.</li><li>Perlengkapan +Rp1.500.000. Utang Usaha +Rp1.500.000.</li><li>Harta +1.500.000 = Utang +1.500.000 + Modal 0. <b>Seimbang ✓</b></li></ol></div>
<div class="contoh"><b>Contoh 2</b><p><i>20 Oktober: membayar listrik dan air Rp600.000.</i></p>
<ol class="langkah"><li>Akun: Kas dan Beban Listrik.</li><li>Kas = harta. Beban listrik mengurangi modal.</li><li>Kas −Rp600.000. Modal −Rp600.000.</li><li>Harta −600.000 = Utang 0 + Modal −600.000. <b>Seimbang ✓</b></li></ol></div>`
      },
      {
        judul: 'Tabel persamaan dasar',
        html: `<p>Akuntan biasanya mencatat pengaruh transaksi dalam <b>tabel persamaan dasar</b>. Ini lima transaksi pertama Laundry Kilat (dalam ribuan rupiah):</p>
<div class="tbl-wrap"><table class="ptbl"><thead><tr><th rowspan="2">Tgl</th><th colspan="4" class="grp">Harta</th><th class="eqc" rowspan="2">=</th><th class="grp">Utang</th><th class="eqc" rowspan="2">+</th><th class="grp">Modal</th></tr>
<tr><th class="r">Kas</th><th class="r">Piutang</th><th class="r">Perlk.</th><th class="r">Peralatan</th><th class="r">Utang Usaha</th><th class="r">Modal</th></tr></thead><tbody>
<tr><td>1/10</td><td class="r up">+20.000</td><td></td><td></td><td></td><td></td><td></td><td></td><td class="r up">+20.000</td></tr>
<tr><td>2/10</td><td class="r dn">−8.000</td><td></td><td></td><td class="r up">+8.000</td><td></td><td></td><td></td><td></td></tr>
<tr class="saldo"><td>Saldo</td><td class="r">12.000</td><td class="r">0</td><td class="r">0</td><td class="r">8.000</td><td></td><td class="r">0</td><td></td><td class="r">20.000</td></tr>
<tr><td>4/10</td><td></td><td></td><td class="r up">+1.500</td><td></td><td></td><td class="r up">+1.500</td><td></td><td></td></tr>
<tr><td>10/10</td><td class="r up">+3.200</td><td></td><td></td><td></td><td></td><td></td><td></td><td class="r up">+3.200</td></tr>
<tr><td>15/10</td><td></td><td class="r up">+1.800</td><td></td><td></td><td></td><td></td><td></td><td class="r up">+1.800</td></tr>
<tr class="saldo"><td>Saldo</td><td class="r">15.200</td><td class="r">1.800</td><td class="r">1.500</td><td class="r">8.000</td><td></td><td class="r">1.500</td><td></td><td class="r">25.000</td></tr>
</tbody></table></div>
<p class="tiny muted" style="margin-top:-4px">Di HP, geser tabel ke kiri untuk melihat kolom utang dan modal.</p>
<p>Periksa saldo akhir: harta 15.200 + 1.800 + 1.500 + 8.000 = <b>26.500</b>. Utang + modal 1.500 + 25.000 = <b>26.500</b>. Seimbang!</p>
<p class="note">Kamu akan melanjutkan lima transaksi berikutnya di aktivitas akhir misi ini sambil melihat timbangannya bergerak.</p>`
      },
      {
        judul: 'Pola pengaruh transaksi',
        html: `<table class="pola"><thead><tr><th>Transaksi</th><th>Harta</th><th>Utang</th><th>Modal</th></tr></thead><tbody>
<tr><td>Setoran modal, pendapatan tunai</td><td class="up">naik</td><td>tetap</td><td class="up">naik</td></tr>
<tr><td>Beli secara kredit, pinjam ke bank</td><td class="up">naik</td><td class="up">naik</td><td>tetap</td></tr>
<tr><td>Beli tunai, terima pelunasan piutang</td><td>tetap*</td><td>tetap</td><td>tetap</td></tr>
<tr><td>Membayar utang</td><td class="dn">turun</td><td class="dn">turun</td><td>tetap</td></tr>
<tr><td>Membayar beban, prive</td><td class="dn">turun</td><td>tetap</td><td class="dn">turun</td></tr>
<tr><td>Beban belum dibayar</td><td>tetap</td><td class="up">naik</td><td class="dn">turun</td></tr>
</tbody></table>
<p class="note">* Satu harta bertambah dan harta lain berkurang sama besar, sehingga total harta tetap.</p>`
      },
      {
        judul: 'Rangkuman Misi 4',
        html: `<ul class="rangkum">
<li><b>Harta = Utang + Modal</b>, selalu seimbang karena setiap harta punya asal.</li>
<li>Pendapatan menambah modal. Beban dan prive mengurangi modal. Prive bukan beban.</li>
<li>Pinjaman bank dan setoran modal bukan pendapatan. Membeli aset dan membayar utang bukan beban.</li>
<li>Analisis transaksi dengan empat langkah: akun, kelompok, tambah/kurang, cek keseimbangan.</li>
<li>Setiap transaksi memengaruhi minimal dua hal (prinsip dualitas).</li>
</ul>
<div class="istilah"><b>Istilah penting</b><span>Harta</span><span>Utang</span><span>Modal</span><span>Pendapatan</span><span>Beban</span><span>Prive</span><span>Piutang</span><span>Dualitas</span></div>`
      }
    ],
    aktivitas: {
      tipe: 'timbangan',
      judul: 'Pembukuan Laundry Kilat, Oktober 2026',
      petunjuk: 'Untuk setiap transaksi, tebak dulu apakah total Harta, Utang, dan Modal naik, turun, atau tetap. Lalu lihat timbangannya.'
    }
  }
];

/* Kamus istilah (dipakai di halaman Kamus) */
const ISTILAH = [
  ['Akuntansi', 'Proses mengolah transaksi keuangan menjadi informasi untuk mengambil keputusan.'],
  ['Transaksi keuangan', 'Kejadian yang mengubah keuangan usaha, dapat diukur dengan uang, dan didukung bukti.'],
  ['Bukti transaksi', 'Dokumen yang menjadi dasar pencatatan, misalnya nota, faktur, dan kuitansi.'],
  ['Nota kontan', 'Bukti pembelian atau penjualan secara tunai.'],
  ['Faktur', 'Bukti pembelian atau penjualan secara kredit.'],
  ['Kuitansi', 'Tanda terima uang yang ditandatangani penerima.'],
  ['Memo internal', 'Perintah tertulis pimpinan kepada bagian akuntansi.'],
  ['Jurnal', 'Catatan transaksi menurut urutan tanggal.'],
  ['Buku besar', 'Kumpulan akun yang menampung transaksi sejenis.'],
  ['Neraca saldo', 'Daftar saldo semua akun buku besar pada akhir periode.'],
  ['Siklus akuntansi', 'Rangkaian kegiatan akuntansi yang berulang setiap periode.'],
  ['Pihak internal', 'Pengguna informasi yang mengelola usaha: pemilik dan manajer.'],
  ['Pihak eksternal', 'Pengguna informasi dari luar usaha: investor, kreditur, pemasok, pelanggan, pemerintah, masyarakat.'],
  ['Kreditur', 'Pihak yang memberi pinjaman, misalnya bank atau koperasi.'],
  ['Investor', 'Pihak yang menanamkan modal untuk mendapat keuntungan.'],
  ['Laporan laba rugi', 'Laporan pendapatan dan beban yang menunjukkan laba atau rugi satu periode.'],
  ['Laporan perubahan modal', 'Laporan yang menunjukkan perubahan modal karena laba, rugi, dan prive.'],
  ['Neraca', 'Laporan posisi keuangan: harta, utang, dan modal pada tanggal tertentu.'],
  ['SAK EMKM', 'Standar akuntansi untuk usaha mikro, kecil, dan menengah yang disusun IAI.'],
  ['Kesatuan usaha', 'Asumsi bahwa keuangan usaha terpisah dari keuangan pribadi pemilik.'],
  ['Kelangsungan usaha', 'Asumsi bahwa usaha akan terus beroperasi (going concern).'],
  ['Periode akuntansi', 'Asumsi bahwa umur usaha dibagi ke periode, misalnya bulan atau tahun.'],
  ['Unit moneter', 'Asumsi bahwa transaksi dicatat dalam satuan uang.'],
  ['Biaya historis', 'Prinsip mencatat aset sebesar harga perolehannya.'],
  ['Objektivitas', 'Prinsip bahwa pencatatan harus didukung bukti yang dapat diperiksa.'],
  ['Materialitas', 'Prinsip bahwa jumlah kecil yang tidak memengaruhi keputusan boleh disederhanakan.'],
  ['Pengakuan pendapatan', 'Prinsip mencatat pendapatan saat jasa selesai diberikan.'],
  ['Penandingan', 'Prinsip mencatat beban pada periode yang sama dengan pendapatan yang dihasilkannya.'],
  ['Konsistensi', 'Prinsip memakai metode yang sama dari periode ke periode.'],
  ['Pengungkapan penuh', 'Prinsip menyajikan semua informasi penting, termasuk dalam catatan laporan.'],
  ['Basis akrual', 'Pencatatan pendapatan dan beban saat terjadi, bukan saat uang berpindah.'],
  ['Basis kas', 'Pencatatan pendapatan dan beban saat uang diterima atau dibayar.'],
  ['Harta (aset)', 'Sumber daya milik usaha yang bermanfaat di masa depan.'],
  ['Utang (kewajiban)', 'Kewajiban usaha kepada pihak lain yang harus dibayar.'],
  ['Modal (ekuitas)', 'Hak pemilik atas harta usaha setelah dikurangi utang.'],
  ['Piutang', 'Hak menagih uang kepada pelanggan yang belum membayar.'],
  ['Pendapatan', 'Penghasilan dari menjual jasa atau barang; menambah modal.'],
  ['Beban', 'Biaya untuk menjalankan usaha; mengurangi modal.'],
  ['Prive', 'Pengambilan harta usaha oleh pemilik untuk keperluan pribadi; mengurangi modal, bukan beban.'],
  ['Dualitas', 'Setiap transaksi memengaruhi minimal dua hal sehingga persamaan tetap seimbang.']
];

/* Akun yang dipakai pada simulasi Misi 4 */
const AKUN_SIM = {
  harta: ['Kas', 'Piutang Usaha', 'Perlengkapan', 'Peralatan'],
  utang: ['Utang Usaha'],
  modal: ['Modal']
};

/* Transaksi Laundry Kilat — Oktober 2026 (Misi 4) */
const TRX_LAUNDRY = [
  { tgl: '1 Okt', t: 'Raka menyetor uang tunai Rp20.000.000 sebagai modal awal usaha.', ub: { 'Kas': 20000000, 'Modal': 20000000 }, ket: 'Setoran modal', arah: [1, 0, 1] },
  { tgl: '2 Okt', t: 'Membeli 2 mesin cuci seharga Rp8.000.000 secara tunai.', ub: { 'Kas': -8000000, 'Peralatan': 8000000 }, ket: 'Pertukaran harta', arah: [0, 0, 0] },
  { tgl: '4 Okt', t: 'Membeli deterjen, pewangi, dan plastik kemasan Rp1.500.000 secara kredit.', ub: { 'Perlengkapan': 1500000, 'Utang Usaha': 1500000 }, ket: 'Pembelian kredit', arah: [1, 1, 0] },
  { tgl: '10 Okt', t: 'Menerima pendapatan jasa laundry secara tunai Rp3.200.000.', ub: { 'Kas': 3200000, 'Modal': 3200000 }, ket: 'Pendapatan', arah: [1, 0, 1] },
  { tgl: '15 Okt', t: 'Menyelesaikan cucian asrama mahasiswa Rp1.800.000, dibayar bulan depan.', ub: { 'Piutang Usaha': 1800000, 'Modal': 1800000 }, ket: 'Pendapatan (piutang)', arah: [1, 0, 1] },
  { tgl: '20 Okt', t: 'Membayar tagihan listrik dan air Rp600.000.', ub: { 'Kas': -600000, 'Modal': -600000 }, ket: 'Beban', arah: [-1, 0, -1] },
  { tgl: '22 Okt', t: 'Membayar sebagian utang kepada toko deterjen Rp1.000.000.', ub: { 'Kas': -1000000, 'Utang Usaha': -1000000 }, ket: 'Pelunasan utang', arah: [-1, -1, 0] },
  { tgl: '25 Okt', t: 'Menerima pelunasan piutang dari asrama mahasiswa Rp1.800.000.', ub: { 'Kas': 1800000, 'Piutang Usaha': -1800000 }, ket: 'Pertukaran harta', arah: [0, 0, 0] },
  { tgl: '30 Okt', t: 'Membayar gaji karyawan Rp1.200.000.', ub: { 'Kas': -1200000, 'Modal': -1200000 }, ket: 'Beban', arah: [-1, 0, -1] },
  { tgl: '31 Okt', t: 'Raka mengambil uang kas Rp500.000 untuk keperluan pribadi.', ub: { 'Kas': -500000, 'Modal': -500000 }, ket: 'Prive', arah: [-1, 0, -1] }
];

/* ---------------- BANK SOAL KUIS ---------------- */
const SOAL = [
  // Misi 1
  { id: 'q01', m: 'm1', q: 'Pengertian akuntansi yang paling tepat adalah …', o: ['Proses mengidentifikasi, mencatat, menggolongkan, mengikhtisarkan, dan melaporkan transaksi keuangan untuk pengambilan keputusan', 'Kegiatan menghitung jumlah uang kas pada akhir bulan', 'Proses menyusun anggaran belanja perusahaan', 'Kegiatan memeriksa kejujuran karyawan'], a: 0, p: 'Akuntansi mencakup seluruh proses dari identifikasi transaksi sampai pelaporan dan penafsiran informasi keuangan.' },
  { id: 'q02', m: 'm1', q: 'Raka memeriksa nota pembelian deterjen sebelum mencatatnya. Kegiatan ini disebut …', o: ['Pencatatan', 'Identifikasi transaksi', 'Pelaporan', 'Penggolongan'], a: 1, p: 'Memeriksa bukti transaksi adalah tahap identifikasi, tahap pertama kegiatan akuntansi.' },
  { id: 'q03', m: 'm1', q: 'Urutan kegiatan akuntansi yang benar adalah …', o: ['Pencatatan → identifikasi → pelaporan → penggolongan', 'Identifikasi → pencatatan → penggolongan → pengikhtisaran → pelaporan', 'Identifikasi → pelaporan → pencatatan → penggolongan', 'Pencatatan → penggolongan → identifikasi → pelaporan'], a: 1, p: 'Bukti diidentifikasi dulu, lalu dicatat, digolongkan, diikhtisarkan, dan dilaporkan.' },
  { id: 'q04', m: 'm1', q: 'Akuntansi disebut "bahasa bisnis" karena …', o: ['Banyak memakai istilah bahasa Inggris', 'Menyampaikan informasi keuangan yang dipahami berbagai pihak untuk mengambil keputusan', 'Hanya digunakan oleh perusahaan besar', 'Wajib dipelajari oleh semua pengusaha'], a: 1, p: 'Laporan keuangan menjadi alat komunikasi kondisi usaha kepada banyak pihak.' },
  { id: 'q05', m: 'm1', q: 'Tujuan utama akuntansi adalah …', o: ['Memperbesar laba perusahaan', 'Mengurangi pajak yang dibayar', 'Menyediakan informasi keuangan yang berguna untuk pengambilan keputusan', 'Menghitung gaji karyawan'], a: 2, p: 'Akuntansi menghasilkan informasi; keputusan diambil oleh pemakainya.' },
  { id: 'q06', m: 'm1', q: 'Mengelompokkan semua transaksi kas ke dalam akun Kas di buku besar termasuk kegiatan …', o: ['Pengikhtisaran', 'Pelaporan', 'Penggolongan', 'Identifikasi'], a: 2, p: 'Penggolongan berarti memindahkan dan mengelompokkan transaksi ke akun-akun buku besar.' },
  { id: 'q27', m: 'm1', q: 'Bukti transaksi yang dibuat penjual saat menjual barang atau jasa secara kredit adalah …', o: ['Kuitansi', 'Faktur', 'Nota kontan', 'Memo internal'], a: 1, p: 'Faktur dipakai untuk transaksi kredit. Nota kontan untuk transaksi tunai, kuitansi sebagai tanda terima uang.' },
  { id: 'q28', m: 'm1', q: 'Bidang akuntansi yang khusus menghitung dan melaporkan kewajiban pajak usaha adalah …', o: ['Akuntansi biaya', 'Akuntansi perpajakan', 'Akuntansi anggaran', 'Pemeriksaan akuntansi'], a: 1, p: 'Akuntansi perpajakan menangani perhitungan, pembayaran, dan pelaporan pajak.' },
  { id: 'q29', m: 'm1', q: 'Kejadian berikut yang BUKAN transaksi keuangan adalah …', o: ['Membayar sewa ruko lewat transfer bank', 'Menerima pembayaran jasa lewat QRIS', 'Menerima pesanan cucian yang belum dikerjakan dan belum dibayar', 'Membeli setrika secara tunai'], a: 2, p: 'Pesanan yang belum dikerjakan dan belum dibayar belum mengubah keadaan keuangan usaha.' },
  // Misi 2
  { id: 'q07', m: 'm2', q: 'Pihak berikut yang termasuk pengguna informasi akuntansi internal adalah …', o: ['Bank', 'Manajer perusahaan', 'Kantor pajak', 'Pemasok'], a: 1, p: 'Manajer berada di dalam perusahaan dan ikut mengelola usaha.' },
  { id: 'q08', m: 'm2', q: 'Sebuah bank akan memberi pinjaman kepada Laundry Kilat. Informasi yang paling dibutuhkan bank adalah …', o: ['Jumlah pelanggan setia', 'Kemampuan usaha membayar kembali pinjaman beserta bunganya', 'Merek mesin cuci yang dipakai', 'Jam buka laundry'], a: 1, p: 'Kreditur berkepentingan atas pengembalian pinjaman dan bunga.' },
  { id: 'q09', m: 'm2', q: 'Pemerintah membutuhkan informasi akuntansi suatu usaha terutama untuk …', o: ['Menentukan harga jual jasa', 'Menetapkan besarnya pajak', 'Memilih karyawan baru', 'Menentukan lokasi usaha'], a: 1, p: 'Laporan keuangan menjadi dasar penetapan pajak dan kebijakan.' },
  { id: 'q10', m: 'm2', q: 'Seorang investor ingin menanamkan modal di Laundry Kilat. Informasi yang ia butuhkan adalah …', o: ['Tingkat keuntungan dan risiko investasi', 'Jadwal piket karyawan', 'Jenis deterjen yang dipakai', 'Nama-nama pelanggan'], a: 0, p: 'Investor menilai keuntungan dan risiko sebelum menanamkan modal.' },
  { id: 'q11', m: 'm2', q: 'Toko deterjen menjual secara kredit kepada Laundry Kilat. Toko itu perlu tahu …', o: ['Kemampuan Laundry Kilat membayar utang tepat waktu', 'Rencana liburan pemilik', 'Jumlah mesin cuci di toko lain', 'Desain logo laundry'], a: 0, p: 'Pemasok berkepentingan agar utang dagang dibayar tepat waktu.' },
  { id: 'q12', m: 'm2', q: 'Pasangan pihak yang SEMUANYA termasuk pengguna eksternal adalah …', o: ['Pemilik dan manajer', 'Kreditur dan pemerintah', 'Manajer dan investor', 'Pemilik dan pemasok'], a: 1, p: 'Kreditur dan pemerintah sama-sama berada di luar perusahaan.' },
  { id: 'q30', m: 'm2', q: 'Laporan keuangan yang menunjukkan apakah usaha memperoleh laba atau rugi dalam satu periode adalah …', o: ['Laporan posisi keuangan (neraca)', 'Laporan laba rugi', 'Laporan perubahan modal', 'Daftar gaji'], a: 1, p: 'Laporan laba rugi membandingkan pendapatan dengan beban dalam satu periode.' },
  { id: 'q31', m: 'm2', q: 'Laundry Kilat memakai metode yang sama setiap bulan sehingga laporan Oktober bisa disandingkan dengan November. Karakteristik informasi yang terpenuhi adalah …', o: ['Dapat dibandingkan', 'Dapat dipahami', 'Tepat waktu', 'Dapat diverifikasi'], a: 0, p: 'Metode yang sama membuat laporan antarperiode dapat dibandingkan.' },
  { id: 'q32', m: 'm2', t: 'angka', q: 'Pendapatan jasa Laundry Kilat bulan Oktober Rp5.000.000, beban listrik Rp600.000, dan beban gaji Rp1.200.000. Berapa laba bersihnya? (tulis angka saja)', a: 3200000, p: 'Laba bersih = 5.000.000 − (600.000 + 1.200.000) = Rp3.200.000.' },
  // Misi 3
  { id: 'q13', m: 'm3', q: 'Raka mencampur uang belanja rumah dengan uang kas laundry. Konsep yang dilanggar adalah …', o: ['Kesatuan usaha (entitas ekonomi)', 'Periode akuntansi', 'Unit moneter', 'Konsistensi'], a: 0, p: 'Keuangan usaha harus terpisah dari keuangan pribadi pemilik.' },
  { id: 'q14', m: 'm3', q: 'Mesin cuci dibeli Rp4.000.000, padahal harga pasarnya Rp4.500.000. Mesin itu dicatat sebesar …', o: ['Rp4.500.000, sesuai harga pasar', 'Rp4.000.000, sesuai prinsip biaya historis', 'Rp4.250.000, rata-rata keduanya', 'Tidak dicatat sampai dijual'], a: 1, p: 'Prinsip biaya historis: aset dicatat sebesar harga perolehannya.' },
  { id: 'q15', m: 'm3', q: 'Laundry Kilat menyusun laporan keuangan setiap akhir bulan. Hal ini sesuai dengan konsep …', o: ['Kelangsungan usaha', 'Objektivitas', 'Periode akuntansi', 'Materialitas'], a: 2, p: 'Umur usaha dibagi ke periode-periode untuk pelaporan.' },
  { id: 'q16', m: 'm3', q: 'Jasa laundry selesai tanggal 30 Oktober, pelanggan membayar 5 November. Pendapatan diakui pada …', o: ['5 November, saat uang diterima', '30 Oktober, sesuai prinsip pengakuan pendapatan', 'Akhir tahun', 'Saat pelanggan datang pertama kali'], a: 1, p: 'Pendapatan diakui saat jasa selesai diberikan (basis akrual).' },
  { id: 'q17', m: 'm3', q: 'Semua transaksi dicatat dalam satuan Rupiah. Ini sesuai dengan konsep …', o: ['Unit moneter', 'Penandingan', 'Pengungkapan penuh', 'Kesatuan usaha'], a: 0, p: 'Konsep unit moneter: pencatatan memakai satuan uang.' },
  { id: 'q18', m: 'm3', q: 'Usaha dianggap akan terus beroperasi dan tidak dibubarkan dalam waktu dekat. Ini disebut konsep …', o: ['Periode akuntansi', 'Kelangsungan usaha (going concern)', 'Biaya historis', 'Materialitas'], a: 1, p: 'Asumsi kelangsungan usaha menjadi dasar penyusutan aset tetap.' },
  { id: 'q33', m: 'm3', q: 'Pencatatan yang mengakui pendapatan saat jasa selesai dan beban saat terjadi, tanpa menunggu uang diterima atau dibayar, disebut …', o: ['Basis kas', 'Basis akrual', 'Biaya historis', 'Konsistensi'], a: 1, p: 'Basis akrual mencatat saat terjadi, bukan saat uang berpindah.' },
  { id: 'q34', m: 'm3', q: 'Gaji karyawan Oktober baru dibayar 1 November, tetapi dicatat sebagai beban Oktober bersama pendapatan Oktober. Prinsip yang diterapkan adalah …', o: ['Penandingan', 'Materialitas', 'Unit moneter', 'Objektivitas'], a: 0, p: 'Beban dicatat pada periode yang sama dengan pendapatan yang dihasilkannya.' },
  { id: 'q35', m: 'm3', q: 'Lembaga yang menyusun Standar Akuntansi Keuangan (SAK) di Indonesia adalah …', o: ['Bank Indonesia', 'Ikatan Akuntan Indonesia (IAI)', 'Otoritas Jasa Keuangan', 'Direktorat Jenderal Pajak'], a: 1, p: 'SAK disusun oleh dewan standar di bawah Ikatan Akuntan Indonesia.' },
  // Misi 4
  { id: 'q19', m: 'm4', q: 'Persamaan dasar akuntansi yang benar adalah …', o: ['Harta = Utang − Modal', 'Harta + Utang = Modal', 'Harta = Utang + Modal', 'Modal = Harta + Utang'], a: 2, p: 'Harta selalu sama dengan sumbernya: utang ditambah modal.' },
  { id: 'q20', m: 'm4', q: 'Membeli perlengkapan secara kredit Rp1.500.000 berakibat …', o: ['Harta bertambah, utang bertambah', 'Harta bertambah, modal bertambah', 'Harta berkurang, utang bertambah', 'Utang bertambah, modal berkurang'], a: 0, p: 'Perlengkapan (harta) bertambah dan utang usaha bertambah dengan jumlah yang sama.' },
  { id: 'q21', m: 'm4', q: 'Membayar beban listrik Rp600.000 secara tunai berakibat …', o: ['Harta berkurang, utang berkurang', 'Harta berkurang, modal berkurang', 'Harta tetap, modal berkurang', 'Harta bertambah, modal berkurang'], a: 1, p: 'Kas berkurang dan beban mengurangi modal.' },
  { id: 'q22', m: 'm4', t: 'angka', q: 'Harta sebuah usaha Rp50.000.000 dan utangnya Rp12.000.000. Berapa modalnya? (tulis angka saja)', a: 38000000, p: 'Modal = Harta − Utang = 50.000.000 − 12.000.000 = Rp38.000.000.' },
  { id: 'q23', m: 'm4', q: 'Membeli peralatan secara tunai Rp8.000.000 berakibat …', o: ['Total harta bertambah Rp8.000.000', 'Peralatan bertambah dan kas berkurang, total harta tetap', 'Modal berkurang Rp8.000.000', 'Utang bertambah Rp8.000.000'], a: 1, p: 'Terjadi pertukaran harta: peralatan naik, kas turun sama besar.' },
  { id: 'q24', m: 'm4', q: 'Pemilik mengambil uang kas untuk keperluan pribadi. Transaksi ini disebut … dan berakibat …', o: ['Beban; harta dan modal berkurang', 'Prive; harta dan modal berkurang', 'Prive; harta berkurang, utang bertambah', 'Pendapatan; harta dan modal bertambah'], a: 1, p: 'Prive mengurangi kas dan modal, tetapi bukan beban usaha.' },
  { id: 'q25', m: 'm4', t: 'angka', q: 'Modal awal Rp20.000.000, pendapatan Rp5.000.000, beban Rp1.800.000, dan prive Rp500.000. Berapa modal akhirnya? (tulis angka saja)', a: 22700000, p: 'Modal akhir = 20.000.000 + 5.000.000 − 1.800.000 − 500.000 = Rp22.700.000.' },
  { id: 'q26', m: 'm4', q: 'Menerima pelunasan piutang dari pelanggan Rp1.800.000 berakibat …', o: ['Kas bertambah, modal bertambah', 'Kas bertambah, piutang berkurang, total harta tetap', 'Kas bertambah, utang berkurang', 'Piutang bertambah, kas berkurang'], a: 1, p: 'Piutang berubah menjadi kas; total harta tidak berubah.' },
  { id: 'q36', m: 'm4', q: 'Laundry Kilat menerima pinjaman dari bank Rp5.000.000. Pengaruhnya terhadap persamaan dasar akuntansi adalah …', o: ['Harta bertambah, modal bertambah', 'Harta bertambah, utang bertambah', 'Harta tetap, utang bertambah', 'Utang bertambah, modal berkurang'], a: 1, p: 'Kas (harta) bertambah dan utang bank bertambah. Pinjaman bukan pendapatan.' }
];

/* ---------------- GAME 1: TIMBANGAN SEIMBANG ---------------- */
const GAME_TIMBANGAN = [
  { t: 'Pemilik menyetor uang tunai sebagai tambahan modal.', a: [1, 0, 1], why: 'Kas naik, modal naik.' },
  { t: 'Membeli peralatan secara tunai.', a: [0, 0, 0], why: 'Peralatan naik, kas turun. Total harta tetap.' },
  { t: 'Membeli perlengkapan secara kredit.', a: [1, 1, 0], why: 'Perlengkapan naik, utang usaha naik.' },
  { t: 'Menerima pendapatan jasa secara tunai.', a: [1, 0, 1], why: 'Kas naik, pendapatan menambah modal.' },
  { t: 'Membayar gaji karyawan.', a: [-1, 0, -1], why: 'Kas turun, beban gaji mengurangi modal.' },
  { t: 'Membayar utang kepada pemasok.', a: [-1, -1, 0], why: 'Kas turun, utang usaha turun.' },
  { t: 'Menerima pinjaman dari bank.', a: [1, 1, 0], why: 'Kas naik, utang bank naik.' },
  { t: 'Pemilik mengambil uang kas untuk keperluan pribadi.', a: [-1, 0, -1], why: 'Kas turun, prive mengurangi modal.' },
  { t: 'Menerima pelunasan piutang dari pelanggan.', a: [0, 0, 0], why: 'Kas naik, piutang turun. Total harta tetap.' },
  { t: 'Jasa selesai dikerjakan, pelanggan akan membayar bulan depan.', a: [1, 0, 1], why: 'Piutang naik, pendapatan menambah modal.' },
  { t: 'Membayar sewa tempat usaha untuk bulan ini.', a: [-1, 0, -1], why: 'Kas turun, beban sewa mengurangi modal.' },
  { t: 'Pemilik menyerahkan sepeda motor pribadinya untuk dipakai usaha sebagai modal.', a: [1, 0, 1], why: 'Kendaraan naik, modal naik.' },
  { t: 'Membeli mesin pengering secara kredit.', a: [1, 1, 0], why: 'Peralatan naik, utang usaha naik.' },
  { t: 'Membayar tagihan listrik dan air.', a: [-1, 0, -1], why: 'Kas turun, beban mengurangi modal.' },
  { t: 'Membeli perlengkapan secara tunai.', a: [0, 0, 0], why: 'Perlengkapan naik, kas turun. Total harta tetap.' },
  { t: 'Menerima tagihan iklan radio lokal bulan ini yang belum dibayar.', a: [0, 1, -1], why: 'Utang naik, beban iklan mengurangi modal. Harta tidak berubah.' }
];

/* ---------------- GAME 2: AKUN KILAT ---------------- */
const KATEGORI_AKUN = ['Harta', 'Utang', 'Modal', 'Pendapatan', 'Beban'];
const GAME_AKUN = [
  ['Kas', 0], ['Piutang Usaha', 0], ['Perlengkapan', 0], ['Peralatan', 0], ['Kendaraan', 0], ['Gedung', 0], ['Tanah', 0],
  ['Sewa Dibayar di Muka', 0, 'Sewa dibayar di muka adalah hak memakai tempat di masa depan, jadi termasuk harta.'],
  ['Asuransi Dibayar di Muka', 0, 'Hak perlindungan di masa depan, jadi termasuk harta.'],
  ['Utang Usaha', 1], ['Utang Bank', 1], ['Utang Gaji', 1],
  ['Pendapatan Diterima di Muka', 1, 'Uang sudah diterima tetapi jasa belum diberikan, jadi masih kewajiban (utang).'],
  ['Modal Raka', 2], ['Prive Raka', 2, 'Prive termasuk kelompok modal dan sifatnya mengurangi modal.'],
  ['Pendapatan Jasa Laundry', 3], ['Pendapatan Bunga', 3], ['Pendapatan Sewa', 3],
  ['Beban Gaji', 4], ['Beban Listrik dan Air', 4], ['Beban Sewa', 4], ['Beban Iklan', 4], ['Beban Perlengkapan', 4],
  ['Beban Penyusutan Peralatan', 4], ['Beban Asuransi', 4]
];

/* ---------------- GAME 3: DETEKTIF PRINSIP ---------------- */
const GAME_DETEKTIF = [
  { a: 'Kesatuan usaha', k: ['Ini tentang siapa pemilik uang.', 'Raka membayar tagihan listrik rumahnya memakai kas laundry.', 'Akuntan mencatatnya sebagai prive, bukan beban usaha.'] },
  { a: 'Biaya historis', k: ['Ini tentang angka yang dicatat untuk sebuah aset.', 'Mesin cuci dibeli Rp4.000.000. Kini harga pasarnya Rp5.000.000.', 'Catatan tetap Rp4.000.000, sesuai harga saat dibeli.'] },
  { a: 'Pengakuan pendapatan', k: ['Ini tentang kapan sesuatu dicatat.', 'Uang dari pelanggan baru diterima bulan depan.', 'Cucian sudah selesai diserahkan, jadi pendapatan dicatat bulan ini.'] },
  { a: 'Penandingan', k: ['Ini tentang pasangan dalam satu periode.', 'Gaji karyawan Oktober baru dibayar 1 November.', 'Gaji itu dicatat sebagai beban Oktober, bersama pendapatan Oktober.'] },
  { a: 'Periode akuntansi', k: ['Ini tentang waktu.', 'Umur usaha dipotong-potong menjadi bagian.', 'Laundry Kilat membuat laporan setiap akhir bulan.'] },
  { a: 'Konsistensi', k: ['Ini tentang kebiasaan yang tidak berubah.', 'Laporan tahun ini harus bisa dibandingkan dengan tahun lalu.', 'Metode penyusutan garis lurus dipakai setiap tahun.'] },
  { a: 'Materialitas', k: ['Ini tentang besar kecilnya jumlah.', 'Sebuah gunting dibeli Rp15.000.', 'Gunting langsung dicatat sebagai beban, tidak dimasukkan ke peralatan.'] },
  { a: 'Kelangsungan usaha', k: ['Ini tentang masa depan usaha.', 'Mesin cuci disusutkan selama 5 tahun.', 'Usaha dianggap akan terus beroperasi, tidak segera dibubarkan.'] },
  { a: 'Unit moneter', k: ['Ini tentang satuan.', 'Keramahan pelayan tidak muncul di laporan.', 'Hanya yang bisa diukur dengan Rupiah yang dicatat.'] },
  { a: 'Objektivitas', k: ['Ini tentang bukti.', 'Pengeluaran tanpa nota ditolak oleh akuntan.', 'Setiap catatan harus didukung dokumen yang dapat diperiksa.'] },
  { a: 'Pengungkapan penuh', k: ['Ini tentang keterbukaan.', 'Ada perkara hukum dengan pemasok yang belum selesai.', 'Hal itu dijelaskan dalam catatan atas laporan keuangan.'] }
];

/* ---------------- GAME 4: JURAGAN SEIMBANG (papan bisnis online) ---------------- */
const JURAGAN = {
  kasAwal: 15000000, gaji: 2000000, bonus: 300000, hadiahPos: 1000000, bonusAudit: 500000,
  pajak: 500000, pinjam: 3000000, batasUtang: 20000000, maksPemain: 4, batasWaktu: 75000,
  grup: { A: 'Kuliner', B: 'Jasa', C: 'Usaha keluarga', D: 'Wisata & niaga' },
  // tipe: mulai | usaha | kartu | pos | pajak | bank | audit | istirahat
  petak: [
    { tipe: 'mulai', n: 'MULAI', s: 'Mulai' },
    { tipe: 'usaha', n: 'Warung Pempek Angso Duo', s: 'Pempek', g: 'A', h: 2000000, w: 500000 },
    { tipe: 'kartu', n: 'Kartu Transaksi', s: 'Kartu' },
    { tipe: 'usaha', n: 'Fotokopi Kampus Mendalo', s: 'Fotokopi', g: 'A', h: 2000000, w: 500000 },
    { tipe: 'pos', n: 'Pos Akuntan', s: 'Pos Akuntan' },
    { tipe: 'bank', n: 'Bank Daerah', s: 'Bank' },
    { tipe: 'usaha', n: 'Laundry Telanaipura', s: 'Laundry', g: 'B', h: 3000000, w: 750000 },
    { tipe: 'usaha', n: 'Kedai Kopi Tanggo Rajo', s: 'Kopi', g: 'B', h: 3000000, w: 750000 },
    { tipe: 'pajak', n: 'Bayar Pajak UMKM', s: 'Pajak' },
    { tipe: 'usaha', n: 'Toko Batik Seberang', s: 'Batik', g: 'B', h: 3500000, w: 900000 },
    { tipe: 'audit', n: 'Pemeriksaan Auditor', s: 'Audit' },
    { tipe: 'usaha', n: 'Bengkel Motor Jelutung', s: 'Bengkel', g: 'C', h: 4000000, w: 1000000 },
    { tipe: 'kartu', n: 'Kartu Transaksi', s: 'Kartu' },
    { tipe: 'usaha', n: 'Katering Kota Baru', s: 'Katering', g: 'C', h: 4000000, w: 1000000 },
    { tipe: 'usaha', n: 'Rental Perahu Danau Sipin', s: 'Perahu', g: 'C', h: 4500000, w: 1100000 },
    { tipe: 'istirahat', n: 'Istirahat di Taman', s: 'Istirahat' },
    { tipe: 'usaha', n: 'Homestay Muaro Jambi', s: 'Homestay', g: 'D', h: 5000000, w: 1250000 },
    { tipe: 'pos', n: 'Pos Akuntan', s: 'Pos Akuntan' },
    { tipe: 'usaha', n: 'Ekspedisi Batanghari', s: 'Ekspedisi', g: 'D', h: 5500000, w: 1400000 },
    { tipe: 'usaha', n: 'Toko Songket Jambi', s: 'Songket', g: 'D', h: 6000000, w: 1500000 }
  ],
  // kartu transaksi: k = kas, a = aset non-kas, u = utang, m = modal (selisih), arah = [harta, utang, modal]
  kartu: [
    { t: 'Pelanggan membayar jasa usahamu secara tunai Rp1.500.000.', k: 1500000, a: 0, u: 0, arah: [1, 0, 1] },
    { t: 'Membayar tagihan listrik dan air Rp600.000.', k: -600000, a: 0, u: 0, arah: [-1, 0, -1] },
    { t: 'Menerima pinjaman dari bank Rp3.000.000.', k: 3000000, a: 0, u: 3000000, arah: [1, 1, 0] },
    { t: 'Membeli mesin baru secara kredit Rp2.000.000.', k: 0, a: 2000000, u: 2000000, arah: [1, 1, 0] },
    { t: 'Membeli perlengkapan secara tunai Rp800.000.', k: -800000, a: 800000, u: 0, arah: [0, 0, 0] },
    { t: 'Mengambil uang kas untuk keperluan pribadi (prive) Rp700.000.', k: -700000, a: 0, u: 0, arah: [-1, 0, -1] },
    { t: 'Menyetor tambahan modal tunai Rp2.000.000.', k: 2000000, a: 0, u: 0, arah: [1, 0, 1] },
    { t: 'Membayar gaji karyawan Rp1.000.000.', k: -1000000, a: 0, u: 0, arah: [-1, 0, -1] },
    { t: 'Menerima pendapatan sewa tenda Rp1.200.000.', k: 1200000, a: 0, u: 0, arah: [1, 0, 1] },
    { t: 'Membayar sebagian utang bank Rp1.000.000.', k: -1000000, a: 0, u: -1000000, arah: [-1, -1, 0], butuhUtang: 1000000 },
    { t: 'Menerima tagihan iklan radio Rp500.000 yang dibayar bulan depan.', k: 0, a: 0, u: 500000, arah: [0, 1, -1] },
    { t: 'Pelanggan membayar uang muka untuk jasa bulan depan Rp1.000.000.', k: 1000000, a: 0, u: 1000000, arah: [1, 1, 0] },
    { t: 'Mendapat hadiah lomba UMKM dari pemerintah daerah Rp1.000.000.', k: 1000000, a: 0, u: 0, arah: [1, 0, 1] },
    { t: 'Mesin rusak, membayar biaya servis Rp400.000.', k: -400000, a: 0, u: 0, arah: [-1, 0, -1] },
    { t: 'Jasa katering selesai, pelanggan akan membayar bulan depan Rp1.500.000.', k: 0, a: 1500000, u: 0, arah: [1, 0, 1] },
    { t: 'Membeli perlengkapan secara kredit Rp600.000.', k: 0, a: 600000, u: 600000, arah: [1, 1, 0] }
  ]
};

/* Soal pilihan ganda untuk Pos Akuntan dan Audit di Juragan Seimbang */
const SOAL_GAME = [
  { q: 'Mencampur uang belanja rumah dengan kas usaha melanggar …', o: ['Kesatuan usaha', 'Konsistensi', 'Materialitas'], a: 0 },
  { q: 'Aset dicatat sebesar harga saat dibeli. Ini prinsip …', o: ['Biaya historis', 'Pengungkapan penuh', 'Periode akuntansi'], a: 0 },
  { q: 'Pendapatan dicatat saat jasa selesai, walau uang belum diterima, sesuai prinsip …', o: ['Pengakuan pendapatan', 'Unit moneter', 'Objektivitas'], a: 0 },
  { q: 'Pihak yang menilai kemampuan usaha membayar pinjaman adalah …', o: ['Kreditur (bank)', 'Pelanggan', 'Masyarakat'], a: 0 },
  { q: 'Bukti untuk transaksi pembelian secara kredit adalah …', o: ['Faktur', 'Kuitansi', 'Nota kontan'], a: 0 },
  { q: 'Laporan yang menunjukkan harta, utang, dan modal pada tanggal tertentu adalah …', o: ['Neraca', 'Laporan laba rugi', 'Laporan perubahan modal'], a: 0 },
  { q: 'Rumus modal adalah …', o: ['Harta − Utang', 'Harta + Utang', 'Utang − Harta'], a: 0 },
  { q: 'Prive menyebabkan …', o: ['Harta dan modal berkurang', 'Harta dan beban bertambah', 'Utang berkurang'], a: 0 },
  { q: 'Usaha dianggap akan terus beroperasi. Ini asumsi …', o: ['Kelangsungan usaha', 'Periode akuntansi', 'Kesatuan usaha'], a: 0 },
  { q: 'Penggaris Rp10.000 langsung dicatat sebagai beban karena prinsip …', o: ['Materialitas', 'Konsistensi', 'Penandingan'], a: 0 },
  { q: 'Kegiatan akuntansi yang pertama adalah …', o: ['Identifikasi transaksi', 'Pelaporan', 'Penggolongan'], a: 0 },
  { q: 'Manajer termasuk pengguna informasi akuntansi …', o: ['Internal', 'Eksternal', 'Bukan pengguna'], a: 0 },
  { q: 'Membayar utang kepada pemasok menyebabkan …', o: ['Harta dan utang berkurang', 'Harta dan modal berkurang', 'Utang bertambah'], a: 0 },
  { q: 'Informasi akuntansi bisa dicek ulang dengan bukti. Karakteristik ini disebut …', o: ['Dapat diverifikasi', 'Tepat waktu', 'Dapat dipahami'], a: 0 },
  { q: 'Laporan keuangan Laundry Kilat disusun setiap bulan sesuai asumsi …', o: ['Periode akuntansi', 'Unit moneter', 'Biaya historis'], a: 0 },
  { q: 'Menerima pinjaman bank menyebabkan …', o: ['Harta dan utang bertambah', 'Harta dan modal bertambah', 'Modal bertambah, utang berkurang'], a: 0 },
  { q: 'Gaji Oktober yang dibayar 1 November dicatat sebagai beban Oktober. Ini prinsip …', o: ['Penandingan', 'Biaya historis', 'Materialitas'], a: 0 },
  { q: 'Laporan yang menunjukkan laba atau rugi satu periode adalah …', o: ['Laporan laba rugi', 'Neraca', 'Daftar aset'], a: 0 }
];

/* ---------------- LEVEL & LENCANA ---------------- */
const LEVEL = [
  { xp: 0, nama: 'Magang' },
  { xp: 200, nama: 'Juru Catat' },
  { xp: 450, nama: 'Asisten Akuntan' },
  { xp: 750, nama: 'Akuntan Junior' },
  { xp: 1050, nama: 'Akuntan Senior' },
  { xp: 1300, nama: 'Kepala Keuangan' }
];

const LENCANA = [
  { id: 'b_m1', nama: 'Pembuka Buku', ket: 'Menyelesaikan Misi 1', ikon: 'book' },
  { id: 'b_m2', nama: 'Peta Pengguna', ket: 'Menyelesaikan Misi 2', ikon: 'people' },
  { id: 'b_m3', nama: 'Penjaga Prinsip', ket: 'Menyelesaikan Misi 3', ikon: 'shield' },
  { id: 'b_m4', nama: 'Si Seimbang', ket: 'Menyelesaikan Misi 4', ikon: 'scale' },
  { id: 'b_tuntas', nama: 'Tuntas', ket: 'Nilai kuis mencapai KKTP', ikon: 'check' },
  { id: 'b_100', nama: 'Sempurna', ket: 'Nilai kuis 100', ikon: 'star' },
  { id: 'b_g1', nama: 'Neraca Emas', ket: 'Skor Timbangan Seimbang ≥ 1.200', ikon: 'scale' },
  { id: 'b_g2', nama: 'Tangan Kilat', ket: 'Skor Akun Kilat ≥ 150', ikon: 'bolt' },
  { id: 'b_g3', nama: 'Detektif Ulung', ket: 'Skor Detektif Prinsip ≥ 1.800', ikon: 'search' },
  { id: 'b_g4', nama: 'Juragan Jambi', ket: 'Modal akhir ≥ Rp21 juta di Juragan Seimbang', ikon: 'trophy' },
  { id: 'b_ref', nama: 'Reflektif', ket: 'Mengirim refleksi belajar', ikon: 'pen' }
];


/* ---------- LOGIKA BERSAMA ---------- */
/* =========================================================
   SEIMBANG — LOGIKA BERSAMA (klien + server)
   Perhitungan XP, level, lencana, persentase misi, pengaturan.
   Ditulis dengan sintaks ES5 sederhana agar aman di Apps Script.
   ========================================================= */

var GAMES = [
  { id: 'timbangan', nama: 'Timbangan Seimbang', target: 1200, badge: 'b_g1', ket: 'Tebak pengaruh transaksi pada Harta, Utang, dan Modal sebelum waktu habis.' },
  { id: 'akun', nama: 'Akun Kilat', target: 150, badge: 'b_g2', ket: 'Golongkan nama akun ke kelompok yang tepat dalam 60 detik.' },
  { id: 'detektif', nama: 'Detektif Prinsip', target: 1800, badge: 'b_g3', ket: 'Ungkap prinsip akuntansi dari petunjuk. Makin sedikit petunjuk, makin besar poin.' },
  { id: 'juragan', nama: 'Juragan Seimbang', target: 21000, min: 15000, badge: 'b_g4', online: true, ket: 'Papan bisnis online bersama teman: beli usaha di Kota Jambi, tarik sewa, catat setiap transaksi, dan kejar modal terbesar.' }
];

var SETTINGS_DEF = [
  ['nama_guru', 'Dr. Herry Sofyandy P., M.Pd., Gr.', 'Nama guru yang tampil di header aplikasi'],
  ['sekolah', '', 'Nama sekolah (boleh kosong)'],
  ['kktp', '75', 'Kriteria ketercapaian tujuan pembelajaran, 0 sampai 100'],
  ['kuis_dibuka', 'TRUE', 'Kuis bisa dikerjakan siswa (TRUE/FALSE)'],
  ['kuis_jumlah', '20', 'Jumlah soal per kuis, 4 sampai 36'],
  ['kuis_durasi', '25', 'Durasi kuis dalam menit'],
  ['kuis_maks', '3', 'Batas percobaan kuis per siswa (0 = tanpa batas)'],
  ['kuis_pembahasan', 'TRUE', 'Tampilkan pembahasan setelah kuis dikirim (TRUE/FALSE)'],
  ['misi_terbuka', 'm1,m2,m3,m4', 'Misi yang dibuka, pisahkan dengan koma'],
  ['game_dibuka', 'TRUE', 'Game bisa dimainkan (TRUE/FALSE)'],
  ['peringkat', 'TRUE', 'Tampilkan papan peringkat kelas (TRUE/FALSE)'],
  ['tamu', 'TRUE', 'Izinkan masuk sebagai tamu tanpa menyimpan nilai (TRUE/FALSE)'],
  ['pengumuman', '', 'Pengumuman yang tampil di beranda siswa'],
  ['fokus', '', 'Ajak siswa ke: m1, m2, m3, m4, kuis, game, atau kosong'],
  ['versi', '1', 'Naik otomatis setiap pengaturan berubah']
];

var SB = {
  bool: function (v) { return v === true || /^(true|ya|y|1|benar|on)$/i.test(String(v == null ? '' : v).trim()); },
  num: function (v, d) { var n = Number(String(v == null ? '' : v).replace(',', '.')); return isFinite(n) && String(v).trim() !== '' ? n : (d === undefined ? 0 : d); },
  clamp: function (n, a, b) { return Math.max(a, Math.min(b, n)); },
  uniqInts: function (arr, max) {
    var seen = {}, out = [];
    (arr || []).forEach(function (x) { var n = Math.floor(Number(x)); if (isFinite(n) && n >= 0 && n < max && !seen[n]) { seen[n] = 1; out.push(n); } });
    return out;
  },

  /* Pengaturan mentah (objek kunci→teks) menjadi objek bertipe */
  parseSettings: function (raw) {
    raw = raw || {};
    var s = {};
    SETTINGS_DEF.forEach(function (d) { s[d[0]] = (raw[d[0]] === undefined || raw[d[0]] === null) ? d[1] : raw[d[0]]; });
    return {
      nama_guru: String(s.nama_guru || ''),
      sekolah: String(s.sekolah || ''),
      kktp: SB.clamp(SB.num(s.kktp, 75), 0, 100),
      kuis_dibuka: SB.bool(s.kuis_dibuka),
      kuis_jumlah: SB.clamp(Math.round(SB.num(s.kuis_jumlah, 20)), 4, 40),
      kuis_durasi: SB.clamp(Math.round(SB.num(s.kuis_durasi, 25)), 3, 180),
      kuis_maks: SB.clamp(Math.round(SB.num(s.kuis_maks, 3)), 0, 20),
      kuis_pembahasan: SB.bool(s.kuis_pembahasan),
      misi_terbuka: String(s.misi_terbuka || '').split(/[,\s]+/).filter(function (x) { return /^m[1-4]$/.test(x); }),
      game_dibuka: SB.bool(s.game_dibuka),
      peringkat: SB.bool(s.peringkat),
      tamu: SB.bool(s.tamu),
      pengumuman: String(s.pengumuman || ''),
      fokus: /^(m[1-4]|kuis|game)$/.test(String(s.fokus)) ? String(s.fokus) : '',
      versi: Math.round(SB.num(s.versi, 1))
    };
  },

  /* Membersihkan data progres kiriman klien */
  cleanData: function (d) {
    d = (d && typeof d === 'object') ? d : {};
    var out = { misi: {}, game: {}, catatan: {}, terakhir: null };
    MISI.forEach(function (m) {
      var p = (d.misi && d.misi[m.id]) || {};
      var o = { baca: SB.uniqInts(p.baca, m.bagian.length), cek: {}, akt: null };
      m.bagian.forEach(function (b, i) {
        if (b.cek && p.cek && p.cek[i]) {
          o.cek[i] = (p.cek[i] || []).slice(0, b.cek.item.length).map(function (v) { return v === 1 || v === true ? 1 : (v === 0 || v === false ? 0 : null); });
        }
      });
      if (p.akt && typeof p.akt === 'object') {
        o.akt = { selesai: !!p.akt.selesai, skor: SB.clamp(SB.num(p.akt.skor, 0), 0, 1), coba: Math.max(0, Math.round(SB.num(p.akt.coba, 0))) };
      }
      out.misi[m.id] = o;
    });
    GAMES.forEach(function (g) { out.game[g.id] = Math.max(0, Math.round(SB.num(d.game && d.game[g.id], 0))); });
    SB.pemantikKeys().forEach(function (k) {
      if (d.catatan && typeof d.catatan[k] === 'string' && d.catatan[k].trim()) out.catatan[k] = d.catatan[k].slice(0, 600);
    });
    if (d.terakhir && typeof d.terakhir === 'object' && /^m[1-4]$/.test(d.terakhir.m)) {
      out.terakhir = { m: d.terakhir.m, b: Math.max(0, Math.round(SB.num(d.terakhir.b, 0))) };
    }
    return out;
  },

  pemantikKeys: function () {
    var keys = [];
    MISI.forEach(function (m) { m.bagian.forEach(function (b, i) { if (b.pemantik) keys.push(m.id + '-' + i); }); });
    return keys;
  },

  /* XP, level, lencana, dan persentase misi.
     kuisTerbaik dan jumlah refleksi berasal dari catatan server. */
  statistik: function (data, kuisTerbaik, refleksiN, kktp) {
    data = SB.cleanData(data);
    kuisTerbaik = SB.num(kuisTerbaik, 0);
    kktp = SB.num(kktp, 75);
    var xp = 0, persen = {}, lencana = [], rincianXp = { materi: 0, latihan: 0, kuis: 0, game: 0, refleksi: 0 };
    MISI.forEach(function (m) {
      var p = data.misi[m.id];
      var aktSelesai = !!(p.akt && p.akt.selesai);
      var done = p.baca.length + (aktSelesai ? 1 : 0);
      persen[m.id] = Math.round(done / (m.bagian.length + 1) * 100);
      rincianXp.materi += p.baca.length * 10;
      Object.keys(p.cek).forEach(function (k) { (p.cek[k] || []).forEach(function (v) { if (v === 1) rincianXp.latihan += 5; }); });
      if (aktSelesai) rincianXp.latihan += 50 + Math.round(50 * p.akt.skor);
      if (persen[m.id] >= 100) lencana.push('b_' + m.id);
    });
    Object.keys(data.catatan).forEach(function (k) { if (data.catatan[k].trim().length >= 5) rincianXp.materi += 5; });
    rincianXp.kuis = Math.round(SB.clamp(kuisTerbaik, 0, 100) * 2);
    if (kuisTerbaik >= kktp && kuisTerbaik > 0) lencana.push('b_tuntas');
    if (kuisTerbaik >= 100) lencana.push('b_100');
    GAMES.forEach(function (g) {
      var best = data.game[g.id] || 0;
      var base = g.min || 0;
      rincianXp.game += SB.clamp(Math.round((best - base) / (g.target - base) * 100), 0, 100);
      if (best >= g.target) lencana.push(g.badge);
    });
    if (refleksiN > 0) { rincianXp.refleksi = 50; lencana.push('b_ref'); }
    xp = rincianXp.materi + rincianXp.latihan + rincianXp.kuis + rincianXp.game + rincianXp.refleksi;
    var li = 0;
    for (var i = 0; i < LEVEL.length; i++) if (xp >= LEVEL[i].xp) li = i;
    var misiSelesai = MISI.filter(function (m) { return persen[m.id] >= 100; }).length;
    return {
      xp: xp, level: li + 1, levelNama: LEVEL[li].nama, xpLevel: LEVEL[li].xp,
      xpNext: li + 1 < LEVEL.length ? LEVEL[li + 1].xp : null,
      nextNama: li + 1 < LEVEL.length ? LEVEL[li + 1].nama : null,
      persen: persen, misiSelesai: misiSelesai, lencana: lencana, rincianXp: rincianXp,
      kuisTerbaik: kuisTerbaik, game: data.game
    };
  },

  /* Angka bebas ("Rp38.000.000", "38 juta" tidak didukung) → bilangan bulat */
  parseAngka: function (v) {
    var s = String(v == null ? '' : v).replace(/[^0-9-]/g, '');
    if (!s || s === '-') return null;
    var n = parseInt(s, 10);
    return isFinite(n) ? n : null;
  }
};


/* ---------- LOGIKA SERVER ---------- */
/* =========================================================
   SEIMBANG — LOGIKA SERVER
   Dipakai oleh Code.gs (Google Apps Script + Google Sheets)
   dan oleh mode demo (penyimpanan di peramban).
   env = { table(nama), secret(), random(), now(), lock(fn), info }
   ========================================================= */

var SBServer = (function () {

  var SCHEMA = {
    Pengaturan: ['kunci', 'nilai', 'keterangan'],
    Akun: ['username', 'nama', 'kelas', 'peran', 'aktif', 'sandi_awal', 'salt', 'hash', 'kuis_tambahan', 'dibuat', 'login_terakhir'],
    Progres: ['username', 'nama', 'kelas', 'xp', 'level', 'misi1', 'misi2', 'misi3', 'misi4', 'kuis_terbaik', 'kuis_percobaan', 'game_timbangan', 'game_akun', 'game_detektif', 'game_juragan', 'lencana', 'refleksi', 'diperbarui', 'data_json'],
    Kuis: ['id', 'waktu', 'username', 'nama', 'kelas', 'percobaan', 'nilai', 'benar', 'jumlah', 'durasi_detik', 'tuntas', 'rincian_json'],
    Refleksi: ['id', 'waktu', 'username', 'nama', 'kelas', 'perasaan', 'paham', 'bingung', 'pertanyaan', 'balasan', 'waktu_balasan']
  };
  var TEXT_COLS = {
    Pengaturan: ['kunci', 'nilai'],
    Akun: ['username', 'kelas', 'sandi_awal', 'salt', 'hash'],
    Progres: ['username', 'kelas', 'data_json'],
    Kuis: ['id', 'username', 'kelas', 'rincian_json'],
    Refleksi: ['id', 'username', 'kelas']
  };

  /* ---------- SHA-256 & HMAC (murni JS, tanpa pustaka) ---------- */
  function sha256(ascii) {
    function rr(v, a) { return (v >>> a) | (v << (32 - a)); }
    var maxWord = Math.pow(2, 32), i, j, result = '';
    var words = [], bitLen = ascii.length * 8;
    var hash = [], k = [], primeCounter = 0, isComposite = {};
    for (var c = 2; primeCounter < 64; c++) {
      if (!isComposite[c]) {
        for (i = 0; i < 313; i += c) isComposite[i] = c;
        hash[primeCounter] = (Math.pow(c, 0.5) * maxWord) | 0;
        k[primeCounter++] = (Math.pow(c, 1 / 3) * maxWord) | 0;
      }
    }
    ascii += '\x80';
    while (ascii.length % 64 - 56) ascii += '\x00';
    for (i = 0; i < ascii.length; i++) {
      j = ascii.charCodeAt(i);
      words[i >> 2] |= j << ((3 - i) % 4) * 8;
    }
    words[words.length] = ((bitLen / maxWord) | 0);
    words[words.length] = bitLen;
    hash = hash.slice(0, 8);
    for (j = 0; j < words.length;) {
      var w = words.slice(j, j += 16), oldHash = hash;
      hash = hash.slice(0, 8);
      for (i = 0; i < 64; i++) {
        var w15 = w[i - 15], w2 = w[i - 2], a = hash[0], e = hash[4];
        var t1 = hash[7] + (rr(e, 6) ^ rr(e, 11) ^ rr(e, 25)) + ((e & hash[5]) ^ ((~e) & hash[6])) + k[i] +
          (w[i] = (i < 16) ? w[i] : (w[i - 16] + (rr(w15, 7) ^ rr(w15, 18) ^ (w15 >>> 3)) + w[i - 7] + (rr(w2, 17) ^ rr(w2, 19) ^ (w2 >>> 10))) | 0);
        var t2 = (rr(a, 2) ^ rr(a, 13) ^ rr(a, 22)) + ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));
        hash = [(t1 + t2) | 0].concat(hash);
        hash[4] = (hash[4] + t1) | 0;
      }
      for (i = 0; i < 8; i++) hash[i] = (hash[i] + oldHash[i]) | 0;
    }
    for (i = 0; i < 8; i++) for (j = 3; j + 1; j--) { var b = (hash[i] >> (j * 8)) & 255; result += ((b < 16) ? '0' : '') + b.toString(16); }
    return result;
  }
  function utf8(s) { return unescape(encodeURIComponent(String(s))); }
  function hexToBin(h) { var s = ''; for (var i = 0; i < h.length; i += 2) s += String.fromCharCode(parseInt(h.substr(i, 2), 16)); return s; }
  function hash(s) { return sha256(utf8(s)); }
  function hmac(key, msg) {
    key = utf8(key); msg = utf8(msg);
    if (key.length > 64) key = hexToBin(sha256(key));
    while (key.length < 64) key += '\x00';
    var ip = '', op = '';
    for (var i = 0; i < 64; i++) { ip += String.fromCharCode(key.charCodeAt(i) ^ 0x36); op += String.fromCharCode(key.charCodeAt(i) ^ 0x5c); }
    return sha256(op + hexToBin(sha256(ip + msg)));
  }
  function hashSandi(salt, pw) {
    var h = hash(salt + ':' + pw);
    for (var i = 0; i < 150; i++) h = sha256(h + ':' + salt);
    return h;
  }

  /* ---------- pembantu ---------- */
  function fail(msg, code) { var e = new Error(msg); e.userMsg = msg; e.code = code || 'gagal'; throw e; }
  function str(v) { return v == null ? '' : String(v); }
  function normUser(u) { return str(u).trim().toLowerCase().replace(/\s+/g, ''); }
  function parseJSON(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }
  function shuffle(arr, rnd) { for (var i = arr.length - 1; i > 0; i--) { var j = Math.floor(rnd() * (i + 1)); var t = arr[i]; arr[i] = arr[j]; arr[j] = t; } return arr; }
  function soalById(id) { for (var i = 0; i < SOAL.length; i++) if (SOAL[i].id === id) return SOAL[i]; return null; }

  function create(env) {
    var T = function (n) { return env.table(n); };
    var rnd = function () { return parseInt(env.random().slice(0, 8), 16) / 0x100000000; };
    var nowMs = function () { return env.now(); };
    var nowIso = function () { return new Date(env.now()).toISOString(); };

    function genSandi(n) {
      var al = 'abcdefghjkmnpqrstuvwxyz23456789', s = '';
      for (var i = 0; i < (n || 6); i++) s += al.charAt(Math.floor(rnd() * al.length));
      return s;
    }

    /* ----- pengaturan ----- */
    function rawSettings() {
      var o = {};
      T('Pengaturan').all().forEach(function (r) { if (r.kunci) o[str(r.kunci)] = r.nilai; });
      return o;
    }
    function settings() { return SB.parseSettings(rawSettings()); }
    function publicSettings(s) {
      return {
        nama_guru: s.nama_guru, sekolah: s.sekolah, kktp: s.kktp, kuis_dibuka: s.kuis_dibuka, kuis_jumlah: s.kuis_jumlah,
        kuis_durasi: s.kuis_durasi, kuis_maks: s.kuis_maks, misi_terbuka: s.misi_terbuka, game_dibuka: s.game_dibuka,
        peringkat: s.peringkat, tamu: s.tamu, pengumuman: s.pengumuman, fokus: s.fokus, versi: s.versi
      };
    }
    function writeSetting(k, v) {
      var t = T('Pengaturan');
      var ket = ''; SETTINGS_DEF.forEach(function (d) { if (d[0] === k) ket = d[2]; });
      if (!t.update('kunci', k, { nilai: v })) t.insert({ kunci: k, nilai: v, keterangan: ket });
    }

    /* ----- akun & sesi ----- */
    function users() { return T('Akun').all(); }
    function findUser(u) { u = normUser(u); var list = users(); for (var i = 0; i < list.length; i++) if (normUser(list[i].username) === u) return list[i]; return null; }
    function makeToken(u, role) {
      var p = 'S|' + u + '|' + role + '|' + (nowMs() + 45 * 86400000);
      return p + '|' + hmac(env.secret(), p).slice(0, 32);
    }
    function cekToken(token, code) {
      var parts = str(token).split('|');
      if (parts.length !== 5 || parts[0] !== 'S') fail('Sesi tidak dikenali. Silakan masuk lagi.', code || 'sesi');
      var p = parts.slice(0, 4).join('|');
      if (hmac(env.secret(), p).slice(0, 32) !== parts[4]) fail('Sesi tidak sah. Silakan masuk lagi.', code || 'sesi');
      if (Number(parts[3]) < nowMs()) fail('Sesi sudah berakhir. Silakan masuk lagi.', code || 'sesi');
      return parts;
    }
    function auth(token, role) {
      var parts = cekToken(token);
      var u = findUser(parts[1]);
      if (!u || str(u.peran) !== parts[2]) fail('Akun tidak ditemukan. Silakan masuk lagi.', 'sesi');
      if (!SB.bool(u.aktif)) fail('Akun ini sedang dinonaktifkan guru.', 'sesi');
      if (role && str(u.peran) !== role) fail('Fitur ini khusus untuk ' + (role === 'guru' ? 'guru' : 'siswa') + '.', 'peran');
      return u;
    }
    function userPublic(u) { return { username: str(u.username), nama: str(u.nama), kelas: str(u.kelas), peran: str(u.peran) }; }

    function buatAkun(o) {
      var salt = env.random().slice(0, 16);
      var row = {
        username: o.username, nama: o.nama, kelas: o.kelas || '', peran: o.peran || 'siswa', aktif: true,
        sandi_awal: o.password, salt: salt, hash: hashSandi(salt, o.password), kuis_tambahan: 0,
        dibuat: nowIso(), login_terakhir: ''
      };
      T('Akun').insert(row);
      return row;
    }
    function usernameDariNama(nama, taken) {
      var base = str(nama).toLowerCase().normalize ? str(nama).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '') : str(nama).toLowerCase();
      base = (base.split(/\s+/)[0] || 'siswa').replace(/[^a-z0-9]/g, '').slice(0, 10) || 'siswa';
      for (var n = 1; n < 1000; n++) { var u = base + (n < 10 ? '0' + n : n); if (!taken[u]) return u; }
      return base + env.random().slice(0, 4);
    }

    /* ----- kuis & progres ----- */
    function kuisRows(username) {
      var u = normUser(username);
      return T('Kuis').all().filter(function (r) { return normUser(r.username) === u; });
    }
    function refleksiRows(username) {
      var u = normUser(username);
      return T('Refleksi').all().filter(function (r) { return normUser(r.username) === u; });
    }
    function kuisInfo(u, s) {
      var rows = kuisRows(u.username);
      var best = 0; rows.forEach(function (r) { best = Math.max(best, SB.num(r.nilai, 0)); });
      var tambahan = Math.max(0, Math.round(SB.num(u.kuis_tambahan, 0)));
      var sisa = s.kuis_maks === 0 ? null : Math.max(0, s.kuis_maks + tambahan - rows.length);
      return {
        dibuka: s.kuis_dibuka, jumlah: s.kuis_jumlah, durasi: s.kuis_durasi, kktp: s.kktp, maks: s.kuis_maks,
        tambahan: tambahan, sisa: sisa, terbaik: best, percobaan: rows.length, pembahasan: s.kuis_pembahasan,
        riwayat: rows.map(function (r) { return { waktu: str(r.waktu), nilai: SB.num(r.nilai), benar: SB.num(r.benar), jumlah: SB.num(r.jumlah) }; })
      };
    }
    function progresRow(username) {
      var u = normUser(username), list = T('Progres').all();
      for (var i = 0; i < list.length; i++) if (normUser(list[i].username) === u) return list[i];
      return null;
    }
    function segarkanProgres(u, dataBaru, s) {
      s = s || settings();
      var row = progresRow(u.username);
      var data = dataBaru ? SB.cleanData(dataBaru) : SB.cleanData(parseJSON(row && row.data_json, {}));
      var ki = kuisInfo(u, s);
      var refN = refleksiRows(u.username).length;
      var st = SB.statistik(data, ki.terbaik, refN, s.kktp);
      var rec = {
        username: str(u.username), nama: str(u.nama), kelas: str(u.kelas), xp: st.xp, level: st.levelNama,
        misi1: st.persen.m1, misi2: st.persen.m2, misi3: st.persen.m3, misi4: st.persen.m4,
        kuis_terbaik: ki.terbaik, kuis_percobaan: ki.percobaan,
        game_timbangan: data.game.timbangan, game_akun: data.game.akun, game_detektif: data.game.detektif, game_juragan: data.game.juragan,
        lencana: st.lencana.length, refleksi: refN, diperbarui: nowIso(), data_json: JSON.stringify(data)
      };
      if (row) T('Progres').update('username', row.username, rec); else T('Progres').insert(rec);
      return { data: data, stat: st, kuis: ki };
    }
    function stateSiswa(u, s) {
      var row = progresRow(u.username);
      var data = SB.cleanData(parseJSON(row && row.data_json, {}));
      var ki = kuisInfo(u, s);
      var refl = refleksiRows(u.username).map(function (r) {
        return { id: str(r.id), waktu: str(r.waktu), perasaan: SB.num(r.perasaan), paham: str(r.paham), bingung: str(r.bingung), pertanyaan: str(r.pertanyaan), balasan: str(r.balasan), waktu_balasan: str(r.waktu_balasan) };
      });
      return {
        user: userPublic(u), settings: publicSettings(s), data: data,
        stat: SB.statistik(data, ki.terbaik, refl.length, s.kktp), kuis: ki, refleksi: refl,
        sandiBawaan: !!str(u.sandi_awal)
      };
    }

    /* ---------- AKSI ---------- */
    var A = {};

    A.ping = function () {
      var s = settings(), info = env.info || {};
      return { settings: publicSettings(s), info: { mode: info.mode, versi: info.versi } };
    };

    A.login = function (d) {
      var u = findUser(d.username);
      if (!u || !u.hash) fail('Username atau kata sandi salah.');
      if (hashSandi(str(u.salt), str(d.password)) !== str(u.hash)) fail('Username atau kata sandi salah.');
      if (!SB.bool(u.aktif)) fail('Akun ini sedang dinonaktifkan guru.');
      env.lock(function () { T('Akun').update('username', u.username, { login_terakhir: nowIso() }); });
      var s = settings();
      var out = { token: makeToken(normUser(u.username), str(u.peran)), user: userPublic(u) };
      if (str(u.peran) === 'guru') { out.settings = s; out.sandiBawaan = !!str(u.sandi_awal); }
      else { var st = stateSiswa(u, s); for (var k in st) out[k] = st[k]; }
      return out;
    };

    A.me = function (d) {
      var u = auth(d.token);
      var s = settings();
      if (str(u.peran) === 'guru') return { user: userPublic(u), settings: s, sandiBawaan: !!str(u.sandi_awal) };
      return stateSiswa(u, s);
    };

    A.sync = function (d) {
      auth(d.token);
      return { settings: publicSettings(settings()) };
    };

    A.gantiSandi = function (d) {
      var u = auth(d.token);
      if (hashSandi(str(u.salt), str(d.lama)) !== str(u.hash)) fail('Kata sandi lama tidak cocok.');
      var baru = str(d.baru);
      if (baru.length < 6) fail('Kata sandi baru minimal 6 karakter.');
      var salt = env.random().slice(0, 16);
      env.lock(function () { T('Akun').update('username', u.username, { salt: salt, hash: hashSandi(salt, baru), sandi_awal: '' }); });
      return { ok: true };
    };

    /* ----- siswa ----- */
    A.simpanProgres = function (d) {
      var u = auth(d.token, 'siswa');
      var json = JSON.stringify(d.data || {});
      if (json.length > 40000) fail('Data progres terlalu besar.');
      var res;
      env.lock(function () {
        // game: simpan skor terbaik, jangan turunkan skor yang sudah tersimpan
        var lama = SB.cleanData(parseJSON((progresRow(u.username) || {}).data_json, {}));
        var baru = SB.cleanData(d.data);
        GAMES.forEach(function (g) { baru.game[g.id] = Math.max(baru.game[g.id] || 0, lama.game[g.id] || 0); });
        MISI.forEach(function (m) {
          var a = lama.misi[m.id], b = baru.misi[m.id];
          b.baca = SB.uniqInts(a.baca.concat(b.baca), m.bagian.length);
          if (a.akt && a.akt.selesai && !(b.akt && b.akt.selesai)) b.akt = a.akt;
          else if (a.akt && a.akt.selesai && b.akt && b.akt.selesai && a.akt.skor > b.akt.skor) b.akt.skor = a.akt.skor;
          Object.keys(a.cek).forEach(function (k) { if (!b.cek[k]) b.cek[k] = a.cek[k]; });
        });
        res = segarkanProgres(u, baru);
      });
      return { stat: res.stat, data: res.data };
    };

    A.kuisMulai = function (d) {
      var u = auth(d.token, 'siswa');
      var s = settings();
      var ki = kuisInfo(u, s);
      if (!s.kuis_dibuka) fail('Kuis sedang ditutup oleh guru.');
      if (ki.sisa === 0) fail('Kesempatan mengerjakan kuis sudah habis. Minta guru menambah kesempatan.');
      // ambil soal merata dari setiap misi
      var grup = {};
      SOAL.forEach(function (q) { (grup[q.m] = grup[q.m] || []).push(q); });
      var ids = Object.keys(grup).sort();
      var N = Math.min(s.kuis_jumlah, SOAL.length), jatah = {}, sisaN = N;
      ids.forEach(function (m) { jatah[m] = 0; shuffle(grup[m], rnd); });
      while (sisaN > 0) {
        var maju = false;
        for (var i = 0; i < ids.length && sisaN > 0; i++) {
          if (jatah[ids[i]] < grup[ids[i]].length) { jatah[ids[i]]++; sisaN--; maju = true; }
        }
        if (!maju) break;
      }
      var pilih = [];
      ids.forEach(function (m) { pilih = pilih.concat(grup[m].slice(0, jatah[m])); });
      var kode = [], soal = [];
      pilih.forEach(function (q) {
        var perm = '';
        if (q.t !== 'angka') perm = shuffle(q.o.map(function (x, i) { return i; }), rnd).join('');
        kode.push(q.id + '.' + perm);
        soal.push({ id: q.id, m: q.m, t: q.t || 'pg', q: q.q, o: q.t === 'angka' ? null : perm.split('').map(function (k) { return q.o[Number(k)]; }) });
      });
      var mulai = nowMs();
      var p = 'K|' + normUser(u.username) + '|' + (ki.percobaan + 1) + '|' + mulai + '|' + s.kuis_durasi + '|' + kode.join(',');
      return { token: p + '|' + hmac(env.secret(), p).slice(0, 32), soal: soal, durasi: s.kuis_durasi, mulai: mulai, berakhir: mulai + s.kuis_durasi * 60000, percobaan: ki.percobaan + 1 };
    };

    A.kuisKirim = function (d) {
      var u = auth(d.token, 'siswa');
      var parts = str(d.kuis).split('|');
      if (parts.length !== 7 || parts[0] !== 'K') fail('Data kuis tidak dikenali.');
      var p = parts.slice(0, 6).join('|'), sig = parts[6];
      if (hmac(env.secret(), p).slice(0, 32) !== sig) fail('Data kuis tidak sah.');
      if (parts[1] !== normUser(u.username)) fail('Kuis ini milik akun lain.');
      var id = sig.slice(0, 14);
      var s = settings();
      var kode = parts[5].split(',');
      var ada = T('Kuis').all().filter(function (r) { return str(r.id) === id; })[0];
      var jawaban = (d.jawaban && typeof d.jawaban === 'object') ? d.jawaban : {};
      if (ada) {
        // kiriman ulang (mis. sinyal putus): pakai jawaban yang sudah tersimpan
        jawaban = {};
        var simpan = {};
        parseJSON(ada.rincian_json, []).forEach(function (x) { simpan[x.id] = x.j; });
        kode.forEach(function (c) {
          var bits = c.split('.'), q = soalById(bits[0]);
          if (!q || !(q.id in simpan)) return;
          if (q.t === 'angka') jawaban[q.id] = simpan[q.id];
          else if (simpan[q.id] !== -1 && simpan[q.id] !== '') jawaban[q.id] = bits[1].split('').map(Number).indexOf(Number(simpan[q.id]));
        });
      }
      var benar = 0, rincian = [], bahas = [], perMisi = {};
      kode.forEach(function (c) {
        var bits = c.split('.'), q = soalById(bits[0]);
        if (!q) return;
        var perm = bits[1] ? bits[1].split('').map(Number) : null;
        var jw = jawaban[q.id], ok = false, asli = -1;
        if (q.t === 'angka') {
          var n = SB.parseAngka(jw);
          ok = n !== null && n === q.a;
          asli = n === null ? '' : n;
        } else if (jw !== undefined && jw !== null && jw !== '') {
          var tampil = Math.floor(Number(jw));
          if (perm && tampil >= 0 && tampil < perm.length) { asli = perm[tampil]; ok = asli === q.a; }
        }
        if (ok) benar++;
        perMisi[q.m] = perMisi[q.m] || { b: 0, n: 0 };
        perMisi[q.m].n++; if (ok) perMisi[q.m].b++;
        rincian.push({ id: q.id, j: asli, b: ok ? 1 : 0 });
        bahas.push({
          id: q.id, m: q.m, t: q.t || 'pg', q: q.q,
          o: perm ? perm.map(function (k) { return q.o[k]; }) : null,
          jawab: q.t === 'angka' ? (jw == null ? '' : str(jw)) : (jw === undefined || jw === null || jw === '' ? -1 : Math.floor(Number(jw))),
          kunci: q.t === 'angka' ? q.a : perm.indexOf(q.a), benar: ok, p: q.p
        });
      });
      var jumlah = rincian.length || 1;
      var nilai = Math.round(benar / jumlah * 100);
      var res;
      env.lock(function () {
        var sudah = T('Kuis').all().filter(function (r) { return str(r.id) === id; })[0];
        if (!sudah) {
          T('Kuis').insert({
            id: id, waktu: nowIso(), username: str(u.username), nama: str(u.nama), kelas: str(u.kelas),
            percobaan: kuisRows(u.username).length + 1, nilai: nilai, benar: benar, jumlah: rincian.length,
            durasi_detik: Math.max(0, Math.round((nowMs() - Number(parts[3])) / 1000)),
            tuntas: nilai >= s.kktp ? 'Ya' : 'Belum', rincian_json: JSON.stringify(rincian)
          });
        } else {
          nilai = SB.num(sudah.nilai); benar = SB.num(sudah.benar);
        }
        res = segarkanProgres(u, null, s);
      });
      return {
        nilai: nilai, benar: benar, jumlah: rincian.length, tuntas: nilai >= s.kktp, kktp: s.kktp, perMisi: perMisi,
        pembahasan: s.kuis_pembahasan ? bahas : null, kuis: res.kuis, stat: res.stat
      };
    };

    A.kirimRefleksi = function (d) {
      var u = auth(d.token, 'siswa');
      var row = {
        id: 'R' + env.random().slice(0, 10), waktu: nowIso(), username: str(u.username), nama: str(u.nama), kelas: str(u.kelas),
        perasaan: SB.clamp(Math.round(SB.num(d.perasaan, 3)), 1, 5),
        paham: str(d.paham).slice(0, 800), bingung: str(d.bingung).slice(0, 800), pertanyaan: str(d.pertanyaan).slice(0, 800),
        balasan: '', waktu_balasan: ''
      };
      if (!row.paham.trim() && !row.bingung.trim() && !row.pertanyaan.trim()) fail('Isi minimal satu kolom refleksi.');
      var res;
      env.lock(function () { T('Refleksi').insert(row); res = segarkanProgres(u, null); });
      return { stat: res.stat, refleksi: stateSiswa(u, settings()).refleksi };
    };

    A.peringkat = function (d) {
      var u = auth(d.token, 'siswa');
      var s = settings();
      if (!s.peringkat) fail('Papan peringkat sedang disembunyikan guru.');
      var aktif = {};
      users().forEach(function (a) { if (str(a.peran) === 'siswa' && SB.bool(a.aktif) && str(a.kelas) === str(u.kelas)) aktif[normUser(a.username)] = a; });
      var list = T('Progres').all().filter(function (r) { return aktif[normUser(r.username)]; })
        .map(function (r) { return { username: normUser(r.username), nama: str(aktif[normUser(r.username)].nama), xp: SB.num(r.xp), level: str(r.level), lencana: SB.num(r.lencana) }; })
        .sort(function (a, b) { return b.xp - a.xp || a.nama.localeCompare(b.nama); });
      var me = normUser(u.username), rank = 0;
      list.forEach(function (r, i) { r.rank = i + 1; if (r.username === me) rank = i + 1; });
      return { kelas: str(u.kelas), top: list.slice(0, 10).map(function (r) { return { rank: r.rank, nama: r.nama, xp: r.xp, level: r.level, aku: r.username === me }; }), rank: rank, total: list.length };
    };

    /* ----- guru ----- */
    A.dasbor = function (d) {
      var g = auth(d.token, 'guru');
      var s = settings();
      var prog = {};
      T('Progres').all().forEach(function (r) { prog[normUser(r.username)] = r; });
      var siswa = users().filter(function (a) { return str(a.peran) === 'siswa'; }).map(function (a) {
        var r = prog[normUser(a.username)] || {};
        return {
          username: str(a.username), nama: str(a.nama), kelas: str(a.kelas), aktif: SB.bool(a.aktif), sandi_awal: str(a.sandi_awal),
          kuis_tambahan: SB.num(a.kuis_tambahan), dibuat: str(a.dibuat), login_terakhir: str(a.login_terakhir),
          xp: SB.num(r.xp), level: str(r.level || LEVEL[0].nama),
          misi: [SB.num(r.misi1), SB.num(r.misi2), SB.num(r.misi3), SB.num(r.misi4)],
          kuis_terbaik: SB.num(r.kuis_terbaik), kuis_percobaan: SB.num(r.kuis_percobaan),
          game: { timbangan: SB.num(r.game_timbangan), akun: SB.num(r.game_akun), detektif: SB.num(r.game_detektif), juragan: SB.num(r.game_juragan) },
          lencana: SB.num(r.lencana), refleksi: SB.num(r.refleksi), diperbarui: str(r.diperbarui),
          data: parseJSON(r.data_json, null)
        };
      });
      var kuis = T('Kuis').all().map(function (r) {
        return { id: str(r.id), waktu: str(r.waktu), username: str(r.username), nama: str(r.nama), kelas: str(r.kelas), percobaan: SB.num(r.percobaan), nilai: SB.num(r.nilai), benar: SB.num(r.benar), jumlah: SB.num(r.jumlah), durasi_detik: SB.num(r.durasi_detik), rincian: parseJSON(r.rincian_json, []) };
      });
      var refleksi = T('Refleksi').all().map(function (r) {
        return { id: str(r.id), waktu: str(r.waktu), username: str(r.username), nama: str(r.nama), kelas: str(r.kelas), perasaan: SB.num(r.perasaan), paham: str(r.paham), bingung: str(r.bingung), pertanyaan: str(r.pertanyaan), balasan: str(r.balasan), waktu_balasan: str(r.waktu_balasan) };
      });
      return {
        settings: s, siswa: siswa, kuis: kuis, refleksi: refleksi, soal: SOAL,
        guru: { username: str(g.username), nama: str(g.nama), sandiBawaan: !!str(g.sandi_awal) }, info: env.info || {}, waktu: nowIso()
      };
    };

    A.ubahPengaturan = function (d) {
      auth(d.token, 'guru');
      var patch = d.patch || {}, boleh = {};
      SETTINGS_DEF.forEach(function (x) { if (x[0] !== 'versi') boleh[x[0]] = 1; });
      env.lock(function () {
        Object.keys(patch).forEach(function (k) {
          if (!boleh[k]) return;
          var v = patch[k];
          if (Array.isArray(v)) v = v.join(',');
          if (typeof v === 'boolean') v = v ? 'TRUE' : 'FALSE';
          writeSetting(k, str(v).slice(0, 1000));
        });
        writeSetting('versi', String(settings().versi + 1));
      });
      return { settings: settings() };
    };

    A.tambahSiswa = function (d) {
      auth(d.token, 'guru');
      var daftar = Array.isArray(d.daftar) ? d.daftar.slice(0, 300) : [];
      if (!daftar.length) fail('Daftar siswa masih kosong.');
      var hasil = [], lewat = [];
      env.lock(function () {
        var taken = {};
        users().forEach(function (a) { taken[normUser(a.username)] = 1; });
        daftar.forEach(function (o) {
          var nama = str(o.nama).trim().replace(/\s+/g, ' ').slice(0, 80);
          if (!nama) return;
          var un = normUser(o.username);
          if (un && !/^[a-z0-9._-]{3,30}$/.test(un)) { lewat.push(nama + ': username hanya boleh huruf kecil, angka, titik, minus (3-30 karakter)'); return; }
          if (un && taken[un]) { lewat.push(nama + ': username ' + un + ' sudah dipakai'); return; }
          if (!un) un = usernameDariNama(nama, taken);
          taken[un] = 1;
          var pw = str(o.password).trim() || genSandi(6);
          buatAkun({ username: un, nama: nama, kelas: str(o.kelas).trim().slice(0, 40), peran: 'siswa', password: pw });
          hasil.push({ nama: nama, kelas: str(o.kelas).trim(), username: un, password: pw });
        });
      });
      return { dibuat: hasil, dilewati: lewat };
    };

    function siswaTarget(d) {
      var u = findUser(d.username);
      if (!u || str(u.peran) !== 'siswa') fail('Siswa tidak ditemukan.');
      return u;
    }

    A.resetSandi = function (d) {
      auth(d.token, 'guru');
      var u = siswaTarget(d), pw = genSandi(6), salt = env.random().slice(0, 16);
      env.lock(function () { T('Akun').update('username', u.username, { salt: salt, hash: hashSandi(salt, pw), sandi_awal: pw }); });
      return { username: str(u.username), password: pw };
    };

    A.ubahSiswa = function (d) {
      auth(d.token, 'guru');
      var u = siswaTarget(d), patch = {};
      if (d.nama !== undefined && str(d.nama).trim()) patch.nama = str(d.nama).trim().slice(0, 80);
      if (d.kelas !== undefined) patch.kelas = str(d.kelas).trim().slice(0, 40);
      if (d.aktif !== undefined) patch.aktif = !!d.aktif;
      if (d.tambahKesempatan) patch.kuis_tambahan = Math.max(0, Math.round(SB.num(u.kuis_tambahan, 0)) + Math.round(SB.num(d.tambahKesempatan, 0)));
      env.lock(function () {
        T('Akun').update('username', u.username, patch);
        if (patch.nama || patch.kelas !== undefined) {
          var row = progresRow(u.username);
          if (row) T('Progres').update('username', row.username, { nama: patch.nama || str(u.nama), kelas: patch.kelas !== undefined ? patch.kelas : str(u.kelas) });
        }
      });
      return { ok: true };
    };

    A.resetProgres = function (d) {
      auth(d.token, 'guru');
      var u = siswaTarget(d);
      env.lock(function () { segarkanProgres(u, {}); });
      return { ok: true };
    };

    A.hapusSiswa = function (d) {
      auth(d.token, 'guru');
      var u = siswaTarget(d);
      env.lock(function () {
        T('Akun').remove('username', u.username);
        T('Progres').remove('username', u.username);
      });
      return { ok: true };
    };

    A.balasRefleksi = function (d) {
      auth(d.token, 'guru');
      var ok;
      env.lock(function () { ok = T('Refleksi').update('id', str(d.id), { balasan: str(d.balasan).slice(0, 800), waktu_balasan: nowIso() }); });
      if (!ok) fail('Refleksi tidak ditemukan.');
      return { ok: true };
    };


    /* =========================================================
       JURAGAN SEIMBANG — papan bisnis online (state di cache)
       ========================================================= */
    var J = JURAGAN;
    var BOT_NAMA = ['Komputer Sari', 'Komputer Bayu', 'Komputer Rani'];
    function rpS(n) { n = Math.round(Number(n) || 0); return (n < 0 ? '−' : '') + 'Rp' + Math.abs(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
    function jrKey(k) { return 'jr:' + k; }
    function jrLoad(kode) { var t = env.cache.get(jrKey(str(kode).toUpperCase())); return t ? JSON.parse(t) : null; }
    function jrSave(r) { r.versi = (r.versi || 0) + 1; r.diubah = nowMs(); env.cache.put(jrKey(r.kode), JSON.stringify(r), 21600); }
    function jrHapus(r) { env.cache.remove(jrKey(r.kode)); }
    function modalP(p) { return p.kas + p.aset - p.utang; }
    function cur(r) { return r.pemain[r.giliran]; }
    function logx(r, t, s) { r.log.push({ t: t, s: s || '', n: r.seq = (r.seq || 0) + 1 }); if (r.log.length > 40) r.log.shift(); }
    function setFase(r, fase, aksi) { r.fase = fase; r.aksi = aksi || null; r.faseMulai = nowMs(); }
    function idxUrut(n) { var a = []; for (var i = 0; i < n; i++) a.push(i); return shuffle(a, rnd); }
    var ARAH_KATA = { '1': 'naik', '0': 'tetap', '-1': 'turun' };
    function kataArah(a) { return 'Harta ' + ARAH_KATA[a[0]] + ', utang ' + ARAH_KATA[a[1]] + ', modal ' + ARAH_KATA[a[2]] + '.'; }

    function pemainId(d) {
      if (d.token) { var parts = cekToken(d.token, 'gagal'); return { pid: normUser(parts[1]), peran: parts[2] }; }
      if (d.tamu && /^[a-z0-9]{6,20}$/.test(str(d.tamu))) return { pid: 'tamu:' + str(d.tamu), tamu: true };
      fail('Masuk dulu dengan akunmu atau sebagai tamu untuk bermain.');
    }
    function namaPemain(id) {
      if (id.tamu) { if (!settings().tamu) fail('Mode tamu sedang dimatikan guru.'); return 'Tamu ' + id.pid.slice(-3).toUpperCase(); }
      var u = findUser(id.pid);
      if (!u || !SB.bool(u.aktif)) fail('Akun tidak ditemukan atau dinonaktifkan.');
      return str(u.nama);
    }
    function pemainBaru(pid, nama, bot, w) { return { pid: pid, nama: nama, bot: !!bot, w: w, pos: 0, kas: J.kasAwal, aset: 0, utang: 0, lewati: false, benar: 0, salah: 0 }; }
    function warnaBebas(r) { for (var w = 0; w < 4; w++) { if (!r.pemain.some(function (p) { return p.w === w; })) return w; } return r.pemain.length % 4; }
    function tambahBot(r) {
      var dipakai = r.pemain.map(function (p) { return p.nama; });
      var nm = BOT_NAMA.filter(function (n) { return dipakai.indexOf(n) < 0; })[0] || ('Komputer ' + (r.pemain.length + 1));
      r.pemain.push(pemainBaru('bot:' + env.random().slice(0, 6), nm, true, warnaBebas(r)));
    }
    function mulaiMain(r) {
      r.status = 'main'; r.ronde = 1; r.giliran = 0; r.dadu = 0;
      r.kartuUrut = idxUrut(J.kartu.length); r.kartuI = 0; r.soalUrut = idxUrut(SOAL_GAME.length); r.soalI = 0;
      logx(r, 'Permainan dimulai! ' + r.maksRonde + ' ronde. Pemenang: modal akhir terbesar.', 'ok');
      setFase(r, 'lempar', null);
    }
    function sewaPetak(r, i) {
      var t = J.petak[i], m = r.milik[i], s = t.w * (m.lv >= 2 ? 2 : 1);
      var grup = J.petak.map(function (x, k) { return k; }).filter(function (k) { return J.petak[k].g === t.g; });
      if (grup.every(function (k) { return r.milik[k] && r.milik[k].p === m.p; })) s *= 2;
      return s;
    }
    function pastikanKas(r, p, butuh) {
      if (p.kas >= butuh) return;
      var pinjam = Math.ceil((butuh - p.kas) / 1000000) * 1000000;
      p.kas += pinjam; p.utang += pinjam;
      logx(r, p.nama + ' kekurangan kas dan meminjam darurat ' + rpS(pinjam) + ' ke bank.', 'bad');
    }
    function tebak(r, teks, arah, why) { setFase(r, 'tebak', { tipe: 'tebak', teks: teks, arah: arah, why: why }); }
    function soalJr(r, jenis) {
      var q = SOAL_GAME[r.soalUrut[r.soalI % r.soalUrut.length]]; r.soalI++;
      var perm = shuffle(q.o.map(function (x, i) { return i; }), rnd);
      setFase(r, 'soal', { tipe: jenis, q: q.q, o: perm.map(function (i) { return q.o[i]; }), kunci: perm.indexOf(q.a) });
    }
    function lempar(r) {
      var p = cur(r), d = 1 + Math.floor(rnd() * 6);
      r.dadu = d; r.daduN = (r.daduN || 0) + 1; p.dari = p.pos;
      var ke = (p.pos + d) % J.petak.length;
      logx(r, p.nama + ' melempar dadu: ' + d + '.');
      if (p.pos + d >= J.petak.length) { p.kas += J.gaji; logx(r, p.nama + ' melewati MULAI dan menerima pendapatan jasa ' + rpS(J.gaji) + '.', 'ok'); }
      p.pos = ke;
      selesaikanPetak(r);
    }
    function selesaikanPetak(r) {
      var p = cur(r), t = J.petak[p.pos];
      if (t.tipe === 'usaha') {
        var m = r.milik[p.pos];
        if (!m) {
          var opsi = [];
          if (p.kas >= t.h) opsi.push('tunai');
          if (p.utang + t.h <= J.batasUtang) opsi.push('kredit');
          opsi.push('lewati');
          setFase(r, 'pilih', { tipe: 'beli', petak: p.pos, opsi: opsi });
        } else if (m.p === r.giliran) {
          if (m.lv < 2 && p.kas >= t.h / 2) setFase(r, 'pilih', { tipe: 'renov', petak: p.pos, opsi: ['renov', 'lewati'] });
          else { logx(r, p.nama + ' singgah di usahanya sendiri, ' + t.n + '.'); akhiriGiliran(r); }
        } else {
          var own = r.pemain[m.p], sewa = sewaPetak(r, p.pos);
          pastikanKas(r, p, sewa);
          p.kas -= sewa; own.kas += sewa;
          logx(r, p.nama + ' membayar sewa ' + t.n + ' ' + rpS(sewa) + ' kepada ' + own.nama + '.', 'bad');
          tebak(r, 'Kamu membayar sewa ' + t.n + ' sebesar ' + rpS(sewa) + ' kepada ' + own.nama + '.', [-1, 0, -1], 'Kas berkurang dan beban sewa mengurangi modal.');
        }
      } else if (t.tipe === 'kartu') {
        var c = null;
        for (var i = 0; i < J.kartu.length; i++) {
          c = J.kartu[r.kartuUrut[r.kartuI % J.kartu.length]]; r.kartuI++;
          if (!c.butuhUtang || p.utang >= c.butuhUtang) break;
        }
        if (c.k < 0) pastikanKas(r, p, -c.k);
        p.kas += c.k; p.aset += c.a; p.utang += c.u;
        logx(r, p.nama + ' mengambil kartu: ' + c.t);
        tebak(r, c.t, c.arah, kataArah(c.arah));
      } else if (t.tipe === 'pos') {
        logx(r, p.nama + ' tiba di Pos Akuntan dan menjawab soal.');
        soalJr(r, 'pos');
      } else if (t.tipe === 'audit') {
        logx(r, p.nama + ' diperiksa auditor. Jawaban salah membuatnya melewati satu giliran.');
        soalJr(r, 'audit');
      } else if (t.tipe === 'pajak') {
        pastikanKas(r, p, J.pajak);
        p.kas -= J.pajak;
        logx(r, p.nama + ' membayar pajak UMKM ' + rpS(J.pajak) + '.', 'bad');
        tebak(r, 'Kamu membayar pajak UMKM ' + rpS(J.pajak) + ' secara tunai.', [-1, 0, -1], 'Kas berkurang dan beban pajak mengurangi modal.');
      } else if (t.tipe === 'bank') {
        var ob = [];
        if (p.utang + J.pinjam <= J.batasUtang) ob.push('pinjam');
        if (p.utang >= J.pinjam && p.kas >= J.pinjam) ob.push('lunasi');
        ob.push('lewati');
        if (ob.length === 1) { logx(r, p.nama + ' mampir ke bank.'); akhiriGiliran(r); }
        else setFase(r, 'pilih', { tipe: 'bank', petak: p.pos, opsi: ob });
      } else {
        logx(r, p.nama + (t.tipe === 'istirahat' ? ' beristirahat sejenak di taman.' : ' berada di petak MULAI.'));
        akhiriGiliran(r);
      }
    }
    function pilihJr(r, opsi) {
      var a = r.aksi;
      if (!a || !a.opsi || a.opsi.indexOf(opsi) < 0) fail('Pilihan itu tidak tersedia.');
      var p = cur(r), t = J.petak[a.petak];
      if (opsi === 'lewati') { logx(r, p.nama + ' memilih tidak melakukan apa-apa.'); akhiriGiliran(r); return; }
      if (a.tipe === 'beli') {
        r.milik[a.petak] = { p: r.giliran, lv: 1 };
        if (opsi === 'tunai') {
          p.kas -= t.h; p.aset += t.h;
          logx(r, p.nama + ' membeli ' + t.n + ' secara tunai ' + rpS(t.h) + '.', 'ok');
          tebak(r, 'Kamu membeli usaha ' + t.n + ' secara tunai ' + rpS(t.h) + '.', [0, 0, 0], 'Kas berkurang dan aset usaha bertambah sama besar, jadi total harta tetap.');
        } else {
          p.aset += t.h; p.utang += t.h;
          logx(r, p.nama + ' membeli ' + t.n + ' secara kredit ' + rpS(t.h) + ' (pinjaman bank).', 'ok');
          tebak(r, 'Kamu membeli usaha ' + t.n + ' ' + rpS(t.h) + ' dengan pinjaman bank.', [1, 1, 0], 'Aset usaha bertambah dan utang bank bertambah sama besar.');
        }
      } else if (a.tipe === 'renov') {
        var biaya = t.h / 2;
        p.kas -= biaya; p.aset += biaya; r.milik[a.petak].lv = 2;
        logx(r, p.nama + ' merenovasi ' + t.n + '. Sewanya kini dua kali lipat.', 'ok');
        tebak(r, 'Kamu merenovasi ' + t.n + ' secara tunai ' + rpS(biaya) + '.', [0, 0, 0], 'Kas berkurang dan nilai aset usaha bertambah sama besar, jadi total harta tetap.');
      } else if (a.tipe === 'bank') {
        if (opsi === 'pinjam') {
          p.kas += J.pinjam; p.utang += J.pinjam;
          logx(r, p.nama + ' meminjam ' + rpS(J.pinjam) + ' ke bank.');
          tebak(r, 'Kamu menerima pinjaman dari bank ' + rpS(J.pinjam) + '.', [1, 1, 0], 'Kas bertambah dan utang bank bertambah. Pinjaman bukan pendapatan.');
        } else {
          p.kas -= J.pinjam; p.utang -= J.pinjam;
          logx(r, p.nama + ' melunasi utang bank ' + rpS(J.pinjam) + '.', 'ok');
          tebak(r, 'Kamu membayar utang bank ' + rpS(J.pinjam) + '.', [-1, -1, 0], 'Kas berkurang dan utang bank berkurang. Membayar utang bukan beban.');
        }
      }
    }
    function jawabJr(r, ans) {
      var a = r.aksi, p = cur(r), out;
      if (r.fase === 'tebak') {
        var ok = Array.isArray(ans) && ans.length === 3 && ans.every(function (v, i) { return Number(v) === a.arah[i]; });
        if (ok) { p.kas += J.bonus; p.benar++; logx(r, p.nama + ' mencatat pengaruhnya dengan tepat. Bonus ketelitian ' + rpS(J.bonus) + '.', 'ok'); }
        else { p.salah++; logx(r, p.nama + ' belum tepat. Seharusnya: ' + kataArah(a.arah), 'bad'); }
        out = { benar: ok, arah: a.arah, why: a.why, bonus: ok ? J.bonus : 0 };
      } else {
        var ok2 = Number(ans) === a.kunci;
        if (ok2) {
          p.benar++;
          var hadiah = a.tipe === 'audit' ? J.bonusAudit : J.hadiahPos;
          p.kas += hadiah;
          logx(r, p.nama + ' menjawab benar dan mendapat hadiah ' + rpS(hadiah) + '.', 'ok');
          out = { benar: true, kunci: a.kunci, hadiah: hadiah };
        } else {
          p.salah++;
          if (a.tipe === 'audit') p.lewati = true;
          logx(r, p.nama + ' menjawab salah. Jawaban benar: ' + a.o[a.kunci] + '.' + (a.tipe === 'audit' ? ' ' + p.nama + ' melewati giliran berikutnya.' : ''), 'bad');
          out = { benar: false, kunci: a.kunci, teksKunci: a.o[a.kunci], lewati: a.tipe === 'audit' };
        }
      }
      akhiriGiliran(r);
      return out;
    }
    function akhiriGiliran(r) {
      var n = r.pemain.length, i = r.giliran;
      for (var guard = 0; guard < n * 3; guard++) {
        i = (i + 1) % n;
        if (i === 0) r.ronde++;
        if (r.ronde > r.maksRonde) { selesaiJr(r); return; }
        if (r.pemain[i].lewati) { r.pemain[i].lewati = false; logx(r, r.pemain[i].nama + ' melewati giliran karena pemeriksaan audit.', 'bad'); continue; }
        break;
      }
      r.giliran = i;
      setFase(r, 'lempar', null);
    }
    function selesaiJr(r) {
      r.status = 'selesai'; setFase(r, 'selesai', null);
      r.hasil = r.pemain.map(function (p, i) { return { i: i, nama: p.nama, bot: p.bot, w: p.w, kas: p.kas, aset: p.aset, utang: p.utang, modal: modalP(p), benar: p.benar, salah: p.salah }; })
        .sort(function (a, b) { return b.modal - a.modal; });
      logx(r, 'Permainan selesai! Juragan terbesar: ' + r.hasil[0].nama + ' dengan modal ' + rpS(r.hasil[0].modal) + '.', 'ok');
      r.pemain.forEach(function (p) {
        if (p.bot || /^tamu:|^bot:/.test(p.pid)) return;
        var u = findUser(p.pid);
        if (!u || str(u.peran) !== 'siswa') return;
        var skor = Math.max(0, Math.round(modalP(p) / 1000));
        var data = SB.cleanData(parseJSON((progresRow(u.username) || {}).data_json, {}));
        if (skor > (data.game.juragan || 0)) { data.game.juragan = skor; segarkanProgres(u, data); }
      });
    }
    function botLangkah(r) {
      var p = cur(r), a = r.aksi;
      if (r.fase === 'lempar') return lempar(r);
      if (r.fase === 'pilih') {
        var o = a.opsi, pick = 'lewati';
        if (a.tipe === 'beli') {
          var h = J.petak[a.petak].h;
          if (o.indexOf('tunai') >= 0 && p.kas - h >= 3000000) pick = 'tunai';
          else if (o.indexOf('kredit') >= 0 && p.utang < 8000000 && rnd() < 0.6) pick = 'kredit';
        } else if (a.tipe === 'renov') { if (p.kas - J.petak[a.petak].h / 2 >= 4000000) pick = 'renov'; }
        else if (a.tipe === 'bank') {
          if (o.indexOf('lunasi') >= 0 && p.kas > 9000000) pick = 'lunasi';
          else if (o.indexOf('pinjam') >= 0 && p.kas < 2000000) pick = 'pinjam';
        }
        return pilihJr(r, pick);
      }
      if (r.fase === 'tebak') {
        var ans = a.arah.slice();
        if (rnd() > 0.72) { var k = Math.floor(rnd() * 3); ans[k] = ans[k] === 0 ? 1 : 0; }
        return jawabJr(r, ans);
      }
      if (r.fase === 'soal') return jawabJr(r, rnd() < 0.7 ? a.kunci : (a.kunci + 1) % a.o.length);
    }
    function perluTick(r) {
      if (!r || r.status !== 'main') return false;
      var p = cur(r), dt = nowMs() - r.faseMulai;
      return p.bot ? dt >= 1500 : dt > J.batasWaktu;
    }
    function tick(r) {
      var p = cur(r);
      if (p.bot) { botLangkah(r); return; }
      logx(r, 'Waktu giliran ' + p.nama + ' habis.', 'bad');
      if (r.fase === 'lempar') lempar(r);
      else if (r.fase === 'pilih') pilihJr(r, 'lewati');
      else if (r.fase === 'tebak') jawabJr(r, null);
      else if (r.fase === 'soal') jawabJr(r, -1);
    }
    function jrView(r, pid) {
      var v = JSON.parse(JSON.stringify(r));
      delete v.kartuUrut; delete v.soalUrut;
      if (v.aksi) { delete v.aksi.arah; delete v.aksi.kunci; delete v.aksi.why; }
      v.saya = -1;
      v.pemain.forEach(function (p, i) { p.modal = modalP(p); if (p.pid === pid) v.saya = i; p.host = p.pid === r.host; p.pid = p.pid.indexOf('bot:') === 0 ? 'bot' : (p.pid === pid ? pid : 'x' + i); });
      v.now = nowMs();
      return v;
    }
    function ambilRuang(kode) {
      var r = jrLoad(kode);
      if (!r) fail('Ruang ' + str(kode).toUpperCase() + ' tidak ditemukan atau sudah berakhir.');
      return r;
    }
    function ubahRuang(kode, pid, fn) {
      var r, out;
      env.lock(function () {
        r = ambilRuang(kode);
        out = fn(r);
        jrSave(r);
      });
      return { ruang: jrView(r, pid), umpan: out || null };
    }

    A.jrBuat = function (d) {
      var id = pemainId(d), nama = namaPemain(id);
      var al = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', kode = '';
      for (var tries = 0; tries < 20; tries++) {
        kode = ''; for (var i = 0; i < 4; i++) kode += al.charAt(Math.floor(rnd() * al.length));
        if (!jrLoad(kode)) break;
      }
      var r = { kode: kode, host: id.pid, status: 'tunggu', dibuat: nowMs(), versi: 0, maksRonde: SB.clamp(Math.round(SB.num(d.ronde, 8)), 3, 15), ronde: 1, giliran: 0,
        pemain: [pemainBaru(id.pid, nama, false, 0)], milik: {}, fase: 'tunggu', aksi: null, dadu: 0, log: [], seq: 0, faseMulai: nowMs(), hasil: null };
      logx(r, nama + ' membuat ruang ' + kode + '.');
      var nb = SB.clamp(Math.round(SB.num(d.bot, 0)), 0, 3);
      for (var b = 0; b < nb; b++) tambahBot(r);
      if (nb > 0 && d.langsung) mulaiMain(r);
      env.lock(function () { jrSave(r); });
      return { ruang: jrView(r, id.pid) };
    };
    A.jrGabung = function (d) {
      var id = pemainId(d), kode = str(d.kode).trim().toUpperCase();
      var r0 = ambilRuang(kode);
      if (r0.pemain.some(function (p) { return p.pid === id.pid; })) return { ruang: jrView(r0, id.pid) };
      var nama = namaPemain(id);
      return ubahRuang(kode, id.pid, function (r) {
        if (r.pemain.some(function (p) { return p.pid === id.pid; })) return;
        if (r.status !== 'tunggu') fail('Permainan di ruang ' + kode + ' sudah dimulai.');
        if (r.pemain.length >= J.maksPemain) fail('Ruang ' + kode + ' sudah penuh (' + J.maksPemain + ' pemain).');
        r.pemain.push(pemainBaru(id.pid, nama, false, warnaBebas(r)));
        logx(r, nama + ' bergabung.', 'ok');
      });
    };
    function hostSaja(r, pid) { if (r.host !== pid) fail('Hanya pembuat ruang yang bisa melakukan ini.'); }
    A.jrBot = function (d) {
      var id = pemainId(d);
      return ubahRuang(d.kode, id.pid, function (r) {
        hostSaja(r, id.pid);
        if (r.status !== 'tunggu') fail('Permainan sudah dimulai.');
        if (r.pemain.length >= J.maksPemain) fail('Ruang sudah penuh.');
        tambahBot(r); logx(r, r.pemain[r.pemain.length - 1].nama + ' ikut bermain.');
      });
    };
    A.jrMulai = function (d) {
      var id = pemainId(d);
      return ubahRuang(d.kode, id.pid, function (r) {
        hostSaja(r, id.pid);
        if (r.status !== 'tunggu') fail('Permainan sudah dimulai.');
        if (r.pemain.length < 2) fail('Butuh minimal 2 pemain. Ajak teman atau tambah pemain komputer.');
        if (d.ronde) r.maksRonde = SB.clamp(Math.round(SB.num(d.ronde, 8)), 3, 15);
        mulaiMain(r);
      });
    };
    A.jrState = function (d) {
      var id = pemainId(d), r = ambilRuang(d.kode);
      if (perluTick(r)) {
        env.lock(function () {
          r = ambilRuang(d.kode);
          if (perluTick(r)) { tick(r); jrSave(r); }
        });
      }
      if (Number(d.v) === r.versi) return { sama: true, now: nowMs() };
      return { ruang: jrView(r, id.pid) };
    };
    function giliranSaya(r, pid, fase) {
      if (r.status !== 'main') fail('Permainan belum berjalan.');
      if (cur(r).pid !== pid) fail('Belum giliranmu.');
      if (fase.indexOf(r.fase) < 0) fail('Langkah ini sudah tidak berlaku. Tunggu sebentar.');
    }
    A.jrLempar = function (d) {
      var id = pemainId(d);
      return ubahRuang(d.kode, id.pid, function (r) { giliranSaya(r, id.pid, ['lempar']); lempar(r); });
    };
    A.jrPilih = function (d) {
      var id = pemainId(d);
      return ubahRuang(d.kode, id.pid, function (r) { giliranSaya(r, id.pid, ['pilih']); pilihJr(r, str(d.opsi)); });
    };
    A.jrJawab = function (d) {
      var id = pemainId(d);
      return ubahRuang(d.kode, id.pid, function (r) { giliranSaya(r, id.pid, ['tebak', 'soal']); return jawabJr(r, d.jawab); });
    };
    A.jrKeluar = function (d) {
      var id = pemainId(d);
      env.lock(function () {
        var r = jrLoad(d.kode); if (!r) return;
        var i = -1; r.pemain.forEach(function (p, k) { if (p.pid === id.pid) i = k; });
        if (i < 0) return;
        var p = r.pemain[i];
        if (r.status === 'tunggu') {
          r.pemain.splice(i, 1);
          logx(r, p.nama + ' keluar dari ruang.');
        } else if (r.status === 'main') {
          p.bot = true; p.nama += ' (komputer)';
          logx(r, p.nama.replace(' (komputer)', '') + ' keluar. Komputer melanjutkan permainannya.', 'bad');
        }
        var manusia = r.pemain.filter(function (q) { return !q.bot; });
        if (!manusia.length) { jrHapus(r); return; }
        if (r.host === id.pid) r.host = manusia[0].pid;
        jrSave(r);
      });
      return { ok: true };
    };

    function handle(req) {
      req = req || {};
      try {
        var fn = A[req.action];
        if (!fn) fail('Aksi tidak dikenal.');
        var out = fn(req.data || {}) || {};
        out.ok = true;
        return out;
      } catch (e) {
        return { ok: false, error: e.userMsg || ('Terjadi kesalahan di server: ' + (e && e.message ? e.message : e)), code: e.code || 'server' };
      }
    }

    return {
      handle: handle,
      // untuk setup & data contoh
      _buatAkun: buatAkun, _writeSetting: writeSetting, _segarkan: segarkanProgres, _findUser: findUser, _settings: settings
    };
  }

  return { create: create, SCHEMA: SCHEMA, TEXT_COLS: TEXT_COLS, sha256: sha256, hmac: hmac, hash: hash };
})();

