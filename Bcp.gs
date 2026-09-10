/**
 * ============================================================================
 * Modul BCP — Point Performance (Post-Fraud)
 * ============================================================================
 * File terpisah dari Code.gs, tapi berbagi spreadsheet dan konstanta yang sama.
 * Apps Script menggabungkan semua file .gs dalam satu project, jadi fungsi
 * di Code.gs (sheetOrCreate_, normalisasiNilai_, dst) bisa dipakai di sini.
 *
 * SUMBER: BCP Summary Case -> tab "Recovery Progress"
 *   Baris 2 : nama blok minggu (Before Fraud, W4 Jun, W1 Jul, ...)
 *   Baris 3 : nama kolom dalam blok (BM, MPP BP, # BP, Mitra Lancar, ...)
 *   Baris 4+: data, satu baris per point
 *
 * Struktur blok berulang tiap minggu dan BERTAMBAH ke kanan setiap update.
 * Script membaca blok secara dinamis, jadi minggu baru otomatis terbawa
 * tanpa perlu mengubah kode.
 *
 * SETUP TAMBAHAN (Script Properties):
 *   BCP_SPREADSHEET_ID = 1eOgwChewd0HETRnIImwEgjJE7wS_TPpkGKdHSLgog-0
 * ============================================================================
 */

// ─── Konstanta BCP ──────────────────────────────────────────────────────────

var ARSIP_BCP     = 'Arsip BCP';
var BCP_TAB       = 'Recovery Progress';

// Baris di source BCP.
var BCP_ROW_BLOK   = 2;   // nama blok minggu
var BCP_ROW_HEADER = 3;   // nama kolom dalam blok
var BCP_ROW_DATA   = 4;   // data mulai di sini

// Kolom identitas yang ada di luar blok mingguan.
var BCP_KOLOM_IDENTITAS = [
  'Point', 'Fraud Month', 'Fraud Week', 'Jumlah Loan Terdampak',
  'Lost Amount', 'Recovery Pre-Investigation'
];

/**
 * Kendali apa saja yang dikirim ke backend. Ubah nilainya lalu Deploy ulang
 * (Manage deployments -> Edit -> New version). Tidak perlu ubah backend.
 */
var BCP_CONFIG = {
  // Kirim data BCP ke branch_data.
  aktif: true,

  // Kirim agregat KPI point (dari Code.gs) ke branch_data.
  // Matikan kalau agregat sudah tidak diperlukan — data BCP tetap jalan.
  kirimAgregatKpi: true,

  // true  = semua blok minggu dikirim (tren penuh, konteks lebih besar)
  // false = hanya Before Fraud + minggu terakhir + blok perbandingan
  semuaMinggu: true,

  // Batas jumlah blok minggu yang dikirim ketika semuaMinggu = true.
  // Mencegah konteks membengkak tanpa batas setelah berbulan-bulan.
  // Blok paling lama yang dibuang lebih dulu; Before Fraud selalu ikut.
  maksBlokMinggu: 3,

  // Role yang boleh menerima data BCP. Kosongkan array untuk semua role.
  // Contoh membatasi: ['BM', 'AM', 'RM']
  roleDiizinkan: []
};

// ─── Menu ───────────────────────────────────────────────────────────────────

function menuBcp_(ui) {
  return ui.createMenu('BCP Point Performance')
    .addItem('Capture BCP', 'captureBcp')
    .addItem('Cek Koneksi BCP', 'cekKoneksiBcp')
    .addItem('Ringkasan BCP', 'ringkasanBcp');
}

function cekKoneksiBcp() {
  var ui = SpreadsheetApp.getUi();
  var id = PropertiesService.getScriptProperties().getProperty('BCP_SPREADSHEET_ID');

  if (!id) {
    ui.alert('BCP_SPREADSHEET_ID belum diisi di Script Properties.');
    return;
  }

  try {
    var src = SpreadsheetApp.openById(id);
    var sheet = src.getSheetByName(BCP_TAB);
    if (!sheet) {
      ui.alert('Tab "' + BCP_TAB + '" tidak ditemukan.\n\nTab tersedia:\n' +
        src.getSheets().map(function (s) { return s.getName(); }).join('\n'));
      return;
    }

    var blok = bacaBlokBcp_(sheet);
    ui.alert(
      'Koneksi OK\n\nSumber: ' + src.getName() +
      '\nTab: ' + BCP_TAB +
      '\nJumlah point: ' + (sheet.getLastRow() - BCP_ROW_DATA + 1) +
      '\n\nBlok minggu terdeteksi (' + blok.length + '):\n' +
      blok.map(function (b) {
        return '  ' + b.nama + '  (kolom ' + b.mulai + '–' + b.akhir + ')';
      }).join('\n')
    );
  } catch (e) {
    ui.alert('Gagal membuka BCP.\n\n' + e.message);
  }
}

// ─── Deteksi blok ───────────────────────────────────────────────────────────

/**
 * Baca baris 2 untuk menemukan batas tiap blok minggu.
 * Nama blok hanya muncul di kolom pertama tiap blok (merged cell),
 * jadi blok dianggap berlanjut sampai muncul nama blok berikutnya.
 */
function bacaBlokBcp_(sheet) {
  var lastCol = sheet.getLastColumn();
  var barisBlok = sheet.getRange(BCP_ROW_BLOK, 1, 1, lastCol).getDisplayValues()[0];

  var blok = [];
  var aktif = null;

  for (var i = 0; i < lastCol; i++) {
    var nama = String(barisBlok[i]).replace(/\s+/g, ' ').trim();
    if (nama) {
      if (aktif) { aktif.akhir = i; blok.push(aktif); }
      aktif = { nama: nama, mulai: i + 1, akhir: lastCol };
    }
  }
  if (aktif) blok.push(aktif);

  // Konversi index 0-based ke nomor kolom 1-based untuk ditampilkan.
  return blok.map(function (b) {
    return { nama: b.nama, mulai: b.mulai, akhir: b.akhir };
  });
}

/** Blok perbandingan dikenali dari kata kunci, bukan pola minggu. */
function isBlokPerbandingan_(nama) {
  return /progressing|vs|perbandingan|preinvestigation/i.test(nama);
}

/** Blok baseline sebelum insiden. */
function isBlokBaseline_(nama) {
  return /before\s*fraud/i.test(nama);
}

// ─── Capture ────────────────────────────────────────────────────────────────

function captureBcp() {
  var ui = SpreadsheetApp.getUi();
  var id = PropertiesService.getScriptProperties().getProperty('BCP_SPREADSHEET_ID');

  if (!id) {
    ui.alert('BCP_SPREADSHEET_ID belum diisi di Script Properties.');
    return;
  }

  try {
    var hasil = captureBcpDari_(id);
    ui.alert('Capture BCP Selesai', hasil, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Capture BCP gagal.\n\n' + e.message);
  }
}

function captureBcpDari_(sourceId) {
  var src = SpreadsheetApp.openById(sourceId);
  var sheet = src.getSheetByName(BCP_TAB);
  if (!sheet) throw new Error('Tab "' + BCP_TAB + '" tidak ditemukan.');

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < BCP_ROW_DATA) throw new Error('Tab BCP kosong.');

  var blok    = bacaBlokBcp_(sheet);
  var headers = sheet.getRange(BCP_ROW_HEADER, 1, 1, lastCol).getDisplayValues()[0];
  var rows    = sheet.getRange(BCP_ROW_DATA, 1, lastRow - BCP_ROW_DATA + 1, lastCol)
                     .getDisplayValues();

  // Kolom identitas berada di luar blok mingguan — cari berdasarkan nama.
  var idxIdentitas = {};
  BCP_KOLOM_IDENTITAS.forEach(function (nama) {
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).replace(/\s+/g, ' ').trim().toLowerCase()
          === nama.toLowerCase()) {
        idxIdentitas[nama] = i;
        return;
      }
    }
  });

  if (idxIdentitas['Point'] === undefined) {
    throw new Error('Kolom "Point" tidak ditemukan di baris ' + BCP_ROW_HEADER + '.');
  }

  // Susun header arsip: identitas, lalu "NamaBlok - NamaKolom" untuk tiap blok.
  var arsipHeaders = ['captured_at'];
  BCP_KOLOM_IDENTITAS.forEach(function (n) {
    if (idxIdentitas[n] !== undefined) arsipHeaders.push(n);
  });

  var petaKolom = [];   // { label, idx }
  blok.forEach(function (b) {
    for (var c = b.mulai - 1; c < b.akhir; c++) {
      var h = String(headers[c]).replace(/\s+/g, ' ').trim();
      if (!h) continue;
      // Lewati kolom identitas yang kebetulan masuk rentang blok.
      if (BCP_KOLOM_IDENTITAS.indexOf(h) >= 0) continue;
      var label = b.nama + ' - ' + h;
      petaKolom.push({ label: label, idx: c });
      arsipHeaders.push(label);
    }
  });

  var now = new Date();
  var out = [];
  rows.forEach(function (r) {
    var point = String(r[idxIdentitas['Point']]).trim();
    if (!point) return;

    var rec = [now];
    BCP_KOLOM_IDENTITAS.forEach(function (n) {
      if (idxIdentitas[n] !== undefined) rec.push(r[idxIdentitas[n]]);
    });
    petaKolom.forEach(function (p) { rec.push(r[p.idx]); });
    out.push(rec);
  });

  if (!out.length) throw new Error('Tidak ada baris point yang terbaca.');

  // Arsip BCP ditulis ulang penuh setiap capture. Berbeda dari arsip KPI
  // yang append per periode: di sini seluruh tabel adalah satu snapshot
  // terbaru yang kolomnya bertambah tiap minggu.
  var arsip = sheetOrCreate_(ARSIP_BCP);
  arsip.clear();
  arsip.getRange(1, 1, 1, arsipHeaders.length)
       .setValues([arsipHeaders]).setFontWeight('bold');
  arsip.setFrozenRows(1);
  arsip.getRange(2, 1, out.length, arsipHeaders.length).setValues(out);

  var namaBlok = blok.map(function (b) { return b.nama; });
  return 'Point tersimpan: ' + out.length +
         '\nKolom total: ' + arsipHeaders.length +
         '\n\nBlok terdeteksi (' + blok.length + '):\n  ' +
         namaBlok.join('\n  ') +
         '\n\nJalankan sync backend untuk mengirim ke Ava.';
}

// ─── Payload untuk doGet ────────────────────────────────────────────────────

/**
 * Bentuk objek per point dari Arsip BCP, siap digabung ke payload branches.
 * Dipanggil dari doGet() di Code.gs.
 */
function buildBcpBranches_() {
  if (!BCP_CONFIG.aktif) return {};

  var sheet = ss_().getSheetByName(ARSIP_BCP);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  var iPoint = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === 'point') { iPoint = i; break; }
  }
  if (iPoint < 0) return {};

  // Tentukan blok mana saja yang dikirim.
  var blokDipakai = pilihBlok_(headers);

  var hasil = {};
  data.forEach(function (r) {
    var point = String(r[iPoint]).trim();
    if (!point) return;

    var rec = {};
    headers.forEach(function (h, idx) {
      var label = String(h).trim();
      if (!label || label === 'captured_at' || idx === iPoint) return;

      var v = r[idx];
      if (v === '' || v === null || v === undefined) return;

      // Kolom identitas selalu ikut.
      if (BCP_KOLOM_IDENTITAS.indexOf(label) >= 0) {
        rec['BCP ' + label] = normalisasiNilai_(v, label);
        return;
      }

      // Kolom blok: hanya kirim yang bloknya terpilih.
      var namaBlok = label.split(' - ')[0];
      if (blokDipakai.indexOf(namaBlok) < 0) return;
      rec['BCP ' + label] = normalisasiNilai_(v, label);
    });

    if (Object.keys(rec).length) hasil[point] = rec;
  });

  return hasil;
}

/**
 * Tentukan blok mana yang dikirim berdasarkan BCP_CONFIG.
 * Baseline dan blok perbandingan selalu ikut karena keduanya adalah
 * acuan untuk menilai pemulihan.
 */
function pilihBlok_(headers) {
  var urut = [];
  headers.forEach(function (h) {
    var label = String(h).trim();
    if (!label || label === 'captured_at') return;
    if (BCP_KOLOM_IDENTITAS.indexOf(label) >= 0) return;
    var nama = label.split(' - ')[0];
    if (nama && urut.indexOf(nama) < 0) urut.push(nama);
  });

  var baseline     = urut.filter(isBlokBaseline_);
  var perbandingan = urut.filter(isBlokPerbandingan_);
  var mingguan     = urut.filter(function (n) {
    return !isBlokBaseline_(n) && !isBlokPerbandingan_(n);
  });

  var dipilih;
  if (BCP_CONFIG.semuaMinggu) {
    // Ambil paling banyak maksBlokMinggu blok terakhir. Urutan kolom di
    // sheet sudah kronologis, jadi ekor array adalah yang terbaru.
    dipilih = mingguan.slice(-BCP_CONFIG.maksBlokMinggu);
  } else {
    dipilih = mingguan.slice(-1);
  }

  return baseline.concat(dipilih).concat(perbandingan);
}

// ─── Ringkasan ──────────────────────────────────────────────────────────────

function ringkasanBcp() {
  var ui = SpreadsheetApp.getUi();
  var sheet = ss_().getSheetByName(ARSIP_BCP);

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('Arsip BCP masih kosong. Jalankan "Capture BCP" dulu.');
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var blokDipakai = pilihBlok_(headers);

  var semuaBlok = [];
  headers.forEach(function (h) {
    var label = String(h).trim();
    if (!label || label === 'captured_at') return;
    if (BCP_KOLOM_IDENTITAS.indexOf(label) >= 0) return;
    var nama = label.split(' - ')[0];
    if (nama && semuaBlok.indexOf(nama) < 0) semuaBlok.push(nama);
  });

  var captured = sheet.getRange(2, 1).getValue();

  ui.alert('Ringkasan BCP',
    'Jumlah point: ' + (sheet.getLastRow() - 1) +
    '\nTerakhir capture: ' + captured +
    '\n\nBlok di arsip (' + semuaBlok.length + '):\n  ' + semuaBlok.join('\n  ') +
    '\n\nDikirim ke Ava (' + blokDipakai.length + '):\n  ' + blokDipakai.join('\n  ') +
    '\n\nsemuaMinggu: ' + BCP_CONFIG.semuaMinggu +
    '\nmaksBlokMinggu: ' + BCP_CONFIG.maksBlokMinggu,
    ui.ButtonSet.OK);
}
