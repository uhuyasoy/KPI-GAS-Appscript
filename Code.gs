/**
 * ============================================================================
 * AI FO Employee Performance Database — Capture & Serve
 * ============================================================================
 * Spreadsheet ini adalah ARSIP PERMANEN. Data di-capture sebagai static values
 * dari KPI FO Tracker, sehingga tetap berdiri walau source dihapus/dicabut.
 *
 * FLOW:
 *   KPI FO Tracker (view-only)
 *     -> [Menu: Capture Bulan Ini] -> Arsip BP / Arsip BM  (static values)
 *     -> [Menu: Hitung Agregat]    -> Agregat Point
 *     -> doGet()                    -> FastAPI /ingest/spreadsheet/sync
 *
 * SETUP:
 *   1. Extensions -> Apps Script -> paste file ini sebagai Code.gs
 *   2. Project Settings -> Script Properties:
 *        SOURCE_SPREADSHEET_ID = 1BmihElRczJKkuxJ_Rc_mQO1TQpAIpBDovIJBW8zXgWg
 *        SYNC_TOKEN            = <samakan dgn SPREADSHEET_SYNC_TOKEN di .env>
 *   3. Reload spreadsheet -> menu "FO Performance" muncul
 *   4. Menu -> "Setup Sheets" (sekali saja)
 *   5. Menu -> "Capture Bulan Ini" -> pilih periode & role
 *   6. Menu -> "Hitung Agregat Point"
 *   7. Deploy -> New deployment -> Web app
 *        Execute as: Me | Access: Anyone with the link
 *   8. Copy URL deployment -> isi SPREADSHEET_SYNC_URL di .env backend
 * ============================================================================
 */

// ─── Konstanta ──────────────────────────────────────────────────────────────

var ARSIP_BP      = 'Arsip BP';
var ARSIP_BM      = 'Arsip BM';
var ARSIP_AM      = 'Arsip AM';
var ARSIP_RM      = 'Arsip RM';
var AGREGAT_POINT = 'Agregat Point';
var CONFIG_SHEET  = '_Config';
var LOG_SHEET     = 'Log Ketidakcocokan';
var HMB_SHEET     = 'Daftar HMB';

/**
 * false = ambil SEMUA karyawan dari source; daftar point hanya dipakai
 *         untuk mengisi kolom kategori_point (post_fraud / bottom_performance).
 * true  = hanya karyawan di point terdaftar yang disimpan (perilaku lama).
 *
 * Tab AM dan RM tidak punya kolom Point, sehingga selalu ikut terambil
 * tanpa memandang nilai ini.
 */
var FILTER_AKTIF = false;

/**
 * Kolom yang disembunyikan dari SEMUA role. Isi dengan potongan nama kolom
 * (case-insensitive, cocok sebagian). Kolom yang namanya mengandung salah
 * satu kata di sini tidak akan dikirim ke Ava.
 *
 * Untuk menyalakan kembali sebuah kolom, hapus katanya dari daftar lalu
 * Deploy ulang (Manage deployments -> Edit -> New version).
 */
var KOLOM_DISEMBUNYIKAN = [
  'total insentif'   // insentif disembunyikan sementara atas permintaan
];

/** true = kolom insentif disembunyikan; false = ditampilkan. */
var SEMBUNYIKAN_INSENTIF = true;

/**
 * Cek apakah sebuah kolom harus disembunyikan dari konteks.
 * Menggabungkan daftar umum dengan toggle khusus insentif.
 */
function kolomDisembunyikan_(namaKolom) {
  var h = String(namaKolom).toLowerCase().trim();
  if (SEMBUNYIKAN_INSENTIF && /insentif/.test(h)) return true;
  if (/\(\d+\)$/.test(h)) return true;          // Buang kolom duplikat "(2)"
  if (/^cek double$/i.test(h)) return true;     // Buang kolom audit
  for (var i = 0; i < KOLOM_DISEMBUNYIKAN.length; i++) {
    if (h.indexOf(KOLOM_DISEMBUNYIKAN[i].toLowerCase()) >= 0) return true;
  }
  return false;
}

var SEMUA_ARSIP = [ARSIP_BP, ARSIP_BM, ARSIP_AM, ARSIP_RM];

/**
 * Role yang datanya dikirim ke backend (Postgres) lewat doGet.
 * AM dan RM sengaja tidak disertakan karena integrasinya masih berjalan —
 * definisi & bobot mereka belum final. Tambahkan 'AM','RM' di sini ketika
 * sudah siap, lalu Deploy ulang. Ini TIDAK menghentikan capture ke arsip;
 * arsip tetap boleh terisi, hanya pengirimannya yang ditahan.
 */
var ROLE_DIKIRIM = ['BP', 'BM'];

/** Cek apakah data sebuah role boleh dikirim ke backend. */
function roleDikirim_(role) {
  return ROLE_DIKIRIM.indexOf(String(role).trim().toUpperCase()) >= 0;
}

/** Tentukan tab arsip berdasarkan role yang terbaca dari nama tab source. */
function arsipUntukRole_(role) {
  if (role === 'BM') return ARSIP_BM;
  if (role === 'AM') return ARSIP_AM;
  if (role === 'RM') return ARSIP_RM;
  return ARSIP_BP;
}

var FILTER_SHEETS = {
  post_fraud:       'Point Post-Fraud',
  bottom_performance: '185 Point Bot Perf'
};

// Hanya sync periode terbaru ke backend (Opsi A).
// Ubah ke true kalau lead sudah minta data historis.
// Riwayat performa bulan lama. Periode terbaru selalu tampil penuh sebagai
// data utama; bulan sebelumnya dikirim sebagai rangkuman ringkas dengan
// awalan "Riwayat <bulan> -". Instruksi KB mengatur agar Ava hanya memakai
// riwayat ketika user menyebut bulannya.
// Baris header di sheet source (Jul - BP / Jul - BM). Data mulai baris 5.
var SOURCE_HEADER_ROW = 4;
var SOURCE_DATA_ROW   = 5;

// Kolom identitas di source (nama header, case-insensitive).
var KEY_NIK   = 'nik';
var KEY_NAMA  = 'nama';
var KEY_POINT = 'point';

// ─── Ganti bagian konfigurasi riwayat (baris ~96) ──────────────────────────
var KIRIM_RIWAYAT = false;

/**
 * Mengembalikan periode H-1 dari bulan berjalan (format "YYYY-MM").
 * Misal: September 2026 -> "2026-08" (Agustus).
 */
function getPeriodeTarget_() {
  var d = new Date();
  d.setDate(1); // amankan bug tanggal 31
  d.setMonth(d.getMonth() - 1);
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  return d.getFullYear() + '-' + m;
}

/**
 * Konversi cell periode (Date object / String) menjadi format "YYYY-MM" murni.
 */
function formatPeriodeCell_(v) {
  if (!v) return '';
  if (v instanceof Date) {
    var m = v.getMonth() + 1;
    return v.getFullYear() + '-' + (m < 10 ? '0' + m : m);
  }
  var s = String(v).trim();
  if (s.length >= 7 && s.charAt(4) === '-') return s.substring(0, 7);
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    var m2 = d.getMonth() + 1;
    return d.getFullYear() + '-' + (m2 < 10 ? '0' + m2 : m2);
  }
  return s;
}

var MONTH_NAMES = {
  'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'mei': '05', 'may': '05',
  'jun': '06', 'jul': '07', 'agu': '08', 'aug': '08', 'sep': '09',
  'okt': '10', 'oct': '10', 'nov': '11', 'des': '12', 'dec': '12'
};

// ─── Config helpers ─────────────────────────────────────────────────────────

function getConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    sourceId: props.getProperty('SOURCE_SPREADSHEET_ID') || '',
    syncToken: props.getProperty('SYNC_TOKEN') || ''
  };
}

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheetOrCreate_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) s = ss_().insertSheet(name);
  return s;
}

// ─── Menu ───────────────────────────────────────────────────────────────────

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('FO Performance')
    .addItem('Setup Sheets', 'setupSheets')
    .addSeparator()
    .addItem('Capture Bulan Ini', 'showCaptureDialog')
    .addItem('Hitung Agregat Point', 'hitungAgregat')
    .addSeparator()
    .addItem('Cek Koneksi Source', 'cekKoneksiSource')
    .addItem('Audit Daftar Point', 'auditDaftarPoint')
    .addItem('Ringkasan Arsip', 'ringkasanArsip')
    .addSeparator()
    .addItem('Preview Data HMB', 'previewHmb')
    .addSubMenu(menuBcp_(ui))
    .addToUi();
}

function setupSheets() {
  SEMUA_ARSIP.concat([AGREGAT_POINT, CONFIG_SHEET, LOG_SHEET, ARSIP_BCP, HMB_SHEET])
    .forEach(function (n) { sheetOrCreate_(n); });

  // Beri header awal untuk Daftar HMB kalau masih kosong.
  var hmb = ss_().getSheetByName(HMB_SHEET);
  if (hmb && hmb.getLastRow() === 0) {
    hmb.getRange(1, 1, 1, 4).setValues([['Role', 'NIK', 'Nama', 'Pulau']])
       .setFontWeight('bold');
    hmb.setFrozenRows(1);
  }

  var cfg = sheetOrCreate_(CONFIG_SHEET);
  if (cfg.getLastRow() === 0) {
    cfg.getRange(1, 1, 1, 4)
      .setValues([['periode', 'role', 'source_tab', 'captured_at']])
      .setFontWeight('bold');
    cfg.setFrozenRows(1);
  }

  var agg = sheetOrCreate_(AGREGAT_POINT);
  if (agg.getLastRow() === 0) {
    agg.getRange(1, 1, 1, 10).setValues([[
      'periode', 'point', 'role', 'jumlah_karyawan', 'rata_rata_skor_kpi',
      'jumlah_kena_boom', 'jumlah_dapat_boost', 'parameter_paling_sering_gagal',
      'kategori_point', 'updated_at'
    ]]).setFontWeight('bold');
    agg.setFrozenRows(1);
  }

  ss_().toast('Sheets siap. Lanjut ke "Capture Bulan Ini".', 'FO Performance');
}

// ─── Cek koneksi ────────────────────────────────────────────────────────────

function cekKoneksiSource() {
  var cfg = getConfig_();
  var ui = SpreadsheetApp.getUi();

  if (!cfg.sourceId) {
    ui.alert('SOURCE_SPREADSHEET_ID belum diisi di Script Properties.');
    return;
  }

  try {
    var src = SpreadsheetApp.openById(cfg.sourceId);
    var names = src.getSheets().map(function (s) { return s.getName(); });
    ui.alert(
      'Koneksi OK\n\nSource: ' + src.getName() +
      '\n\nTab tersedia:\n' + names.join('\n')
    );
  } catch (e) {
    ui.alert('Gagal buka source.\n\n' + e.message +
      '\n\nPastikan akun ini punya akses view ke spreadsheet tsb.');
  }
}

// ─── Capture ────────────────────────────────────────────────────────────────

function showCaptureDialog() {
  var cfg = getConfig_();
  var ui = SpreadsheetApp.getUi();

  if (!cfg.sourceId) {
    ui.alert('SOURCE_SPREADSHEET_ID belum diisi di Script Properties.');
    return;
  }

  var resp = ui.prompt(
    'Capture Data',
    'Masukkan nama tab source, pisahkan koma kalau lebih dari satu.\n' +
    'Contoh: Jul - BP, Jul - BM, Jul - AM, Jul - RM\n\n' +
    'Tab dengan ribuan baris sebaiknya dijalankan satu per satu\n' +
    'agar tidak menabrak batas waktu eksekusi Apps Script (6 menit).',
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var tabs = resp.getResponseText().split(',').map(function (t) {
    return t.trim();
  }).filter(Boolean);

  if (!tabs.length) {
    ui.alert('Tidak ada tab yang dimasukkan.');
    return;
  }

  var mulai = new Date().getTime();
  var hasil = [];

  tabs.forEach(function (tabName) {
    // Sisakan margin sebelum batas 6 menit supaya hasil sebagian tetap
    // tersimpan dan tidak hilang karena eksekusi dihentikan paksa.
    if (new Date().getTime() - mulai > 4.5 * 60 * 1000) {
      hasil.push('DILEWATI ' + tabName +
        ': mendekati batas waktu. Jalankan tab ini terpisah.');
      return;
    }
    try {
      hasil.push(captureTab_(cfg.sourceId, tabName));
    } catch (e) {
      hasil.push('GAGAL ' + tabName + ': ' + e.message);
    }
  });

  var detik = Math.round((new Date().getTime() - mulai) / 1000);
  hasil.push('\nWaktu eksekusi: ' + detik + ' detik.');

  ui.alert('Hasil Capture', hasil.join('\n\n'), ui.ButtonSet.OK);
}

/**
 * Baca satu tab source, filter berdasar daftar point, tulis static values
 * ke tab arsip yang sesuai. Dedup by periode + NIK.
 */
function captureTab_(sourceId, tabName) {
  var src = SpreadsheetApp.openById(sourceId);
  var sheet = src.getSheetByName(tabName);
  if (!sheet) throw new Error('Tab "' + tabName + '" tidak ditemukan di source.');

  var periode = parsePeriode_(tabName);
  var role    = parseRole_(tabName);
  var arsipName = arsipUntukRole_(role);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < SOURCE_DATA_ROW) throw new Error('Tab "' + tabName + '" kosong.');

  // getDisplayValues() -> ambil teks apa adanya (persen tetap "68.64%"),
  // sekaligus memutus semua formula. Ini yang bikin arsip independen.
  var headers = sheet.getRange(SOURCE_HEADER_ROW, 1, 1, lastCol)
                     .getDisplayValues()[0];
  var rows = sheet.getRange(SOURCE_DATA_ROW, 1, lastRow - SOURCE_DATA_ROW + 1, lastCol)
                  .getDisplayValues();

  // Peta kolom -> index. Kolom tanpa header (mis. kolom D) di-skip.
  var colMap = {};
  var usedCols = [];
  headers.forEach(function (h, i) {
    var clean = String(h).replace(/\s+/g, ' ').trim();
    if (!clean) return;
    // Kalau ada header duplikat, beri suffix biar tidak saling timpa.
    var key = clean;
    var n = 2;
    while (colMap.hasOwnProperty(key)) { key = clean + ' (' + n + ')'; n++; }
    colMap[key] = i;
    usedCols.push(key);
  });

  var idxNik   = findCol_(colMap, KEY_NIK);
  var idxNama  = findCol_(colMap, KEY_NAMA);
  var idxPoint = findCol_(colMap, KEY_POINT);

  if (idxNik === null) throw new Error('Kolom NIK tidak ditemukan di "' + tabName + '".');

  // Point hanya ada di tab BP dan BM. AM memakai Area, RM memakai Regional
  // sebagai satuan wilayah terkecil, jadi ketiadaan Point bukan error.
  var punyaPoint = (idxPoint !== null);

  var filterMap = bacaDaftarPoint_();

  // Susun baris arsip: periode, role, kategori_point, lalu semua kolom source.
  var arsipHeaders = ['periode', 'role', 'kategori_point'].concat(usedCols);

  var daftarPoint = Object.keys(filterMap);

  var out = [];
  var mismatch = {};        // point mentah -> { jumlah, contoh NIK }
  var kosongPoint = 0;
  var jumlahBerkategori = 0;

  rows.forEach(function (r) {
    var nik = String(r[idxNik]).trim();
    if (!nik) return;

    // FILTER_AKTIF = false: semua karyawan diambil, daftar point hanya
    // dipakai untuk menandai kategori. Baris tanpa point (AM/RM) atau yang
    // point-nya di luar daftar tetap masuk arsip dengan kategori kosong.
    var kategori = '';
    if (punyaPoint) {
      var pointRaw = String(r[idxPoint]).trim();
      if (!pointRaw) {
        kosongPoint++;
      } else {
        kategori = filterMap[normalizePoint_(pointRaw)] || '';
        if (kategori) {
          jumlahBerkategori++;
        } else {
          // Tetap dicatat sebagai informasi, bukan sebagai penolakan.
          if (!mismatch[pointRaw]) mismatch[pointRaw] = { n: 0, niks: [] };
          mismatch[pointRaw].n++;
          if (mismatch[pointRaw].niks.length < 3) mismatch[pointRaw].niks.push(nik);
        }
      }
    }

    if (FILTER_AKTIF && punyaPoint && !kategori) return;

    var rec = [periode, role, kategori];
    usedCols.forEach(function (k) { rec.push(r[colMap[k]]); });
    out.push(rec);
  });

  var totalDiluarDaftar = Object.keys(mismatch).reduce(function (a, k) {
    return a + mismatch[k].n;
  }, 0);

  // Log hanya relevan ketika filter aktif; saat nonaktif, "tidak cocok"
  // adalah kondisi normal untuk mayoritas point nasional.
  if (FILTER_AKTIF) {
    catatMismatch_(periode, role, tabName, mismatch, daftarPoint);
  }

  if (!out.length) {
    throw new Error('Tidak ada baris terbaca di "' + tabName +
      '". Total baris source: ' + rows.length);
  }

  tulisArsip_(arsipName, arsipHeaders, out, periode, role);
  catatConfig_(periode, role, tabName);

  var ringkas = tabName + '\n  Periode: ' + periode + ' | Role: ' + role +
                '\n  Tersimpan: ' + out.length + ' baris';
  if (punyaPoint) {
    ringkas += '\n  Berkategori (post_fraud / bottom_performance): ' + jumlahBerkategori +
               '\n  Di luar daftar kategori: ' + totalDiluarDaftar;
    if (kosongPoint) ringkas += '\n  Baris tanpa nama point: ' + kosongPoint;
  } else {
    ringkas += '\n  Tab ini tidak punya kolom Point — kategori dikosongkan.';
  }
  if (FILTER_AKTIF) {
    ringkas += '\n  FILTER AKTIF: hanya point terdaftar yang disimpan.';
  }

  return ringkas;
}

/**
 * Catat setiap nama point yang tidak cocok ke tab log, lengkap dengan
 * kandidat terdekat dari daftar filter supaya beda ejaan mudah dikenali.
 */
function catatMismatch_(periode, role, tabName, mismatch, daftarPoint) {
  var sheet = sheetOrCreate_(LOG_SHEET);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 7).setValues([[
      'waktu', 'periode', 'role', 'source_tab',
      'nama_point_di_source', 'jumlah_baris', 'kandidat_mirip_di_daftar'
    ]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 220);
    sheet.setColumnWidth(7, 300);
  }

  // Hapus baris log lama untuk periode+role ini supaya tidak menumpuk.
  if (sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][1]) === periode && String(data[i][2]) === role) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  var keys = Object.keys(mismatch);
  if (!keys.length) return;

  var now = new Date();
  var rows = keys.sort(function (a, b) {
    return mismatch[b].n - mismatch[a].n;
  }).map(function (p) {
    return [
      now, periode, role, tabName, p, mismatch[p].n,
      cariKandidat_(p, daftarPoint)
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
}

/**
 * Cari nama point di daftar filter yang paling mirip dengan nama di source.
 * Dipakai untuk membedakan "salah ketik" dari "memang di luar daftar".
 */
function cariKandidat_(pointRaw, daftarPoint) {
  var target = normalizePoint_(pointRaw);
  var skor = [];

  daftarPoint.forEach(function (kandidat) {
    var d = jarakLevenshtein_(target, kandidat);
    var maxLen = Math.max(target.length, kandidat.length) || 1;
    var mirip = 1 - (d / maxLen);
    if (mirip >= 0.6) skor.push({ nama: kandidat, mirip: mirip });
  });

  if (!skor.length) return 'tidak ada yang mirip — kemungkinan memang di luar daftar';

  skor.sort(function (a, b) { return b.mirip - a.mirip; });
  return skor.slice(0, 3).map(function (s) {
    return s.nama + ' (' + Math.round(s.mirip * 100) + '%)';
  }).join('  |  ');
}

/** Jarak edit antara dua string. */
function jarakLevenshtein_(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;

  var prev = [];
  for (var j = 0; j <= b.length; j++) prev[j] = j;

  for (var i = 1; i <= a.length; i++) {
    var cur = [i];
    for (var k = 1; k <= b.length; k++) {
      var cost = a.charAt(i - 1) === b.charAt(k - 1) ? 0 : 1;
      cur[k] = Math.min(cur[k - 1] + 1, prev[k] + 1, prev[k - 1] + cost);
    }
    prev = cur;
  }
  return prev[b.length];
}

/**
 * Tulis ke tab arsip. Kalau periode+role sudah ada, baris lama dihapus
 * dulu supaya capture ulang bersifat replace, bukan duplikat.
 */
function tulisArsip_(arsipName, headers, rows, periode, role) {
  var sheet = sheetOrCreate_(arsipName);

  // Urutan kolom apa adanya dari capture ini. Disimpan sebelum 'headers'
  // berpotensi di-merge dengan header lama di bawah.
  var capturedOrder = headers.slice();

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // Header source bisa bertambah kolom antar bulan. Kalau beda, perluas.
    var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var merged = existing.slice();
    headers.forEach(function (h) {
      if (merged.indexOf(h) === -1) merged.push(h);
    });
    if (merged.length !== existing.length) {
      sheet.getRange(1, 1, 1, merged.length).setValues([merged]).setFontWeight('bold');
    }
    headers = merged;
  }

  var headerNow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Hapus baris lama untuk periode+role ini.
  if (sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var iPer = headerNow.indexOf('periode');
    var iRol = headerNow.indexOf('role');
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][iPer]) === periode && String(data[i][iRol]) === role) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  // Petakan tiap baris ke urutan header aktual di sheet. Kolom yang tidak
  // ada di capture ini diisi kosong, jadi bulan dengan jumlah kolom berbeda
  // tetap sejajar.
  var aligned = rows.map(function (r) {
    var map = {};
    capturedOrder.forEach(function (h, i) { map[h] = r[i]; });
    return headerNow.map(function (h) {
      return map.hasOwnProperty(h) ? map[h] : '';
    });
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, aligned.length, headerNow.length)
       .setValues(aligned);
}

function catatConfig_(periode, role, sourceTab) {
  var sheet = sheetOrCreate_(CONFIG_SHEET);
  var now = new Date();

  if (sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === periode && String(data[i][1]) === role) {
        sheet.getRange(i + 2, 3, 1, 2).setValues([[sourceTab, now]]);
        return;
      }
    }
  }
  sheet.appendRow([periode, role, sourceTab, now]);
}

// ─── Daftar point (filter) ──────────────────────────────────────────────────

/**
 * Baca kedua tab filter, kembalikan map: normalized_point -> kategori.
 * Kalau satu point ada di dua daftar, kategorinya digabung.
 */
function bacaDaftarPoint_() {
  var map = {};

  Object.keys(FILTER_SHEETS).forEach(function (kategori) {
    var name = FILTER_SHEETS[kategori];
    var sheet = ss_().getSheetByName(name);
    if (!sheet || sheet.getLastRow() === 0) return;

    var values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
    values.forEach(function (row) {
      var v = String(row[0]).trim();
      if (!v) return;
      // Lewati kemungkinan baris header.
      if (/^(point|nama point|cabang)$/i.test(v)) return;

      var key = normalizePoint_(v);
      map[key] = map[key] ? (map[key] + ',' + kategori) : kategori;
    });
  });

  return map;
}

/** Normalisasi nama point supaya "01 Kembaran" == "01  kembaran". */
function normalizePoint_(s) {
  return String(s).toLowerCase().replace(/\s+/g, ' ').trim();
}

// ─── Agregat per point ──────────────────────────────────────────────────────

function hitungAgregat() {
  var hasil = [];
  // Hanya BP dan BM yang punya kolom Point. AM beroperasi di level Area
  // dan RM di level Regional, sehingga tidak bisa diagregasi per point.
  [{ sheet: ARSIP_BP, role: 'BP' }, { sheet: ARSIP_BM, role: 'BM' }]
    .forEach(function (cfg) {
      hasil = hasil.concat(agregatDariArsip_(cfg.sheet));
    });

  var sheet = sheetOrCreate_(AGREGAT_POINT);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  if (!hasil.length) {
    SpreadsheetApp.getUi().alert('Belum ada data arsip untuk diagregasi.');
    return;
  }

  sheet.getRange(2, 1, hasil.length, hasil[0].length).setValues(hasil);
  ss_().toast(hasil.length + ' baris agregat dihitung.', 'FO Performance');
}

function agregatDariArsip_(arsipName) {
  var sheet = ss_().getSheetByName(arsipName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  var iPer   = headers.indexOf('periode');
  var iRole  = headers.indexOf('role');
  var iKat   = headers.indexOf('kategori_point');
  var iPoint = findColIdx_(headers, KEY_POINT);
  var iSkor  = findColIdx_(headers, 'total skor kpi');
  var iBoom  = findColIdx_(headers, 'kena boom');
  var iBoost = findColIdx_(headers, 'dapat boost');
  var iParam = findColIdx_(headers, 'parameter yang belum tercapai');

  var groups = {};
  data.forEach(function (r) {
    var key = r[iPer] + '||' + String(r[iPoint]).trim() + '||' + r[iRole];
    if (!groups[key]) {
      groups[key] = {
        periode: r[iPer], point: String(r[iPoint]).trim(), role: r[iRole],
        kategori: r[iKat], skor: [], boom: 0, boost: 0, params: {}
      };
    }
    var g = groups[key];

    if (iSkor >= 0) {
      var v = skorKeSkala_(r[iSkor]);
      if (v !== null) g.skor.push(v);
    }
    if (iBoom >= 0 && /^ya|yes$/i.test(String(r[iBoom]).trim())) g.boom++;
    if (iBoost >= 0 && /^ya|yes$/i.test(String(r[iBoost]).trim())) g.boost++;
    if (iParam >= 0) {
      String(r[iParam]).split(',').forEach(function (p) {
        var t = p.trim();
        if (!t || /^semua parameter belum tercapai$/i.test(t)) return;
        g.params[t] = (g.params[t] || 0) + 1;
      });
    }
  });

  var now = new Date();
  return Object.keys(groups).map(function (k) {
    var g = groups[k];
    var avg = g.skor.length
      ? (g.skor.reduce(function (a, b) { return a + b; }, 0) / g.skor.length)
      : '';

    var topParam = '';
    var max = 0;
    Object.keys(g.params).forEach(function (p) {
      if (g.params[p] > max) { max = g.params[p]; topParam = p; }
    });
    if (topParam) topParam += ' (' + max + ' org)';

    return [
      g.periode, g.point, g.role,
      g.skor.length ? g.skor.length : 0,
      avg === '' ? '' : (avg.toFixed(2) + '%'),
      g.boom, g.boost, topParam, g.kategori, now
    ];
  });
}

// ─── doGet: serve JSON ke FastAPI ───────────────────────────────────────────

function doGet(e) {
  var cfg = getConfig_();
  var token = (e && e.parameter && e.parameter.token) || '';
  if (cfg.syncToken && token !== cfg.syncToken) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var scope = String((e.parameter.scope || 'all')).toLowerCase();
  var page = Math.max(1, parseInt(e.parameter.page || '1', 10) || 1);
  var limit = parseInt(e.parameter.limit || '1000', 10) || 1000;
  if (limit < 1) limit = 1000;
  if (limit > 2000) limit = 2000;

  var out = { page: page, limit: limit };

  if (scope === 'users' || scope === 'all') {
    var ringkasan = RINGKASAN_TIM_AKTIF ? ringkasanTimPerNik_() : null;
    var allUsers = buildUsersFull_(ringkasan);
    out.users = potongPage_(allUsers, page, limit);
    out.users_total = allUsers.length;
  }

  if (scope === 'branches' || scope === 'all') {
    var allBranches = buildBranchesFull_();
    out.branches = potongPage_(allBranches, page, limit);
    out.branches_total = allBranches.length;
  }

  out.row_count = (out.users_total || 0) + (out.branches_total || 0);
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}


function potongPage_(arr, page, limit) {
  var mulai = (page - 1) * limit;
  return arr.slice(mulai, mulai + limit);
}

function buildUsersFull_(ringkasan) {
  var branches = []; // tidak dipakai, hanya untuk reuse
  var bcp = {};
  var users = buildUsers_(ringkasan);
  if (ringkasan && ringkasan.semua) {
    var hmbUsers = buildHmbUsers_(ringkasan.semua, ringkasan.regToPulau);
    users = users.concat(hmbUsers);
  }
  return users;
}

// Isi lama doGet() untuk branches dipindah ke sini tanpa diubah.
function buildBranchesFull_() {
  var branches = BCP_CONFIG.kirimAgregatKpi ? buildBranches_() : [];
  var bcp = buildBcpBranches_();
  var byPoint = {};
  branches.forEach(function (b) { byPoint[b.point] = b; });
  Object.keys(bcp).forEach(function (point) {
    if (!byPoint[point]) byPoint[point] = { point: point, nama_cabang: point };
    var rec = bcp[point];
    Object.keys(rec).forEach(function (k) { byPoint[point][k] = rec[k]; });
  });
  return Object.keys(byPoint).map(function (p) { return byPoint[p]; });
}

/**
 * Satu objek per NIK. Struktur sudah future-proof: 'riwayat' siap diisi
 * periode terbaru jadi data utama; bulan lama jadi rangkuman riwayat.
 */
function buildUsers_(ringkasan) {
  var byNik = {};
  var targetPeriode = getPeriodeTarget_();
  SEMUA_ARSIP.forEach(function (arsipName) {
    var sheet = ss_().getSheetByName(arsipName);
    if (!sheet || sheet.getLastRow() < 2) return;
    var roleArsip = arsipName.replace(/^Arsip\s+/i, '').trim().toUpperCase();
    if (!roleDikirim_(roleArsip)) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var iPer  = headers.indexOf('periode');
    var iNik  = findColIdx_(headers, KEY_NIK);
    var iNama = findColIdx_(headers, KEY_NAMA);
    if (iNik < 0 || iPer < 0) return;
    var adaTarget = data.some(function (r) { 
      return formatPeriodeCell_(r[iPer]) === targetPeriode; 
    });
    var latest = '';
    if (!adaTarget) {
      data.forEach(function (r) {
        var p = formatPeriodeCell_(r[iPer]);
        if (p > latest) latest = p;
      });
    }
    var periodeDipakai = adaTarget ? targetPeriode : latest;
    data.forEach(function (r) {
      var periode = formatPeriodeCell_(r[iPer]);
      if (periode !== periodeDipakai) return;
      var nik = String(r[iNik]).trim();
      if (!nik) return;
      var kpi = {};
      var skip = [iPer, iNik, iNama];
      headers.forEach(function (h, i) {
        if (skip.indexOf(i) >= 0 || !h) return;
        var v = r[i];
        if (v === '' || v === null) return;
        kpi[h] = v;
      });
      byNik[nik] = {
        nik: nik,
        nama: String(r[iNama] || '').trim(),
        periode: periode,
        kpi: kpi
      };
    });
  });
  // Gunakan ringkasan yang sudah dioper dari doGet
  var timPerNik = (ringkasan && ringkasan.perNik) ? ringkasan.perNik : {};
  return Object.keys(byNik).map(function (nik) {
    var u = byNik[nik];
    var rec = {
      username: nik,
      full_name: u.nama,
      periode_kpi: u.periode || ''
    };
    Object.keys(u.kpi).forEach(function (h) {
      var v = u.kpi[h];
      if (v === '' || v === null || v === undefined) return;
      if (kolomDisembunyikan_(h)) return;
      var label = labelKpi_(h);
      if (!label) return;
      rec[label] = normalisasiNilai_(v, h);
    });
    if (rec['Kategori Point'] === '' || rec['Kategori Point'] === undefined) {
      delete rec['Kategori Point'];
    }
    var tim = timPerNik[nik];
    if (tim) {
      Object.keys(tim).forEach(function (k) { rec[k] = tim[k]; });
    }
    return rec;
  });
}

/**
 * Rapikan nilai supaya konsisten dan tidak ambigu.
 *
 * Persentase ditulis dengan simbol % dan dua desimal tetap. Pencegahan
 * salah tafsir (54.40% dibaca sebagai 0.544) ditangani lewat aturan
 * penulisan di KB, bukan dengan mengubah simbolnya.
 *
 * Rupiah diberi pemisah ribuan agar tidak terbaca sebagai angka polos.
 */
function normalisasiNilai_(v, namaKolom) {
  var s = String(v).trim();
  if (!s) return s;
  var nama = String(namaKolom || '');
  if (/%$/.test(s)) {
    if (/[\-\/]/.test(s) && (s.match(/%/g) || []).length > 1) return s;
    var angka = s.replace(/%/g, '').replace(/,/g, '.').trim();
    var np = parseFloat(angka);
    if (!isNaN(np)) return np.toFixed(2) + '%';
    return s;
  }
  var sClean = s.replace(/,/g, '.');
  var scoreDsb = /score|skor|^gap to target %|cohort|grouping|pencapaian|growth|quality|celengan|ppob/i.test(nama);
  var pastiPersen = scoreDsb ||
                    /%|rate|renewal|repayment|achievement|majelis anggota|celengan|ppob|profit|flow rate|retention|audit rating/i.test(nama);
  var hitungan = !scoreDsb && (
                   /rangking|nik|rank|jumlah|cek double/i.test(nama) ||
                   /^(target |gap to target )?new majelis cair per bulan$/i.test(nama)
                 );
  if (!hitungan && /^-?\d*\.?\d+$/.test(sClean)) {
    var n = parseFloat(sClean);
    if (!isNaN(n)) {
      var isDecimalRatio = sClean.indexOf('.') >= 0 && Math.abs(n) <= 10;
      if (pastiPersen || isDecimalRatio) {
        var persen = (Math.abs(n) <= 10) ? n * 100 : n;
        return persen.toFixed(2) + '%';
      }
    }
  }
  var isRupiah = /insentif|amount|disbursement|lost|recovery|collect/i.test(nama);
  if (isRupiah && /^-?\d+$/.test(sClean)) {
    return 'Rp' + Number(sClean).toLocaleString('id-ID');
  }
  if (!nama && /^-?\d{7,}$/.test(sClean)) {
    return 'Rp' + Number(sClean).toLocaleString('id-ID');
  }
  return s;
}

function labelKpi_(header) {
  var h = String(header).replace(/\s+/g, ' ').trim();

  // Sudah ada di user_context dari JWT — jangan diulang.
  if (['Point', 'Area', 'Regional', 'Nama'].indexOf(h) >= 0) return null;

  if (h === 'Jabatan' || h === 'Pulau') return h;
  if (h === 'role') return 'Role';
  if (h === 'kategori_point') return 'Kategori Point';

  // Backend sudah memberi tag pada blok metrik, jadi prefix "KPI" di tiap
  // baris menjadi mubazir ("KPI Total Skor KPI"). Kolom dipakai apa adanya.
  return h;
}

/** Satu objek per point, dari tab Agregat Point (periode terbaru saja). */
function buildBranches_() {
  var sheet = ss_().getSheetByName(AGREGAT_POINT);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  var iPer = headers.indexOf('periode');
  var iPoint = headers.indexOf('point');
  var iRole = headers.indexOf('role');

  var targetPeriode = getPeriodeTarget_();
  var adaTarget = data.some(function (r) { 
    return formatPeriodeCell_(r[iPer]) === targetPeriode; 
  });
  var latest = '';
  if (!adaTarget) {
    data.forEach(function (r) {
      var p = formatPeriodeCell_(r[iPer]);
      if (p > latest) latest = p;
    });
  }
  var periodeDipakai = adaTarget ? targetPeriode : latest;

  var byPoint = {};
  data.forEach(function (r) {
    if (formatPeriodeCell_(r[iPer]) !== periodeDipakai) return;
    var point = String(r[iPoint]).trim();
    if (!point) return;

    if (!byPoint[point]) {
      byPoint[point] = { point: point, nama_cabang: point, periode_kpi: periodeDipakai };
    }

    var role = String(r[iRole]).trim();
    headers.forEach(function (h, i) {
      if (i === iPer || i === iPoint || i === iRole || !h) return;
      if (r[i] === '' || r[i] === null) return;
      var label = String(h).replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
      if (kolomDisembunyikan_(label)) return;
      byPoint[point]['Point ' + role + ' - ' + label] = normalisasiNilai_(r[i], label);
    });
  });

  return Object.keys(byPoint).map(function (p) { return byPoint[p]; });
}

// ─── Utilities ──────────────────────────────────────────────────────────────

/**
 * Cari index kolom di colMap. Exact match diprioritaskan supaya 'point'
 * tidak tertangkap oleh header lain yang kebetulan mengandung kata itu.
 */
function findCol_(colMap, needle) {
  var keys = Object.keys(colMap);
  var n = needle.toLowerCase();
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase().trim() === n) return colMap[keys[i]];
  }
  for (var j = 0; j < keys.length; j++) {
    if (keys[j].toLowerCase().indexOf(n) >= 0) return colMap[keys[j]];
  }
  return null;
}

/**
 * Cari index kolom. Exact match diprioritaskan agar 'point' tidak salah
 * menangkap 'kategori_point', dan 'nama' tidak menangkap 'nama_cabang'.
 */
function findColIdx_(headers, needle) {
  var n = needle.toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().replace(/\s+/g, ' ').trim() === n) return i;
  }
  for (var j = 0; j < headers.length; j++) {
    if (String(headers[j]).toLowerCase().indexOf(n) >= 0) return j;
  }
  return -1;
}

/** "Jul - BP" -> "2026-07". Tahun diambil dari nama file source bila ada. */
/** "2026-07" -> "Juli 2026". Untuk label riwayat yang enak dibaca. */
function namaPeriode_(p) {
  var nama = {
    '01': 'Januari', '02': 'Februari', '03': 'Maret', '04': 'April',
    '05': 'Mei', '06': 'Juni', '07': 'Juli', '08': 'Agustus',
    '09': 'September', '10': 'Oktober', '11': 'November', '12': 'Desember'
  };
  var m = String(p).match(/^(\d{4})-(\d{2})$/);
  if (!m) return String(p);
  return (nama[m[2]] || m[2]) + ' ' + m[1];
}

function parsePeriode_(tabName) {
  var lower = tabName.toLowerCase();
  var bulan = '';
  Object.keys(MONTH_NAMES).forEach(function (m) {
    if (!bulan && lower.indexOf(m) >= 0) bulan = MONTH_NAMES[m];
  });
  if (!bulan) bulan = '00';

  var tahun = String(new Date().getFullYear());
  var cfg = getConfig_();
  if (cfg.sourceId) {
    try {
      var nama = SpreadsheetApp.openById(cfg.sourceId).getName();
      var m = nama.match(/20\d{2}/);
      if (m) tahun = m[0];
    } catch (e) { /* fallback ke tahun berjalan */ }
  }
  return tahun + '-' + bulan;
}

/** "Jul - BM" -> "BM". Default BP. */
function parseRole_(tabName) {
  var t = tabName.toUpperCase();
  if (/\bBM\b/.test(t)) return 'BM';
  if (/\bAM\b/.test(t)) return 'AM';
  if (/\bRM\b/.test(t)) return 'RM';
  return 'BP';
}

/** "68.64%" -> 68.64 ; "" -> null */
function parsePersen_(v) {
  return skorKeSkala_(v);
}

/**
 * Samakan skor ke skala persen 0–100 apa pun format sumbernya
 * ("61.22%", "0.6122", atau "61.22"). Dipakai untuk agregat point.
 */

function skorKeSkala_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var s = String(v).trim();
  var adaPersen = /%$/.test(s);
  var n = parseFloat(s.replace(/%/g, '').replace(/,/g, '.').trim());
  if (isNaN(n)) return null;
  if (adaPersen) return n;
  return (Math.abs(n) <= 10) ? n * 100 : n;
}


// ─── Ringkasan & audit ──────────────────────────────────────────────────────

/**
 * Bandingkan daftar point di tab filter dengan point yang benar-benar
 * muncul di source. Menjawab dua pertanyaan sekaligus:
 *   - point mana di daftar lu yang tidak punya satu pun karyawan di source
 *   - point mana di source yang tidak ada di daftar lu
 */
function auditDaftarPoint() {
  var cfg = getConfig_();
  var ui = SpreadsheetApp.getUi();

  if (!cfg.sourceId) {
    ui.alert('SOURCE_SPREADSHEET_ID belum diisi di Script Properties.');
    return;
  }

  var resp = ui.prompt(
    'Audit Daftar Point',
    'Masukkan nama tab source yang mau dicek, pisahkan koma.\n' +
    'Contoh: Jul - BP, Jul - BM',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var tabs = resp.getResponseText().split(',').map(function (t) {
    return t.trim();
  }).filter(Boolean);
  if (!tabs.length) return;

  var filterMap = bacaDaftarPoint_();
  var daftarPoint = Object.keys(filterMap);
  if (!daftarPoint.length) {
    ui.alert('Tab filter masih kosong. Isi dulu "' +
      FILTER_SHEETS.post_fraud + '" dan "' +
      FILTER_SHEETS.bottom_performance + '".');
    return;
  }

  var src = SpreadsheetApp.openById(cfg.sourceId);
  var pointDiSource = {};   // normalized -> nama mentah pertama yang ditemui

  tabs.forEach(function (tabName) {
    var sheet = src.getSheetByName(tabName);
    if (!sheet || sheet.getLastRow() < SOURCE_DATA_ROW) return;

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(SOURCE_HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
    var idxPoint = findColIdx_(headers, KEY_POINT);
    if (idxPoint < 0) return;

    var col = sheet.getRange(SOURCE_DATA_ROW, idxPoint + 1,
                             sheet.getLastRow() - SOURCE_DATA_ROW + 1, 1)
                   .getDisplayValues();
    col.forEach(function (r) {
      var v = String(r[0]).trim();
      if (!v) return;
      var k = normalizePoint_(v);
      if (!pointDiSource[k]) pointDiSource[k] = v;
    });
  });

  var tidakAdaDiSource = daftarPoint.filter(function (p) {
    return !pointDiSource.hasOwnProperty(p);
  });
  var tidakAdaDiDaftar = Object.keys(pointDiSource).filter(function (p) {
    return !filterMap.hasOwnProperty(p);
  });

  // Tulis hasil ke tab log agar bisa ditelusuri, bukan sekadar popup.
  var sheet = sheetOrCreate_(LOG_SHEET);
  var now = new Date();
  var rows = [];

  tidakAdaDiSource.forEach(function (p) {
    rows.push([now, 'AUDIT', filterMap[p], tabs.join(' + '),
      p + '  [ada di daftar, tidak ada di source]', 0,
      cariKandidat_(p, Object.keys(pointDiSource))]);
  });
  tidakAdaDiDaftar.forEach(function (p) {
    rows.push([now, 'AUDIT', '-', tabs.join(' + '),
      pointDiSource[p] + '  [ada di source, tidak ada di daftar]', 0,
      cariKandidat_(p, daftarPoint)]);
  });

  if (rows.length) {
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 7).setValues([[
        'waktu', 'periode', 'role', 'source_tab',
        'nama_point_di_source', 'jumlah_baris', 'kandidat_mirip_di_daftar'
      ]]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  }

  ui.alert('Hasil Audit',
    'Point di daftar lu: ' + daftarPoint.length +
    '\nPoint ditemukan di source: ' + Object.keys(pointDiSource).length +
    '\n\nAda di daftar tapi tidak ada di source: ' + tidakAdaDiSource.length +
    '\nAda di source tapi tidak ada di daftar: ' + tidakAdaDiDaftar.length +
    (rows.length ? '\n\nDetail ditulis ke tab "' + LOG_SHEET + '".'
                 : '\n\nSemua nama point cocok.'),
    ui.ButtonSet.OK);
}

function ringkasanArsip() {
  var lines = [];
  SEMUA_ARSIP.forEach(function (name) {
    var sheet = ss_().getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) {
      lines.push(name + ': kosong');
      return;
    }
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var iPer = headers.indexOf('periode');
    var data = sheet.getRange(2, iPer + 1, sheet.getLastRow() - 1, 1).getValues();

    var count = {};
    data.forEach(function (r) {
      var p = String(r[0]).trim();
      count[p] = (count[p] || 0) + 1;
    });
    var detail = Object.keys(count).sort().map(function (p) {
      return '  ' + p + ': ' + count[p] + ' baris';
    }).join('\n');
    lines.push(name + ':\n' + detail);
  });

  SpreadsheetApp.getUi().alert('Ringkasan Arsip', lines.join('\n\n'),
    SpreadsheetApp.getUi().ButtonSet.OK);
}
