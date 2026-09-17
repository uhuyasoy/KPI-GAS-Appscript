/**
 * ============================================================================
 * Code.gs — AI Trainer KPI sync (sumber: PROD-DB-KPI-FO)
 * ============================================================================
 * Sumber data BARU: satu spreadsheet PROD-DB-KPI-FO berisi:
 *   - DB_Employee            : identitas + wilayah (NIK -> Point/Area/Regional/Pulau)
 *   - DB_Perf_BP/BM/AM/RM     : performa per role, kolom P01_.. P05_.., BOOM_.., BOOST_..
 *   - Daftar HMB              : didaftarkan manual (dipakai Team.gs)
 *
 * Alur: baca DB_Perf_<role> -> join ke DB_Employee via NIK -> ambil minggu
 * terbaru per NIK -> normalisasi nilai -> doGet kirim JSON ke backend.
 *
 * SETUP (Script Properties):
 *   PROD_SPREADSHEET_ID = <id spreadsheet PROD-DB-KPI-FO>
 *   SYNC_TOKEN          = <samakan dengan SPREADSHEET_SYNC_TOKEN di .env>
 * ============================================================================
 */

// ─── Nama tab sumber ─────────────────────────────────────────────────────────

var EMP_SHEET = 'DB_Employee';
var PERF_SHEET = {
  BP: 'DB_Perf_BP',
  BM: 'DB_Perf_BM',
  AM: 'DB_Perf_AM',
  RM: 'DB_Perf_RM'
};
var HMB_SHEET = 'Daftar HMB';

// ─── Konfigurasi via tab _Status (command center) ────────────────────────────
//
// Semua perilaku pengiriman diatur dari tab _Status, bukan hard-code. Command
// center (web app) menulis ke tab ini; doGet & builder membacanya. Kalau tab
// belum ada, dipakai DEFAULT_CONFIG di bawah.

var STATUS_SHEET = '_Status';

var DEFAULT_CONFIG = {
  periode_final: '',          // "" = pakai periode terbaru yang ada (belum ada gating)
  role_dikirim: 'BP,BM',      // daftar role dipisah koma
  sembunyikan_insentif: 'false',
  kolom_disembunyikan: '',    // daftar kata dipisah koma
  kirim_riwayat: 'false',
  kirim_agregat_point: 'false',
  ringkasan_tim_aktif: 'true',    // ringkasan tim untuk atasan
  granular_aktif: 'false',        // data lengkap tiap bawahan di konteks atasan
  granular_untuk_role: 'BM',      // role penerima granular, dipisah koma
  granular_maks_bawahan: '10'     // batas bawahan granular per atasan (0 = tanpa batas)
};

// Cache config selama satu eksekusi agar tidak baca sheet berulang.
var _configCache = null;

function getKonfig_() {
  if (_configCache) return _configCache;
  var cfg = {};
  Object.keys(DEFAULT_CONFIG).forEach(function (k) { cfg[k] = DEFAULT_CONFIG[k]; });

  var sheet = ss_().getSheetByName(STATUS_SHEET);
  if (sheet && sheet.getLastRow() >= 2) {
    var v = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
    for (var r = 1; r < v.length; r++) {
      var key = String(v[r][0]).trim().toLowerCase();
      if (key && DEFAULT_CONFIG.hasOwnProperty(key)) {
        cfg[key] = String(v[r][1]).trim();
      }
    }
  }
  _configCache = cfg;
  return cfg;
}

// ─── Command center: tulis config, password, deteksi periode ─────────────────

/** Tulis satu key config ke tab _Status (dibuat bila belum ada). */
function setKonfig_(key, value) {
  key = String(key).trim().toLowerCase();
  if (!DEFAULT_CONFIG.hasOwnProperty(key)) throw new Error('Key config tidak dikenal: ' + key);

  var sheet = ss_().getSheetByName(STATUS_SHEET);
  if (!sheet) {
    sheet = ss_().insertSheet(STATUS_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // isi default supaya lengkap
    var rows = Object.keys(DEFAULT_CONFIG).map(function (k) { return [k, DEFAULT_CONFIG[k]]; });
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  var v = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  for (var r = 1; r < v.length; r++) {
    if (String(v[r][0]).trim().toLowerCase() === key) {
      sheet.getRange(r + 1, 2).setValue(value);
      _configCache = null;
      return;
    }
  }
  sheet.appendRow([key, value]);
  _configCache = null;
}

/** Verifikasi password command center terhadap Script Property CC_PASSWORD. */
function cekPassword_(input) {
  var pw = PropertiesService.getScriptProperties().getProperty('CC_PASSWORD') || '';
  if (!pw) return false;                 // belum diset -> tolak semua
  return String(input).trim() === String(pw).trim();
}

/**
 * Diagnosa password — JALANKAN DARI EDITOR (Run), lihat hasil di Execution log.
 * Tidak menampilkan password, hanya apakah ter-set dan panjangnya, supaya aman.
 */
function cekSetupPassword() {
  var pw = PropertiesService.getScriptProperties().getProperty('CC_PASSWORD');
  if (pw === null || pw === undefined) {
    Logger.log('CC_PASSWORD BELUM ADA di Script Properties. Tambahkan dulu.');
    return;
  }
  if (String(pw).trim() === '') {
    Logger.log('CC_PASSWORD ADA tapi KOSONG/berisi spasi saja.');
    return;
  }
  Logger.log('CC_PASSWORD ter-set. Panjang: ' + pw.length +
    ' | ada spasi di ujung: ' + (pw !== pw.trim() ? 'YA (masalah!)' : 'tidak'));
  // Cek juga PROD id sekalian
  var pid = PropertiesService.getScriptProperties().getProperty('PROD_SPREADSHEET_ID');
  Logger.log('PROD_SPREADSHEET_ID: ' + (pid ? 'ter-set' : 'BELUM ADA'));
}

/** Daftar periode (label) yang ada di semua tab DB_Perf, urut terbaru dulu. */
function daftarPeriode_() {
  var set = {};
  ['BP', 'BM', 'AM', 'RM'].forEach(function (role) {
    var sheet = prodSS_().getSheetByName(PERF_SHEET[role]);
    if (!sheet || sheet.getLastRow() < 2) return;
    var lastCol = sheet.getLastColumn();
    var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                 .map(function (h) { return String(h).trim().toLowerCase(); });
    var iPer = header.indexOf('perioda_tanggal');
    if (iPer < 0) return;
    var col = sheet.getRange(2, iPer + 1, sheet.getLastRow() - 1, 1).getValues();
    col.forEach(function (r) {
      var label = String(r[0] || '').trim();
      if (!label) return;
      var key = parsePeriodeMingguan_(label).key;
      set[key] = label;
    });
  });
  return Object.keys(set).sort().reverse().map(function (k) {
    return { key: k, label: set[k] };
  });
}

/** Hitung jumlah baris per role untuk sebuah periode (untuk cek kelengkapan). */
/** Hitung detail kelengkapan: mentah, lolos validasi, dan HMB. */
function kelengkapanPeriode_(label) {
  var targetKey = parsePeriodeMingguan_(label).key;
  var emp = bacaEmployee_();
  var hasil = { roles: {}, total_raw: 0, total_lolos: 0, hmb: 0, total_kirim: 0 };

  ['BP', 'BM', 'AM', 'RM'].forEach(function (role) {
    var sheet = prodSS_().getSheetByName(PERF_SHEET[role]);
    if (!sheet || sheet.getLastRow() < 2) {
      hasil.roles[role] = { raw: 0, lolos: 0, skip: 0 };
      return;
    }

    var values = sheet.getDataRange().getValues();
    var header = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
    var iPer = header.indexOf('perioda_tanggal');
    var iNik = header.indexOf('nik');

    if (iPer < 0 || iNik < 0) {
      hasil.roles[role] = { raw: 0, lolos: 0, skip: 0 };
      return;
    }
    
    var raw = 0, lolos = 0, seenNik = {};
    for (var r = 1; r < values.length; r++) {
      if (parsePeriodeMingguan_(String(values[r][iPer] || '').trim()).key !== targetKey) continue;
      raw++;
      var nik = normNik_(values[r][iNik]);
      if (nik && emp[nik] && !seenNik[nik]) {
        seenNik[nik] = true;
        lolos++;
      }
    }
    hasil.roles[role] = { raw: raw, lolos: lolos, skip: raw - lolos };
    hasil.total_raw += raw;
    hasil.total_lolos += lolos;
  });

  // Hitung HMB jika tab ada
  try {
    var daftarHmb = (typeof bacaDaftarHmb_ === 'function') ? bacaDaftarHmb_() : [];
    hasil.hmb = daftarHmb.length;
  } catch (e) {
    hasil.hmb = 0;
  }

  hasil.total_kirim = hasil.total_lolos + hasil.hmb;
  return hasil;
}

function roleDikirim_(role) {
  var daftar = getKonfig_().role_dikirim.split(',').map(function (s) {
    return s.trim().toUpperCase();
  });
  return daftar.indexOf(String(role).trim().toUpperCase()) >= 0;
}

function kolomDisembunyikan_(namaKolom) {
  var cfg = getKonfig_();
  var h = String(namaKolom).toLowerCase();
  if (/^true$/i.test(cfg.sembunyikan_insentif) && /insentif/.test(h)) return true;
  var extra = cfg.kolom_disembunyikan.split(',').map(function (s) { return s.trim().toLowerCase(); });
  for (var i = 0; i < extra.length; i++) {
    if (extra[i] && h.indexOf(extra[i]) >= 0) return true;
  }
  return false;
}

function kirimRiwayat_() { return /^true$/i.test(getKonfig_().kirim_riwayat); }
function periodeFinal_() { return getKonfig_().periode_final; }

// Config untuk Team.gs (ringkasan tim & granular) — dibaca dari _Status.
function cfgRingkasanTimAktif_() { return /^true$/i.test(getKonfig_().ringkasan_tim_aktif); }
function cfgGranularAktif_() { return /^true$/i.test(getKonfig_().granular_aktif); }
function cfgGranularRole_() {
  return getKonfig_().granular_untuk_role.split(',').map(function (s) {
    return s.trim().toUpperCase();
  }).filter(Boolean);
}
function cfgGranularMaks_() {
  var n = parseInt(getKonfig_().granular_maks_bawahan, 10);
  return isNaN(n) ? 0 : n;
}

// KIRIM_AGREGAT_POINT dipakai buildBranchesFull_.
function get_KIRIM_AGREGAT_POINT_() { return /^true$/i.test(getKonfig_().kirim_agregat_point); }

var RIWAYAT_MINGGU = 4;
var RIWAYAT_KOLOM = [
  'Skor_KPI', 'Ranking', 'Grouping_Skor_KPI',
  'Boom', 'Boost', 'Parameter_Unreached', 'KPI_Status'
];

// ─── Helpers dasar ───────────────────────────────────────────────────────────

function getConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    prodId: props.getProperty('PROD_SPREADSHEET_ID') || '',
    syncToken: props.getProperty('SYNC_TOKEN') || ''
  };
}

/**
 * Spreadsheet aktif. Di web app tidak ada "active spreadsheet", jadi selalu
 * buka lewat PROD_SPREADSHEET_ID. getActiveSpreadsheet() hanya dipakai sebagai
 * fallback saat dijalankan dari editor/menu (mis. saat ID belum diisi).
 */
/**
 * Spreadsheet OPERASIONAL — tempat tab _Status, Log Sync, Preview dibuat.
 * Ini spreadsheet milik script, TERPISAH dari PROD sumber data.
 *
 * Urutan: OPS_SPREADSHEET_ID (kalau diisi) -> spreadsheet aktif (saat
 * dijalankan dari editor/menu). Di web app, OPS_SPREADSHEET_ID wajib diisi
 * karena tidak ada spreadsheet aktif.
 */
function ss_() {
  var props = PropertiesService.getScriptProperties();
  var opsId = props.getProperty('OPS_SPREADSHEET_ID') || '';
  if (opsId) {
    try { return SpreadsheetApp.openById(opsId); } catch (e) { /* fallback */ }
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('OPS_SPREADSHEET_ID belum diisi di Script Properties. Isi dengan ID spreadsheet tempat script ini menempel (untuk _Status & Log Sync).');
}

/**
 * Spreadsheet PROD — sumber data KPI (DB_Perf, DB_Employee). Hanya dibaca,
 * tidak pernah ditulis. Wajib PROD_SPREADSHEET_ID.
 */
function prodSS_() {
  var cfg = getConfig_();
  if (cfg.prodId) {
    try { return SpreadsheetApp.openById(cfg.prodId); } catch (e) {
      throw new Error('Gagal membuka PROD (PROD_SPREADSHEET_ID). Cek ID & akses.');
    }
  }
  throw new Error('PROD_SPREADSHEET_ID belum diisi di Script Properties.');
}

function sheetOrCreate_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) s = ss_().insertSheet(name);
  return s;
}

/** Normalisasi NIK agar join tidak gagal karena string vs float. */
function normNik_(v) {
  if (v === null || v === undefined) return '';
  var s = String(v).trim();
  if (!s) return '';
  return s.replace(/\.0+$/, '');   // buang ".0" artefak angka
}

// ─── Periode mingguan ────────────────────────────────────────────────────────

var BULAN_MAP = {
  'januari': '01', 'februari': '02', 'maret': '03', 'april': '04',
  'mei': '05', 'juni': '06', 'juli': '07', 'agustus': '08',
  'september': '09', 'oktober': '10', 'november': '11', 'desember': '12',
  'january': '01', 'february': '02', 'march': '03', 'may': '05',
  'june': '06', 'july': '07', 'august': '08', 'october': '10', 'december': '12'
};

/** "Week 1 - September 2026" -> { key:"2026-09-W01", label:"Week 1 - September 2026" } */
function parsePeriodeMingguan_(s) {
  var str = String(s || '').trim();
  var m = str.match(/week\s*(\d+)\s*-\s*([a-z]+)\s*(\d{4})/i);
  if (!m) return { key: str, label: str };
  var w = ('0' + m[1]).slice(-2);
  var bl = BULAN_MAP[m[2].toLowerCase()] || '00';
  return { key: m[3] + '-' + bl + '-W' + w, label: str };
}

// ─── Baca DB_Employee ────────────────────────────────────────────────────────

/** Peta NIK -> {point, area, regional, pulau, position}. */
function bacaEmployee_() {
  var sheet = prodSS_().getSheetByName(EMP_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var header = values[0].map(function (h) { return String(h).trim().toLowerCase(); });

  var idx = {};
  header.forEach(function (h, i) { idx[h] = i; });
  var iNik = idx['nik'], iPoint = idx['point'], iArea = idx['area'],
      iReg = idx['regional'], iPulau = idx['pulau'], iPos = idx['position'];

  var peta = {};
  for (var r = 1; r < values.length; r++) {
    var nik = normNik_(values[r][iNik]);
    if (!nik) continue;

    var position = iPos != null ? String(values[r][iPos] || '').trim() : '';
    // Buang karyawan Head Office — bukan FO, tidak relevan untuk KPI FO.
    if (/^ho$/i.test(position)) continue;

    peta[nik] = {
      point: iPoint != null ? String(values[r][iPoint] || '').trim() : '',
      area: iArea != null ? String(values[r][iArea] || '').trim() : '',
      regional: iReg != null ? String(values[r][iReg] || '').trim() : '',
      pulau: iPulau != null ? String(values[r][iPulau] || '').trim() : '',
      position: position
    };
  }
  return peta;
}

// ─── Baca DB_Perf_<role> ─────────────────────────────────────────────────────

/**
 * Baca satu tab performa. Kembalikan { byNik: {nik: {periode->record}}, ... }.
 * Tiap record: { periodeKey, periodeLabel, nama, role, kpi:{kolom->nilai} }.
 * Kolom identitas & meta dipisah dari kpi.
 */
function bacaPerf_(role) {
  var tabName = PERF_SHEET[role];
  var sheet = prodSS_().getSheetByName(tabName);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  var header = values[0].map(function (h) { return String(h).trim(); });

  var idx = {};
  header.forEach(function (h, i) { idx[h.toLowerCase()] = i; });
  var iNik = idx['nik'], iNama = idx['fullname'], iRole = idx['role'],
      iPer = idx['perioda_tanggal'];

  // Kolom yang bukan KPI (identitas + meta yang ditangani khusus).
  var kolomMeta = ['nik', 'fullname', 'role', 'perioda_tanggal'];

  var byNik = {};
  for (var r = 1; r < values.length; r++) {
    var nik = normNik_(values[r][iNik]);
    if (!nik) continue;

    var per = parsePeriodeMingguan_(values[r][iPer]);
    var kpi = {};
    header.forEach(function (h, i) {
      if (kolomMeta.indexOf(h.toLowerCase()) >= 0) return;
      var v = values[r][i];
      if (v === '' || v === null || v === undefined) return;
      kpi[h] = v;
    });

    if (!byNik[nik]) byNik[nik] = { nama: '', role: role, snapshots: {} };
    byNik[nik].nama = String(values[r][iNama] || '').trim();
    byNik[nik].snapshots[per.key] = { label: per.label, kpi: kpi };
  }
  return byNik;
}

// ─── Normalisasi nilai ───────────────────────────────────────────────────────

/**
 * Rapikan nilai untuk konteks Ava.
 * - Desimal rasio (<=3) jadi persen: 0.98 -> 98.00%, 1.2 -> 120.00%.
 * - Kolom Rupiah (_IDR, insentif, amount) jadi Rp dengan pemisah ribuan.
 * - Kolom count (NOA, ranking) dibiarkan angka.
 */
function normalisasiNilai_(v, namaKolom) {
  var s = String(v).trim();
  if (!s) return s;
  var nama = String(namaKolom || '');

  // Kolom teks/kategori: jangan pernah diubah jadi angka meski mengandung
  // angka atau '%'. Contoh: Grouping "1.<80%", Status "1. Need Improvement",
  // Parameter_Unreached (daftar teks), periode, nama.
  if (/grouping|status|parameter|perioda|fullname|^role$|unreached/i.test(nama)) {
    return s;
  }

  // Sudah bertanda persen di ujung -> seragamkan (hanya bila murni angka+%).
  if (/^-?\d*\.?\d+\s*%$/.test(s)) {
    var np = parseFloat(s.replace('%', '').replace(/,/g, '.').trim());
    return isNaN(np) ? s : np.toFixed(2) + '%';
  }

  // Rupiah: kolom IDR / insentif / amount.
  var isRupiah = /_idr|insentif|amount|disbursement.*idr/i.test(nama);
  if (isRupiah && /^-?\d+(\.\d+)?$/.test(s)) {
    var rp = Math.round(parseFloat(s));
    return 'Rp' + Number(rp).toLocaleString('id-ID');
  }

  // Count murni: Ranking, dan Disbursement NOA (target/capaian/gap berupa jumlah).
  if (/ranking/i.test(nama)) return s;
  if (/_noa\b/i.test(nama) && /target|capaian|gap/i.test(nama)) return s;

  // Angka desimal -> persen bila rasio; angka besar dibiarkan.
  if (/^-?\d*\.?\d+$/.test(s)) {
    var n = parseFloat(s);
    if (isNaN(n)) return s;
    return (Math.abs(n) <= 3 ? n * 100 : n).toFixed(2) + '%';
  }

  return s;
}

/**
 * Rapikan nama kolom KPI jadi label enak baca.
 * "P01_Score_Repayment_Rate_DPD_0" -> "Score Repayment Rate DPD 0"
 * "BOOST_Capaian_Mitra_Celengan_50K" -> "Boost Capaian Mitra Celengan 50K"
 * Kolom meta identitas dikembalikan null (sudah dari DB_Employee/JWT).
 */
function labelKpi_(header) {
  var h = String(header).trim();
  var low = h.toLowerCase();

  // Sudah tersedia dari DB_Employee / JWT.
  if (['point', 'area', 'regional', 'pulau', 'nama', 'fullname'].indexOf(low) >= 0) return null;

  // Buang prefix kode P01_/P02_/BOOM_/BOOST_, ganti underscore jadi spasi.
  var t = h.replace(/^P\d+_/i, '').replace(/^BOOM_/i, 'Boom ').replace(/^BOOST_/i, 'Boost ');
  t = t.replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
  return t;
}

// ─── Bangun record user ──────────────────────────────────────────────────────

/**
 * Kumpulkan semua karyawan (role yang dikirim) dengan periode terbaru sebagai
 * data utama. Dipakai buildUsers_ dan juga sebagai bahan Team.gs (via
 * bacaSemuaKaryawanProd_).
 */
function kumpulkanKaryawan_(rolesFilter, stats) {
  var emp = bacaEmployee_();
  var roles = rolesFilter || ['BP', 'BM', 'AM', 'RM'];
  var hasil = [];

  // Gating periode final. Bila diset (mis. "Week 1 - September 2026"), hanya
  // snapshot periode itu ke bawah yang boleh dikirim — minggu lebih baru yang
  // belum ditandai final diabaikan. Kosong = tanpa gating (pakai terbaru).
  var finalLabel = periodeFinal_();
  var finalKey = finalLabel ? parsePeriodeMingguan_(finalLabel).key : '';

  roles.forEach(function (role) {
    var byNik = bacaPerf_(role);
    var total = 0, terkirim = 0, skipNoEmp = 0, skipKosong = 0, skipBelumFinal = 0;
    var contohNoEmp = [];

    Object.keys(byNik).forEach(function (nik) {
      total++;
      var u = byNik[nik];
      var keys = Object.keys(u.snapshots).sort();   // sortable
      if (!keys.length) { skipKosong++; return; }

      // Terapkan gating: buang snapshot yang lebih baru dari periode final.
      if (finalKey) {
        keys = keys.filter(function (k) { return k <= finalKey; });
        if (!keys.length) { skipBelumFinal++; return; }
      }

      // Wilayah = acuan penempatan terupdate dari DB_Employee. Karyawan yang
      // tidak ada di DB_Employee non-HO (termasuk HO yang sudah dibuang, atau
      // NIK tak terdaftar) tidak dikirim — tidak punya penempatan valid.
      var w = emp[nik];
      if (!w) {
        skipNoEmp++;
        if (contohNoEmp.length < 5) contohNoEmp.push(nik + (u.nama ? ' (' + u.nama + ')' : ''));
        return;
      }

      var latest = keys[keys.length - 1];
      var snap = u.snapshots[latest];

      hasil.push({
        nik: nik,
        nama: u.nama,
        role: role,
        position: w.position || '',
        point: w.point || '',
        area: w.area || '',
        regional: w.regional || '',
        pulau: w.pulau || '',
        periodeKey: latest,
        periodeLabel: snap.label,
        kpi: snap.kpi,
        skor: skorKeSkala_(snap.kpi['Skor_KPI']),
        snapshots: u.snapshots
      });
      terkirim++;
    });

    if (stats) {
      stats.push({
        role: role,
        dikirim: roleDikirim_(role),
        total: total,
        lolos: terkirim,
        skip_no_employee: skipNoEmp,
        skip_tanpa_periode: skipKosong,
        skip_belum_final: skipBelumFinal,
        contoh_no_employee: contohNoEmp
      });
    }
  });

  return hasil;
}

/** Skor ke skala 0-100 apa pun formatnya (untuk ringkasan tim). */
function skorKeSkala_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var s = String(v).trim();
  var ada = /%$/.test(s);
  var n = parseFloat(s.replace('%', '').replace(/,/g, '.').trim());
  if (isNaN(n)) return null;
  if (ada) return n;
  return (Math.abs(n) <= 3) ? n * 100 : n;
}

/** Payload users untuk backend (hanya role yang dikirim). */
function buildUsers_(ringkasan, stats) {
  var rolesKirim = ['BP', 'BM', 'AM', 'RM'].filter(roleDikirim_);
  var karyawan = kumpulkanKaryawan_(rolesKirim, stats);
  var timPerNik = (ringkasan && ringkasan.perNik) ? ringkasan.perNik : {};

  return karyawan.map(function (k) {
    var rec = {
      username: k.nik,
      full_name: k.nama,
      role: k.role,
      position: k.position,
      periode_kpi: k.periodeLabel,
      point: k.point,
      area: k.area,
      regional: k.regional,
      pulau: k.pulau
    };

    // KPI periode terbaru.
    Object.keys(k.kpi).forEach(function (kolom) {
      if (kolomDisembunyikan_(kolom)) return;
      var label = labelKpi_(kolom);
      if (!label) return;
      rec[label] = normalisasiNilai_(k.kpi[kolom], kolom);
    });

    // Riwayat minggu lama (bila diaktifkan).
    if (kirimRiwayat_()) {
      var keys = Object.keys(k.snapshots).sort();
      var older = keys.slice(0, -1).slice(-RIWAYAT_MINGGU);
      older.forEach(function (pk) {
        var snap = k.snapshots[pk];
        RIWAYAT_KOLOM.forEach(function (kolom) {
          if (snap.kpi[kolom] === undefined || snap.kpi[kolom] === '') return;
          if (kolomDisembunyikan_(kolom)) return;
          rec['Riwayat ' + snap.label + ' - ' + labelKpi_(kolom)] =
            normalisasiNilai_(snap.kpi[kolom], kolom);
        });
      });
    }

    // Ringkasan tim bila user ini atasan.
    var tim = timPerNik[k.nik];
    if (tim) Object.keys(tim).forEach(function (kk) { rec[kk] = tim[kk]; });

    return rec;
  });
}

/**
 * Bahan untuk Team.gs: semua karyawan (semua role, tanpa filter kirim) dengan
 * wilayah dan skor. Team.gs butuh lihat seluruh hierarki walau yang dikirim
 * hanya sebagian role.
 */
function bacaSemuaKaryawanProd_() {
  return kumpulkanKaryawan_(['BP', 'BM', 'AM', 'RM']).map(function (k) {
    return {
      nik: k.nik, nama: k.nama, role: k.role,
      point: k.point, area: k.area, regional: k.regional, pulau: k.pulau,
      skor: k.skor, kpi: k.kpi
    };
  });
}

// ─── doGet ───────────────────────────────────────────────────────────────────

function doGet(e) {
  // Routing: request backend (ada token) -> JSON. Selain itu -> command center.
  var params = (e && e.parameter) || {};
  if (params.token || params.scope) {
    return doGetData_(e);
  }
  // Halaman command center.
  return HtmlService.createHtmlOutputFromFile('CommandCenter')
    .setTitle('FO KPI Command Center')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─── API dipanggil dari halaman (google.script.run) ──────────────────────────

function cc_login(password) {
  return { ok: cekPassword_(password) };
}

/** Ambil semua yang dibutuhkan halaman: config saat ini + daftar periode. */
function cc_getStatus(password) {
  if (!cekPassword_(password)) return { error: 'unauthorized' };
  var periode = daftarPeriode_();
  var cfg = getKonfig_();
  // kelengkapan untuk periode terbaru & yang dipilih final
  var lengkapTerbaru = periode.length ? kelengkapanPeriode_(periode[0].label) : {};
  var lengkapFinal = cfg.periode_final ? kelengkapanPeriode_(cfg.periode_final) : {};
  return {
    ok: true,
    config: cfg,
    periode: periode,
    kelengkapan_terbaru: lengkapTerbaru,
    kelengkapan_final: lengkapFinal
  };
}

/** Simpan config dari halaman. patch = objek {key: value}. */
function cc_saveConfig(password, patch) {
  if (!cekPassword_(password)) return { error: 'unauthorized' };
  var diubah = [];
  Object.keys(patch || {}).forEach(function (k) {
    if (DEFAULT_CONFIG.hasOwnProperty(k)) {
      setKonfig_(k, String(patch[k]));
      diubah.push(k);
    }
  });
  catatLog_('commandCenter', 'CONFIG_UPDATE', 'diubah: ' + diubah.join(', '), {});
  return { ok: true, diubah: diubah, config: getKonfig_() };
}

/** Cek kelengkapan satu periode (dipanggil saat user pilih di dropdown). */
function cc_cekKelengkapan(password, label) {
  if (!cekPassword_(password)) return { error: 'unauthorized' };
  return { ok: true, label: label, kelengkapan: kelengkapanPeriode_(label) };
}

// ─── doGet data (backend) ────────────────────────────────────────────────────

function doGetData_(e) {
  var cfg = getConfig_();
  var token = (e && e.parameter && e.parameter.token) || '';
  if (cfg.syncToken && token !== cfg.syncToken) {
    catatLog_('doGet', 'DITOLAK', 'token tidak valid', {});
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
  var mulai = new Date().getTime();

  try {
    var statsUser = [];
    if (scope === 'users' || scope === 'all') {
      var ringkasan = cfgRingkasanTimAktif_()
        ? ringkasanTimPerNik_() : null;
      var allUsers = buildUsersFull_(ringkasan, statsUser);
      out.users = potongPage_(allUsers, page, limit);
      out.users_total = allUsers.length;
    }

    if (scope === 'branches' || scope === 'all') {
      var allBranches = buildBranchesFull_();
      out.branches = potongPage_(allBranches, page, limit);
      out.branches_total = allBranches.length;
    }

    out.row_count = (out.users_total || 0) + (out.branches_total || 0);

    // Catat log hanya di halaman pertama supaya tidak menumpuk saat paginasi.
    if (page === 1) {
      var durasi = ((new Date().getTime() - mulai) / 1000).toFixed(1) + 's';
      catatLog_('doGet', 'BERHASIL',
        'scope=' + scope + ' | users=' + (out.users_total || 0) +
        ' branches=' + (out.branches_total || 0) + ' | ' + durasi,
        { stats: statsUser });
    }

    return ContentService
      .createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    catatLog_('doGet', 'GAGAL', String(err && err.message ? err.message : err), {});
    return ContentService
      .createTextOutput(JSON.stringify({ error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Catat satu baris ke tab "Log Sync". Kolom: waktu, sumber, status, ringkasan,
 * lalu per-role detail (dikirim/lolos/skip) supaya mudah ditelusuri.
 * Menyimpan maksimal ~500 baris terakhir agar tidak membengkak.
 */
function catatLog_(sumber, status, ringkasan, extra) {
  try {
    var sheet = sheetOrCreate_('Log Sync');
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 5).setValues([[
        'waktu', 'sumber', 'status', 'ringkasan', 'detail_per_role'
      ]]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(4, 320);
      sheet.setColumnWidth(5, 500);
    }

    var detail = '';
    if (extra && extra.stats && extra.stats.length) {
      detail = extra.stats.map(function (s) {
        var t = s.role + (s.dikirim ? '[kirim]' : '[tahan]') +
                ': lolos ' + s.lolos + '/' + s.total;
        if (s.skip_no_employee) {
          t += ', skip ' + s.skip_no_employee + ' (tak ada di DB_Employee';
          if (s.contoh_no_employee && s.contoh_no_employee.length) {
            t += ': ' + s.contoh_no_employee.join(', ');
          }
          t += ')';
        }
        if (s.skip_tanpa_periode) t += ', skip ' + s.skip_tanpa_periode + ' (tanpa periode)';
        return t;
      }).join('  |  ');
    }

    sheet.appendRow([new Date(), sumber, status, ringkasan, detail]);

    // Pangkas kalau lebih dari 500 baris data (+1 header).
    var maxRows = 501;
    if (sheet.getLastRow() > maxRows) {
      sheet.deleteRows(2, sheet.getLastRow() - maxRows);
    }
  } catch (e) {
    // Log gagal tidak boleh menggagalkan sync. Diamkan.
  }
}

function potongPage_(arr, page, limit) {
  var mulai = (page - 1) * limit;
  return arr.slice(mulai, mulai + limit);
}

function buildUsersFull_(ringkasan, stats) {
  var users = buildUsers_(ringkasan, stats);
  if (ringkasan && ringkasan.semua) {
    var hmbUsers = buildHmbUsers_(ringkasan.semua, ringkasan.regToPulau);
    users = users.concat(hmbUsers);
  }
  return users;
}

function buildBranchesFull_() {
  // Agregat point per BP/BM. BCP sudah tidak dipakai.
  if (!get_KIRIM_AGREGAT_POINT_()) return [];
  return buildBranches_();
}

// ─── Agregat point (dari data PROD, per point untuk BP/BM) ────────────────────

/**
 * Agregat per point dari karyawan BP & BM periode terbaru: jumlah, rata-rata
 * skor per role, jumlah kena boom, dapat boost. Dipakai buildBranchesFull_.
 */
function buildBranches_() {
  var karyawan = kumpulkanKaryawan_(['BP', 'BM']);
  var perPoint = {};

  karyawan.forEach(function (k) {
    if (!k.point) return;
    if (!perPoint[k.point]) {
      perPoint[k.point] = { point: k.point, nama_cabang: k.point, orang: [] };
    }
    perPoint[k.point].orang.push(k);
  });

  return Object.keys(perPoint).map(function (p) {
    var blok = perPoint[p];
    var rec = { point: blok.point, nama_cabang: blok.nama_cabang };

    ['BP', 'BM'].forEach(function (role) {
      var arr = blok.orang.filter(function (o) { return o.role === role; });
      if (!arr.length) return;
      rec['Point Jumlah ' + role] = arr.length;
      var skor = arr.filter(function (o) { return o.skor !== null; })
                    .map(function (o) { return o.skor; });
      if (skor.length) {
        var avg = skor.reduce(function (a, b) { return a + b; }, 0) / skor.length;
        rec['Point Rata-rata Skor ' + role] = avg.toFixed(2) + '%';
      }
      rec['Point Kena Boom ' + role] = arr.filter(function (o) {
        return /^ya|yes$/i.test(String(o.kpi['Boom'] || '').trim());
      }).length;
      rec['Point Dapat Boost ' + role] = arr.filter(function (o) {
        return /^ya|yes$/i.test(String(o.kpi['Boost'] || '').trim());
      }).length;
    });

    return rec;
  });
}

// ─── Menu ────────────────────────────────────────────────────────────────────

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('FO Performance')
    .addItem('Cek Koneksi Sumber', 'cekKoneksiSumber')
    .addItem('Ringkasan Data', 'ringkasanData')
    .addItem('Preview Payload User', 'previewPayloadUser')
    .addItem('Tes Sync (catat ke Log)', 'tesSyncManual')
    .addItem('Buka Log Sync', 'bukaLogSync');
  if (typeof previewHmb === 'function') menu.addItem('Preview Data HMB', 'previewHmb');
  menu.addToUi();
}

function cekKoneksiSumber() {
  var ui = SpreadsheetApp.getUi();
  var ss = prodSS_();
  var lines = ['Sumber: ' + ss.getName(), ''];
  var need = [EMP_SHEET].concat(Object.keys(PERF_SHEET).map(function (r) { return PERF_SHEET[r]; }));
  need.forEach(function (t) {
    var s = ss.getSheetByName(t);
    lines.push((s ? '✓ ' : '✗ ') + t + (s ? ' (' + (s.getLastRow() - 1) + ' baris)' : ' TIDAK ADA'));
  });
  ui.alert('Cek Koneksi Sumber', lines.join('\n'), ui.ButtonSet.OK);
}

function ringkasanData() {
  var ui = SpreadsheetApp.getUi();
  var emp = bacaEmployee_();
  var lines = ['Employee ter-index: ' + Object.keys(emp).length, ''];
  ['BP', 'BM', 'AM', 'RM'].forEach(function (role) {
    var byNik = bacaPerf_(role);
    var niks = Object.keys(byNik);
    var cocok = niks.filter(function (n) { return emp[n]; }).length;
    var kirim = roleDikirim_(role) ? ' [DIKIRIM]' : ' [ditahan]';
    lines.push(role + ': ' + niks.length + ' orang, ' + cocok + ' cocok wilayah' + kirim);
  });
  ui.alert('Ringkasan Data', lines.join('\n'), ui.ButtonSet.OK);
}

/**
 * Jalankan proses build seperti doGet, tapi manual dari menu — supaya bisa
 * lihat hasil dan log tanpa menunggu backend memanggil. Mencatat ke Log Sync.
 */
function tesSyncManual() {
  var ui = SpreadsheetApp.getUi();
  var mulai = new Date().getTime();
  try {
    var stats = [];
    var ringkasan = cfgRingkasanTimAktif_()
      ? ringkasanTimPerNik_() : null;
    var users = buildUsersFull_(ringkasan, stats);
    var branches = buildBranchesFull_();
    var durasi = ((new Date().getTime() - mulai) / 1000).toFixed(1) + 's';

    catatLog_('tesManual', 'BERHASIL',
      'users=' + users.length + ' branches=' + branches.length + ' | ' + durasi,
      { stats: stats });

    var ringkasStr = stats.map(function (s) {
      return s.role + (s.dikirim ? ' [kirim]' : ' [tahan]') +
             ': ' + s.lolos + '/' + s.total +
             (s.skip_no_employee ? ' (skip ' + s.skip_no_employee + ' tak ada di DB_Employee)' : '');
    }).join('\n');

    ui.alert('Tes Sync Berhasil',
      'Total user: ' + users.length + '\nTotal branch: ' + branches.length +
      '\nDurasi: ' + durasi + '\n\nPer role:\n' + ringkasStr +
      '\n\nDetail lengkap ada di tab "Log Sync".', ui.ButtonSet.OK);
  } catch (err) {
    catatLog_('tesManual', 'GAGAL', String(err && err.message ? err.message : err), {});
    ui.alert('Tes Sync GAGAL', String(err) + '\n\nDicatat di tab "Log Sync".', ui.ButtonSet.OK);
  }
}

function bukaLogSync() {
  var sheet = ss_().getSheetByName('Log Sync');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Belum ada log. Jalankan "Tes Sync" dulu, atau tunggu backend memanggil.');
    return;
  }
  ss_().setActiveSheet(sheet);
}

function previewPayloadUser() {
  var ui = SpreadsheetApp.getUi();
  var users = buildUsers_(null);
  if (!users.length) { ui.alert('Tidak ada user (cek ROLE_DIKIRIM & data).'); return; }

  var sheet = sheetOrCreate_('Preview Payload');
  sheet.clear();
  var contoh = users[0];
  var baris = [['Field', 'Nilai (' + contoh.full_name + ' / ' + contoh.role + ')']];
  Object.keys(contoh).forEach(function (k) {
    baris.push([k, String(contoh[k])]);
  });
  sheet.getRange(1, 1, baris.length, 2).setValues(baris);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  ss_().setActiveSheet(sheet);
  ui.alert('Preview Payload',
    'Total user dikirim: ' + users.length +
    '\nContoh baris pertama ada di tab "Preview Payload".', ui.ButtonSet.OK);
}
