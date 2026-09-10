/**
 * ============================================================================
 * Modul Ringkasan Tim — Visibilitas Atasan & HMB
 * ============================================================================
 * File ketiga dalam project Apps Script yang sama (Code.gs, BCP.gs, ini).
 *
 * TUJUAN: memberi tiap atasan ringkasan performa bawahannya, dan memberi
 * HMB ringkasan per regional di pulaunya. Semua di sisi GAS — backend tidak
 * diubah. Ringkasan ditempelkan ke baris user yang bersangkutan di payload
 * buildUsers_, sebagai field "Tim - ...".
 *
 * HIERARKI (disimpulkan dari wilayah, dikonfirmasi bisa dari wilayah):
 *   BM  membawahi BP           di point yang sama
 *   AM  membawahi BM, BP       di area yang sama
 *   RM  membawahi AM, BM, BP   di regional yang sama
 *   HMB melihat per regional   di pulau yang sama
 *
 * CAKUPAN RINGKASAN (dikonfirmasi):
 *   - Jumlah bawahan per role
 *   - Rata-rata skor tim
 *   - Jumlah di bawah target
 *   - Nama 3 terbawah dan 3 teratas
 *   - Dihitung dari periode terbaru saja
 *
 * CATATAN: ini ringkasan, bukan data granular. Atasan tidak bisa menanyakan
 * skor spesifik satu bawahan kecuali bawahan itu muncul di daftar 3 teratas
 * atau terbawah. Query granular on-demand memerlukan perubahan backend dan
 * sengaja tidak dikerjakan di sini.
 * ============================================================================
 */

// ─── Konfigurasi ────────────────────────────────────────────────────────────

/** true = ringkasan tim dikirim ke atasan. false = tidak sama sekali. */
var RINGKASAN_TIM_AKTIF = true;

/**
 * Ambang "di bawah target". Total Skor KPI dinyatakan dalam persen pada
 * data Q3; skor di bawah nilai ini dihitung sebagai belum mencapai target.
 */
var AMBANG_BAWAH_TARGET = 100;   // persen

/** Berapa nama yang ditampilkan di daftar teratas dan terbawah. */
var JUMLAH_SOROTAN = 3;

/** Kolom skor yang dipakai untuk semua perhitungan ringkasan. */
var KOLOM_SKOR = 'Total Skor KPI';

/**
 * ── GRANULAR: data lengkap tiap bawahan di konteks atasan (Cara 1) ──
 *
 * Ketika aktif untuk sebuah role, atasan role itu menerima data granular
 * SETIAP bawahannya di konteks, sehingga bisa ditanya per individu.
 *
 * PERINGATAN SKALA: RM dan HMB bisa punya ratusan bawahan. Mengaktifkan
 * granular untuk mereka membuat konteks membengkak di SETIAP query, bukan
 * hanya saat menanyakan bawahan. Biarkan false kecuali untuk pengujian.
 *
 * Toggle utama: set false untuk mematikan granular sepenuhnya dalam sekejap.
 */
var GRANULAR_AKTIF = true;

/** Role mana yang menerima data granular bawahan. Hapus dari daftar untuk mematikan. */
var GRANULAR_UNTUK_ROLE = ['BM'];   // aman untuk tes; tambah 'AM','RM' dengan sadar

/**
 * Kolom bawahan yang dibawa ke konteks granular. Dibatasi agar konteks
 * tidak meledak. Kosongkan array ([]) untuk membawa SEMUA kolom.
 */
var GRANULAR_KOLOM = [
  'Total Skor KPI', 'Rangking', 'Grouping Skor KPI',
  'Kena Boom?', 'Dapat Boost?', 'Parameter yang belum tercapai'
];

/**
 * Batas jumlah bawahan yang di-inject granular per atasan.
 *
 * Cara mengatur, dari paling sederhana:
 *   - Angka biasa (mis. 40): berlaku untuk semua role yang granular-nya aktif.
 *   - 0 atau -1: TANPA BATAS — semua bawahan di-inject berapa pun jumlahnya.
 *   - Objek per-role: batas berbeda tiap role, mis.
 *       { BM: 0, AM: 60, RM: 100 }   (BM tanpa batas, AM maks 60, RM maks 100)
 *     Role yang tidak disebut memakai GRANULAR_MAKS_DEFAULT.
 *
 * PERINGATAN: tanpa batas untuk RM/HMB bisa membuat konteks sangat besar di
 * setiap query mereka. Aman untuk BM. Untuk role besar, pakai angka.
 */
var GRANULAR_MAKS_BAWAHAN = 10;        // 0 = tanpa batas (untuk testing penuh)

/** Dipakai bila GRANULAR_MAKS_BAWAHAN berupa objek dan role tidak tercantum. */
var GRANULAR_MAKS_DEFAULT = 40;

/**
 * Kembalikan batas efektif untuk sebuah role.
 * Mengembalikan 0 berarti tanpa batas.
 */
function batasGranular_(role) {
  var b = GRANULAR_MAKS_BAWAHAN;

  // Bentuk objek per-role.
  if (b && typeof b === 'object') {
    if (b.hasOwnProperty(role)) b = b[role];
    else b = GRANULAR_MAKS_DEFAULT;
  }

  var n = Number(b);
  if (isNaN(n) || n < 0) return 0;   // nilai aneh atau negatif -> tanpa batas
  return n;                          // 0 = tanpa batas
}

// ─── Pembacaan arsip jadi baris per orang ───────────────────────────────────

/**
 * Kumpulkan seluruh karyawan (periode terbaru) dari empat arsip menjadi
 * daftar objek datar dengan wilayah dan skor. Ini bahan dasar semua ringkasan.
 */
function bacaSemuaKaryawan_() {
  var orangMap = {}; // key: nik -> object orang (otomatis dedup NIK)
  var targetPeriode = (typeof getPeriodeTarget_ === 'function') ? getPeriodeTarget_() : '';
  SEMUA_ARSIP.forEach(function (arsipName) {
    var sheet = ss_().getSheetByName(arsipName);
    if (!sheet || sheet.getLastRow() < 2) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var iPer   = headers.indexOf('periode');
    var iRole  = headers.indexOf('role');
    var iNik   = findColIdx_(headers, KEY_NIK);
    var iNama  = findColIdx_(headers, KEY_NAMA);
    var iPoint = findColIdx_(headers, 'point');
    var iArea  = findColIdx_(headers, 'area');
    var iReg   = findColIdx_(headers, 'regional');
    var iPulau = findColIdx_(headers, 'pulau');
    var iSkor  = findColIdx_(headers, KOLOM_SKOR.toLowerCase());
    var adaTarget = targetPeriode && data.some(function (r) { 
      var p = typeof formatPeriodeCell_ === 'function' ? formatPeriodeCell_(r[iPer]) : String(r[iPer]).trim();
      return p === targetPeriode; 
    });
    var latest = '';
    if (!adaTarget) {
      data.forEach(function (r) {
        var p = typeof formatPeriodeCell_ === 'function' ? formatPeriodeCell_(r[iPer]) : String(r[iPer]).trim();
        if (p > latest) latest = p;
      });
    }
    var periodeDipakai = adaTarget ? targetPeriode : latest;
    data.forEach(function (r) {
      var periode = typeof formatPeriodeCell_ === 'function' ? formatPeriodeCell_(r[iPer]) : String(r[iPer]).trim();
      if (periode !== periodeDipakai) return;
      var nik = iNik >= 0 ? String(r[iNik]).trim() : '';
      if (!nik) return;
      var skipIdx = [iPer, iRole, iNik, iNama, iPoint, iArea, iReg, iPulau];
      var kpi = {};
      headers.forEach(function (h, i) {
        if (skipIdx.indexOf(i) >= 0) return;
        if (!h || h === 'kategori_point') return;
        var v = r[i];
        if (v === '' || v === null) return;
        kpi[String(h).trim()] = v;
      });
      // Simpan ke map. Jika ada NIK yang sama di bulan ini, baris bawah menggantikan baris atas
      orangMap[nik] = {
        nik: nik,
        nama: iNama >= 0 ? String(r[iNama]).trim() : '',
        role: String(r[iRole]).trim(),
        point: iPoint >= 0 ? String(r[iPoint]).trim() : '',
        area: iArea >= 0 ? String(r[iArea]).trim() : '',
        regional: iReg >= 0 ? String(r[iReg]).trim() : '',
        pulau: iPulau >= 0 ? String(r[iPulau]).trim() : '',
        skor: iSkor >= 0 ? skorKeSkala_(r[iSkor]) : null,
        kpi: kpi
      };
    });
  });
  return Object.keys(orangMap).map(function (k) { return orangMap[k]; });
}


/**
 * Peta regional -> pulau, disimpulkan dari baris BP/BM yang punya kolom
 * Pulau. Dipakai untuk mengelompokkan regional ke pulau bagi HMB.
 */
function petaRegionalKePulau_(semuaOrang) {
  var peta = {};
  semuaOrang.forEach(function (o) {
    if (o.regional && o.pulau && !peta[o.regional]) {
      peta[o.regional] = o.pulau;
    }
  });
  return peta;
}

// ─── Perhitungan ringkasan sekelompok orang ─────────────────────────────────

/**
 * Hitung ringkasan dari sekumpulan orang: jumlah per role, rata-rata skor,
 * jumlah di bawah target, serta sorotan teratas dan terbawah.
 * Mengembalikan objek field datar siap ditempel ke payload.
 */
function ringkasKelompok_(orang, prefix) {
  var rec = {};
  if (!orang.length) return rec;

  // Jumlah per role.
  var perRole = {};
  orang.forEach(function (o) {
    perRole[o.role] = (perRole[o.role] || 0) + 1;
  });
  Object.keys(perRole).sort().forEach(function (role) {
    rec[prefix + 'Jumlah ' + role] = perRole[role];
  });
  rec[prefix + 'Jumlah Total'] = orang.length;

  // Rata-rata skor dihitung TERPISAH per role. Skor BP/BM berskala persen
  // besar (60–120%), sedangkan AM/RM berskala berbeda (belasan persen),
  // sehingga rata-rata gabungan akan menyesatkan.
  var berskor = orang.filter(function (o) { return o.skor !== null; });
  if (berskor.length) {
    var perRoleSkor = {};
    berskor.forEach(function (o) {
      (perRoleSkor[o.role] = perRoleSkor[o.role] || []).push(o.skor);
    });
    Object.keys(perRoleSkor).sort().forEach(function (role) {
      var arr = perRoleSkor[role];
      var avg = arr.reduce(function (a, b) { return a + b; }, 0) / arr.length;
      rec[prefix + 'Rata-rata Skor ' + role] = avg.toFixed(2) + '%';
      var bawah = arr.filter(function (s) { return s < AMBANG_BAWAH_TARGET; }).length;
      rec[prefix + 'Di Bawah Target ' + role] = bawah + ' dari ' + arr.length;
    });

    // Sorotan teratas/terbawah tetap lintas role, tapi label menyertakan
    // role sehingga perbedaan skala tetap terbaca jelas.
    var urut = berskor.slice().sort(function (a, b) { return a.skor - b.skor; });
    rec[prefix + 'Terbawah'] = urut.slice(0, JUMLAH_SOROTAN).map(fmtSorotan_).join('; ');
    rec[prefix + 'Teratas'] = urut.slice(-JUMLAH_SOROTAN).reverse().map(fmtSorotan_).join('; ');
  }

  return rec;
}

function fmtSorotan_(o) {
  var nama = o.nama || o.nik;
  return nama + ' (' + o.skor.toFixed(2) + '%, ' + o.role + ')';
}

// ─── Penempelan ke payload user ─────────────────────────────────────────────

/**
 * Bangun peta nik -> field ringkasan tim, untuk semua atasan dan HMB.
 * Dipanggil dari buildUsers_ di Code.gs; hasilnya digabung ke record user.
 */
function ringkasanTimPerNik_() {
  if (!RINGKASAN_TIM_AKTIF) return {};
  var semua = bacaSemuaKaryawan_();
  if (!semua.length) return {};
  var regToPulau = petaRegionalKePulau_(semua);
  var hasil = {};
  // INDEX CEPAT: Kelompokkan bawahan sekali saja di awal
  var bpByPoint = {};
  var staffByArea = {};
  var staffByReg = {};
  semua.forEach(function (o) {
    if (o.role === 'BP' && o.point) {
      (bpByPoint[o.point] = bpByPoint[o.point] || []).push(o);
    }
    if ((o.role === 'BP' || o.role === 'BM') && o.area) {
      (staffByArea[o.area] = staffByArea[o.area] || []).push(o);
    }
    if ((o.role === 'BP' || o.role === 'BM' || o.role === 'AM') && o.regional) {
      (staffByReg[o.regional] = staffByReg[o.regional] || []).push(o);
    }
  });
  semua.forEach(function (atasan) {
    // BP tidak punya bawahan, langsung skip (menghemat 90% waktu eksekusi)
    if (atasan.role !== 'BM' && atasan.role !== 'AM' && atasan.role !== 'RM') return;
    var bawahan = [];
    if (atasan.role === 'BM' && atasan.point) {
      bawahan = (bpByPoint[atasan.point] || []).filter(function (o) { return o.nik !== atasan.nik; });
    } else if (atasan.role === 'AM' && atasan.area) {
      bawahan = (staffByArea[atasan.area] || []).filter(function (o) { return o.nik !== atasan.nik; });
    } else if (atasan.role === 'RM' && atasan.regional) {
      bawahan = (staffByReg[atasan.regional] || []).filter(function (o) { return o.nik !== atasan.nik; });
    }
    if (!bawahan.length) return;
    var rec = ringkasKelompok_(bawahan, 'Tim - ');
    if (GRANULAR_AKTIF && GRANULAR_UNTUK_ROLE.indexOf(atasan.role) >= 0) {
      var batas = batasGranular_(atasan.role);
      if (batas > 0 && bawahan.length > batas) {
        rec['Tim - Catatan Detail'] =
          'Detail per individu tidak ditampilkan karena jumlah bawahan (' +
          bawahan.length + ') melebihi batas ' + batas +
          '. Hanya ringkasan yang tersedia.';
      } else {
        tempelGranular_(rec, bawahan);
      }
    }
    hasil[atasan.nik] = rec;
  });
  return { perNik: hasil, semua: semua, regToPulau: regToPulau };
}
/**
 * Tempelkan data granular tiap bawahan ke record atasan, dengan awalan
 * "Bawahan <nama> - <kolom>". Kolom dibatasi GRANULAR_KOLOM, insentif
 * tetap disembunyikan mengikuti toggle di Code.gs.
 */
function tempelGranular_(rec, bawahan) {
  bawahan.forEach(function (b) {
    var nama = b.nama || b.nik;
    var prefix = 'Bawahan ' + nama + ' (' + b.role + ') - ';

    Object.keys(b.kpi).forEach(function (kolom) {
      // Hormati toggle insentif dari Code.gs.
      if (typeof kolomDisembunyikan_ === 'function' && kolomDisembunyikan_(kolom)) return;

      // Kalau GRANULAR_KOLOM diisi, hanya kolom itu yang dibawa.
      if (GRANULAR_KOLOM.length) {
        var cocok = GRANULAR_KOLOM.some(function (k) {
          return kolom.toLowerCase() === k.toLowerCase();
        });
        if (!cocok) return;
      }

      var v = b.kpi[kolom];
      if (v === '' || v === null || v === undefined) return;
      rec[prefix + kolom] = (typeof normalisasiNilai_ === 'function')
        ? normalisasiNilai_(v, kolom) : v;
    });
  });
}

/**
 * Tentukan bawahan seorang atasan berdasarkan wilayah.
 * BM: BP di point sama. AM: BM+BP di area sama. RM: AM+BM+BP di regional sama.
 */
function bawahanDari_(atasan, semua) {
  var role = atasan.role;
  if (role === 'BM') {
    return semua.filter(function (o) {
      return o.nik !== atasan.nik && // <- User BM tidak akan masuk jadi bawahan sendiri
             o.role === 'BP' && 
             o.point && o.point === atasan.point;
    });
  }
  if (role === 'AM') {
    return semua.filter(function (o) {
      return o.nik !== atasan.nik &&
             (o.role === 'BP' || o.role === 'BM') &&
             o.area && o.area === atasan.area;
    });
  }
  if (role === 'RM') {
    return semua.filter(function (o) {
      return o.nik !== atasan.nik &&
             (o.role === 'BP' || o.role === 'BM' || o.role === 'AM') &&
             o.regional && o.regional === atasan.regional;
    });
  }
  return [];
}
// ─── HMB ─────────────────────────────────────────────────────────────────────

/**
 * Baca daftar HMB dari tab manual. Struktur kolom: Role, NIK, Nama, Pulau.
 * Mengembalikan array objek HMB.
 */
function bacaDaftarHmb_() {
  var sheet = ss_().getSheetByName(HMB_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
                     .map(function (h) { return String(h).toLowerCase().trim(); });
  var iRole  = headers.indexOf('role');
  var iNik   = headers.indexOf('nik');
  var iNama  = headers.indexOf('nama');
  var iPulau = headers.indexOf('pulau');

  // Hanya NIK yang wajib. Pulau boleh kosong (mis. Head HMB yang nasional).
  if (iNik < 0) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
                  .getDisplayValues();
  var out = [];
  data.forEach(function (r) {
    var nik = String(r[iNik]).trim();
    if (!nik) return;

    // Kolom Pulau bisa berisi lebih dari satu pulau, dipisah koma atau garis
    // miring — untuk HMB yang menangani beberapa pulau sekaligus.
    var pulauRaw = iPulau >= 0 ? String(r[iPulau]).trim() : '';
    var pulauList = pulauRaw.split(/[,/;]+/).map(function (p) {
      return p.trim();
    }).filter(Boolean);

    out.push({
      nik: nik,
      nama: iNama >= 0 ? String(r[iNama]).trim() : '',
      role: iRole >= 0 ? String(r[iRole]).trim() : 'HMB',
      pulau: pulauRaw,          // teks asli, untuk ditampilkan
      pulauList: pulauList      // daftar ternormalisasi, untuk pencocokan
    });
  });
  return out;
}

/**
 * Kenali Head HMB dari kolom Role. Head HMB cakupannya nasional dan
 * biasanya tidak terikat satu pulau.
 */
function isHeadHmb_(role) {
  return /head/i.test(String(role));
}

/**
 * Bangun record user untuk tiap HMB dan Head HMB.
 * - HMB: ringkasan seluruh pulaunya + rincian per regional di pulau itu.
 * - Head HMB: agregat nasional + rincian per pulau.
 * Keduanya tidak punya data KPI sendiri.
 */
function buildHmbUsers_(semua, regToPulau) {
  var daftar = bacaDaftarHmb_();
  if (!daftar.length) return [];

  // Kelompokkan sekali: per regional dan per pulau.
  var perRegional = {};
  var perPulau = {};
  semua.forEach(function (o) {
    if (o.regional) (perRegional[o.regional] = perRegional[o.regional] || []).push(o);
    var pulau = o.pulau || (o.regional ? regToPulau[o.regional] : '');
    if (pulau) (perPulau[pulau] = perPulau[pulau] || []).push(o);
  });

  return daftar.map(function (hmb) {
    if (isHeadHmb_(hmb.role)) {
      return buildHeadHmb_(hmb, semua, perPulau);
    }
    return buildHmbSatuPulau_(hmb, perRegional, regToPulau);
  });
}

/**
 * Record HMB. Menangani satu atau beberapa pulau.
 * - Satu pulau: agregat pulau + rincian per regional.
 * - Beberapa pulau: agregat gabungan semua pulaunya + rincian per pulau,
 *   lalu rincian per regional di dalamnya.
 */
function buildHmbSatuPulau_(hmb, perRegional, regToPulau) {
  var daftarPulau = (hmb.pulauList && hmb.pulauList.length)
    ? hmb.pulauList : (hmb.pulau ? [hmb.pulau] : []);

  var rec = {
    username: hmb.nik,
    full_name: hmb.nama,
    jabatan: hmb.role || 'HMB',
    pulau: hmb.pulau
  };

  if (!daftarPulau.length) {
    rec['Catatan'] = 'Pulau belum diisi untuk HMB ini.';
    return rec;
  }

  var multi = daftarPulau.length > 1;

  // Kumpulkan orang di tiap pulau yang ditangani.
  var orangPerPulau = {};    // pulau -> array orang
  daftarPulau.forEach(function (pulau) {
    var regDiPulau = Object.keys(perRegional).filter(function (reg) {
      return regToPulau[reg] === pulau;
    }).sort();
    var orang = [];
    regDiPulau.forEach(function (reg) {
      orang = orang.concat(perRegional[reg]);
    });
    orangPerPulau[pulau] = { orang: orang, regional: regDiPulau };
  });

  var totalOrang = Object.keys(orangPerPulau).reduce(function (a, p) {
    return a + orangPerPulau[p].orang.length;
  }, 0);
  if (!totalOrang) {
    rec['Catatan'] = 'Belum ada data KPI untuk pulau: ' + daftarPulau.join(', ');
    return rec;
  }

  // Untuk multi-pulau, tambahkan agregat gabungan lebih dulu.
  if (multi) {
    var gabungan = [];
    daftarPulau.forEach(function (p) {
      gabungan = gabungan.concat(orangPerPulau[p].orang);
    });
    var ringkasGabungan = ringkasKelompok_(gabungan, 'Gabungan - ');
    Object.keys(ringkasGabungan).forEach(function (k) { rec[k] = ringkasGabungan[k]; });
  }

  // Rincian per pulau, lalu per regional di dalamnya.
  daftarPulau.forEach(function (pulau) {
    var blok = orangPerPulau[pulau];
    if (!blok.orang.length) return;

    // Untuk satu pulau, prefix cukup "Pulau - " (tanpa nama, agar ringkas).
    // Untuk banyak pulau, sertakan nama pulau supaya tidak tertukar.
    var prefixPulau = multi ? ('Pulau ' + pulau + ' - ') : 'Pulau - ';
    var ringkasPulau = ringkasKelompok_(blok.orang, prefixPulau);
    Object.keys(ringkasPulau).forEach(function (k) { rec[k] = ringkasPulau[k]; });

    blok.regional.forEach(function (reg) {
      var ringkasReg = ringkasKelompok_(perRegional[reg], 'Regional ' + reg + ' - ');
      Object.keys(ringkasReg).forEach(function (k) { rec[k] = ringkasReg[k]; });
    });
  });

  return rec;
}

/** Record Head HMB: agregat nasional + rincian per pulau. */
function buildHeadHmb_(head, semua, perPulau) {
  var rec = {
    username: head.nik,
    full_name: head.nama,
    jabatan: head.role || 'Head HMB',
    cakupan: 'Nasional'
  };

  if (!semua.length) {
    rec['Catatan'] = 'Belum ada data KPI nasional.';
    return rec;
  }

  // Agregat nasional (semua role, semua wilayah).
  var ringkasNasional = ringkasKelompok_(semua, 'Nasional - ');
  Object.keys(ringkasNasional).forEach(function (k) { rec[k] = ringkasNasional[k]; });

  // Rincian per pulau.
  Object.keys(perPulau).sort().forEach(function (pulau) {
    var ringkasPulau = ringkasKelompok_(perPulau[pulau], 'Pulau ' + pulau + ' - ');
    Object.keys(ringkasPulau).forEach(function (k) { rec[k] = ringkasPulau[k]; });
  });

  return rec;
}

// ─── Preview (untuk testing) ─────────────────────────────────────────────────

/**
 * Tampilkan data yang akan diterima tiap HMB / Head HMB, tanpa perlu deploy
 * dan login. Menulis hasilnya ke tab "Preview HMB" agar mudah dibaca.
 */
function previewHmb() {
  var ui = SpreadsheetApp.getUi();

  var semua = bacaSemuaKaryawan_();
  if (!semua.length) {
    ui.alert('Arsip karyawan masih kosong. Jalankan capture dulu.');
    return;
  }

  var regToPulau = petaRegionalKePulau_(semua);
  var hmbUsers = buildHmbUsers_(semua, regToPulau);

  if (!hmbUsers.length) {
    ui.alert('Tab "' + HMB_SHEET + '" masih kosong. Isi dulu daftar HMB-nya.');
    return;
  }

  var sheet = sheetOrCreate_('Preview HMB');
  sheet.clear();

  var baris = [['NIK', 'Nama', 'Jabatan', 'Field', 'Nilai']];
  hmbUsers.forEach(function (u) {
    var meta = ['username', 'full_name', 'jabatan'];
    Object.keys(u).forEach(function (k) {
      if (meta.indexOf(k) >= 0) return;
      baris.push([u.username, u.full_name, u.jabatan, k, String(u[k])]);
    });
    baris.push(['', '', '', '', '']);   // pemisah antar HMB
  });

  sheet.getRange(1, 1, baris.length, 5).setValues(baris);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 5);

  ss_().setActiveSheet(sheet);
  ui.alert('Preview HMB',
    hmbUsers.length + ' HMB/Head HMB diproses. ' +
    'Hasilnya ada di tab "Preview HMB".', ui.ButtonSet.OK);
}
