/**
 * Backend Aplikasi Pembayaran Komite SMP Negeri 3 Jayapura
 * Revisi 1.02 / backend rev002
 * Google Apps Script + Google Spreadsheet
 * Tahun Pelajaran 2026/2027
 * Nominal Komite: Rp50.000 / bulan / siswa
 *
 * Fokus rev002:
 * - Format kelas ringkas 7A s.d. 7J + kompatibilitas format lama VII A.
 * - Data 12 bulan tetap tersedia sekali saat bootstrap untuk rekap lokal yang instan.
 * - Pembayaran multi-bulan melalui antrean frontend, disimpan per bulan di Spreadsheet.
 * - Tanggal pembayaran dapat dipilih Bendahara dan dipertahankan saat sinkronisasi offline.
 * - ID transaksi tetap internal untuk idempotensi/log, tidak perlu ditampilkan di frontend.
 */

const APP = Object.freeze({
  TZ: 'Asia/Jayapura',
  SHEETS: {
    SISWA: 'SISWA',
    KELAS: 'KELAS',
    PENGGUNA: 'PENGGUNA',
    PEMBAYARAN: 'PEMBAYARAN',
    PENGATURAN: 'PENGATURAN',
    LOG: 'LOG_AKTIVITAS'
  },
  HEADERS: {
    SISWA: ['ID_SISWA', 'NISN', 'NAMA_SISWA', 'KELAS', 'STATUS'],
    KELAS: ['ID_KELAS', 'NAMA_KELAS', 'NAMA_WALI', 'USERNAME_WALI', 'STATUS'],
    PENGGUNA: ['ID_USER', 'USERNAME', 'PASSWORD', 'NAMA', 'ROLE', 'KELAS', 'STATUS'],
    PEMBAYARAN: ['ID_TRANSAKSI', 'ID_SISWA', 'NISN', 'NAMA_SISWA', 'KELAS', 'BULAN', 'TAHUN', 'NOMINAL', 'TANGGAL_BAYAR', 'WAKTU_BAYAR', 'PETUGAS', 'STATUS'],
    PENGATURAN: ['KEY', 'VALUE'],
    LOG: ['TANGGAL_WAKTU', 'USER', 'AKSI', 'ID_TRANSAKSI', 'ID_SISWA']
  },
  SESSION_HOURS: 12,
  CACHE_SECONDS: 180
});

let EXEC_STATIC_CACHE_VERSION = null;
let EXEC_PAYMENT_CACHE_VERSION = null;

/**
 * Jalankan fungsi ini SATU KALI pada pemasangan awal.
 * Aman dijalankan ulang: tidak menghapus data siswa/transaksi yang sudah ada.
 */
function setupAplikasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Google Spreadsheet yang akan dipakai sebagai backend.');

  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  if (!props.getProperty('API_SECRET')) props.setProperty('API_SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!props.getProperty('STATIC_CACHE_VERSION')) props.setProperty('STATIC_CACHE_VERSION', '1');
  if (!props.getProperty('PAYMENT_CACHE_VERSION')) props.setProperty('PAYMENT_CACHE_VERSION', '1');

  try { ss.setSpreadsheetTimeZone(APP.TZ); } catch (err) {}
  try { ss.setSpreadsheetLocale('id_ID'); } catch (err) {}

  const siswa = ensureSheet_(ss, APP.SHEETS.SISWA, APP.HEADERS.SISWA);
  const kelas = ensureSheet_(ss, APP.SHEETS.KELAS, APP.HEADERS.KELAS);
  const pengguna = ensureSheet_(ss, APP.SHEETS.PENGGUNA, APP.HEADERS.PENGGUNA);
  const pembayaran = ensureSheet_(ss, APP.SHEETS.PEMBAYARAN, APP.HEADERS.PEMBAYARAN);
  const pengaturan = ensureSheet_(ss, APP.SHEETS.PENGATURAN, APP.HEADERS.PENGATURAN);
  const log = ensureSheet_(ss, APP.SHEETS.LOG, APP.HEADERS.LOG);

  siswa.getRange('B:B').setNumberFormat('@');
  pengguna.getRange('B:C').setNumberFormat('@');
  pembayaran.getRange('C:C').setNumberFormat('@');
  pembayaran.getRange('H:H').setNumberFormat('[$Rp-id-ID]#,##0');
  pembayaran.getRange('I:I').setNumberFormat('dd/MM/yyyy');
  pembayaran.getRange('J:J').setNumberFormat('HH:mm:ss');
  log.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');

  const classRows = [];
  for (let i = 0; i < 10; i++) {
    const letter = String.fromCharCode(65 + i);
    classRows.push(['K07' + letter, '7' + letter, 'Isi Nama Wali 7' + letter, 'wali7' + letter.toLowerCase(), 'AKTIF']);
  }
  upsertDefaultsByKey_(kelas, APP.HEADERS.KELAS, classRows, 0);

  const userRows = [['U001', 'bendahara', '123456', 'Bendahara Komite', 'BENDAHARA', '', 'AKTIF']];
  for (let i = 0; i < 10; i++) {
    const letter = String.fromCharCode(65 + i);
    const lower = letter.toLowerCase();
    userRows.push([
      'U' + String(i + 2).padStart(3, '0'),
      'wali7' + lower,
      '7' + lower + '123',
      'Wali Kelas 7' + letter,
      'WALI',
      '7' + letter,
      'AKTIF'
    ]);
  }
  upsertDefaultsByKey_(pengguna, APP.HEADERS.PENGGUNA, userRows, 0);

  const settingRows = [
    ['NAMA_SEKOLAH', 'SMP Negeri 3 Jayapura'],
    ['TAHUN_PELAJARAN', '2026/2027'],
    ['JENJANG', '7'],
    ['NOMINAL_KOMITE', 50000],
    ['BULAN_MULAI', 'JULI 2026'],
    ['BULAN_AKHIR', 'JUNI 2027'],
    ['NAMA_APLIKASI', 'Komite SMP Negeri 3 Jayapura']
  ];
  upsertDefaultsByKey_(pengaturan, APP.HEADERS.PENGATURAN, settingRows, 0);
  migrateClassFormats_(kelas, pengguna, siswa);
  const jenjangRow = findSettingRow_(pengaturan, 'JENJANG');
  if (jenjangRow > 0) pengaturan.getRange(jenjangRow, 2).setValue('7');
  pengaturan.getRange('B:B').setNumberFormat('@');
  const nominalCell = findSettingRow_(pengaturan, 'NOMINAL_KOMITE');
  if (nominalCell > 0) pengaturan.getRange(nominalCell, 2).setNumberFormat('[$Rp-id-ID]#,##0');

  setWidths_(siswa, [100, 135, 260, 110, 95]);
  setWidths_(kelas, [100, 115, 240, 150, 95]);
  setWidths_(pengguna, [90, 150, 140, 230, 120, 110, 95]);
  setWidths_(pembayaran, [180, 100, 135, 250, 110, 115, 90, 130, 130, 110, 160, 100]);
  setWidths_(pengaturan, [220, 280]);
  setWidths_(log, [170, 160, 220, 190, 110]);

  const required = new Set(Object.values(APP.SHEETS));
  ss.getSheets().forEach(sh => {
    if (!required.has(sh.getName()) && sh.getLastRow() === 0 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(sh); } catch (err) {}
    }
  });

  bumpStaticCacheVersion_();
  bumpPaymentCacheVersion_();
  SpreadsheetApp.flush();
  ss.toast('Setup selesai. Sheet, header, kelas, akun awal, dan pengaturan sudah dibuat.', 'Komite SMPN 3 Jayapura', 8);
  return 'SETUP_SELESAI';
}

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'ping').trim();
    if (action === 'config') return jsonOutput_({ ok: true, data: publicConfig_() });
    return jsonOutput_({
      ok: true,
      message: 'Backend Komite SMP Negeri 3 Jayapura aktif.',
      revision: '1.02',
      time: Utilities.formatDate(new Date(), APP.TZ, 'dd/MM/yyyy HH:mm:ss')
    });
  } catch (err) {
    return jsonOutput_({ ok: false, message: errorMessage_(err) });
  }
}

function doPost(e) {
  try {
    requireSetup_();
    const payloadText = e && e.parameter ? e.parameter.payload : '';
    if (!payloadText) throw new Error('Payload permintaan tidak ditemukan.');

    let payload;
    try { payload = JSON.parse(payloadText); }
    catch (err) { throw new Error('Format payload tidak valid.'); }

    const action = String(payload.action || '').trim();
    if (!action) throw new Error('Action tidak boleh kosong.');

    let data;
    switch (action) {
      case 'publicConfig': data = publicConfig_(); break;
      case 'loginStaff': data = loginStaff_(payload); break;
      case 'loginStudent': data = loginStudent_(payload); break;
      case 'staffBootstrap': data = staffBootstrap_(payload); break;
      case 'dashboard': data = dashboard_(payload); break;
      case 'listStudents': data = listStudents_(payload); break;
      case 'studentDetail': data = studentDetail_(payload); break;
      case 'addPayment': data = addPayment_(payload); break;
      case 'cancelPayment': data = cancelPayment_(payload); break;
      case 'summary': data = summary_(payload); break;
      default: throw new Error('Action tidak dikenal: ' + action);
    }

    return jsonOutput_({ ok: true, data: data });
  } catch (err) {
    return jsonOutput_({ ok: false, message: errorMessage_(err) });
  }
}

// ========================= AUTH =========================

function loginStaff_(payload) {
  const username = clean_(payload.username).toLowerCase();
  const password = clean_(payload.password);
  if (!username || !password) throw new Error('Username dan password wajib diisi.');

  const users = rowsAsObjects_(getSheet_(APP.SHEETS.PENGGUNA));
  const user = users.find(r =>
    clean_(r.USERNAME).toLowerCase() === username &&
    clean_(r.PASSWORD) === password &&
    clean_(r.STATUS).toUpperCase() === 'AKTIF'
  );
  if (!user) throw new Error('Username atau password salah.');

  const role = clean_(user.ROLE).toUpperCase();
  if (!['BENDAHARA', 'WALI'].includes(role)) throw new Error('Role akun tidak dikenali.');
  const expectedRole = clean_(payload.expectedRole).toUpperCase();
  if (expectedRole && ['BENDAHARA', 'WALI'].includes(expectedRole) && role !== expectedRole) {
    throw new Error('Jenis login tidak sesuai dengan akun. Pilih ' + (role === 'WALI' ? 'Wali Kelas' : 'Bendahara') + '.');
  }

  const session = {
    role: role,
    idUser: clean_(user.ID_USER),
    username: clean_(user.USERNAME),
    nama: clean_(user.NAMA),
    kelas: normalizeClassName_(user.KELAS)
  };

  const token = createSession_(session);
  logActivity_(session.username, 'LOGIN_' + role, '', '');
  const bootstrap = buildStaffBootstrapForSession_(session, { mode: 'MONTH' });

  return {
    token: token,
    user: session,
    config: bootstrap.config,
    bootstrap: bootstrap
  };
}

function loginStudent_(payload) {
  const nisn = clean_(payload.nisn);
  if (!nisn) throw new Error('NISN wajib diisi.');

  const students = activeStudents_();
  const student = students.find(r => clean_(r.NISN) === nisn);
  if (!student) throw new Error('NISN tidak ditemukan.');

  const session = {
    role: 'SISWA',
    idSiswa: clean_(student.ID_SISWA),
    nisn: clean_(student.NISN),
    nama: clean_(student.NAMA_SISWA),
    kelas: normalizeClassName_(student.KELAS)
  };

  return {
    token: createSession_(session),
    user: session,
    config: publicConfig_(),
    detail: buildStudentDetail_(student)
  };
}

function createSession_(data) {
  const secret = PropertiesService.getScriptProperties().getProperty('API_SECRET');
  if (!secret) throw new Error('API secret belum tersedia. Jalankan setupAplikasi().');

  const now = Date.now();
  const body = Object.assign({}, data, {
    iat: now,
    exp: now + APP.SESSION_HOURS * 60 * 60 * 1000
  });
  const payload64 = Utilities.base64EncodeWebSafe(JSON.stringify(body), Utilities.Charset.UTF_8).replace(/=+$/g, '');
  const sigBytes = Utilities.computeHmacSha256Signature(payload64, secret, Utilities.Charset.UTF_8);
  const sig64 = Utilities.base64EncodeWebSafe(sigBytes).replace(/=+$/g, '');
  return payload64 + '.' + sig64;
}

function requireSession_(token, allowedRoles) {
  const raw = clean_(token);
  if (!raw || raw.indexOf('.') < 0) throw new Error('Sesi tidak valid. Silakan login kembali.');

  const parts = raw.split('.');
  if (parts.length !== 2) throw new Error('Sesi tidak valid. Silakan login kembali.');

  const secret = PropertiesService.getScriptProperties().getProperty('API_SECRET');
  const expected = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(parts[0], secret, Utilities.Charset.UTF_8)
  ).replace(/=+$/g, '');
  if (expected !== parts[1]) throw new Error('Sesi tidak valid. Silakan login kembali.');

  let session;
  try { session = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString()); }
  catch (err) { throw new Error('Sesi tidak valid. Silakan login kembali.'); }

  if (!session.exp || Date.now() > Number(session.exp)) throw new Error('Sesi telah berakhir. Silakan login kembali.');
  if (allowedRoles && !allowedRoles.includes(clean_(session.role).toUpperCase())) throw new Error('Anda tidak memiliki izin untuk tindakan ini.');
  return session;
}

// ========================= API DATA =========================

function publicConfig_() {
  requireSetup_();
  const s = settings_();
  const periods = periods_();
  return {
    appName: s.NAMA_APLIKASI || 'Komite SMP Negeri 3 Jayapura',
    schoolName: s.NAMA_SEKOLAH || 'SMP Negeri 3 Jayapura',
    schoolYear: s.TAHUN_PELAJARAN || '2026/2027',
    level: '7',
    amount: Number(s.NOMINAL_KOMITE || 50000),
    startMonth: s.BULAN_MULAI || 'JULI 2026',
    endMonth: s.BULAN_AKHIR || 'JUNI 2027',
    periods: periods,
    classes: classes_(),
    currentPeriod: currentPeriod_(periods),
    revision: '1.02'
  };
}

function staffBootstrap_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA', 'WALI']);
  if (payload.force === true || clean_(payload.force).toLowerCase() === 'true') bumpAllCacheVersions_();
  return buildStaffBootstrapForSession_(session, payload);
}

function buildStaffBootstrapForSession_(session, payload) {
  const config = publicConfig_();
  const periods = config.periods;
  const scope = scopeFromPayload_(payload, periods);
  const kelas = session.role === 'WALI' ? normalizeClassName_(session.kelas) : normalizeClassName_(payload.kelas);
  const rawStudents = activeStudents_();
  const paidSets = paidSetsForPeriods_(periods);
  const students = buildStudentList_(session, rawStudents, paidSets, periods);
  const summary = aggregateSummary_(scope, kelas, rawStudents, paidSets, true, Number(config.amount || 50000));

  return {
    config: config,
    students: students,
    dashboard: { user: session, scope: scope, stats: summaryStatsOnly_(summary), config: config },
    summary: { scope: scope, kelas: kelas || 'SEMUA', summary: summary }
  };
}

function dashboard_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA', 'WALI']);
  if (payload.force === true || clean_(payload.force).toLowerCase() === 'true') bumpAllCacheVersions_();
  const periods = periods_();
  const scope = scopeFromPayload_(payload, periods);
  const kelas = session.role === 'WALI' ? normalizeClassName_(session.kelas) : normalizeClassName_(payload.kelas);
  const rawStudents = activeStudents_();
  const paidSets = paidSetsForPeriods_(periods);
  const amount = Number(settings_().NOMINAL_KOMITE || 50000);
  const summary = aggregateSummary_(scope, kelas, rawStudents, paidSets, false, amount);
  return { user: session, scope: scope, stats: summaryStatsOnly_(summary), config: publicConfig_() };
}

function listStudents_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA', 'WALI']);
  if (payload.force === true || clean_(payload.force).toLowerCase() === 'true') bumpAllCacheVersions_();
  const query = clean_(payload.query).toLowerCase();
  const kelas = session.role === 'WALI' ? normalizeClassName_(session.kelas) : normalizeClassName_(payload.kelas);
  const periods = periods_();
  const paidSets = paidSetsForPeriods_(periods);
  let result = buildStudentList_(session, activeStudents_(), paidSets, periods);

  if (kelas) result = result.filter(s => normalizeClassName_(s.kelas) === normalizeClassName_(kelas));
  if (query) result = result.filter(s => s.nama.toLowerCase().includes(query) || s.nisn.toLowerCase().includes(query));
  return result;
}

function studentDetail_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA', 'WALI', 'SISWA']);
  if (payload.force === true || clean_(payload.force).toLowerCase() === 'true') bumpAllCacheVersions_();

  let idSiswa = clean_(payload.idSiswa);
  if (session.role === 'SISWA') idSiswa = session.idSiswa;
  if (!idSiswa) throw new Error('ID siswa tidak ditemukan.');

  const student = activeStudents_().find(r => clean_(r.ID_SISWA) === idSiswa);
  if (!student) throw new Error('Data siswa tidak ditemukan atau tidak aktif.');
  if (session.role === 'WALI' && normalizeClassName_(student.KELAS) !== normalizeClassName_(session.kelas)) {
    throw new Error('Anda hanya dapat melihat siswa di kelas Anda.');
  }
  return buildStudentDetail_(student);
}

function summary_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA', 'WALI']);
  if (payload.force === true || clean_(payload.force).toLowerCase() === 'true') bumpAllCacheVersions_();
  const periods = periods_();
  const scope = scopeFromPayload_(payload, periods);
  const kelas = session.role === 'WALI' ? normalizeClassName_(session.kelas) : normalizeClassName_(payload.kelas);
  const rawStudents = activeStudents_();
  const paidSets = paidSetsForPeriods_(periods);
  const amount = Number(settings_().NOMINAL_KOMITE || 50000);
  const summary = aggregateSummary_(scope, kelas, rawStudents, paidSets, true, amount);
  return { scope: scope, kelas: kelas || 'SEMUA', summary: summary };
}

// ========================= PEMBAYARAN =========================

function addPayment_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA']);
  const idSiswa = clean_(payload.idSiswa);
  const bulan = clean_(payload.bulan).toUpperCase();
  const tahun = Number(payload.tahun);
  const requestedTrx = normalizeClientTransactionId_(payload.clientTransactionId);
  const selectedDate = normalizePaymentDate_(payload.tanggalPembayaran);
  if (!idSiswa || !bulan || !tahun) throw new Error('Siswa dan bulan pembayaran wajib dipilih.');

  const periods = periods_();
  const period = periods.find(p => p.bulan === bulan && Number(p.tahun) === tahun);
  if (!period) throw new Error('Periode pembayaran tidak valid.');

  const student = activeStudents_().find(r => clean_(r.ID_SISWA) === idSiswa);
  if (!student) throw new Error('Data siswa tidak ditemukan atau tidak aktif.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const paymentSheet = getSheet_(APP.SHEETS.PEMBAYARAN);
    const values = paymentSheet.getDataRange().getValues();
    const headers = values.length ? values[0].map(clean_) : APP.HEADERS.PEMBAYARAN;
    const idx = headerIndexes_(headers);

    if (requestedTrx && idx.ID_TRANSAKSI >= 0) {
      for (let r = 1; r < values.length; r++) {
        if (clean_(values[r][idx.ID_TRANSAKSI]) === requestedTrx) {
          const status = idx.STATUS >= 0 ? clean_(values[r][idx.STATUS]).toUpperCase() : '';
          if (status === 'BATAL') throw new Error('Transaksi ini pernah diterima tetapi sudah dibatalkan.');
          return paymentRowToResponse_(values[r], idx, true);
        }
      }
    }

    for (let r = 1; r < values.length; r++) {
      if (
        clean_(values[r][idx.ID_SISWA]) === idSiswa &&
        clean_(values[r][idx.BULAN]).toUpperCase() === bulan &&
        Number(values[r][idx.TAHUN]) === tahun &&
        clean_(values[r][idx.STATUS]).toUpperCase() === 'LUNAS'
      ) {
        throw new Error('Pembayaran Komite bulan ' + titleCase_(bulan) + ' ' + tahun + ' sudah tercatat lunas.');
      }
    }

    const now = new Date();
    const trx = requestedTrx || transactionId_(now);
    const amount = Number(settings_().NOMINAL_KOMITE || 50000);
    const dateText = selectedDate || Utilities.formatDate(now, APP.TZ, 'yyyy-MM-dd');
    const timeText = Utilities.formatDate(now, APP.TZ, 'HH:mm:ss');
    const className = normalizeClassName_(student.KELAS);

    paymentSheet.appendRow([
      trx,
      idSiswa,
      clean_(student.NISN),
      clean_(student.NAMA_SISWA),
      className,
      bulan,
      tahun,
      amount,
      dateText,
      timeText,
      session.username,
      'LUNAS'
    ]);

    logActivity_(session.username, 'BAYAR_' + bulan + '_' + tahun, trx, idSiswa);
    SpreadsheetApp.flush();
    bumpPaymentCacheVersion_();

    return {
      idTransaksi: trx,
      student: {
        idSiswa: idSiswa,
        nama: clean_(student.NAMA_SISWA),
        nisn: clean_(student.NISN),
        kelas: className
      },
      bulan: bulan,
      tahun: tahun,
      nominal: amount,
      tanggalBayar: displayDate_(dateText),
      waktuBayar: timeText,
      status: 'LUNAS',
      alreadyProcessed: false
    };
  } finally {
    lock.releaseLock();
  }
}

function cancelPayment_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA']);
  const trx = clean_(payload.idTransaksi);
  if (!trx) throw new Error('ID transaksi wajib diisi.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet_(APP.SHEETS.PEMBAYARAN);
    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error('Transaksi tidak ditemukan.');
    const headers = values[0].map(clean_);
    const idx = headerIndexes_(headers);
    if (idx.ID_TRANSAKSI < 0 || idx.STATUS < 0) throw new Error('Struktur sheet PEMBAYARAN tidak valid.');

    let foundRow = -1;
    let studentId = '';
    for (let r = 1; r < values.length; r++) {
      if (clean_(values[r][idx.ID_TRANSAKSI]) === trx) {
        foundRow = r + 1;
        studentId = idx.ID_SISWA >= 0 ? clean_(values[r][idx.ID_SISWA]) : '';
        if (clean_(values[r][idx.STATUS]).toUpperCase() === 'BATAL') throw new Error('Transaksi ini sudah dibatalkan.');
        break;
      }
    }
    if (foundRow < 0) throw new Error('Transaksi tidak ditemukan.');

    sh.getRange(foundRow, idx.STATUS + 1).setValue('BATAL');
    logActivity_(session.username, 'BATAL_PEMBAYARAN', trx, studentId);
    SpreadsheetApp.flush();
    bumpPaymentCacheVersion_();
    return { idTransaksi: trx, status: 'BATAL', idSiswa: studentId };
  } finally {
    lock.releaseLock();
  }
}

function normalizeClientTransactionId_(value) {
  const s = clean_(value).toUpperCase();
  if (!s) return '';
  if (!/^TRX-[A-Z0-9-]{8,60}$/.test(s)) throw new Error('ID transaksi perangkat tidak valid.');
  return s;
}

function paymentRowToResponse_(row, idx, alreadyProcessed) {
  return {
    idTransaksi: clean_(row[idx.ID_TRANSAKSI]),
    student: {
      idSiswa: clean_(row[idx.ID_SISWA]),
      nama: clean_(row[idx.NAMA_SISWA]),
      nisn: clean_(row[idx.NISN]),
      kelas: normalizeClassName_(row[idx.KELAS])
    },
    bulan: clean_(row[idx.BULAN]).toUpperCase(),
    tahun: Number(row[idx.TAHUN]),
    nominal: Number(row[idx.NOMINAL] || 0),
    tanggalBayar: displayDate_(row[idx.TANGGAL_BAYAR]),
    waktuBayar: displayTime_(row[idx.WAKTU_BAYAR]),
    status: clean_(row[idx.STATUS]).toUpperCase(),
    alreadyProcessed: !!alreadyProcessed
  };
}

// ========================= BUSINESS LOGIC =========================

function buildStudentList_(session, rawStudents, paidSets, periods) {
  let students = rawStudents.slice();
  const waliClass = normalizeClassName_(session.kelas);
  if (session.role === 'WALI') students = students.filter(s => normalizeClassName_(s.KELAS) === waliClass);
  students.sort((a, b) => clean_(a.NAMA_SISWA).localeCompare(clean_(b.NAMA_SISWA), 'id'));

  return students.map(s => {
    const id = clean_(s.ID_SISWA);
    let paidCount = 0;
    const paidPeriods = [];
    periods.forEach(p => {
      const key = periodKey_(p.bulan, p.tahun);
      if (paidSets[key]?.has(id)) { paidCount++; paidPeriods.push(key); }
    });
    return {
      idSiswa: id,
      nisn: clean_(s.NISN),
      nama: clean_(s.NAMA_SISWA),
      kelas: normalizeClassName_(s.KELAS),
      paidMonths: paidCount,
      unpaidMonths: Math.max(0, periods.length - paidCount),
      paidPeriods: paidPeriods
    };
  });
}

function buildStudentDetail_(student) {
  const periods = periods_();
  const studentPayments = studentPayments_(clean_(student.ID_SISWA));
  const paid = {};
  const amount = Number(settings_().NOMINAL_KOMITE || 50000);

  studentPayments.forEach(p => {
    if (clean_(p.STATUS).toUpperCase() !== 'LUNAS') return;
    paid[periodKey_(clean_(p.BULAN).toUpperCase(), Number(p.TAHUN))] = p;
  });

  const periodStatus = periods.map(p => {
    const key = periodKey_(p.bulan, p.tahun);
    const tx = paid[key];
    return {
      bulan: p.bulan,
      tahun: p.tahun,
      label: p.label,
      status: tx ? 'LUNAS' : (isPeriodDue_(p, periods) ? 'BELUM_BAYAR' : 'BELUM_JATUH_TEMPO'),
      nominal: tx ? Number(tx.NOMINAL || 0) : amount,
      idTransaksi: tx ? clean_(tx.ID_TRANSAKSI) : ''
    };
  });

  const history = studentPayments
    .filter(p => clean_(p.STATUS).toUpperCase() === 'LUNAS')
    .map(p => ({
      idTransaksi: clean_(p.ID_TRANSAKSI),
      bulan: clean_(p.BULAN).toUpperCase(),
      tahun: Number(p.TAHUN),
      nominal: Number(p.NOMINAL || 0),
      tanggalBayar: displayDate_(p.TANGGAL_BAYAR),
      waktuBayar: displayTime_(p.WAKTU_BAYAR),
      petugas: clean_(p.PETUGAS),
      status: 'LUNAS'
    }))
    .sort((a, b) => periodIndex_(b.bulan, b.tahun, periods) - periodIndex_(a.bulan, a.tahun, periods));

  const paidCount = periodStatus.filter(p => p.status === 'LUNAS').length;
  const paidAmount = history.reduce((sum, x) => sum + Number(x.nominal || 0), 0);

  return {
    student: {
      idSiswa: clean_(student.ID_SISWA),
      nisn: clean_(student.NISN),
      nama: clean_(student.NAMA_SISWA),
      kelas: normalizeClassName_(student.KELAS)
    },
    paidMonths: paidCount,
    totalMonths: periods.length,
    paidAmount: paidAmount,
    periods: periodStatus,
    history: history
  };
}

function aggregateSummary_(scope, kelas, rawStudents, paidSets, includeList, amount) {
  const normalizedClass = normalizeClassName_(kelas);
  const students = rawStudents.filter(s => !normalizedClass || normalizeClassName_(s.KELAS) === normalizedClass);
  const selectedPeriods = scope.periods;
  const duePeriods = selectedPeriods.filter(p => isPeriodDue_(p, periods_()));
  const dueKeys = new Set(duePeriods.map(p => periodKey_(p.bulan, p.tahun)));

  let paidSlots = 0;
  let unpaidDueSlots = 0;
  let futureSlots = 0;
  let fullyPaidStudents = 0;

  const list = students.map(s => {
    const id = clean_(s.ID_SISWA);
    let paidCount = 0;
    let paidDueCount = 0;
    selectedPeriods.forEach(p => {
      const key = periodKey_(p.bulan, p.tahun);
      const isPaid = !!paidSets[key]?.has(id);
      if (isPaid) {
        paidCount++;
        if (dueKeys.has(key)) paidDueCount++;
      }
    });

    const dueCount = duePeriods.length;
    const unpaidDue = Math.max(0, dueCount - paidDueCount);
    const futureUnpaid = Math.max(0, selectedPeriods.length - dueCount - (paidCount - paidDueCount));
    paidSlots += paidCount;
    unpaidDueSlots += unpaidDue;
    futureSlots += futureUnpaid;
    if (paidCount === selectedPeriods.length && selectedPeriods.length) fullyPaidStudents++;

    let status = 'BELUM_JATUH_TEMPO';
    if (paidCount === selectedPeriods.length && selectedPeriods.length) status = 'LUNAS';
    else if (unpaidDue > 0) status = 'BELUM_BAYAR';
    else if (paidCount > 0) status = 'SEBAGIAN';

    return {
      idSiswa: id,
      nisn: clean_(s.NISN),
      nama: clean_(s.NAMA_SISWA),
      kelas: normalizeClassName_(s.KELAS),
      paidMonths: paidCount,
      totalMonths: selectedPeriods.length,
      unpaidDueMonths: unpaidDue,
      futureMonths: futureUnpaid,
      paidAmount: paidCount * amount,
      status: status
    };
  });

  const totalSlots = students.length * selectedPeriods.length;
  const result = {
    totalStudents: students.length,
    totalSlots: totalSlots,
    paid: paidSlots,
    unpaid: unpaidDueSlots,
    future: futureSlots,
    fullyPaidStudents: fullyPaidStudents,
    percentage: totalSlots ? Math.round((paidSlots / totalSlots) * 100) : 0,
    totalAmount: paidSlots * amount
  };

  if (includeList) result.students = list.sort((a, b) => a.nama.localeCompare(b.nama, 'id'));
  return result;
}

function summaryStatsOnly_(summary) {
  return {
    totalStudents: summary.totalStudents,
    totalSlots: summary.totalSlots,
    paid: summary.paid,
    unpaid: summary.unpaid,
    future: summary.future,
    fullyPaidStudents: summary.fullyPaidStudents,
    percentage: summary.percentage,
    totalAmount: summary.totalAmount
  };
}

function scopeFromPayload_(payload, periods) {
  const mode = normalizeScopeMode_(payload.mode);
  if (mode === 'SEM1') return { mode: mode, label: 'Semester 1', periods: periods.slice(0, 6), anchor: periods[0] };
  if (mode === 'SEM2') return { mode: mode, label: 'Semester 2', periods: periods.slice(6, 12), anchor: periods[6] };
  if (mode === 'YEAR') return { mode: mode, label: 'Tahun Ajaran ' + clean_(settings_().TAHUN_PELAJARAN || '2026/2027'), periods: periods.slice(), anchor: periods[0] };

  const period = normalizePeriod_(payload.bulan, payload.tahun, periods) || currentPeriod_(periods);
  return { mode: 'MONTH', label: period.label, periods: [period], anchor: period };
}

function normalizeScopeMode_(value) {
  const mode = clean_(value).toUpperCase();
  return ['MONTH', 'SEM1', 'SEM2', 'YEAR'].includes(mode) ? mode : 'MONTH';
}

// ========================= CACHE =========================

function scriptCache_() {
  return CacheService.getScriptCache();
}

function staticCacheVersion_() {
  if (EXEC_STATIC_CACHE_VERSION !== null) return EXEC_STATIC_CACHE_VERSION;
  const props = PropertiesService.getScriptProperties();
  let value = Number(props.getProperty('STATIC_CACHE_VERSION') || 1);
  if (!value) value = 1;
  EXEC_STATIC_CACHE_VERSION = value;
  return value;
}

function paymentCacheVersion_() {
  if (EXEC_PAYMENT_CACHE_VERSION !== null) return EXEC_PAYMENT_CACHE_VERSION;
  const props = PropertiesService.getScriptProperties();
  let value = Number(props.getProperty('PAYMENT_CACHE_VERSION') || 1);
  if (!value) value = 1;
  EXEC_PAYMENT_CACHE_VERSION = value;
  return value;
}

function bumpStaticCacheVersion_() {
  const props = PropertiesService.getScriptProperties();
  const next = Number(props.getProperty('STATIC_CACHE_VERSION') || 1) + 1;
  props.setProperty('STATIC_CACHE_VERSION', String(next));
  EXEC_STATIC_CACHE_VERSION = next;
  return next;
}

function bumpPaymentCacheVersion_() {
  const props = PropertiesService.getScriptProperties();
  const next = Number(props.getProperty('PAYMENT_CACHE_VERSION') || 1) + 1;
  props.setProperty('PAYMENT_CACHE_VERSION', String(next));
  EXEC_PAYMENT_CACHE_VERSION = next;
  return next;
}

function bumpAllCacheVersions_() {
  bumpStaticCacheVersion_();
  bumpPaymentCacheVersion_();
}

function getStaticCacheJson_(name) {
  const raw = scriptCache_().get('S' + staticCacheVersion_() + ':' + name);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (err) { return null; }
}

function putStaticCacheJson_(name, value) {
  try { scriptCache_().put('S' + staticCacheVersion_() + ':' + name, JSON.stringify(value), APP.CACHE_SECONDS); }
  catch (err) {}
}

function getPaymentCacheJson_(name) {
  const raw = scriptCache_().get('P' + paymentCacheVersion_() + ':' + name);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (err) { return null; }
}

function putPaymentCacheJson_(name, value) {
  try { scriptCache_().put('P' + paymentCacheVersion_() + ':' + name, JSON.stringify(value), APP.CACHE_SECONDS); }
  catch (err) {}
}

function settings_() {
  const cached = getStaticCacheJson_('settings');
  if (cached) return cached;
  const result = {};
  rowsAsObjects_(getSheet_(APP.SHEETS.PENGATURAN)).forEach(r => {
    const key = clean_(r.KEY);
    if (key) result[key] = r.VALUE;
  });
  putStaticCacheJson_('settings', result);
  return result;
}

function classes_() {
  const cached = getStaticCacheJson_('classes');
  if (cached) return cached.map(c => ({ id: clean_(c.id), nama: normalizeClassName_(c.nama) }));
  const sh = getSheet_(APP.SHEETS.KELAS);
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0].map(clean_);
  const iName = headers.indexOf('NAMA_KELAS');
  const iStatus = headers.indexOf('STATUS');
  const iId = headers.indexOf('ID_KELAS');
  let changed = false;
  const result = [];
  for (let r = 1; r < values.length; r++) {
    if (iStatus >= 0 && clean_(values[r][iStatus]).toUpperCase() !== 'AKTIF') continue;
    const normalized = normalizeClassName_(iName >= 0 ? values[r][iName] : '');
    if (!normalized) continue;
    if (iName >= 0 && clean_(values[r][iName]) !== normalized) {
      values[r][iName] = normalized;
      changed = true;
    }
    result.push({ id: iId >= 0 ? clean_(values[r][iId]) : '', nama: normalized });
  }
  if (changed && iName >= 0) {
    sh.getRange(2, iName + 1, values.length - 1, 1).setValues(values.slice(1).map(r => [r[iName]]));
    SpreadsheetApp.flush();
  }
  putStaticCacheJson_('classes', result);
  return result;
}

function activeStudents_() {
  const cached = getStaticCacheJson_('activeStudents');
  if (cached) return cached;
  const result = readAndNormalizeStudents_().filter(r => clean_(r.STATUS).toUpperCase() === 'AKTIF');
  putStaticCacheJson_('activeStudents', result);
  return result;
}

function readAndNormalizeStudents_() {
  const sheet = getSheet_(APP.SHEETS.SISWA);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(clean_);
  const iId = headers.indexOf('ID_SISWA');
  const iNisn = headers.indexOf('NISN');
  const iNama = headers.indexOf('NAMA_SISWA');
  const iKelas = headers.indexOf('KELAS');
  const iStatus = headers.indexOf('STATUS');
  if (iId < 0 || iNisn < 0 || iNama < 0 || iStatus < 0) return rowsToObjects_(values);

  const used = new Set();
  let maxNumber = 0;
  for (let r = 1; r < values.length; r++) {
    const id = clean_(values[r][iId]);
    if (!id) continue;
    used.add(id);
    const m = id.match(/^S(\d+)$/i);
    if (m) maxNumber = Math.max(maxNumber, Number(m[1]));
  }

  let next = maxNumber + 1;
  let changedId = false;
  let changedStatus = false;
  let changedClass = false;
  for (let r = 1; r < values.length; r++) {
    const hasStudent = clean_(values[r][iNisn]) || clean_(values[r][iNama]);
    if (!hasStudent) continue;

    if (!clean_(values[r][iId])) {
      let id;
      do { id = 'S' + String(next++).padStart(4, '0'); } while (used.has(id));
      used.add(id);
      values[r][iId] = id;
      changedId = true;
    }
    if (!clean_(values[r][iStatus])) {
      values[r][iStatus] = 'AKTIF';
      changedStatus = true;
    }
    if (iKelas >= 0) {
      const normalizedClass = normalizeClassName_(values[r][iKelas]);
      if (normalizedClass && clean_(values[r][iKelas]) !== normalizedClass) {
        values[r][iKelas] = normalizedClass;
        changedClass = true;
      }
    }
  }

  const rowCount = values.length - 1;
  if (rowCount > 0 && changedId) sheet.getRange(2, iId + 1, rowCount, 1).setValues(values.slice(1).map(r => [r[iId]]));
  if (rowCount > 0 && changedStatus) sheet.getRange(2, iStatus + 1, rowCount, 1).setValues(values.slice(1).map(r => [r[iStatus]]));
  if (rowCount > 0 && changedClass && iKelas >= 0) sheet.getRange(2, iKelas + 1, rowCount, 1).setValues(values.slice(1).map(r => [r[iKelas]]));
  if (changedId || changedStatus || changedClass) SpreadsheetApp.flush();

  return rowsToObjects_(values);
}

function paidSetsForPeriods_(periods) {
  const result = {};
  let missing = false;
  periods.forEach(p => {
    const key = periodKey_(p.bulan, p.tahun);
    const cached = getPaymentCacheJson_('paid:' + key);
    if (Array.isArray(cached)) result[key] = new Set(cached);
    else missing = true;
  });
  if (!missing) return result;

  const built = {};
  periods.forEach(p => built[periodKey_(p.bulan, p.tahun)] = new Set());
  rowsAsObjects_(getSheet_(APP.SHEETS.PEMBAYARAN)).forEach(p => {
    if (clean_(p.STATUS).toUpperCase() !== 'LUNAS') return;
    const key = periodKey_(clean_(p.BULAN).toUpperCase(), Number(p.TAHUN));
    if (built[key]) built[key].add(clean_(p.ID_SISWA));
  });

  Object.keys(built).forEach(key => putPaymentCacheJson_('paid:' + key, Array.from(built[key])));
  return built;
}

function studentPayments_(idSiswa) {
  const cacheName = 'student:' + clean_(idSiswa);
  const cached = getPaymentCacheJson_(cacheName);
  if (Array.isArray(cached)) return cached;
  const result = rowsAsObjects_(getSheet_(APP.SHEETS.PEMBAYARAN))
    .filter(r => clean_(r.ID_SISWA) === clean_(idSiswa));
  putPaymentCacheJson_(cacheName, result);
  return result;
}

// ========================= PERIOD =========================

function periods_() {
  const yearText = clean_(settings_().TAHUN_PELAJARAN || '2026/2027');
  const parts = yearText.split('/');
  const y1 = Number(parts[0]) || 2026;
  const y2 = Number(parts[1]) || (y1 + 1);
  const first = ['JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'];
  const second = ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI'];
  return first.map(b => ({ bulan: b, tahun: y1, label: titleCase_(b) + ' ' + y1 }))
    .concat(second.map(b => ({ bulan: b, tahun: y2, label: titleCase_(b) + ' ' + y2 })));
}

function currentPeriod_(periods) {
  const y = Number(Utilities.formatDate(new Date(), APP.TZ, 'yyyy'));
  const m = Number(Utilities.formatDate(new Date(), APP.TZ, 'M'));
  const names = ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'];
  const found = periods.find(p => p.tahun === y && p.bulan === names[m - 1]);
  if (found) return found;

  const nowKey = y * 100 + m;
  const firstKey = periodNumericKey_(periods[0]);
  const lastKey = periodNumericKey_(periods[periods.length - 1]);
  return nowKey < firstKey ? periods[0] : (nowKey > lastKey ? periods[periods.length - 1] : periods[0]);
}

function isPeriodDue_(period, periods) {
  const nowY = Number(Utilities.formatDate(new Date(), APP.TZ, 'yyyy'));
  const nowM = Number(Utilities.formatDate(new Date(), APP.TZ, 'M'));
  const nowKey = nowY * 100 + nowM;
  const pKey = periodNumericKey_(period);
  const first = periodNumericKey_(periods[0]);
  if (nowKey < first) return false;
  return pKey <= nowKey;
}

function periodNumericKey_(p) {
  const monthMap = { JANUARI:1, FEBRUARI:2, MARET:3, APRIL:4, MEI:5, JUNI:6, JULI:7, AGUSTUS:8, SEPTEMBER:9, OKTOBER:10, NOVEMBER:11, DESEMBER:12 };
  return Number(p.tahun) * 100 + monthMap[clean_(p.bulan).toUpperCase()];
}

function normalizePeriod_(bulan, tahun, periods) {
  const b = clean_(bulan).toUpperCase();
  const t = Number(tahun);
  if (!b || !t) return null;
  return periods.find(p => p.bulan === b && p.tahun === t) || null;
}

function periodKey_(bulan, tahun) {
  return clean_(bulan).toUpperCase() + '-' + Number(tahun);
}

function periodIndex_(bulan, tahun, periods) {
  return periods.findIndex(p => p.bulan === clean_(bulan).toUpperCase() && p.tahun === Number(tahun));
}

// ========================= SHEET HELPERS =========================

function requireSetup_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Backend belum disiapkan. Jalankan setupAplikasi() terlebih dahulu.');
}

function getSpreadsheet_() {
  requireSetup_();
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(id);
}

function getSheet_(name) {
  const sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) throw new Error('Sheet ' + name + ' tidak ditemukan. Jalankan setupAplikasi().');
  return sh;
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#0B3A67')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);
  return sh;
}

function setWidths_(sheet, widths) {
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

function upsertDefaultsByKey_(sheet, headers, rows, keyIndex) {
  const current = sheet.getDataRange().getValues();
  const existing = new Set();
  for (let i = 1; i < current.length; i++) existing.add(clean_(current[i][keyIndex]));
  const toAppend = rows.filter(r => !existing.has(clean_(r[keyIndex])));
  if (toAppend.length) sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
}

function rowsAsObjects_(sheet) {
  return rowsToObjects_(sheet.getDataRange().getValues());
}

function rowsToObjects_(values) {
  if (!values || values.length <= 1) return [];
  const headers = values[0].map(clean_);
  return values.slice(1)
    .filter(row => row.some(v => clean_(v) !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function headerIndexes_(headers) {
  const result = {};
  APP.HEADERS.PEMBAYARAN.forEach(name => result[name] = headers.indexOf(name));
  return result;
}

function findSettingRow_(sheet, key) {
  const values = sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), 1).getValues();
  for (let i = 1; i < values.length; i++) if (clean_(values[i][0]) === key) return i + 1;
  return -1;
}

function logActivity_(user, action, trx, studentId) {
  const sh = getSheet_(APP.SHEETS.LOG);
  sh.appendRow([new Date(), clean_(user), clean_(action), clean_(trx), clean_(studentId)]);
}

// ========================= UTILITIES =========================

function migrateClassFormats_(kelasSheet, penggunaSheet, siswaSheet) {
  const targets = [
    {sheet: kelasSheet, header: 'NAMA_KELAS'},
    {sheet: penggunaSheet, header: 'KELAS'},
    {sheet: siswaSheet, header: 'KELAS'}
  ];
  targets.forEach(item => {
    const values = item.sheet.getDataRange().getValues();
    if (values.length <= 1) return;
    const headers = values[0].map(clean_);
    const idx = headers.indexOf(item.header);
    if (idx < 0) return;
    let changed = false;
    for (let r = 1; r < values.length; r++) {
      const normalized = normalizeClassName_(values[r][idx]);
      if (normalized && clean_(values[r][idx]) !== normalized) {
        values[r][idx] = normalized;
        changed = true;
      }
    }
    if (changed) item.sheet.getRange(2, idx + 1, values.length - 1, 1).setValues(values.slice(1).map(r => [r[idx]]));
  });
}

function normalizeClassName_(value) {
  let s = clean_(value).toUpperCase();
  if (!s) return '';
  s = s.replace(/[._-]+/g, ' ').replace(/\s+/g, ' ').trim();
  let m = s.match(/^VII\s*([A-J])$/);
  if (m) return '7' + m[1];
  m = s.match(/^7\s*([A-J])$/);
  if (m) return '7' + m[1];
  return s;
}

function normalizePaymentDate_(value) {
  const s = clean_(value);
  if (!s) return '';
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Tanggal pembayaran tidak valid.');
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const test = new Date(Date.UTC(y, mo - 1, d, 12, 0, 0));
  if (test.getUTCFullYear() !== y || test.getUTCMonth() + 1 !== mo || test.getUTCDate() !== d) throw new Error('Tanggal pembayaran tidak valid.');
  const today = Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd');
  if (s > today) throw new Error('Tanggal pembayaran tidak boleh melebihi hari ini.');
  return s;
}

function transactionId_(date) {
  const stamp = Utilities.formatDate(date, APP.TZ, 'yyMMdd-HHmmss');
  const random = Math.floor(1000 + Math.random() * 9000);
  return 'TRX-' + stamp + '-' + random;
}

function clean_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function titleCase_(text) {
  const s = clean_(text).toLowerCase();
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : '';
}

function displayDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, APP.TZ, 'dd/MM/yyyy');
  }
  const s = clean_(value);
  const match = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return match[3] + '/' + match[2] + '/' + match[1];
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
  return iso ? iso[3] + '/' + iso[2] + '/' + iso[1] : s;
}

function displayTime_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, APP.TZ, 'HH:mm:ss');
  }
  return clean_(value);
}

function errorMessage_(err) {
  return err && err.message ? String(err.message) : String(err || 'Terjadi kesalahan.');
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
