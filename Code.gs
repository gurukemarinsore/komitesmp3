/**
 * Backend Aplikasi Pembayaran Komite SMP Negeri 3 Jayapura
 * Google Apps Script + Google Spreadsheet
 * Tahun Pelajaran 2026/2027
 * Nominal Komite: Rp50.000 / bulan / siswa
 *
 * CARA AWAL:
 * 1. Buat Google Spreadsheet kosong.
 * 2. Extensions > Apps Script.
 * 3. Ganti isi Code.gs dengan file ini.
 * 4. Jalankan fungsi setupAplikasi() satu kali.
 * 5. Deploy sebagai Web app.
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
  SESSION_HOURS: 12
});

/**
 * Jalankan fungsi ini SATU KALI dari editor Apps Script.
 * Aman dijalankan ulang: tidak menghapus data siswa/transaksi yang sudah ada.
 */
function setupAplikasi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Google Spreadsheet yang akan dipakai sebagai backend.');

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  if (!PropertiesService.getScriptProperties().getProperty('API_SECRET')) {
    PropertiesService.getScriptProperties().setProperty('API_SECRET', Utilities.getUuid() + Utilities.getUuid());
  }

  try { ss.setSpreadsheetTimeZone(APP.TZ); } catch (err) {}
  try { ss.setSpreadsheetLocale('id_ID'); } catch (err) {}

  const siswa = ensureSheet_(ss, APP.SHEETS.SISWA, APP.HEADERS.SISWA);
  const kelas = ensureSheet_(ss, APP.SHEETS.KELAS, APP.HEADERS.KELAS);
  const pengguna = ensureSheet_(ss, APP.SHEETS.PENGGUNA, APP.HEADERS.PENGGUNA);
  const pembayaran = ensureSheet_(ss, APP.SHEETS.PEMBAYARAN, APP.HEADERS.PEMBAYARAN);
  const pengaturan = ensureSheet_(ss, APP.SHEETS.PENGATURAN, APP.HEADERS.PENGATURAN);
  const log = ensureSheet_(ss, APP.SHEETS.LOG, APP.HEADERS.LOG);

  // Format kolom yang harus diperlakukan sebagai teks agar angka nol di depan tidak hilang.
  siswa.getRange('B:B').setNumberFormat('@');
  pengguna.getRange('B:C').setNumberFormat('@');
  pembayaran.getRange('C:C').setNumberFormat('@');
  pembayaran.getRange('H:H').setNumberFormat('[$Rp-id-ID]#,##0');
  pembayaran.getRange('I:I').setNumberFormat('dd/MM/yyyy');
  pembayaran.getRange('J:J').setNumberFormat('HH:mm:ss');
  log.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');

  // Data kelas VII A s.d. VII J.
  const classRows = [];
  for (let i = 0; i < 10; i++) {
    const letter = String.fromCharCode(65 + i);
    classRows.push([
      'K07' + letter,
      'VII ' + letter,
      'Isi Nama Wali VII ' + letter,
      'wali7' + letter.toLowerCase(),
      'AKTIF'
    ]);
  }
  upsertDefaultsByKey_(kelas, APP.HEADERS.KELAS, classRows, 0);

  // Akun awal. Password sengaja disimpan terbaca sesuai kebutuhan pengguna.
  const userRows = [[
    'U001', 'bendahara', '123456', 'Bendahara Komite', 'BENDAHARA', '', 'AKTIF'
  ]];
  for (let i = 0; i < 10; i++) {
    const letter = String.fromCharCode(65 + i);
    const lower = letter.toLowerCase();
    userRows.push([
      'U' + String(i + 2).padStart(3, '0'),
      'wali7' + lower,
      '7' + lower + '123',
      'Wali Kelas VII ' + letter,
      'WALI',
      'VII ' + letter,
      'AKTIF'
    ]);
  }
  upsertDefaultsByKey_(pengguna, APP.HEADERS.PENGGUNA, userRows, 0);

  const settingRows = [
    ['NAMA_SEKOLAH', 'SMP Negeri 3 Jayapura'],
    ['TAHUN_PELAJARAN', '2026/2027'],
    ['JENJANG', 'VII'],
    ['NOMINAL_KOMITE', 50000],
    ['BULAN_MULAI', 'JULI 2026'],
    ['BULAN_AKHIR', 'JUNI 2027'],
    ['NAMA_APLIKASI', 'Komite SMP Negeri 3 Jayapura']
  ];
  upsertDefaultsByKey_(pengaturan, APP.HEADERS.PENGATURAN, settingRows, 0);
  pengaturan.getRange('B:B').setNumberFormat('@');
  const nominalCell = findSettingRow_(pengaturan, 'NOMINAL_KOMITE');
  if (nominalCell > 0) pengaturan.getRange(nominalCell, 2).setNumberFormat('[$Rp-id-ID]#,##0');

  // Rapikan lebar kolom.
  setWidths_(siswa, [100, 135, 260, 110, 95]);
  setWidths_(kelas, [100, 115, 240, 150, 95]);
  setWidths_(pengguna, [90, 150, 140, 230, 120, 110, 95]);
  setWidths_(pembayaran, [180, 100, 135, 250, 110, 115, 90, 130, 130, 110, 160, 100]);
  setWidths_(pengaturan, [220, 280]);
  setWidths_(log, [170, 160, 220, 190, 110]);

  // Hapus sheet bawaan yang benar-benar kosong agar hanya tersisa sheet aplikasi.
  const required = new Set(Object.values(APP.SHEETS));
  ss.getSheets().forEach(sh => {
    if (!required.has(sh.getName()) && sh.getLastRow() === 0 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(sh); } catch (err) {}
    }
  });

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
    try {
      payload = JSON.parse(payloadText);
    } catch (err) {
      throw new Error('Format payload tidak valid.');
    }

    const action = String(payload.action || '').trim();
    if (!action) throw new Error('Action tidak boleh kosong.');

    let data;
    switch (action) {
      case 'publicConfig':
        data = publicConfig_();
        break;
      case 'loginStaff':
        data = loginStaff_(payload);
        break;
      case 'loginStudent':
        data = loginStudent_(payload);
        break;
      case 'dashboard':
        data = dashboard_(payload);
        break;
      case 'listStudents':
        data = listStudents_(payload);
        break;
      case 'studentDetail':
        data = studentDetail_(payload);
        break;
      case 'addPayment':
        data = addPayment_(payload);
        break;
      case 'cancelPayment':
        data = cancelPayment_(payload);
        break;
      case 'summary':
        data = summary_(payload);
        break;
      default:
        throw new Error('Action tidak dikenal: ' + action);
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
    kelas: clean_(user.KELAS)
  };

  logActivity_(session.username, 'LOGIN_' + role, '', '');
  return {
    token: createSession_(session),
    user: session,
    config: publicConfig_()
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
    kelas: clean_(student.KELAS)
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
  try {
    session = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  } catch (err) {
    throw new Error('Sesi tidak valid. Silakan login kembali.');
  }

  if (!session.exp || Date.now() > Number(session.exp)) throw new Error('Sesi telah berakhir. Silakan login kembali.');
  if (allowedRoles && !allowedRoles.includes(clean_(session.role).toUpperCase())) throw new Error('Anda tidak memiliki izin untuk tindakan ini.');
  return session;
}

// ========================= API DATA =========================

function publicConfig_() {
  requireSetup_();
  const s = settings_();
  const periods = periods_();
  const classes = rowsAsObjects_(getSheet_(APP.SHEETS.KELAS))
    .filter(r => clean_(r.STATUS).toUpperCase() === 'AKTIF')
    .map(r => ({ id: clean_(r.ID_KELAS), nama: clean_(r.NAMA_KELAS) }));

  return {
    appName: s.NAMA_APLIKASI || 'Komite SMP Negeri 3 Jayapura',
    schoolName: s.NAMA_SEKOLAH || 'SMP Negeri 3 Jayapura',
    schoolYear: s.TAHUN_PELAJARAN || '2026/2027',
    level: s.JENJANG || 'VII',
    amount: Number(s.NOMINAL_KOMITE || 50000),
    startMonth: s.BULAN_MULAI || 'JULI 2026',
    endMonth: s.BULAN_AKHIR || 'JUNI 2027',
    periods: periods,
    classes: classes,
    currentPeriod: currentPeriod_(periods)
  };
}

function dashboard_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA', 'WALI']);
  const periods = periods_();
  const requested = normalizePeriod_(payload.bulan, payload.tahun, periods) || currentPeriod_(periods);
  const kelas = session.role === 'WALI' ? session.kelas : clean_(payload.kelas);
  const stats = paymentSummary_(requested.bulan, requested.tahun, kelas);

  return {
    user: session,
    period: requested,
    stats: stats,
    config: publicConfig_()
  };
}

function listStudents_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA', 'WALI']);
  const query = clean_(payload.query).toLowerCase();
  const kelas = session.role === 'WALI' ? session.kelas : clean_(payload.kelas);
  const payments = paidMap_();

  let students = activeStudents_();
  if (kelas) students = students.filter(r => clean_(r.KELAS) === kelas);
  if (query) {
    students = students.filter(r =>
      clean_(r.NAMA_SISWA).toLowerCase().includes(query) ||
      clean_(r.NISN).toLowerCase().includes(query)
    );
  }

  students.sort((a, b) => clean_(a.NAMA_SISWA).localeCompare(clean_(b.NAMA_SISWA), 'id'));

  return students.map(s => {
    const id = clean_(s.ID_SISWA);
    const paidCount = Object.keys(payments[id] || {}).length;
    return {
      idSiswa: id,
      nisn: clean_(s.NISN),
      nama: clean_(s.NAMA_SISWA),
      kelas: clean_(s.KELAS),
      paidMonths: paidCount,
      unpaidMonths: Math.max(0, periods_().length - paidCount)
    };
  });
}

function studentDetail_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA', 'WALI', 'SISWA']);
  let idSiswa = clean_(payload.idSiswa);
  if (session.role === 'SISWA') idSiswa = session.idSiswa;
  if (!idSiswa) throw new Error('ID siswa tidak ditemukan.');

  const student = activeStudents_().find(r => clean_(r.ID_SISWA) === idSiswa);
  if (!student) throw new Error('Data siswa tidak ditemukan atau tidak aktif.');

  if (session.role === 'WALI' && clean_(student.KELAS) !== clean_(session.kelas)) {
    throw new Error('Anda hanya dapat melihat siswa di kelas Anda.');
  }

  return buildStudentDetail_(student);
}

function addPayment_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA']);
  const idSiswa = clean_(payload.idSiswa);
  const bulan = clean_(payload.bulan).toUpperCase();
  const tahun = Number(payload.tahun);
  if (!idSiswa || !bulan || !tahun) throw new Error('Siswa dan bulan pembayaran wajib dipilih.');

  const period = periods_().find(p => p.bulan === bulan && Number(p.tahun) === tahun);
  if (!period) throw new Error('Periode pembayaran tidak valid.');

  const student = activeStudents_().find(r => clean_(r.ID_SISWA) === idSiswa);
  if (!student) throw new Error('Data siswa tidak ditemukan atau tidak aktif.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const paymentSheet = getSheet_(APP.SHEETS.PEMBAYARAN);
    const existing = rowsAsObjects_(paymentSheet).find(r =>
      clean_(r.ID_SISWA) === idSiswa &&
      clean_(r.BULAN).toUpperCase() === bulan &&
      Number(r.TAHUN) === tahun &&
      clean_(r.STATUS).toUpperCase() === 'LUNAS'
    );
    if (existing) throw new Error('Pembayaran Komite bulan ' + titleCase_(bulan) + ' ' + tahun + ' sudah tercatat lunas.');

    const now = new Date();
    const trx = transactionId_(now);
    const amount = Number(settings_().NOMINAL_KOMITE || 50000);
    const dateText = Utilities.formatDate(now, APP.TZ, 'yyyy-MM-dd');
    const timeText = Utilities.formatDate(now, APP.TZ, 'HH:mm:ss');

    paymentSheet.appendRow([
      trx,
      idSiswa,
      clean_(student.NISN),
      clean_(student.NAMA_SISWA),
      clean_(student.KELAS),
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

    return {
      idTransaksi: trx,
      student: {
        idSiswa: idSiswa,
        nama: clean_(student.NAMA_SISWA),
        nisn: clean_(student.NISN),
        kelas: clean_(student.KELAS)
      },
      bulan: bulan,
      tahun: tahun,
      nominal: amount,
      tanggalBayar: Utilities.formatDate(now, APP.TZ, 'dd/MM/yyyy'),
      waktuBayar: timeText,
      status: 'LUNAS'
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
    const idxTrx = headers.indexOf('ID_TRANSAKSI');
    const idxStatus = headers.indexOf('STATUS');
    const idxStudent = headers.indexOf('ID_SISWA');
    if (idxTrx < 0 || idxStatus < 0) throw new Error('Struktur sheet PEMBAYARAN tidak valid.');

    let foundRow = -1;
    let studentId = '';
    for (let r = 1; r < values.length; r++) {
      if (clean_(values[r][idxTrx]) === trx) {
        foundRow = r + 1;
        studentId = idxStudent >= 0 ? clean_(values[r][idxStudent]) : '';
        if (clean_(values[r][idxStatus]).toUpperCase() === 'BATAL') throw new Error('Transaksi ini sudah dibatalkan.');
        break;
      }
    }
    if (foundRow < 0) throw new Error('Transaksi tidak ditemukan.');

    sh.getRange(foundRow, idxStatus + 1).setValue('BATAL');
    logActivity_(session.username, 'BATAL_PEMBAYARAN', trx, studentId);
    SpreadsheetApp.flush();
    return { idTransaksi: trx, status: 'BATAL' };
  } finally {
    lock.releaseLock();
  }
}

function summary_(payload) {
  const session = requireSession_(payload.token, ['BENDAHARA', 'WALI']);
  const periods = periods_();
  const period = normalizePeriod_(payload.bulan, payload.tahun, periods) || currentPeriod_(periods);
  const kelas = session.role === 'WALI' ? session.kelas : clean_(payload.kelas);
  const summary = paymentSummary_(period.bulan, period.tahun, kelas, true);
  return { period: period, kelas: kelas || 'SEMUA', summary: summary };
}

// ========================= BUSINESS LOGIC =========================

function buildStudentDetail_(student) {
  const periods = periods_();
  const allPayments = rowsAsObjects_(getSheet_(APP.SHEETS.PEMBAYARAN));
  const studentPayments = allPayments.filter(r => clean_(r.ID_SISWA) === clean_(student.ID_SISWA));
  const paid = {};

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
      nominal: tx ? Number(tx.NOMINAL || 0) : Number(settings_().NOMINAL_KOMITE || 50000),
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
      kelas: clean_(student.KELAS)
    },
    paidMonths: paidCount,
    totalMonths: periods.length,
    paidAmount: paidAmount,
    periods: periodStatus,
    history: history
  };
}

function paymentSummary_(bulan, tahun, kelas, includeList) {
  const students = activeStudents_().filter(s => !kelas || clean_(s.KELAS) === clean_(kelas));
  const payments = rowsAsObjects_(getSheet_(APP.SHEETS.PEMBAYARAN));
  const paidIds = new Set(
    payments
      .filter(p =>
        clean_(p.BULAN).toUpperCase() === clean_(bulan).toUpperCase() &&
        Number(p.TAHUN) === Number(tahun) &&
        clean_(p.STATUS).toUpperCase() === 'LUNAS'
      )
      .map(p => clean_(p.ID_SISWA))
  );

  const paid = students.filter(s => paidIds.has(clean_(s.ID_SISWA))).length;
  const total = students.length;
  const amount = Number(settings_().NOMINAL_KOMITE || 50000);
  const result = {
    totalStudents: total,
    paid: paid,
    unpaid: Math.max(0, total - paid),
    percentage: total ? Math.round((paid / total) * 100) : 0,
    totalAmount: paid * amount
  };

  if (includeList) {
    result.students = students
      .map(s => ({
        idSiswa: clean_(s.ID_SISWA),
        nisn: clean_(s.NISN),
        nama: clean_(s.NAMA_SISWA),
        kelas: clean_(s.KELAS),
        status: paidIds.has(clean_(s.ID_SISWA)) ? 'LUNAS' : 'BELUM_BAYAR'
      }))
      .sort((a, b) => a.nama.localeCompare(b.nama, 'id'));
  }
  return result;
}

function paidMap_() {
  const result = {};
  rowsAsObjects_(getSheet_(APP.SHEETS.PEMBAYARAN)).forEach(p => {
    if (clean_(p.STATUS).toUpperCase() !== 'LUNAS') return;
    const id = clean_(p.ID_SISWA);
    if (!id) return;
    if (!result[id]) result[id] = {};
    result[id][periodKey_(clean_(p.BULAN).toUpperCase(), Number(p.TAHUN))] = true;
  });
  return result;
}

function activeStudents_() {
  const sh = getSheet_(APP.SHEETS.SISWA);
  normalizeStudentRows_(sh);
  return rowsAsObjects_(sh).filter(r => clean_(r.STATUS).toUpperCase() === 'AKTIF');
}

// Memudahkan pengisian data: bila ID_SISWA atau STATUS dikosongkan,
// backend mengisinya otomatis saat data pertama kali dibaca.
function normalizeStudentRows_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0].map(clean_);
  const iId = headers.indexOf('ID_SISWA');
  const iNisn = headers.indexOf('NISN');
  const iNama = headers.indexOf('NAMA_SISWA');
  const iStatus = headers.indexOf('STATUS');
  if (iId < 0 || iNisn < 0 || iNama < 0 || iStatus < 0) return;

  const used = new Set();
  let maxNumber = 0;
  for (let r = 1; r < values.length; r++) {
    const id = clean_(values[r][iId]);
    if (id) {
      used.add(id);
      const m = id.match(/^S(\d+)$/i);
      if (m) maxNumber = Math.max(maxNumber, Number(m[1]));
    }
  }

  let next = maxNumber + 1;
  for (let r = 1; r < values.length; r++) {
    const hasStudent = clean_(values[r][iNisn]) || clean_(values[r][iNama]);
    if (!hasStudent) continue;

    if (!clean_(values[r][iId])) {
      let id;
      do { id = 'S' + String(next++).padStart(4, '0'); } while (used.has(id));
      used.add(id);
      sheet.getRange(r + 1, iId + 1).setValue(id);
      values[r][iId] = id;
    }
    if (!clean_(values[r][iStatus])) {
      sheet.getRange(r + 1, iStatus + 1).setValue('AKTIF');
      values[r][iStatus] = 'AKTIF';
    }
  }
}

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
  const values = sheet.getDataRange().getValues();
  if (!values.length || values.length === 1) return [];
  const headers = values[0].map(clean_);
  return values.slice(1)
    .filter(row => row.some(v => clean_(v) !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function settings_() {
  const result = {};
  rowsAsObjects_(getSheet_(APP.SHEETS.PENGATURAN)).forEach(r => {
    const key = clean_(r.KEY);
    if (key) result[key] = r.VALUE;
  });
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
  return match ? match[3] + '/' + match[2] + '/' + match[1] : s;
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
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
