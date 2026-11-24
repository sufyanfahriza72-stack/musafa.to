const CONFIG = {
  ADMIN_CREDENTIALS: ['MUSAFA-ADMIN', 'ADMIN-MUSAFA', 'SUFYAN-ADMIN'],
  SECURITY_PIN: 'FT',
  TTE_SECRET: 'MUSAFA-TTE-FT',
  TTE_SPECIMEN_URL: 'https://lh3.googleusercontent.com/d/1mDzoyRWgLuyOutFn81FJd2XQ_JcAQtCo',
  TIMEZONE: 'Asia/Jakarta'
};

const SHEETS = {
  AGENDA: {
    name: 'Agenda',
    headers: ['ID', 'Judul', 'Lokasi', 'JatuhTempo', 'Durasi', 'Prioritas', 'Status', 'DisposisiKepada', 'TamuKhusus', 'Notulensi', 'CreatedAt']
  },
  PELAJARAN: {
    name: 'Pelajaran',
    headers: ['ID', 'NamaMapel', 'NamaGuru', 'Hari', 'WaktuMulai', 'WaktuSelesai']
  },
  UJIAN: {
    name: 'Ujian',
    headers: ['ID', 'Hari', 'Jam', 'NamaMapel', 'StatusAktif']
  },
  TUGAS: {
    name: 'Tugas',
    headers: ['ID', 'Deadline', 'NamaMapel', 'NamaGuru', 'DetailTugas', 'Selesai']
  },
  NOTULENSI: {
    name: 'Notulensi',
    headers: ['ID', 'AgendaID', 'Judul', 'Isi', 'TokenTTE', 'Penandatangan', 'Jabatan', 'Tanggal']
  }
};

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (props) return SpreadsheetApp.openById(props);
  throw new Error('Spreadsheet aktif tidak ditemukan. Buka Code.gs dari menu Extensions > Apps Script pada spreadsheet.');
}

function getOrCreateSheet_(meta) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(meta.name);
  if (!sh) {
    sh = ss.insertSheet(meta.name);
  }
  if (sh.getLastRow() === 0 && meta.headers && meta.headers.length > 0) {
    sh.appendRow(meta.headers);
  }
  return sh;
}

function formatCellValue_(val) {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    const year = val.getFullYear();
    if (year === 1899 || year === 1900) {
      return Utilities.formatDate(val, CONFIG.TIMEZONE, 'HH:mm');
    }
    return Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  }
  return String(val).trim();
}

function readSheetObjects_(meta) {
  const sh = getOrCreateSheet_(meta);
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), meta.headers.length);
  if (lastRow < 2) return [];

  const rawHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map((h) => String(h || '').trim());
  const headers = meta.headers.slice();
  rawHeaders.forEach((h) => {
    if (h && !headers.includes(h)) headers.push(h);
  });

  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const result = [];

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const hasAnyValue = row.some((c) => c !== '' && c !== null && c !== undefined);
    if (!hasAnyValue) continue;

    const obj = {};
    for (let c = 0; c < rawHeaders.length; c++) {
      const key = rawHeaders[c] || meta.headers[c];
      if (!key) continue;
      obj[key] = formatCellValue_(row[c]);
    }
    meta.headers.forEach((h) => {
      if (!(h in obj)) obj[h] = '';
    });
    if (obj.ID) result.push(obj);
  }

  return result;
}

function writeRowByHeaders_(sh, headers, recordObj) {
  const rowValues = headers.map((col) => {
    const v = recordObj[col];
    return v === undefined || v === null ? '' : v;
  });

  const targetId = String(recordObj.ID || '').trim();
  const lastRow = sh.getLastRow();

  if (targetId && lastRow >= 2) {
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === targetId) {
        sh.getRange(i + 2, 1, 1, headers.length).setValues([rowValues]);
        return rowValues;
      }
    }
  }

  sh.appendRow(rowValues);
  return rowValues;
}

function verifyPin_(pin) {
  const entered = String(pin || '').trim().toUpperCase();
  const expected = String(CONFIG.SECURITY_PIN || '').trim().toUpperCase();
  return entered !== '' && entered === expected;
}

function verifyLogin(credential) {
  const val = String(credential || '').trim();
  if (!val) {
    return { success: false, message: 'Kredensial akses wajib diisi.' };
  }
  const allowed = CONFIG.ADMIN_CREDENTIALS.map((x) => String(x).trim().toUpperCase());
  if (allowed.includes(val.toUpperCase())) {
    return { success: true, role: 'admin', message: 'Login Admin berhasil.' };
  }
  return { success: false, message: 'Kredensial tidak valid!' };
}

function getAllData() {
  try {
    const agenda = readSheetObjects_(SHEETS.AGENDA);
    const pelajaran = readSheetObjects_(SHEETS.PELAJARAN);
    const ujian = readSheetObjects_(SHEETS.UJIAN);
    const tugas = readSheetObjects_(SHEETS.TUGAS);
    const notulensi = readSheetObjects_(SHEETS.NOTULENSI);

    return {
      success: true,
      agenda: agenda,
      pelajaran: pelajaran,
      ujian: ujian,
      tugas: tugas,
      notulensi: notulensi
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal membaca data Spreadsheet: ' + err.message,
      agenda: [],
      pelajaran: [],
      ujian: [],
      tugas: [],
      notulensi: []
    };
  }
}

function saveRecord(sheetName, payload, pin) {
  try {
    if (!verifyPin_(pin)) {
      return { success: false, message: 'PIN Keamanan salah!' };
    }

    const key = String(sheetName || '').toUpperCase();
    const meta = SHEETS[key];
    if (!meta) {
      return { success: false, message: 'Sheet tidak dikenal: ' + sheetName };
    }

    const sh = getOrCreateSheet_(meta);
    const data = Object.assign({}, payload || {});
    const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    if (!data.ID || String(data.ID).startsWith('TMP-')) {
      const prefixMap = {
        AGENDA: 'AGE',
        PELAJARAN: 'PEL',
        UJIAN: 'UJN',
        TUGAS: 'TGS',
        NOTULENSI: 'NOT'
      };
      data.ID = (prefixMap[key] || key.slice(0, 3)) + '-' + Date.now();
    }

    if (key === 'AGENDA') {
      data.JatuhTempo = String(data.JatuhTempo || '').replace('T', ' ').trim();
      if (!data.CreatedAt) data.CreatedAt = nowStr;
    } else if (key === 'TUGAS') {
      data.Deadline = String(data.Deadline || '').replace('T', ' ').trim();
      if (!data.Selesai) data.Selesai = 'Belum';
    } else if (key === 'PELAJARAN') {
      data.WaktuMulai = String(data.WaktuMulai || '').trim();
      data.WaktuSelesai = String(data.WaktuSelesai || '').trim();
    } else if (key === 'UJIAN') {
      data.StatusAktif = String(data.StatusAktif || 'NON AKTIF').trim().toUpperCase();
    }

    writeRowByHeaders_(sh, meta.headers, data);
    SpreadsheetApp.flush();

    return {
      success: true,
      id: data.ID,
      message: 'Data ' + meta.name + ' berhasil disimpan ke Spreadsheet.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan data: ' + err.message };
  }
}

function deleteRecord(sheetName, id, pin) {
  try {
    if (!verifyPin_(pin)) {
      return { success: false, message: 'PIN Keamanan salah!' };
    }

    const key = String(sheetName || '').toUpperCase();
    const meta = SHEETS[key];
    if (!meta) {
      return { success: false, message: 'Sheet tidak dikenal: ' + sheetName };
    }

    const sh = getOrCreateSheet_(meta);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Data tidak ditemukan.' };

    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]).trim() === String(id).trim()) {
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { success: true, message: 'Data ' + meta.name + ' berhasil dihapus.' };
      }
    }

    return { success: false, message: 'ID data tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus data: ' + err.message };
  }
}

function toggleUjianStatus(id, nextStatus, pin) {
  try {
    if (!verifyPin_(pin)) {
      return { success: false, message: 'PIN Keamanan salah!' };
    }

    const sh = getOrCreateSheet_(SHEETS.UJIAN);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Jadwal ujian tidak ditemukan.' };

    const range = sh.getRange(2, 1, lastRow - 1, SHEETS.UJIAN.headers.length);
    const values = range.getValues();
    const statusColIdx = SHEETS.UJIAN.headers.indexOf('StatusAktif');

    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(id).trim()) {
        values[i][statusColIdx] = String(nextStatus || 'NON AKTIF').toUpperCase();
        range.setValues(values);
        SpreadsheetApp.flush();
        return { success: true, message: 'Status ujian diubah menjadi ' + values[i][statusColIdx] + '.' };
      }
    }

    return { success: false, message: 'ID ujian tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: 'Gagal mengubah status ujian: ' + err.message };
  }
}

function saveNotulensi(agendaId, isiNotulensi) {
  try {
    if (!agendaId) return { success: false, message: 'AgendaID kosong.' };
    const sh = getOrCreateSheet_(SHEETS.AGENDA);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Agenda tidak ditemukan.' };

    const range = sh.getRange(2, 1, lastRow - 1, SHEETS.AGENDA.headers.length);
    const values = range.getValues();
    const notulensiColIdx = SHEETS.AGENDA.headers.indexOf('Notulensi');

    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(agendaId).trim()) {
        values[i][notulensiColIdx] = String(isiNotulensi || '');
        range.setValues(values);
        SpreadsheetApp.flush();
        return { success: true, message: 'Notulensi agenda tersimpan.' };
      }
    }

    return { success: false, message: 'Agenda tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan notulensi agenda: ' + err.message };
  }
}

function saveNotulensiArchive(record) {
  try {
    const sh = getOrCreateSheet_(SHEETS.NOTULENSI);
    const rec = Object.assign({}, record || {});
    const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy, HH:mm:ss');

    if (!rec.ID) {
      rec.ID = 'NOT-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd') + '-' + Math.floor(1000 + Math.random() * 9000);
    }
    if (!rec.Penandatangan) {
      rec.Penandatangan = 'Muhammad Sufyan Fahriza';
    }
    if (!rec.Tanggal) {
      rec.Tanggal = nowStr;
    }

    writeRowByHeaders_(sh, SHEETS.NOTULENSI.headers, rec);
    SpreadsheetApp.flush();

    return {
      success: true,
      id: rec.ID,
      token: rec.TokenTTE || '',
      message: 'Arsip Notulensi berhasil disimpan ke Spreadsheet.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan arsip notulensi: ' + err.message };
  }
}

function deleteNotulensiArchive(id, pin) {
  try {
    if (!verifyPin_(pin)) {
      return { success: false, message: 'PIN Keamanan salah!' };
    }

    const sh = getOrCreateSheet_(SHEETS.NOTULENSI);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Arsip notulensi kosong.' };

    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]).trim() === String(id).trim()) {
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { success: true, message: 'Arsip notulensi berhasil dihapus.' };
      }
    }

    return { success: false, message: 'ID notulensi tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus arsip notulensi: ' + err.message };
  }
}

function computeTteTokenForSlot_(slot) {
  const input = CONFIG.TTE_SECRET + '-' + slot;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  const hex = digest
    .map((b) => {
      const v = (b < 0 ? b + 256 : b).toString(16);
      return v.length === 1 ? '0' + v : v;
    })
    .join('')
    .toUpperCase();
  return hex.slice(0, 6);
}

function getAdminTteToken(credential) {
  const login = verifyLogin(credential);
  if (!login.success) {
    return { success: false, message: 'Akses ditolak. Khusus Admin.' };
  }
  const slot = Math.floor(Date.now() / 60000);
  const token = computeTteTokenForSlot_(slot);
  return {
    success: true,
    token: token,
    expiresInSeconds: 60 - new Date().getSeconds()
  };
}

function verifyTteToken(token) {
  const clean = String(token || '').trim().toUpperCase();
  if (!clean) {
    return { success: false, message: 'Token TTE tidak boleh kosong.' };
  }
  const currentSlot = Math.floor(Date.now() / 60000);
  const validTokens = [
    computeTteTokenForSlot_(currentSlot),
    computeTteTokenForSlot_(currentSlot - 1)
  ];

  if (validTokens.includes(clean)) {
    return {
      success: true,
      valid: true,
      ttdUrl: CONFIG.TTE_SPECIMEN_URL,
      signer: 'Muhammad Sufyan Fahriza',
      roleTitle: 'Eselon III (Pejabat Rohis)',
      message: 'Token TTE terverifikasi sah.'
    };
  }

  return {
    success: false,
    valid: false,
    message: 'Token TTE tidak valid atau sudah kedaluwarsa.'
  };
}

function dispatchAction_(actionName, args) {
  const params = Array.isArray(args) ? args : [];
  switch (String(actionName || '').trim()) {
    case 'verifyLogin':
      return verifyLogin(params[0]);
    case 'getAllData':
      return getAllData();
    case 'saveRecord':
      return saveRecord(params[0], params[1], params[2]);
    case 'deleteRecord':
      return deleteRecord(params[0], params[1], params[2]);
    case 'toggleUjianStatus':
      return toggleUjianStatus(params[0], params[1], params[2]);
    case 'saveNotulensi':
      return saveNotulensi(params[0], params[1]);
    case 'saveNotulensiArchive':
      return saveNotulensiArchive(params[0]);
    case 'deleteNotulensiArchive':
      return deleteNotulensiArchive(params[0], params[1]);
    case 'getAdminTteToken':
      return getAdminTteToken(params[0]);
    case 'verifyTteToken':
      return verifyTteToken(params[0]);
    default:
      return { success: false, message: 'Action tidak dikenal: ' + actionName };
  }
}

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  let result;

  if (p.action) {
    let parsedArgs = [];
    if (p.args) {
      try {
        parsedArgs = JSON.parse(p.args);
      } catch (err) {
        parsedArgs = [];
      }
    }
    result = dispatchAction_(p.action, parsedArgs);
  } else {
    result = getAllData();
  }

  const jsonStr = JSON.stringify(result);
  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + jsonStr + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let payload = {};
  if (e && e.postData && e.postData.contents) {
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (err) {
      payload = {};
    }
  }
  const p = (e && e.parameter) ? e.parameter : {};
  const actionName = payload.action || p.action || 'getAllData';
  const args = Array.isArray(payload.args) ? payload.args : [];
  const result = dispatchAction_(actionName, args);

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}