const SPREADSHEET_ID = '1dfbhCNZnUwYWuYssXSBMEB72Zer_ERDuK9iyVoNEqEc';
const MAIN_FOLDER_ID = '1e8BLN-VMMYHtTKPzRMmpzbWjMeGXNR-h';
const LOGO_FILE_ID   = '1uhCJGvX8gNfJ7yZpVr2bfCbElsmx_n33';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Starlite Time Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper untuk include file HTML/CSS/JS di Google Apps Script (tanpa folder).
 */
function include(filename) {
  const cleanName = String(filename || '').replace(/\.html$/, '');
  return HtmlService.createHtmlOutputFromFile(cleanName).getContent();
}

/**
 * Deteksi Otomatis Zona Waktu berdasarkan Provinsi Sales:
 * - Jawa Barat, Jawa Tengah, Jawa Timur, Sumatera, Kalbar, Kalteng -> WIB (Asia/Jakarta)
 * - Bali, NTB, NTT, Kalsel, Kaltim, Kaltara, Sulawesi -> WITA (Asia/Makassar)
 * - Maluku, Papua -> WIT (Asia/Jayapura)
 */
function getTimezoneByProvinsi_(provinsi) {
  const p = String(provinsi || '').toLowerCase();
  const witList = ['maluku', 'papua', 'ambon', 'jayapura', 'sorong', 'merauke', 'timika', 'ternate'];
  const witaList = [
    'bali', 'nusa tenggara', 'ntb', 'ntt', 'kalimantan selatan', 'kalsel', 
    'kalimantan timur', 'kaltim', 'kalimantan utara', 'kaltara', 'sulawesi', 
    'sulsel', 'sulut', 'sulteng', 'sultra', 'sulbar', 'gorontalo', 'makassar', 
    'manado', 'denpasar', 'balikpapan', 'samarinda', 'banjarmasin', 'lombok', 'kupang'
  ];

  for (let i = 0; i < witList.length; i++) {
    if (p.indexOf(witList[i]) !== -1) return { tz: 'Asia/Jayapura', label: 'WIT' };
  }
  for (let j = 0; j < witaList.length; j++) {
    if (p.indexOf(witaList[j]) !== -1) return { tz: 'Asia/Makassar', label: 'WITA' };
  }
  // Default seluruh Pulau Jawa (Jabar, Jateng, Jatim), Sumatera, Kalbar, Kalteng adalah WIB
  return { tz: 'Asia/Jakarta', label: 'WIB' };
}

function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let userSheet = ss.getSheetByName('USERS');
  if (!userSheet) {
    userSheet = ss.insertSheet('USERS');
    userSheet.appendRow(['ID_User', 'Username', 'Password', 'Role', 'Nama_Sales', 'Status', 'Provinsi', 'Stasiun']);
    userSheet.appendRow(['USR-001', 'admin', 'admin123', 'Admin', 'Administrator', 'Inhouse', 'Jawa Barat', 'HQ Bandung']);
    userSheet.appendRow(['USR-002', 'sales1', '123456', 'Sales', 'Budi Santoso', 'Inhouse', 'Jawa Barat', 'Stasiun Bandung']);
    userSheet.getRange('A1:H1').setFontWeight('bold').setBackground('#1c2434').setFontColor('#ffffff');
  }
  
  let absenSheet = ss.getSheetByName('ABSENSI');
  if (!absenSheet) {
    absenSheet = ss.insertSheet('ABSENSI');
    absenSheet.appendRow([
      'ID_Absen', 'Tanggal', 'Username', 'Nama_Sales', 'Status_Sales', 'Provinsi', 'Stasiun',
      'Jam_In', 'Status_In', 'Koordinat_In', 'Maps_In', 'Foto_In',
      'Jam_Out', 'Status_Out', 'Koordinat_Out', 'Maps_Out', 'Foto_Out',
      'Stasiun_In', 'Target_Visit', 'Visit_Kemana_In', 'Target_Closing_HC',
      'Stasiun_Out', 'Visit_Kemana_Out', 'Aktual_WO', 'Aktual_HC', 'Pending_WO', 'CID_List'
    ]);
    absenSheet.getRange('A1:AA1').setFontWeight('bold').setBackground('#1c2434').setFontColor('#ffffff');
  }

  let settingSheet = ss.getSheetByName('SETTINGS');
  if (!settingSheet) {
    settingSheet = ss.insertSheet('SETTINGS');
    settingSheet.appendRow(['Key', 'Value', 'Updated_At']);
    settingSheet.appendRow(['MAX_CLOCK_IN', '09:00', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')]);
    settingSheet.appendRow(['MIN_CLOCK_OUT', '18:00', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')]);
    settingSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#1c2434').setFontColor('#ffffff');
  }

  let stasiunSheet = ss.getSheetByName('STASIUN');
  if (!stasiunSheet) {
    stasiunSheet = ss.insertSheet('STASIUN');
    stasiunSheet.appendRow(['ID_Stasiun', 'Kota', 'Nama_Stasiun', 'Updated_At']);
    stasiunSheet.appendRow(['STN-001', 'Bandung', 'Stasiun Bandung', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')]);
    stasiunSheet.appendRow(['STN-002', 'Bandung', 'Stasiun Kiaracondong', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')]);
    stasiunSheet.appendRow(['STN-003', 'Jakarta', 'Stasiun Gambir', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')]);
    stasiunSheet.appendRow(['STN-004', 'Jakarta', 'Stasiun Pasar Senen', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')]);
    stasiunSheet.appendRow(['STN-005', 'Surabaya', 'Stasiun Gubeng', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')]);
    stasiunSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#1c2434').setFontColor('#ffffff');
  }

  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }

  upgradeAbsensiSheetHeaders_();
  setupGpsTracksSheet_();

  try {
    DriveApp.getFolderById(MAIN_FOLDER_ID).getName();
  } catch (e) {}

  return "Setup Database Berhasil!";
}

function upgradeAbsensiSheetHeaders_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('ABSENSI');
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  const requiredHeaders = [
    'ID_Absen', 'Tanggal', 'Username', 'Nama_Sales', 'Status_Sales', 'Provinsi', 'Stasiun',
    'Jam_In', 'Status_In', 'Koordinat_In', 'Maps_In', 'Foto_In',
    'Jam_Out', 'Status_Out', 'Koordinat_Out', 'Maps_Out', 'Foto_Out',
    'Stasiun_In', 'Target_Visit', 'Visit_Kemana_In', 'Target_Closing_HC',
    'Stasiun_Out', 'Visit_Kemana_Out', 'Aktual_WO', 'Aktual_HC', 'Pending_WO', 'CID_List'
  ];
  if (lastCol < requiredHeaders.length) {
    const missingHeaders = requiredHeaders.slice(lastCol);
    sheet.getRange(1, lastCol + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    sheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight('bold').setBackground('#1c2434').setFontColor('#ffffff');
  }
}

function setupGpsTracksSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let trackSheet = ss.getSheetByName('GPS_TRACKS');
  if (!trackSheet) {
    trackSheet = ss.insertSheet('GPS_TRACKS');
    trackSheet.appendRow([
      'ID_Absen', 'Tanggal', 'Username', 'Nama_Sales',
      'Status_Koneksi', 'Last_Sync', 'Total_Titik', 'Total_Jarak_KM', 'Waypoints_JSON'
    ]);
    trackSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#1c2434').setFontColor('#ffffff');
  }
}

function setupCheck_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss.getSheetByName('USERS') || !ss.getSheetByName('ABSENSI') || !ss.getSheetByName('SETTINGS') || !ss.getSheetByName('STASIUN')) {
    setupDatabase();
  } else {
    upgradeAbsensiSheetHeaders_();
    setupGpsTracksSheet_();
  }
}

function getWorkTimeSettings() {
  setupCheck_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('SETTINGS');
  const data = sheet.getDataRange().getValues();
  let config = { maxClockIn: '09:00', minClockOut: '18:00' };

  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const val = formatHHMM_(data[i][1]);
    if (key === 'MAX_CLOCK_IN' && val) config.maxClockIn = val;
    if (key === 'MIN_CLOCK_OUT' && val) config.minClockOut = val;
  }
  return config;
}

function saveWorkTimeSettings(maxClockIn, minClockOut, callerRole) {
  if (String(callerRole || '').toLowerCase() !== 'admin') {
    return { success: false, message: 'Akses Ditolak: Hanya Admin yang dapat mengubah aturan jam kerja & sistem.' };
  }
  setupCheck_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('SETTINGS');
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  sheet.clearContents();
  sheet.appendRow(['Key', 'Value', 'Updated_At']);
  sheet.appendRow(['MAX_CLOCK_IN', "'" + String(maxClockIn).trim(), nowStr]);
  sheet.appendRow(['MIN_CLOCK_OUT', "'" + String(minClockOut).trim(), nowStr]);

  return {
    success: true,
    message: `Aturan jam kerja diperbarui (Batas Clock-In: ${maxClockIn}, Minimal Clock-Out: ${minClockOut}).`,
    settings: { maxClockIn: String(maxClockIn).trim(), minClockOut: String(minClockOut).trim() }
  };
}

function getLogoUrl() {
  try {
    const file = DriveApp.getFileById(LOGO_FILE_ID);
    const base64 = Utilities.base64Encode(file.getBlob().getBytes());
    return 'data:' + file.getMimeType() + ';base64,' + base64;
  } catch (e) {
    return '';
  }
}

// ================= AUTHENTICATION =================
function loginUser(username, password) {
  setupCheck_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = ss.getSheetByName('USERS').getDataRange().getValues();
  const settings = getWorkTimeSettings();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(username).trim().toLowerCase() && 
        String(data[i][2]).trim() === String(password).trim()) {
      return {
        success: true,
        settings: settings,
        user: {
          id: String(data[i][0]),
          username: String(data[i][1]),
          role: String(data[i][3]),
          nama: String(data[i][4]),
          status: String(data[i][5]),
          provinsi: String(data[i][6]),
          stasiun: String(data[i][7])
        }
      };
    }
  }
  return { success: false, message: 'Username atau Password salah.' };
}

/**
 * Otomatis memverifikasi apakah akun pengguna masih ada dan aktif di Sheet USERS.
 * Dipanggil ketika web dibuka atau di-refresh agar jika admin menghapus akun,
 * sesi sales langsung terdeteksi tidak valid dan otomatis logout.
 */
function verifyUserSession(username, userId) {
  setupCheck_();
  if (!username) {
    return { valid: false, message: 'Username tidak valid.' };
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('USERS');
  if (!userSheet) {
    return { valid: false, message: 'Database USERS tidak ditemukan.' };
  }
  const data = userSheet.getDataRange().getValues();
  const settings = getWorkTimeSettings();

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0]).trim();
    const rowUsername = String(data[i][1]).trim().toLowerCase();
    if (rowUsername === String(username).trim().toLowerCase()) {
      if (userId && rowId !== String(userId).trim()) {
        continue;
      }
      return {
        valid: true,
        settings: settings,
        user: {
          id: rowId,
          username: String(data[i][1]),
          role: String(data[i][3]),
          nama: String(data[i][4]),
          status: String(data[i][5]),
          provinsi: String(data[i][6]),
          stasiun: String(data[i][7])
        }
      };
    }
  }

  return {
    valid: false,
    message: 'Akun Anda tidak ditemukan atau telah dinonaktifkan/dihapus oleh Administrator.'
  };
}

// ================= SALES ATTENDANCE LOGIC =================
function getSalesTodayStatus(username, nama) {
  setupCheck_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('USERS');
  const absenSheet = ss.getSheetByName('ABSENSI');
  
  // Cari user & provinsi user untuk menentukan zona waktu lokalnya (WIB/WITA/WIT)
  const userData = userSheet.getDataRange().getValues();
  let userProvinsi = 'Jawa Barat';
  let userFound = false;
  for (let u = 1; u < userData.length; u++) {
    let rowUname = String(userData[u][1]).trim().toLowerCase();
    let rowNama = String(userData[u][4]).trim().toLowerCase();
    if (rowUname === String(username).trim().toLowerCase()) {
      if (!nama || rowNama === String(nama).trim().toLowerCase()) {
        userProvinsi = String(userData[u][6]);
        userFound = true;
        break;
      }
    }
  }

  // Jika akun tidak ditemukan (misal sudah dihapus admin saat user masih buka web), tolak akses
  if (!userFound) {
    return {
      success: false,
      userNotFound: true,
      message: 'Akun Anda tidak ditemukan atau telah dihapus oleh Administrator.'
    };
  }
  const tzInfo = getTimezoneByProvinsi_(userProvinsi);

  const data = absenSheet.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), tzInfo.tz, 'yyyy-MM-dd');
  const nowTime  = Utilities.formatDate(new Date(), tzInfo.tz, 'HH:mm:ss');
  const settings = getWorkTimeSettings();

  let record = null;
  for (let i = data.length - 1; i >= 1; i--) {
    let rowDate = formatRowDate_(data[i][1]);
    let rowUname = String(data[i][2]).toLowerCase();
    let rowNama = String(data[i][3]).toLowerCase();
    if (rowDate === todayStr && rowUname === String(username).toLowerCase() && (!nama || rowNama === String(nama).toLowerCase())) {
      record = {
        idAbsen: String(data[i][0] || ''),
        tanggal: rowDate,
        jamIn: formatRowTime_(data[i][7]),
        statusIn: String(data[i][8] || ''),
        jamOut: formatRowTime_(data[i][12]),
        statusOut: String(data[i][13] || ''),
        stasiunIn: String(data[i][17] || ''),
        targetVisit: formatNumberOrBlank_(data[i][18]),
        visitKemanaIn: String(data[i][19] || ''),
        targetClosingHc: formatNumberOrBlank_(data[i][20]),
        stasiunOut: String(data[i][21] || ''),
        visitKemanaOut: String(data[i][22] || ''),
        aktualWo: formatNumberOrBlank_(data[i][23]),
        aktualHc: formatNumberOrBlank_(data[i][24]),
        pendingWo: formatNumberOrBlank_(data[i][25]),
        cidList: String(data[i][26] || '')
      };
      break;
    }
  }

  let history = [];
  for (let i = data.length - 1; i >= 1; i--) {
    let rowUname = String(data[i][2]).toLowerCase();
    let rowNama = String(data[i][3]).toLowerCase();
    if (rowUname === String(username).toLowerCase() && (!nama || rowNama === String(nama).toLowerCase())) {
      history.push({
        idAbsen: String(data[i][0] || ''),
        tanggal: formatRowDate_(data[i][1]),
        jamIn: formatRowTime_(data[i][7]),
        statusIn: String(data[i][8] || ''),
        koordinatIn: String(data[i][9] || ''),
        mapsIn: String(data[i][10] || ''),
        fotoIn: String(data[i][11] || ''),
        jamOut: formatRowTime_(data[i][12]),
        statusOut: String(data[i][13] || ''),
        koordinatOut: String(data[i][14] || ''),
        mapsOut: String(data[i][15] || ''),
        fotoOut: String(data[i][16] || ''),
        stasiunIn: String(data[i][17] || ''),
        targetVisit: formatNumberOrBlank_(data[i][18]),
        visitKemanaIn: String(data[i][19] || ''),
        targetClosingHc: formatNumberOrBlank_(data[i][20]),
        stasiunOut: String(data[i][21] || ''),
        visitKemanaOut: String(data[i][22] || ''),
        aktualWo: formatNumberOrBlank_(data[i][23]),
        aktualHc: formatNumberOrBlank_(data[i][24]),
        pendingWo: formatNumberOrBlank_(data[i][25]),
        cidList: String(data[i][26] || '')
      });
    }
  }

  return {
    serverDate: todayStr,
    serverTime: nowTime,
    timezoneLabel: tzInfo.label,
    settings: settings,
    todayRecord: record,
    history: history,
    stations: getStationsData_()
  };
}

function submitAttendance(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  
  try {
    setupCheck_();
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Verifikasi bahwa user masih aktif di sheet USERS
    const userSheet = ss.getSheetByName('USERS');
    const usersData = userSheet.getDataRange().getValues();
    let userValid = false;
    for (let u = 1; u < usersData.length; u++) {
      if (String(usersData[u][1]).trim().toLowerCase() === String(payload.username).trim().toLowerCase()) {
        userValid = true;
        break;
      }
    }
    if (!userValid) {
      return { 
        success: false, 
        userNotFound: true,
        message: 'Ditolak: Akun Anda tidak aktif atau telah dihapus oleh Administrator.' 
      };
    }

    const sheet = ss.getSheetByName('ABSENSI');
    const data = sheet.getDataRange().getValues();
    const settings = getWorkTimeSettings();
    
    // Gunakan Zona Waktu sesuai Provinsi Sales (WIB / WITA / WIT)
    const tzInfo = getTimezoneByProvinsi_(payload.provinsi);
    const now = new Date();
    const todayStr = Utilities.formatDate(now, tzInfo.tz, 'yyyy-MM-dd');
    const timeStr  = Utilities.formatDate(now, tzInfo.tz, 'HH:mm:ss');
    const hourNum  = parseInt(Utilities.formatDate(now, tzInfo.tz, 'HH'), 10);
    const minNum   = parseInt(Utilities.formatDate(now, tzInfo.tz, 'mm'), 10);
    const secNum   = parseInt(Utilities.formatDate(now, tzInfo.tz, 'ss'), 10);
    const nowTotalSeconds = (hourNum * 3600) + (minNum * 60) + secNum;

    const [maxInH, maxInM] = settings.maxClockIn.split(':').map(Number);
    const maxInSeconds = (maxInH * 3600) + ((maxInM || 0) * 60);

    const [minOutH, minOutM] = settings.minClockOut.split(':').map(Number);
    const minOutSeconds = (minOutH * 3600) + ((minOutM || 0) * 60);

    let existingRowIndex = -1;
    for (let i = data.length - 1; i >= 1; i--) {
      let rowDate = formatRowDate_(data[i][1]);
      let rowUname = String(data[i][2]).toLowerCase();
      let rowNama = String(data[i][3]).toLowerCase();
      if (rowDate === todayStr && rowUname === String(payload.username).toLowerCase() && (!payload.nama || rowNama === String(payload.nama).toLowerCase())) {
        existingRowIndex = i + 1;
        break;
      }
    }

    if (payload.type === 'IN' && existingRowIndex !== -1) {
      return { success: false, message: `Anda sudah melakukan Clock-In pada hari ini (${todayStr}).` };
    }

    if (payload.type === 'OUT') {
      if (existingRowIndex === -1) {
        return { success: false, message: `Ditolak! Anda belum melakukan Clock-In pada hari ini (${todayStr}).` };
      }
      const existingOut = sheet.getRange(existingRowIndex, 13).getValue();
      if (existingOut !== '') {
        return { success: false, message: 'Anda sudah melakukan Clock-Out hari ini.' };
      }
      if (nowTotalSeconds < minOutSeconds) {
        return { 
          success: false, 
          message: `Clock-Out belum bisa dilakukan. Minimal pukul ${settings.minClockOut} ${tzInfo.label} (Waktu lokal provinsi ${payload.provinsi} saat ini: ${timeStr} ${tzInfo.label}).` 
        };
      }
    }

    const photoUrl = savePhotoToDrive_(payload.imageBase64, payload.nama, todayStr, payload.type, timeStr);
    const mapsUrl  = `https://www.google.com/maps?q=${payload.lat},${payload.lng}`;
    const coordStr = `${payload.lat}, ${payload.lng}`;

    if (payload.type === 'IN') {
      const isLate = nowTotalSeconds > maxInSeconds;
      const statusIn = isLate ? 'Late' : 'On Time';
      const idAbsen = 'ABS-' + Utilities.formatDate(now, tzInfo.tz, 'yyyyMMddHHmmss');

      const stasiunInStr = Array.isArray(payload.stasiunIn) 
        ? payload.stasiunIn.join(', ') 
        : String(payload.stasiunIn || payload.stasiun || '');
      const targetVisit = payload.targetVisit !== undefined && payload.targetVisit !== '' ? Number(payload.targetVisit) : '';
      const visitKemanaIn = String(payload.visitKemanaIn || '');
      const targetClosingHc = payload.targetClosingHc !== undefined && payload.targetClosingHc !== '' ? Number(payload.targetClosingHc) : '';

      sheet.appendRow([
        idAbsen,
        "'" + todayStr,
        payload.username,
        payload.nama,
        payload.status,
        payload.provinsi,
        payload.stasiun,
        "'" + timeStr,
        statusIn,
        coordStr,
        mapsUrl,
        photoUrl,
        '', '', '', '', '', // Jam_Out, Status_Out, Koordinat_Out, Maps_Out, Foto_Out
        stasiunInStr,
        targetVisit,
        visitKemanaIn,
        targetClosingHc,
        '', '', '', '', '', '' // Stasiun_Out, Visit_Kemana_Out, Aktual_WO, Aktual_HC, Pending_WO, CID_List
      ]);

      return { 
        success: true, 
        idAbsen: idAbsen,
        message: `Clock-In berhasil pukul ${timeStr} ${tzInfo.label} (${statusIn}).` 
      };

    } else if (payload.type === 'OUT') {
      const statusOut = 'Completed';

      const stasiunOutStr = Array.isArray(payload.stasiunOut) 
        ? payload.stasiunOut.join(', ') 
        : String(payload.stasiunOut || payload.stasiun || '');
      const visitKemanaOut = String(payload.visitKemanaOut || '');
      const aktualWo = payload.aktualWo !== undefined && payload.aktualWo !== '' ? Number(payload.aktualWo) : '';
      const aktualHc = payload.aktualHc !== undefined && payload.aktualHc !== '' ? Number(payload.aktualHc) : '';
      const pendingWo = payload.pendingWo !== undefined && payload.pendingWo !== '' ? Number(payload.pendingWo) : '';
      const cidListStr = Array.isArray(payload.cidList) 
        ? payload.cidList.filter(c => String(c).trim() !== '').join(', ') 
        : String(payload.cidList || '');

      sheet.getRange(existingRowIndex, 13, 1, 5).setValues([[
        "'" + timeStr,
        statusOut,
        coordStr,
        mapsUrl,
        photoUrl
      ]]);

      sheet.getRange(existingRowIndex, 22, 1, 6).setValues([[
        stasiunOutStr,
        visitKemanaOut,
        aktualWo,
        aktualHc,
        pendingWo,
        cidListStr
      ]]);

      // Update status tracking GPS menjadi COMPLETED jika ada
      try {
        const targetIdAbsen = String(data[existingRowIndex - 1][0]);
        const trackSheet = ss.getSheetByName('GPS_TRACKS');
        if (trackSheet) {
          const tData = trackSheet.getDataRange().getValues();
          for (let t = 1; t < tData.length; t++) {
            if (String(tData[t][0]) === targetIdAbsen) {
              trackSheet.getRange(t + 1, 5).setValue('COMPLETED');
              break;
            }
          }
        }
      } catch (errTrack) {}

      return { 
        success: true, 
        message: `Clock-Out berhasil pukul ${timeStr} ${tzInfo.label}. Kehadiran hari ini lengkap!` 
      };
    }

  } catch (err) {
    return { success: false, message: 'Kesalahan sistem: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function savePhotoToDrive_(base64Data, salesName, dateStr, type, timeStr) {
  let mainFolder;
  try {
    mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    mainFolder.getFoldersByName(salesName);
  } catch (err) {
    const fallbackFolders = DriveApp.getFoldersByName('STARLITE_FOTO_ABSENSI');
    mainFolder = fallbackFolders.hasNext() ? fallbackFolders.next() : DriveApp.createFolder('STARLITE_FOTO_ABSENSI');
  }
  
  let userFolder;
  try {
    const userFolders = mainFolder.getFoldersByName(salesName);
    userFolder = userFolders.hasNext() ? userFolders.next() : mainFolder.createFolder(salesName);
  } catch (err) {
    const fallbackFolders = DriveApp.getFoldersByName('STARLITE_FOTO_ABSENSI');
    mainFolder = fallbackFolders.hasNext() ? fallbackFolders.next() : DriveApp.createFolder('STARLITE_FOTO_ABSENSI');
    const userFolders = mainFolder.getFoldersByName(salesName);
    userFolder = userFolders.hasNext() ? userFolders.next() : mainFolder.createFolder(salesName);
  }

  const dateFolders = userFolder.getFoldersByName(dateStr);
  const dateFolder = dateFolders.hasNext() ? dateFolders.next() : userFolder.createFolder(dateStr);

  const splitBase64 = base64Data.split(',');
  const contentType = splitBase64[0].match(/:(.*?);/)[1];
  const byteCharacters = Utilities.base64Decode(splitBase64[1]);
  const cleanTime = timeStr.replace(/:/g, '-');
  const fileName = `CLOCK_${type}_${salesName}_${cleanTime}.jpg`;
  const blob = Utilities.newBlob(byteCharacters, contentType, fileName);
  
  const file = dateFolder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  
  return file.getUrl();
}

// Helper untuk mengekstrak File ID dari URL Google Drive dan memindahkannya ke Trash (Sampah)
function trashDriveFileByUrl_(fileUrl) {
  if (!fileUrl) return false;
  try {
    const match = String(fileUrl).match(/[-\w]{25,}/);
    if (match && match[0]) {
      DriveApp.getFileById(match[0]).setTrashed(true);
      return true;
    }
  } catch (e) {
    // Abaikan jika file sudah terhapus sebelumnya
  }
  return false;
}

// Helper untuk menghapus seluruh Folder milik Sales di Google Drive
function trashSalesFolderInDrive_(salesName) {
  if (!salesName) return 0;
  let count = 0;

  // 1. Cek di dalam MAIN_FOLDER_ID
  try {
    const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    const subFolders = mainFolder.getFoldersByName(salesName);
    while (subFolders.hasNext()) {
      subFolders.next().setTrashed(true);
      count++;
    }
  } catch (e) {}

  // 2. Cek juga di dalam folder fallback STARLITE_FOTO_ABSENSI jika ada
  try {
    const fbIter = DriveApp.getFoldersByName('STARLITE_FOTO_ABSENSI');
    while (fbIter.hasNext()) {
      const fbFolder = fbIter.next();
      const userFolders = fbFolder.getFoldersByName(salesName);
      while (userFolders.hasNext()) {
        userFolders.next().setTrashed(true);
        count++;
      }
    }
  } catch (e) {}

  return count;
}

// ================= ADMIN CRUD & DASHBOARD LOGIC =================
function getAdminDashboardData(filterDate) {
  setupCheck_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('USERS');
  const absenSheet = ss.getSheetByName('ABSENSI');
  const settings = getWorkTimeSettings();
  
  const todayStr = filterDate || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  
  const usersData = userSheet.getDataRange().getValues();
  const absenData = absenSheet.getDataRange().getValues();

  const trackSheet = ss.getSheetByName('GPS_TRACKS');
  const trackMap = {};
  if (trackSheet) {
    const tData = trackSheet.getDataRange().getValues();
    for (let t = 1; t < tData.length; t++) {
      const tId = String(tData[t][0]).trim();
      const tUname = String(tData[t][2]).trim().toLowerCase();
      const tTgl = formatRowDate_(tData[t][1]);
      const statusKon = String(tData[t][4] || '');
      const dist = Number(tData[t][7] || 0);
      const ptsCount = Number(tData[t][6] || 0);
      const tObj = { statusKoneksi: statusKon, totalJarakKm: dist, totalTitik: ptsCount };
      if (tId) trackMap[tId] = tObj;
      trackMap[tUname + '_' + tTgl] = tObj;
    }
  }

  let usersList = [];
  let salesOnlyList = [];
  let validUsernamesSet = {};

  for (let i = 1; i < usersData.length; i++) {
    if (!usersData[i][0]) continue;
    let uname = String(usersData[i][1]).trim();
    let role = String(usersData[i][3]).trim();
    let u = {
      id: String(usersData[i][0]),
      username: uname,
      password: String(usersData[i][2]),
      role: role,
      nama: String(usersData[i][4]).trim(),
      status: String(usersData[i][5]).trim(),
      provinsi: String(usersData[i][6]).trim(),
      stasiun: String(usersData[i][7]).trim()
    };
    usersList.push(u);
    validUsernamesSet[uname.toLowerCase()] = true;
    if (role.toLowerCase() === 'sales') {
      salesOnlyList.push(u);
    }
  }

  let allLogs = [];
  let todayLogs = [];

  for (let i = absenData.length - 1; i >= 1; i--) {
    if (!absenData[i][0]) continue;
    let unameKey = String(absenData[i][2]).toLowerCase();
    
    if (!validUsernamesSet[unameKey]) continue;

    let rowDate = formatRowDate_(absenData[i][1]);
    let jamInStr = formatRowTime_(absenData[i][7]);
    let jamOutStr = formatRowTime_(absenData[i][12]);

    let logItem = {
      id: String(absenData[i][0]),
      tanggal: rowDate,
      username: String(absenData[i][2]),
      nama: String(absenData[i][3]),
      statusSales: String(absenData[i][4]),
      provinsi: String(absenData[i][5]),
      stasiun: String(absenData[i][6]),
      jamIn: jamInStr,
      statusIn: String(absenData[i][8] || ''),
      koordinatIn: String(absenData[i][9] || ''),
      mapsIn: String(absenData[i][10] || ''),
      fotoIn: String(absenData[i][11] || ''),
      jamOut: jamOutStr,
      statusOut: jamOutStr ? 'Completed' : 'Belum Clock-Out',
      koordinatOut: String(absenData[i][14] || ''),
      mapsOut: String(absenData[i][15] || ''),
      fotoOut: String(absenData[i][16] || ''),
      stasiunIn: String(absenData[i][17] || ''),
      targetVisit: formatNumberOrBlank_(absenData[i][18]),
      visitKemanaIn: String(absenData[i][19] || ''),
      targetClosingHc: formatNumberOrBlank_(absenData[i][20]),
      stasiunOut: String(absenData[i][21] || ''),
      visitKemanaOut: String(absenData[i][22] || ''),
      aktualWo: formatNumberOrBlank_(absenData[i][23]),
      aktualHc: formatNumberOrBlank_(absenData[i][24]),
      pendingWo: formatNumberOrBlank_(absenData[i][25]),
      cidList: String(absenData[i][26] || '')
    };

    const tInfo = trackMap[logItem.id] || trackMap[unameKey + '_' + rowDate] || null;
    logItem.trackingStatus = tInfo ? tInfo.statusKoneksi : (jamOutStr ? 'COMPLETED' : (jamInStr ? 'CONNECTED' : 'NO_DATA'));
    logItem.totalJarakKm = tInfo ? tInfo.totalJarakKm : 0;
    logItem.totalTitik = tInfo ? tInfo.totalTitik : 0;

    allLogs.push(logItem);
    if (rowDate === todayStr) {
      todayLogs.push(logItem);
    }
  }

  // Evaluasi kehadiran per individu sales hari ini menggunakan composite key (username + nama)
  const attendedSalesSet = new Set(todayLogs.map(l => (String(l.username || '').trim().toLowerCase() + '|' + String(l.nama || '').trim().toLowerCase())));
  const belumAbsenList = salesOnlyList.filter(s => !attendedSalesSet.has(s.username.toLowerCase() + '|' + s.nama.toLowerCase()));
  const stationsList = getStationsData_();

  return {
    selectedDate: todayStr,
    settings: settings,
    users: usersList,
    salesOnly: salesOnlyList,
    todayLogs: todayLogs,
    allLogs: allLogs,
    belumAbsen: belumAbsenList,
    stations: stationsList
  };
}

// ================= GPS TRACKING (LIVE ROUTE RECORDING) =================
function calculateTotalDistanceKm_(pts) {
  if (!pts || pts.length < 2) return 0;
  let totalKm = 0;
  let lastLat = null, lastLng = null;
  for (let i = 0; i < pts.length; i++) {
    const p = pts[i];
    if (p && typeof p.lat === 'number' && typeof p.lng === 'number') {
      if (lastLat !== null && lastLng !== null) {
        const dLat = (p.lat - lastLat) * Math.PI / 180;
        const dLon = (p.lng - lastLng) * Math.PI / 180;
        const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
                  Math.cos(lastLat * Math.PI / 180) * Math.cos(p.lat * Math.PI / 180) *
                  Math.sin(dLon / 2) * Math.sin(dLon / 2);
        const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
        const d = 6371 * c; // KM
        if (d < 50) { // filter GPS anomaly teleport (> 50 km jumping)
          totalKm += d;
        }
      }
      lastLat = p.lat;
      lastLng = p.lng;
    }
  }
  return Math.round(totalKm * 100) / 100;
}

function syncGpsTracking(payload) {
  if (!payload || !payload.idAbsen) {
    return { success: false, message: 'ID Absen tidak valid' };
  }
  setupCheck_();
  setupGpsTracksSheet_();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const trackSheet = ss.getSheetByName('GPS_TRACKS');
  const data = trackSheet.getDataRange().getValues();

  const idAbsen = String(payload.idAbsen).trim();
  const tanggal = String(payload.tanggal || '').trim();
  const username = String(payload.username || '').trim();
  const nama = String(payload.nama || '').trim();
  const statusKoneksi = String(payload.statusKoneksi || 'CONNECTED');
  const newPoints = Array.isArray(payload.points) ? payload.points : [];
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  let existingRowIndex = -1;
  let existingWaypoints = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === idAbsen) {
      existingRowIndex = i + 1;
      try {
        existingWaypoints = JSON.parse(String(data[i][8] || '[]'));
      } catch (e) {
        existingWaypoints = [];
      }
      break;
    }
  }

  let mergedWaypoints = existingWaypoints.slice();
  if (newPoints.length > 0) {
    for (let p of newPoints) {
      const last = mergedWaypoints[mergedWaypoints.length - 1];
      if (!last || last.time !== p.time || p.type === 'DISCONNECTED' || p.type === 'RECONNECTED' || p.type === 'CLOCK_OUT') {
        mergedWaypoints.push(p);
      }
    }
  }

  const totalDistKm = calculateTotalDistanceKm_(mergedWaypoints);
  const pointsJson = JSON.stringify(mergedWaypoints);

  if (existingRowIndex > 0) {
    trackSheet.getRange(existingRowIndex, 5, 1, 5).setValues([[
      statusKoneksi,
      nowStr,
      mergedWaypoints.length,
      totalDistKm,
      pointsJson
    ]]);
  } else {
    trackSheet.appendRow([
      idAbsen,
      tanggal,
      username,
      nama,
      statusKoneksi,
      nowStr,
      mergedWaypoints.length,
      totalDistKm,
      pointsJson
    ]);
  }

  return {
    success: true,
    totalTitik: mergedWaypoints.length,
    totalJarakKm: totalDistKm,
    statusKoneksi: statusKoneksi
  };
}

function getGpsTrackingData(idAbsen, username, tanggal) {
  setupCheck_();
  setupGpsTracksSheet_();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const trackSheet = ss.getSheetByName('GPS_TRACKS');
  const absenSheet = ss.getSheetByName('ABSENSI');

  let trackRecord = null;
  const trackData = trackSheet.getDataRange().getValues();
  for (let i = 1; i < trackData.length; i++) {
    const rowId = String(trackData[i][0]).trim();
    const rowUname = String(trackData[i][2]).trim().toLowerCase();
    const rowTgl = formatRowDate_(trackData[i][1]);

    if ((idAbsen && rowId === String(idAbsen).trim()) || 
        (username && tanggal && rowUname === String(username).toLowerCase() && rowTgl === String(tanggal))) {
      let pts = [];
      try {
        pts = JSON.parse(String(trackData[i][8] || '[]'));
      } catch (e) {
        pts = [];
      }
      trackRecord = {
        idAbsen: rowId,
        tanggal: rowTgl,
        username: String(trackData[i][2]),
        nama: String(trackData[i][3]),
        statusKoneksi: String(trackData[i][4] || 'CONNECTED'),
        lastSync: String(trackData[i][5] || ''),
        totalTitik: Number(trackData[i][6] || pts.length),
        totalJarakKm: Number(trackData[i][7] || 0),
        waypoints: pts
      };
      break;
    }
  }

  let absenDetail = null;
  const absenData = absenSheet.getDataRange().getValues();
  for (let j = 1; j < absenData.length; j++) {
    const aId = String(absenData[j][0]).trim();
    const aUname = String(absenData[j][2]).trim().toLowerCase();
    const aTgl = formatRowDate_(absenData[j][1]);

    if ((idAbsen && aId === String(idAbsen).trim()) ||
        (username && tanggal && aUname === String(username).toLowerCase() && aTgl === String(tanggal))) {
      absenDetail = {
        idAbsen: aId,
        tanggal: aTgl,
        username: String(absenData[j][2]),
        nama: String(absenData[j][3]),
        statusSales: String(absenData[j][4]),
        provinsi: String(absenData[j][5]),
        stasiun: String(absenData[j][6]),
        jamIn: formatRowTime_(absenData[j][7]),
        statusIn: String(absenData[j][8] || ''),
        koordinatIn: String(absenData[j][9] || ''),
        mapsIn: String(absenData[j][10] || ''),
        fotoIn: String(absenData[j][11] || ''),
        jamOut: formatRowTime_(absenData[j][12]),
        statusOut: String(absenData[j][13] || ''),
        koordinatOut: String(absenData[j][14] || ''),
        mapsOut: String(absenData[j][15] || ''),
        fotoOut: String(absenData[j][16] || ''),
        stasiunIn: String(absenData[j][17] || ''),
        targetVisit: formatNumberOrBlank_(absenData[j][18]),
        visitKemanaIn: String(absenData[j][19] || ''),
        targetClosingHc: formatNumberOrBlank_(absenData[j][20]),
        stasiunOut: String(absenData[j][21] || ''),
        visitKemanaOut: String(absenData[j][22] || ''),
        aktualWo: formatNumberOrBlank_(absenData[j][23]),
        aktualHc: formatNumberOrBlank_(absenData[j][24]),
        pendingWo: formatNumberOrBlank_(absenData[j][25]),
        cidList: String(absenData[j][26] || '')
      };
      break;
    }
  }

  // Graceful fallback synthesis jika data belum punya log titik rute terinci
  if (!trackRecord && absenDetail) {
    let syntheticPts = [];
    if (absenDetail.koordinatIn && absenDetail.koordinatIn.includes(',')) {
      const parts = absenDetail.koordinatIn.split(',').map(s => parseFloat(s.trim()));
      if (!isNaN(parts[0]) && !isNaN(parts[1])) {
        syntheticPts.push({
          type: 'CLOCK_IN',
          lat: parts[0],
          lng: parts[1],
          time: absenDetail.jamIn || '08:00',
          acc: 15,
          note: 'Titik Clock-In'
        });
      }
    }
    if (absenDetail.koordinatOut && absenDetail.koordinatOut.includes(',')) {
      const parts = absenDetail.koordinatOut.split(',').map(s => parseFloat(s.trim()));
      if (!isNaN(parts[0]) && !isNaN(parts[1])) {
        syntheticPts.push({
          type: 'CLOCK_OUT',
          lat: parts[0],
          lng: parts[1],
          time: absenDetail.jamOut || '18:00',
          acc: 15,
          note: 'Titik Clock-Out'
        });
      }
    }
    trackRecord = {
      idAbsen: absenDetail.idAbsen,
      tanggal: absenDetail.tanggal,
      username: absenDetail.username,
      nama: absenDetail.nama,
      statusKoneksi: absenDetail.jamOut ? 'COMPLETED' : 'CONNECTED',
      lastSync: absenDetail.jamOut ? absenDetail.jamOut : absenDetail.jamIn,
      totalTitik: syntheticPts.length,
      totalJarakKm: calculateTotalDistanceKm_(syntheticPts),
      waypoints: syntheticPts
    };
  }

  return {
    success: true,
    tracking: trackRecord,
    absenDetail: absenDetail
  };
}

function getAllSalesTrackingData(dateStr) {
  setupCheck_();
  setupGpsTracksSheet_();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const trackSheet = ss.getSheetByName('GPS_TRACKS');
  const absenSheet = ss.getSheetByName('ABSENSI');

  const todayStr = dateStr || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');

  const trackData = trackSheet ? trackSheet.getDataRange().getValues() : [];
  const absenData = absenSheet.getDataRange().getValues();

  // Absen map per user untuk tanggal terpilih
  const absenMap = {};
  for (let j = 1; j < absenData.length; j++) {
    const aId = String(absenData[j][0]).trim();
    const aTgl = formatRowDate_(absenData[j][1]);
    const aUname = String(absenData[j][2]).trim().toLowerCase();
    const aNama = String(absenData[j][3]).trim().toLowerCase();

    if (aTgl === todayStr) {
      const aObj = {
        idAbsen: aId,
        tanggal: aTgl,
        username: String(absenData[j][2]),
        nama: String(absenData[j][3]),
        statusSales: String(absenData[j][4]),
        provinsi: String(absenData[j][5]),
        stasiun: String(absenData[j][6]),
        jamIn: formatRowTime_(absenData[j][7]),
        statusIn: String(absenData[j][8] || ''),
        koordinatIn: String(absenData[j][9] || ''),
        fotoIn: String(absenData[j][11] || ''),
        jamOut: formatRowTime_(absenData[j][12]),
        statusOut: String(absenData[j][13] || ''),
        koordinatOut: String(absenData[j][14] || ''),
        fotoOut: String(absenData[j][16] || ''),
        stasiunIn: String(absenData[j][17] || ''),
        targetVisit: formatNumberOrBlank_(absenData[j][18]),
        visitKemanaIn: String(absenData[j][19] || ''),
        targetClosingHc: formatNumberOrBlank_(absenData[j][20]),
        stasiunOut: String(absenData[j][21] || ''),
        visitKemanaOut: String(absenData[j][22] || ''),
        aktualWo: formatNumberOrBlank_(absenData[j][23]),
        aktualHc: formatNumberOrBlank_(absenData[j][24]),
        pendingWo: formatNumberOrBlank_(absenData[j][25])
      };
      absenMap[aId] = aObj;
      absenMap[aUname + '|' + aNama] = aObj;
      if (!absenMap[aUname]) absenMap[aUname] = aObj;
    }
  }

  const salesTracks = [];
  const trackedUserKeys = {};

  for (let i = 1; i < trackData.length; i++) {
    const rowTgl = formatRowDate_(trackData[i][1]);
    if (rowTgl === todayStr) {
      const rId = String(trackData[i][0]).trim();
      const rUname = String(trackData[i][2]).trim().toLowerCase();
      const rNama = String(trackData[i][3]).trim().toLowerCase();
      const userKey = rUname + '|' + rNama;
      let pts = [];
      try {
        pts = JSON.parse(String(trackData[i][8] || '[]'));
      } catch (e) {
        pts = [];
      }

      const absenInfo = absenMap[rId] || absenMap[userKey] || absenMap[rUname] || {};
      trackedUserKeys[userKey] = true;
      trackedUserKeys[rId] = true;

      // Graceful fallback jika pts kosong di GPS_TRACKS namun absen memiliki koordinat
      if (pts.length === 0) {
        if (absenInfo.koordinatIn && absenInfo.koordinatIn.includes(',')) {
          const parts = absenInfo.koordinatIn.split(',').map(s => parseFloat(s.trim()));
          if (!isNaN(parts[0]) && !isNaN(parts[1])) {
            pts.push({ type: 'CLOCK_IN', lat: parts[0], lng: parts[1], time: absenInfo.jamIn || '08:00', acc: 15, note: 'Titik Clock-In' });
          }
        }
        if (absenInfo.koordinatOut && absenInfo.koordinatOut.includes(',')) {
          const parts = absenInfo.koordinatOut.split(',').map(s => parseFloat(s.trim()));
          if (!isNaN(parts[0]) && !isNaN(parts[1])) {
            pts.push({ type: 'CLOCK_OUT', lat: parts[0], lng: parts[1], time: absenInfo.jamOut || '18:00', acc: 15, note: 'Titik Clock-Out' });
          }
        }
      }

      salesTracks.push({
        idAbsen: rId,
        tanggal: rowTgl,
        username: String(trackData[i][2]),
        nama: String(trackData[i][3]) || absenInfo.nama || rUname,
        statusKoneksi: String(trackData[i][4] || (absenInfo.jamOut ? 'COMPLETED' : 'CONNECTED')),
        lastSync: String(trackData[i][5] || ''),
        totalTitik: Number(trackData[i][6] || pts.length),
        totalJarakKm: Number(trackData[i][7] || calculateTotalDistanceKm_(pts)),
        waypoints: pts,
        absen: absenInfo
      });
    }
  }

  // Graceful synthesis untuk data absen yang belum ada baris di GPS_TRACKS
  Object.keys(absenMap).forEach(key => {
    if (key.includes('|') || !key.startsWith('ABS_')) return;
    const a = absenMap[key];
    const uKey = a.username.toLowerCase() + '|' + a.nama.toLowerCase();
    if (!trackedUserKeys[uKey] && !trackedUserKeys[a.idAbsen]) {
      trackedUserKeys[uKey] = true;
      trackedUserKeys[a.idAbsen] = true;
      let synPts = [];
      if (a.koordinatIn && a.koordinatIn.includes(',')) {
        const parts = a.koordinatIn.split(',').map(s => parseFloat(s.trim()));
        if (!isNaN(parts[0]) && !isNaN(parts[1])) {
          synPts.push({ type: 'CLOCK_IN', lat: parts[0], lng: parts[1], time: a.jamIn || '08:00', acc: 15, note: 'Titik Clock-In' });
        }
      }
      if (a.koordinatOut && a.koordinatOut.includes(',')) {
        const parts = a.koordinatOut.split(',').map(s => parseFloat(s.trim()));
        if (!isNaN(parts[0]) && !isNaN(parts[1])) {
          synPts.push({ type: 'CLOCK_OUT', lat: parts[0], lng: parts[1], time: a.jamOut || '18:00', acc: 15, note: 'Titik Clock-Out' });
        }
      }
      salesTracks.push({
        idAbsen: a.idAbsen,
        tanggal: a.tanggal,
        username: a.username,
        nama: a.nama,
        statusKoneksi: a.jamOut ? 'COMPLETED' : 'CONNECTED',
        lastSync: a.jamOut || a.jamIn,
        totalTitik: synPts.length,
        totalJarakKm: calculateTotalDistanceKm_(synPts),
        waypoints: synPts,
        absen: a
      });
    }
  });

  return {
    success: true,
    date: todayStr,
    tracks: salesTracks
  };
}

function getStationsData() {
  setupCheck_();
  return getStationsData_();
}

function getStationsData_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const stasiunSheet = ss.getSheetByName('STASIUN');
  if (!stasiunSheet) return [];
  const stData = stasiunSheet.getDataRange().getValues();
  let stationsList = [];
  for (let s = 1; s < stData.length; s++) {
    if (!stData[s][0]) continue;
    stationsList.push({
      id: String(stData[s][0]),
      kota: String(stData[s][1]),
      namaStasiun: String(stData[s][2]),
      updatedAt: stData[s][3] ? formatRowDate_(stData[s][3]) : ''
    });
  }
  return stationsList;
}

function saveStationAdmin(stationObj, callerRole) {
  if (String(callerRole || '').toLowerCase() !== 'admin') {
    return { success: false, message: 'Akses Ditolak: Hanya Admin yang dapat mengelola database stasiun.' };
  }
  setupCheck_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('STASIUN');
  if (!sheet) {
    return { success: false, message: 'Sheet STASIUN tidak ditemukan.' };
  }

  const kota = String(stationObj.kota || '').trim();
  const namaStasiun = String(stationObj.namaStasiun || '').trim();
  if (!kota || !namaStasiun) {
    return { success: false, message: 'Kota dan Nama Stasiun wajib diisi.' };
  }

  const data = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  if (stationObj.id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(stationObj.id)) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[kota, namaStasiun, nowStr]]);
        return { 
          success: true, 
          message: `Data stasiun "${namaStasiun}" berhasil diperbarui.`,
          stations: getStationsData_()
        };
      }
    }
    return { success: false, message: 'Data stasiun tidak ditemukan.' };
  } else {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).toLowerCase().trim() === namaStasiun.toLowerCase()) {
        return { success: false, message: `Stasiun "${namaStasiun}" sudah ada di database.` };
      }
    }
    const newId = 'STN-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMddHHmmss');
    sheet.appendRow([newId, kota, namaStasiun, nowStr]);
    return { 
      success: true, 
      message: `Stasiun "${namaStasiun}" (${kota}) berhasil ditambahkan.`,
      stations: getStationsData_()
    };
  }
}

function deleteStationAdmin(stationId, callerRole) {
  if (String(callerRole || '').toLowerCase() !== 'admin') {
    return { success: false, message: 'Akses Ditolak: Hanya Admin yang dapat menghapus stasiun.' };
  }
  setupCheck_();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('STASIUN');
  if (!sheet) {
    return { success: false, message: 'Sheet STASIUN tidak ditemukan.' };
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(stationId)) {
      const stnName = String(data[i][2]);
      sheet.deleteRow(i + 1);
      return { 
        success: true, 
        message: `Stasiun "${stnName}" berhasil dihapus.`,
        stations: getStationsData_()
      };
    }
  }
  return { success: false, message: 'Data stasiun tidak ditemukan.' };
}

function saveUserAdmin(userObj, callerRole) {
  const role = String(callerRole || '').toLowerCase();
  if (role !== 'admin') {
    return { success: false, message: 'Akses Ditolak: Hanya Admin yang diizinkan menambah atau mengubah akun.' };
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('USERS');
  const absenSheet = ss.getSheetByName('ABSENSI');
  const data = sheet.getDataRange().getValues();

  if (userObj.id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userObj.id)) {
        const oldUsername = String(data[i][1]);
        const oldNama = String(data[i][4]);
        sheet.getRange(i + 1, 2, 1, 7).setValues([[
          userObj.username,
          userObj.password,
          userObj.role,
          userObj.nama,
          userObj.status,
          userObj.provinsi,
          userObj.stasiun
        ]]);

        // Jika Nama Sales diubah, ubah juga nama foldernya di Google Drive agar tetap rapi
        if (oldNama !== userObj.nama) {
          try {
            const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
            const folders = mainFolder.getFoldersByName(oldNama);
            while (folders.hasNext()) {
              folders.next().setName(userObj.nama);
            }
          } catch (e) {}
        }

        const absenData = absenSheet.getDataRange().getValues();
        for (let j = 1; j < absenData.length; j++) {
          if (String(absenData[j][2]).toLowerCase() === oldUsername.toLowerCase()) {
            absenSheet.getRange(j + 1, 3, 1, 5).setValues([[
              userObj.username,
              userObj.nama,
              userObj.status,
              userObj.provinsi,
              userObj.stasiun
            ]]);
          }
        }
        return { success: true, message: 'Data akun & riwayat berhasil diperbarui.' };
      }
    }
  } else {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === String(userObj.username).toLowerCase()) {
        return { success: false, message: 'Username sudah terdaftar. Gunakan username lain.' };
      }
    }
    const newId = 'USR-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMddHHmmss');
    sheet.appendRow([
      newId,
      userObj.username,
      userObj.password,
      userObj.role,
      userObj.nama,
      userObj.status,
      userObj.provinsi,
      userObj.stasiun
    ]);
    return { success: true, message: 'Akun baru berhasil ditambahkan.' };
  }
}

/**
 * FULL CASCADE DELETE:
 * 1. Hapus akun di Sheet USERS
 * 2. Hapus semua file foto Clock-In & Clock-Out milik user di Google Drive
 * 3. Hapus Folder Nama Sales beserta sub-folder tanggalnya di Google Drive
 * 4. Hapus semua baris riwayat absensi milik user di Sheet ABSENSI
 */
function deleteUserAdmin(userId, callerRole) {
  const role = String(callerRole || '').toLowerCase();
  if (role !== 'admin') {
    return { success: false, message: 'Akses Ditolak: Hanya Admin yang diizinkan menghapus akun.' };
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('USERS');
  const absenSheet = ss.getSheetByName('ABSENSI');
  const userData = userSheet.getDataRange().getValues();

  let targetUsername = null;
  let targetSalesName = null;

  for (let i = 1; i < userData.length; i++) {
    if (String(userData[i][0]) === String(userId)) {
      targetUsername = String(userData[i][1]).trim().toLowerCase();
      targetSalesName = String(userData[i][4]).trim();
      userSheet.deleteRow(i + 1);
      break;
    }
  }

  if (!targetUsername) {
    return { success: false, message: 'Akun tidak ditemukan.' };
  }

  // Hapus semua baris milik username di Sheet ABSENSI + hapus file fotonya satu per satu
  const absenData = absenSheet.getDataRange().getValues();
  let deletedLogs = 0;
  let deletedPhotos = 0;

  for (let j = absenData.length - 1; j >= 1; j--) {
    if (String(absenData[j][2]).trim().toLowerCase() === targetUsername) {
      const fotoInUrl  = absenData[j][11]; // Kolom L (Foto_In)
      const fotoOutUrl = absenData[j][16]; // Kolom Q (Foto_Out)

      if (trashDriveFileByUrl_(fotoInUrl)) deletedPhotos++;
      if (trashDriveFileByUrl_(fotoOutUrl)) deletedPhotos++;

      absenSheet.deleteRow(j + 1);
      deletedLogs++;
    }
  }

  // Hapus juga Folder Induk milik Sales tersebut di Google Drive
  trashSalesFolderInDrive_(targetSalesName);

  // Hapus semua data pelacakan rute milik user di Sheet GPS_TRACKS
  try {
    const trackSheet = ss.getSheetByName('GPS_TRACKS');
    if (trackSheet) {
      const tData = trackSheet.getDataRange().getValues();
      for (let k = tData.length - 1; k >= 1; k--) {
        if (String(tData[k][2]).trim().toLowerCase() === targetUsername) {
          trackSheet.deleteRow(k + 1);
        }
      }
    }
  } catch (eTrack) {}

  return { 
    success: true, 
    message: `Akun "${targetSalesName}" dihapus beserta ${deletedLogs} data absensi, rekaman GPS track, dan seluruh file fotonya di Google Drive.` 
  };
}

function formatRowDate_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  return String(val).replace(/^'/, '').trim();
}

function formatRowTime_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'Asia/Jakarta', 'HH:mm:ss');
  }
  return String(val).replace(/^'/, '').trim();
}

function formatHHMM_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'Asia/Jakarta', 'HH:mm');
  }
  let s = String(val).replace(/^'/, '').trim();
  let parts = s.split(':');
  if (parts.length >= 2) {
    return parts[0].padStart(2, '0') + ':' + parts[1].padStart(2, '0');
  }
  return s;
}

function formatNumberOrBlank_(val) {
  if (val === '' || val === null || val === undefined) return '';
  const n = Number(val);
  return isNaN(n) ? String(val) : n;
}
