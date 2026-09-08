// ============================================================
// WEB VERIFIKASI MITRA - BACKEND (Google Apps Script)
// Tempel file ini di Apps Script Editor (Extensions > Apps Script)
// pada Google Sheet "Data Pendaftar"
// ============================================================

var SHEET_LOGIN = 'Data Login Mitra';   // Nama Mitra | Username | Password
var SHEET_MITRA = 'Data Mitra';         // Nama Pelanggan | Nomor Telepon | Alamat | Nama Mitra | Stasiun | Latitude | Longitude | Status
var SESSION_DURATION = 21600;           // 6 jam (detik)
var HEADER_MITRA = ['Nama Pelanggan', 'Nomor Telepon', 'Alamat', 'Nama Mitra', 'Stasiun', 'Latitude', 'Longitude', 'Status'];

// ---------- DEBUG: jalankan manual dari editor untuk cek nama tab ----------
// Pilih fungsi ini di dropdown atas Apps Script editor, klik Run,
// lalu lihat hasilnya di menu View > Logs (atau Ctrl+Enter)
function debugCekNamaSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var names = sheets.map(function(s) { return "'" + s.getName() + "'"; });
  Logger.log('Nama-nama tab yang terdeteksi: ' + names.join(', '));
  Logger.log('Mencari SHEET_LOGIN = "' + SHEET_LOGIN + '" -> ditemukan: ' + (ss.getSheetByName(SHEET_LOGIN) !== null));
  Logger.log('Mencari SHEET_MITRA = "' + SHEET_MITRA + '" -> ditemukan: ' + (ss.getSheetByName(SHEET_MITRA) !== null));
}

// ---------- HALAMAN WEB / API JSON ----------
// Kalau diakses langsung tanpa parameter ?action=..., tampilkan halaman HTML (untuk tes langsung dari Apps Script).
// Kalau ada ?action=..., dijalankan sebagai API JSON (dipakai oleh frontend yang di-hosting terpisah, misal GitHub Pages).
// Frontend memanggil pakai fetch() method GET dengan parameter di URL -
// ini sengaja dipakai (bukan POST) karena GET tidak kena masalah redirect/CORS preflight di Apps Script Web App.
function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.action) {
    return handleApiRequest_(params);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Verifikasi Mitra')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Tetap disediakan untuk kompatibilitas kalau ada yang memanggil via POST.
function doPost(e) {
  var params = {};
  if (e && e.postData && e.postData.contents) {
    try { params = JSON.parse(e.postData.contents); } catch (err) {}
  }
  return handleApiRequest_(params);
}

function handleApiRequest_(params) {
  var out;
  try {
    var action = params.action;
    var result;
    var data = parseIfJson_(params.data);
    var rowIndex = params.rowIndex ? parseInt(params.rowIndex, 10) : null;

    switch (action) {
      case 'login':
        result = loginMitra(params.username, params.password);
        break;
      case 'logout':
        result = logoutMitra(params.token);
        break;
      case 'getData':
        result = getPelangganData(params.token);
        break;
      case 'update':
        result = updatePelanggan(params.token, rowIndex, data);
        break;
      case 'add':
        result = addPelanggan(params.token, data);
        break;
      case 'delete':
        result = deletePelanggan(params.token, rowIndex);
        break;
      default:
        result = { success: false, message: 'Aksi tidak dikenali: ' + action };
    }
    out = result;
  } catch (err) {
    out = { success: false, message: 'Terjadi kesalahan server: ' + err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// 'data' bisa datang sebagai string JSON (dari query GET) atau sudah berupa object (dari POST JSON)
function parseIfJson_(value) {
  if (typeof value === 'string') {
    try { return JSON.parse(value); } catch (err) { return value; }
  }
  return value;
}

// ---------- LOGIN / LOGOUT ----------
function loginMitra(username, password) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_LOGIN);
  if (!sheet) {
    return { success: false, message: 'Sheet "' + SHEET_LOGIN + '" tidak ditemukan. Cek nama tab di spreadsheet ini (jalankan debugCekNamaSheet di Apps Script untuk melihat daftar nama tab yang benar).' };
  }
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var namaMitra = data[i][0];
    var user = data[i][1];
    var pass = data[i][2];
    if (!user) continue;
    if (String(user).trim() === String(username).trim() &&
        String(pass).trim() === String(password).trim()) {
      var token = Utilities.getUuid();
      CacheService.getScriptCache().put(token, namaMitra, SESSION_DURATION);
      return { success: true, token: token, namaMitra: namaMitra };
    }
  }
  return { success: false, message: 'Username atau password salah.' };
}

function logoutMitra(token) {
  CacheService.getScriptCache().remove(token);
  return { success: true };
}

function getNamaMitraFromToken_(token) {
  return CacheService.getScriptCache().get(token);
}

// ---------- AMBIL DATA PELANGGAN (hanya milik mitra yang login) ----------
function getPelangganData(token) {
  var namaMitra = getNamaMitraFromToken_(token);
  if (!namaMitra) return { success: false, message: 'Sesi habis, silakan login ulang.' };

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_MITRA);
  if (!sheet) {
    return { success: false, message: 'Sheet "' + SHEET_MITRA + '" tidak ditemukan. Cek nama tab di spreadsheet ini.' };
  }
  ensureHeaders_(sheet);
  var data = sheet.getDataRange().getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (String(data[i][3]).trim() === String(namaMitra).trim()) {
      result.push({
        rowIndex: i + 1, // nomor baris asli di sheet, dipakai utk update
        namaPelanggan: data[i][0],
        noHp: data[i][1],
        alamat: data[i][2],
        namaMitra: data[i][3],
        stasiun: data[i][4],
        latitude: data[i][5],
        longitude: data[i][6],
        status: data[i][7] || ''
      });
    }
  }
  return { success: true, namaMitra: namaMitra, data: result };
}

// Mengisi header kolom yang masih kosong, tanpa menimpa header yang sudah ada
function ensureHeaders_(sheet) {
  var range = sheet.getRange(1, 1, 1, HEADER_MITRA.length);
  var header = range.getValues()[0];
  var changed = false;
  for (var i = 0; i < HEADER_MITRA.length; i++) {
    if (!header[i]) {
      header[i] = HEADER_MITRA[i];
      changed = true;
    }
  }
  if (changed) range.setValues([header]);
}

// ---------- UPDATE DATA PELANGGAN ----------
function updatePelanggan(token, rowIndex, updatedData) {
  var namaMitra = getNamaMitraFromToken_(token);
  if (!namaMitra) return { success: false, message: 'Sesi habis, silakan login ulang.' };

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_MITRA);
  var rowOwner = sheet.getRange(rowIndex, 4).getValue();
  if (String(rowOwner).trim() !== String(namaMitra).trim()) {
    return { success: false, message: 'Akses ditolak: data ini bukan milik mitra Anda.' };
  }

  sheet.getRange(rowIndex, 1).setValue(updatedData.namaPelanggan);
  sheet.getRange(rowIndex, 2).setValue(updatedData.noHp);
  sheet.getRange(rowIndex, 3).setValue(updatedData.alamat);
  sheet.getRange(rowIndex, 5).setValue(updatedData.stasiun);
  sheet.getRange(rowIndex, 6).setValue(updatedData.latitude);
  sheet.getRange(rowIndex, 7).setValue(updatedData.longitude);

  return { success: true };
}

// ---------- TAMBAH DATA PELANGGAN BARU ----------
function addPelanggan(token, newData) {
  var namaMitra = getNamaMitraFromToken_(token);
  if (!namaMitra) return { success: false, message: 'Sesi habis, silakan login ulang.' };

  if (!newData.namaPelanggan) {
    return { success: false, message: 'Nama Pelanggan wajib diisi.' };
  }

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_MITRA);
  ensureHeaders_(sheet);

  sheet.appendRow([
    newData.namaPelanggan,
    newData.noHp,
    newData.alamat,
    namaMitra,                 // Nama Mitra otomatis dari sesi login, tidak bisa dipalsukan
    newData.stasiun,
    newData.latitude,
    newData.longitude,
    ''
  ]);

  return { success: true };
}

// ---------- HAPUS DATA PELANGGAN ----------
function deletePelanggan(token, rowIndex) {
  var namaMitra = getNamaMitraFromToken_(token);
  if (!namaMitra) return { success: false, message: 'Sesi habis, silakan login ulang.' };

  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_MITRA);
  var rowOwner = sheet.getRange(rowIndex, 4).getValue();
  if (String(rowOwner).trim() !== String(namaMitra).trim()) {
    return { success: false, message: 'Akses ditolak: data ini bukan milik mitra Anda.' };
  }

  sheet.deleteRow(rowIndex);
  return { success: true };
}