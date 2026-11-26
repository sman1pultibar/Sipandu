/**
 * ==============================================================================
 * SIPANDU - Aplikasi Digital Pengelolaan Panitia Ujian Sekolah
 * FILE: code.gs
 * DESKRIPSI: Web App Controller, Routing, & Template Evaluator
 * STANDAR: Google Apps Script Web Architecture (Modular Single Page Application)
 * ==============================================================================
 */

/**
 * Endpoint Utama Web App (HTTP GET)
 * Merender antarmuka pengguna SPA dengan konfigurasi viewport responsif.
 */
function doGet(e) {
  // Pastikan database telah diinisialisasi jika belum ada
  ensureDatabaseReady();
  
  const template = HtmlService.createTemplateFromFile('index');
  
  // Inject konfigurasi awal agar dapat dibaca langsung oleh frontend
  template.appTitle = 'Sipandu - Pengelolaan Panitia Ujian Sekolah';
  template.logoLoginUrl = DEFAULT_LOGO_URL;
  
  const output = template.evaluate()
    .setTitle('Sipandu - Aplikasi Digital Pengelolaan Panitia Ujian Sekolah')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
  return output;
}

/**
 * Endpoint Eksternal API (HTTP POST) untuk Hosting Eksternal (Vercel, GitHub Pages, dll)
 * Mengizinkan pertukaran data JSON via fetch() dari Vercel ke Google Sheets.
 */
function doPost(e) {
  try {
    ensureDatabaseReady();
    let contents = {};
    if (e && e.postData && e.postData.contents) {
      try {
        contents = JSON.parse(e.postData.contents);
      } catch (pe) {
        contents = {};
      }
    }
    const action = contents.action || (e && e.parameter ? e.parameter.action : '');
    const payload = contents.payload || {};
    
    const result = apiRouter(action, payload);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Server error: ' + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper untuk menyertakan berkas HTML parsial (Modular CSS & JS)
 * Menghindari penumpukan kode pada satu berkas raksasa.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Memastikan lembar kerja telah siap dipakai sebelum aplikasi dibuka
 */
function ensureDatabaseReady() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.PENGATURAN) || !ss.getSheetByName(SHEETS.PENGGUNA)) {
    setupDatabase();
  }
}

/**
 * Universal Router API untuk menangani seluruh permintaan dari antarmuka pengguna
 * Pola Request-Response aman dengan error handling terpusat.
 */
function apiRouter(action, payload) {
  try {
    switch (action) {
      // Autentikasi & Akun
      case 'login':
        return apiLogin(payload);
      case 'changePassword':
        return apiChangePassword(payload);
        
      // Pengaturan Sistem & Identitas Sekolah
      case 'getSettings':
        return apiGetSettings();
      case 'saveSettings':
        return apiSaveSettings(payload);
        
      // Dashboard Metrik & Informasi
      case 'getDashboard':
        return apiGetDashboard();
        
      // Master Data Siswa
      case 'getStudents':
        return apiGetStudents(payload);
      case 'saveStudent':
        return apiSaveStudent(payload);
      case 'deleteStudent':
        return apiDeleteStudent(payload);
      case 'importStudents':
        return apiImportStudents(payload);
        
      // Master Data Panitia
      case 'getPanitia':
        return apiGetPanitia();
      case 'savePanitia':
        return apiSavePanitia(payload);
      case 'deletePanitia':
        return apiDeletePanitia(payload);
      case 'importPanitia':
        return apiImportPanitia(payload);
        
      // Master Data Pengawas
      case 'getPengawas':
        return apiGetPengawas();
      case 'savePengawas':
        return apiSavePengawas(payload);
      case 'deletePengawas':
        return apiDeletePengawas(payload);
      case 'importPengawas':
        return apiImportPengawas(payload);
        
      // Pembagian Ruang Ujian
      case 'generateRuang':
        return apiGenerateRuang(payload);
      case 'getDistribusiRuang':
        return apiGetDistribusiRuang(payload);
      case 'updateDistribusiKursi':
        return apiUpdateDistribusiKursi(payload);
        
      // Modul Cetak Dokumen Administrasi
      case 'getDataCetak':
        return apiGetDataCetak(payload);
        
      // Upload Media & Google Drive
      case 'uploadPhoto':
        return apiUploadPhoto(payload);
      case 'batchMatchPhotos':
        return apiBatchMatchPhotos(payload);
        
      // Backup & Restore Database
      case 'backupDatabase':
        return apiBackupDatabase();
      case 'restoreDatabase':
        return apiRestoreDatabase(payload);
      case 'resetDatabase':
        return apiResetDatabase(payload);
        
      default:
        return { success: false, message: 'Aksi API "' + action + '" tidak dikenali.' };
    }
  } catch (err) {
    Logger.log('Error API Router [' + action + ']: ' + err.toString());
    return { success: false, message: 'Terjadi kesalahan sistem: ' + err.toString() };
  }
}
