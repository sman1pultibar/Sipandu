/**
 * ==============================================================================
 * SIPANDU - Aplikasi Digital Pengelolaan Panitia Ujian Sekolah
 * FILE: setup.gs
 * DESKRIPSI: Modul Inisialisasi Database & Safe Migration (Idempoten & Non-Destruktif)
 * STANDAR: ZettBOT Engineering Standard (Index-Based Column Mapping, No Underscore)
 * ==============================================================================
 */

// Konstanta Nama Sheet (Baku, Tanpa Karakter Underscore)
const SHEETS = {
  PENGATURAN: 'Pengaturan',
  PENGGUNA: 'Pengguna',
  DATA_SISWA: 'DataSiswa',
  DATA_PANITIA: 'DataPanitia',
  DATA_PENGAWAS: 'DataPengawas',
  DISTRIBUSI_RUANG: 'DistribusiRuang'
};

// Pemetaan Indeks Kolom Baku (0-based) agar kebal perubahan teks header
const COL_PENGATURAN = {
  KUNCI: 0,
  NILAI: 1
};

const COL_PENGGUNA = {
  ID_USER: 0,
  USERNAME: 1,
  PASSWORD_HASH: 2,
  SALT: 3,
  NAMA_LENGKAP: 4,
  TERAKHIR_LOGIN: 5
};

const COL_SISWA = {
  ID_SISWA: 0,
  NISN: 1,
  NIS: 2,
  NAMA_SISWA: 3,
  KELAS: 4,
  JENIS_KELAMIN: 5,
  NO_PESERTA: 6,
  FOTO_URL: 7
};

const COL_PANITIA = {
  ID_PANITIA: 0,
  NIP: 1,
  NAMA_PANITIA: 2,
  JABATAN: 3,
  NO_HP: 4,
  FOTO_URL: 5
};

const COL_PENGAWAS = {
  ID_PENGAWAS: 0,
  NIP: 1,
  NAMA_PENGAWAS: 2,
  ASAL_SEKOLAH: 3,
  RUANG_TUGAS: 4,
  SESI_TUGAS: 5,
  FOTO_URL: 6
};

const COL_DISTRIBUSI = {
  ID_DISTRIBUSI: 0,
  ID_SISWA: 1,
  NO_PESERTA: 2,
  NAMA_SISWA: 3,
  KELAS: 4,
  NAMA_RUANG: 5,
  NO_MEJA: 6,
  SESI: 7
};

// URL Logo Login Baku & Folder Google Drive
const DEFAULT_LOGO_URL = 'https://lh3.googleusercontent.com/d/1nqJ02h0xff2v_wgAQLAGi2_6KiewIf-k';
const DEFAULT_DRIVE_FOLDER_ID = '12e7ZcpAkpw5iBdddGZbzOZy7pFhmxEn-';

/**
 * Fungsi Utama Inisialisasi & Safe Migration
 * Aman dijalankan berulang kali tanpa merusak atau menghapus data riil yang sudah ada.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Trigger otorisasi Google Drive (DriveApp) saat pertama kali dijalankan
  try {
    DriveApp.getRootFolder();
  } catch (e) {
    Logger.log('DriveApp Auth Check: ' + e.toString());
  }
  
  // 1. Inisialisasi Sheet Pengaturan
  initSheet(ss, SHEETS.PENGATURAN, ['Kunci', 'Nilai'], function(sheet) {
    const defaultSettings = [
      ['NamaSekolah', 'SMK Negeri 1 Indonesia'],
      ['HeaderKop1', 'PEMERINTAH PROVINSI / DAERAH'],
      ['HeaderKop2', 'DINAS PENDIDIKAN DAN KEBUDAYAAN'],
      ['TahunAjaran', '2026/2027'],
      ['Semester', 'Genap'],
      ['NamaUjian', 'Asesmen Akhir Semester Genap'],
      ['KapasitasRuang', '20'],
      ['PrefixNoPeserta', 'UJN-2026-'],
      ['KepalaSekolah', 'Drs. H. Ahmad Fauzi, M.Pd.'],
      ['NipKepalaSekolah', '197501012000031001'],
      ['KetuaPanitia', 'Budi Santoso, S.Kom.'],
      ['NipKetuaPanitia', '198502022010011005'],
      ['TempatTanggal', 'Jakarta, 15 Mei 2026'],
      ['LogoDinas', DEFAULT_LOGO_URL],
      ['LogoSekolah', DEFAULT_LOGO_URL],
      ['SistemSesi', 'Tunggal'],
      ['FolderDriveId', DEFAULT_DRIVE_FOLDER_ID]
    ];
    sheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  });

  // 2. Inisialisasi Sheet Pengguna (Default Admin: username 'admin', password 'admin123')
  initSheet(ss, SHEETS.PENGGUNA, ['IdUser', 'Username', 'PasswordHash', 'Salt', 'NamaLengkap', 'TerakhirLogin'], function(sheet) {
    const salt = 'Sip4ndu!Salt99';
    const rawPass = 'admin123';
    const hash = computeSha256(rawPass + salt);
    const defaultUser = [
      ['USR-001', 'admin', hash, salt, 'Administrator Panitia', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')]
    ];
    sheet.getRange(2, 1, 1, 6).setValues(defaultUser);
  });

  // 3. Inisialisasi Sheet DataSiswa
  initSheet(ss, SHEETS.DATA_SISWA, ['IdSiswa', 'NISN', 'NIS', 'NamaSiswa', 'Kelas', 'JenisKelamin', 'NoPeserta', 'FotoUrl'], function(sheet) {
    const sampleStudents = [
      ['SISWA-0001', '0061234561', '24001', 'Ahmad Dani Pratama', 'XII RPL 1', 'Laki-laki', 'UJN-2026-0001', DEFAULT_LOGO_URL],
      ['SISWA-0002', '0061234562', '24002', 'Annisa Rahmawati', 'XII RPL 1', 'Perempuan', 'UJN-2026-0002', DEFAULT_LOGO_URL],
      ['SISWA-0003', '0061234563', '24003', 'Bagus Wicaksono', 'XII RPL 1', 'Laki-laki', 'UJN-2026-0003', DEFAULT_LOGO_URL],
      ['SISWA-0004', '0061234564', '24004', 'Citra Dewi Lestari', 'XII TKJ 1', 'Perempuan', 'UJN-2026-0004', DEFAULT_LOGO_URL],
      ['SISWA-0005', '0061234565', '24005', 'Dimas Arya Putra', 'XII TKJ 1', 'Laki-laki', 'UJN-2026-0005', DEFAULT_LOGO_URL]
    ];
    sheet.getRange(2, 1, sampleStudents.length, 8).setValues(sampleStudents);
  });

  // Pastikan Kolom NISN dan NIS diformat sebagai Plain Text (@) agar angka 0 di depan tidak pernah hilang
  const sheetSiswa = ss.getSheetByName(SHEETS.DATA_SISWA);
  if (sheetSiswa) {
    sheetSiswa.getRange("B:C").setNumberFormat('@');
  }

  // 4. Inisialisasi Sheet DataPanitia
  initSheet(ss, SHEETS.DATA_PANITIA, ['IdPanitia', 'NIP', 'NamaPanitia', 'Jabatan', 'NoHp', 'FotoUrl'], function(sheet) {
    const samplePanitia = [
      ['PAN-001', '198502022010011005', 'Budi Santoso, S.Kom.', 'Ketua Panitia', '081234567890', DEFAULT_LOGO_URL],
      ['PAN-002', '198904152015022003', 'Siti Rahayu, S.Pd.', 'Sekretaris', '081234567891', DEFAULT_LOGO_URL],
      ['PAN-003', '199208202019031008', 'Rian Hidayat, S.E.', 'Bendahara', '081234567892', DEFAULT_LOGO_URL]
    ];
    sheet.getRange(2, 1, samplePanitia.length, 6).setValues(samplePanitia);
  });

  const sheetPanitia = ss.getSheetByName(SHEETS.DATA_PANITIA);
  if (sheetPanitia) {
    sheetPanitia.getRange("B:B").setNumberFormat('@');
  }

  // 5. Inisialisasi Sheet DataPengawas
  initSheet(ss, SHEETS.DATA_PENGAWAS, ['IdPengawas', 'NIP', 'NamaPengawas', 'AsalSekolah', 'RuangTugas', 'SesiTugas', 'FotoUrl'], function(sheet) {
    const samplePengawas = [
      ['PGW-001', '198003102005011004', 'Drs. Hendro Wibowo', 'SMK Negeri 1 Indonesia', 'Ruang 01', 'Sesi 1', DEFAULT_LOGO_URL],
      ['PGW-002', '198407182009022006', 'Dra. Endang Purwanti', 'SMK Negeri 1 Indonesia', 'Ruang 02', 'Sesi 1', DEFAULT_LOGO_URL],
      ['PGW-003', '199011252014031002', 'Agus Salim, M.Pd.', 'SMK Negeri 1 Indonesia', 'Ruang 01', 'Sesi 2', DEFAULT_LOGO_URL]
    ];
    sheet.getRange(2, 1, samplePengawas.length, 7).setValues(samplePengawas);
  });

  const sheetPengawas = ss.getSheetByName(SHEETS.DATA_PENGAWAS);
  if (sheetPengawas) {
    sheetPengawas.getRange("B:B").setNumberFormat('@');
  }

  // 6. Inisialisasi Sheet DistribusiRuang
  initSheet(ss, SHEETS.DISTRIBUSI_RUANG, ['IdDistribusi', 'IdSiswa', 'NoPeserta', 'NamaSiswa', 'Kelas', 'NamaRuang', 'NoMeja', 'Sesi'], function(sheet) {
    // Awal dibuat kosong dengan header saja
  });

  // Rapikan sheet pertama jika ada sheet default "Sheet1"
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {
      // Abaikan jika tidak diizinkan
    }
  }

  // Otomatis pulihkan angka 0 di depan untuk NISN (10 digit standar Kemdikbud)
  fixExistingNisnLeadingZeros();

  // Otomatis buatkan subfolder Foto Siswa, Foto Panitia, dan Foto Pengawas di Google Drive
  setupDriveFolders();

  Logger.log('Inisialisasi Database Sipandu Berhasil Diselesaikan.');
  return { success: true, message: 'Database Sipandu berhasil disiapkan dan dimigrasikan.' };
}

/**
 * Helper Pembuatan Sheet Idempoten (Safe Migration)
 */
function initSheet(ss, sheetName, headers, populateCallback) {
  let sheet = ss.getSheetByName(sheetName);
  const isNew = !sheet;
  
  if (isNew) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeaderRow(sheet, headers.length);
    if (populateCallback) {
      populateCallback(sheet);
    }
  } else {
    // Sinkronisasi header (Safe Migration): Pastikan header lama tidak hilang, tambahkan kolom baru jika ada
    const existingLastCol = Math.max(sheet.getLastColumn(), 1);
    const existingHeaders = sheet.getRange(1, 1, 1, existingLastCol).getValues()[0];
    
    headers.forEach(function(h, idx) {
      if (idx >= existingHeaders.length || existingHeaders[idx] !== h) {
        sheet.getRange(1, idx + 1).setValue(h);
      }
    });
    formatHeaderRow(sheet, headers.length);
  }
  return sheet;
}

/**
 * Styling Header Row Profesional
 */
function formatHeaderRow(sheet, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground('#0f172a'); // Deep Navy
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontFamily('Arial');
  headerRange.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

/**
 * Algoritma SHA-256 Hashing menggunakan Google Apps Script Utilities
 */
function computeSha256(input) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    let byteHex = hashVal.toString(16);
    if (byteHex.length === 1) byteHex = '0' + byteHex;
    txtHash += byteHex;
  }
  return txtHash;
}

/**
 * Fungsi eksplisit untuk memicu jendela izin akses Google Drive (OAuth Scope)
 * sekaligus membuat subfolder foto dan memulihkan format NISN.
 */
function authorizeDrive() {
  const root = DriveApp.getRootFolder();
  Logger.log('Otorisasi Google Drive Berhasil: ' + root.getName());
  const driveResult = setupDriveFolders();
  const nisnResult = fixExistingNisnLeadingZeros();
  return 'Otorisasi Google Drive Berhasil! Subfolder Foto Siswa siap & NISN diperiksa.';
}

/**
 * Menyiapkan subfolder "Foto Siswa", "Foto Panitia", dan "Foto Pengawas"
 * di dalam folder induk Google Drive (atau fallback Sipandu_Media).
 */
function setupDriveFolders() {
  let parentFolder = null;
  try {
    parentFolder = DriveApp.getFolderById(DEFAULT_DRIVE_FOLDER_ID);
  } catch (e) {
    Logger.log('Folder induk eksternal tidak dapat diakses langsung: ' + e.toString());
  }

  if (!parentFolder) {
    const root = DriveApp.getRootFolder();
    const existing = root.getFoldersByName('Sipandu_Media');
    if (existing.hasNext()) {
      parentFolder = existing.next();
    } else {
      parentFolder = root.createFolder('Sipandu_Media');
    }
  }

  const subfolders = ['Foto Siswa', 'Foto Panitia', 'Foto Pengawas'];
  const createdFolders = {};

  subfolders.forEach(function(fName) {
    let folder = null;
    const it = parentFolder.getFoldersByName(fName);
    if (it.hasNext()) {
      folder = it.next();
    } else {
      try {
        folder = parentFolder.createFolder(fName);
      } catch (err) {
        Logger.log('Gagal membuat subfolder ' + fName + ': ' + err.toString());
      }
    }
    if (folder) {
      try {
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {}
      createdFolders[fName] = {
        id: folder.getId(),
        url: folder.getUrl()
      };
      Logger.log('Subfolder siap: ' + fName + ' -> ' + folder.getUrl());
    }
  });

  return { success: true, folders: createdFolders };
}

/**
 * Memperbaiki seluruh baris NISN di sheet DataSiswa yang kehilangan angka 0 di depan (misal: 86951614 -> 0086951614)
 * Menyesuaikan standar NISN Kemdikbud RI (tepat 10 digit plain text).
 */
function fixExistingNisnLeadingZeros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_SISWA);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true, count: 0 };

  const numRows = sheet.getLastRow() - 1;
  // Set format kolom B (NISN) dan C (NIS) ke Plain Text (@)
  sheet.getRange(2, 2, numRows, 2).setNumberFormat('@');

  const range = sheet.getRange(2, 2, numRows, 1);
  const values = range.getValues();
  let fixedCount = 0;

  for (let i = 0; i < values.length; i++) {
    let raw = String(values[i][0] || '').trim();
    if (raw && /^\d+$/.test(raw) && raw.length < 10) {
      values[i][0] = raw.padStart(10, '0');
      fixedCount++;
    }
  }

  if (fixedCount > 0) {
    range.setValues(values);
  }
  Logger.log('Perbaikan NISN berhasil: ' + fixedCount + ' data siswa dipulihkan angka nol di depannya.');
  return { success: true, fixedCount: fixedCount };
}
