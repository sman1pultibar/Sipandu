/**
 * ==============================================================================
 * SIPANDU - Aplikasi Digital Pengelolaan Panitia Ujian Sekolah
 * FILE: api.gs
 * DESKRIPSI: Business Logic, CRUD Engine, Generator Ruang, Media CDN & Backup
 * STANDAR: ZettBOT Engineering Architecture (Index-Based, Safe Lock, Anti-Crash)
 * ==============================================================================
 */

// ==============================================================================
// 1. MODUL AUTENTIKASI & KEAMANAN AKUN
// ==============================================================================

function apiLogin(payload) {
  if (!payload || !payload.username || !payload.password) {
    return { success: false, message: 'Username dan password wajib diisi.' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PENGGUNA);
  if (!sheet) return { success: false, message: 'Tabel pengguna belum siap.' };
  
  const data = sheet.getDataRange().getValues();
  const targetUsername = String(payload.username).trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const username = String(row[COL_PENGGUNA.USERNAME]).trim().toLowerCase();
    
    if (username === targetUsername) {
      const storedHash = String(row[COL_PENGGUNA.PASSWORD_HASH]);
      const salt = String(row[COL_PENGGUNA.SALT]);
      const computedHash = computeSha256(payload.password + salt);
      
      if (storedHash === computedHash) {
        // Perbarui waktu terakhir login
        const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
        sheet.getRange(i + 1, COL_PENGGUNA.TERAKHIR_LOGIN + 1).setValue(nowStr);
        
        // Buat token sesi acak aman
        const token = computeSha256(username + nowStr + Math.random());
        
        return {
          success: true,
          message: 'Login berhasil.',
          user: {
            idUser: row[COL_PENGGUNA.ID_USER],
            username: row[COL_PENGGUNA.USERNAME],
            namaLengkap: row[COL_PENGGUNA.NAMA_LENGKAP]
          },
          token: token
        };
      } else {
        return { success: false, message: 'Kata sandi tidak sesuai.' };
      }
    }
  }
  
  return { success: false, message: 'Username tidak terdaftar.' };
}

function apiChangePassword(payload) {
  if (!payload || !payload.username || !payload.oldPassword || !payload.newPassword) {
    return { success: false, message: 'Semua kolom password wajib diisi.' };
  }
  
  if (payload.newPassword.length < 6) {
    return { success: false, message: 'Password baru minimal 6 karakter.' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PENGGUNA);
  const data = sheet.getDataRange().getValues();
  const targetUsername = String(payload.username).trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[COL_PENGGUNA.USERNAME]).trim().toLowerCase() === targetUsername) {
      const storedHash = String(row[COL_PENGGUNA.PASSWORD_HASH]);
      const salt = String(row[COL_PENGGUNA.SALT]);
      
      if (computeSha256(payload.oldPassword + salt) !== storedHash) {
        return { success: false, message: 'Password lama Anda salah.' };
      }
      
      // Buat salt baru dan simpan password baru
      const newSalt = Utilities.getUuid().substring(0, 10);
      const newHash = computeSha256(payload.newPassword + newSalt);
      
      sheet.getRange(i + 1, COL_PENGGUNA.PASSWORD_HASH + 1).setValue(newHash);
      sheet.getRange(i + 1, COL_PENGGUNA.SALT + 1).setValue(newSalt);
      
      return { success: true, message: 'Password berhasil diganti. Silakan gunakan password baru pada login berikutnya.' };
    }
  }
  
  return { success: false, message: 'Akun pengguna tidak ditemukan.' };
}

// ==============================================================================
// 2. MODUL PENGATURAN SISTEM & IDENTITAS SEKOLAH
// ==============================================================================

function apiGetSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PENGATURAN);
  if (!sheet) return { success: false, message: 'Sheet Pengaturan tidak ditemukan.' };
  
  const data = sheet.getDataRange().getValues();
  const settings = {};
  
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][COL_PENGATURAN.KUNCI]).trim();
    const val = data[i][COL_PENGATURAN.NILAI];
    if (key) {
      settings[key] = val;
    }
  }
  
  // Normalisasi URL gambar ke format Google CDN lh3
  if (settings.LogoDinas) settings.LogoDinas = formatToLh3Url(settings.LogoDinas);
  if (settings.LogoSekolah) settings.LogoSekolah = formatToLh3Url(settings.LogoSekolah);
  
  return { success: true, settings: settings };
}

function apiSaveSettings(payload) {
  if (!payload) return { success: false, message: 'Data pengaturan kosong.' };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PENGATURAN);
  const data = sheet.getDataRange().getValues();
  
  const existingKeys = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][COL_PENGATURAN.KUNCI]).trim();
    if (key) {
      existingKeys[key] = i + 1; // Baris 1-based di sheet
    }
  }
  
  // Terapkan perubahan dan tambah baris baru jika kunci belum ada
  for (const k in payload) {
    if (payload.hasOwnProperty(k)) {
      let val = payload[k];
      if (k === 'LogoDinas' || k === 'LogoSekolah') {
        val = formatToLh3Url(val);
      }
      
      if (existingKeys[k]) {
        sheet.getRange(existingKeys[k], COL_PENGATURAN.NILAI + 1).setValue(val);
      } else {
        sheet.appendRow([k, val]);
      }
    }
  }
  
  return { success: true, message: 'Pengaturan ujian berhasil disimpan.' };
}

// ==============================================================================
// 3. MODUL DASHBOARD RINGKASAN METRIK
// ==============================================================================

function apiGetDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Hitung jumlah siswa
  const siswaSheet = ss.getSheetByName(SHEETS.DATA_SISWA);
  const totalSiswa = siswaSheet ? Math.max(siswaSheet.getLastRow() - 1, 0) : 0;
  
  // Hitung jumlah panitia
  const panitiaSheet = ss.getSheetByName(SHEETS.DATA_PANITIA);
  const totalPanitia = panitiaSheet ? Math.max(panitiaSheet.getLastRow() - 1, 0) : 0;
  
  // Hitung jumlah pengawas
  const pengawasSheet = ss.getSheetByName(SHEETS.DATA_PENGAWAS);
  const totalPengawas = pengawasSheet ? Math.max(pengawasSheet.getLastRow() - 1, 0) : 0;
  
  // Hitung jumlah ruang unik dari DistribusiRuang
  const distSheet = ss.getSheetByName(SHEETS.DISTRIBUSI_RUANG);
  let totalRuang = 0;
  if (distSheet && distSheet.getLastRow() > 1) {
    const distData = distSheet.getRange(2, COL_DISTRIBUSI.NAMA_RUANG + 1, distSheet.getLastRow() - 1, 1).getValues();
    const uniqueRooms = {};
    for (let i = 0; i < distData.length; i++) {
      const r = String(distData[i][0]).trim();
      if (r) uniqueRooms[r] = true;
    }
    totalRuang = Object.keys(uniqueRooms).length;
  }
  
  const settingsRes = apiGetSettings();
  
  return {
    success: true,
    stats: {
      totalSiswa: totalSiswa,
      totalRuang: totalRuang,
      totalPanitia: totalPanitia,
      totalPengawas: totalPengawas
    },
    settings: settingsRes.settings || {}
  };
}

// ==============================================================================
// 4. MODUL MASTER DATA SISWA (PAGINATION, SEARCH, CRUD, IMPORT)
// ==============================================================================

function apiGetStudents(payload) {
  payload = payload || {};
  const page = parseInt(payload.page, 10) || 1;
  const limit = parseInt(payload.limit, 10) || 15;
  const search = payload.search ? String(payload.search).toLowerCase().trim() : '';
  const filterKelas = payload.kelas ? String(payload.kelas).trim() : '';
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_SISWA);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, data: [], total: 0, page: 1, limit: limit, totalPages: 0, classes: [] };
  }
  
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const allClassesMap = {};
  const filtered = [];
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const idSiswa = String(row[COL_SISWA.ID_SISWA] || '');
    let nisn = String(row[COL_SISWA.NISN] || '').trim();
    // Jamin format NISN 10 digit standar Kemdikbud RI
    if (nisn && /^\d+$/.test(nisn) && nisn.length < 10) {
      nisn = nisn.padStart(10, '0');
    }
    const nis = String(row[COL_SISWA.NIS] || '').trim();
    const namaSiswa = String(row[COL_SISWA.NAMA_SISWA] || '');
    const kelas = String(row[COL_SISWA.KELAS] || '');
    const jenisKelamin = String(row[COL_SISWA.JENIS_KELAMIN] || '');
    const noPeserta = String(row[COL_SISWA.NO_PESERTA] || '');
    let fotoUrl = String(row[COL_SISWA.FOTO_URL] || DEFAULT_LOGO_URL);
    fotoUrl = formatToLh3Url(fotoUrl);
    
    if (kelas) allClassesMap[kelas] = true;
    
    // Filter kelas jika dipilih
    if (filterKelas && kelas !== filterKelas) {
      continue;
    }
    
    // Filter pencarian
    if (search) {
      const match = namaSiswa.toLowerCase().indexOf(search) !== -1 ||
                    nisn.indexOf(search) !== -1 ||
                    nis.indexOf(search) !== -1 ||
                    kelas.toLowerCase().indexOf(search) !== -1 ||
                    noPeserta.toLowerCase().indexOf(search) !== -1;
      if (!match) continue;
    }
    
    filtered.push({
      rowIndex: i + 2,
      idSiswa: idSiswa,
      nisn: nisn,
      nis: nis,
      namaSiswa: namaSiswa,
      kelas: kelas,
      jenisKelamin: jenisKelamin,
      noPeserta: noPeserta,
      fotoUrl: fotoUrl
    });
  }
  
  const total = filtered.length;
  const totalPages = Math.ceil(total / limit) || 1;
  const startIdx = (page - 1) * limit;
  const pagedData = filtered.slice(startIdx, startIdx + limit);
  const classesList = Object.keys(allClassesMap).sort();
  
  return {
    success: true,
    data: pagedData,
    total: total,
    page: page,
    limit: limit,
    totalPages: totalPages,
    classes: classesList
  };
}

function apiSaveStudent(payload) {
  if (!payload || !payload.namaSiswa || !payload.kelas) {
    return { success: false, message: 'Nama siswa dan kelas wajib diisi.' };
  }
  
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.DATA_SISWA);
    let foto = payload.fotoUrl ? formatToLh3Url(payload.fotoUrl) : DEFAULT_LOGO_URL;
    
    // Mode Update jika ID Siswa disertakan
    let cleanNisn = String(payload.nisn || '').trim();
    if (cleanNisn && /^\d+$/.test(cleanNisn) && cleanNisn.length < 10) {
      cleanNisn = cleanNisn.padStart(10, '0');
    }
    const cleanNis = String(payload.nis || '').trim();

    if (payload.idSiswa) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][COL_SISWA.ID_SISWA]) === String(payload.idSiswa)) {
          sheet.getRange(i + 1, COL_SISWA.NISN + 1).setNumberFormat('@').setValue(cleanNisn);
          sheet.getRange(i + 1, COL_SISWA.NIS + 1).setNumberFormat('@').setValue(cleanNis);
          sheet.getRange(i + 1, COL_SISWA.NAMA_SISWA + 1).setValue(payload.namaSiswa);
          sheet.getRange(i + 1, COL_SISWA.KELAS + 1).setValue(payload.kelas);
          sheet.getRange(i + 1, COL_SISWA.JENIS_KELAMIN + 1).setValue(payload.jenisKelamin || '');
          sheet.getRange(i + 1, COL_SISWA.NO_PESERTA + 1).setValue(payload.noPeserta || '');
          sheet.getRange(i + 1, COL_SISWA.FOTO_URL + 1).setValue(foto);
          
          return { success: true, message: 'Data siswa berhasil diperbarui.' };
        }
      }
    }
    
    // Mode Tambah Baru: Generate ID Sekuensial & No Peserta Otomatis
    const newId = generateNextId(sheet, COL_SISWA.ID_SISWA, 'SISWA-', 4);
    let noPeserta = payload.noPeserta;
    if (!noPeserta) {
      const settings = apiGetSettings().settings || {};
      const prefix = settings.PrefixNoPeserta || 'UJN-2026-';
      noPeserta = generateNextId(sheet, COL_SISWA.NO_PESERTA, prefix, 4);
    }
    
    sheet.appendRow([
      newId,
      cleanNisn,
      cleanNis,
      payload.namaSiswa,
      payload.kelas,
      payload.jenisKelamin || 'Laki-laki',
      noPeserta,
      foto
    ]);
    sheet.getRange(sheet.getLastRow(), COL_SISWA.NISN + 1, 1, 2).setNumberFormat('@');
    
    return { success: true, message: 'Siswa baru berhasil ditambahkan.' };
  } finally {
    lock.releaseLock();
  }
}

function apiDeleteStudent(payload) {
  if (!payload || !payload.idSiswa) return { success: false, message: 'ID Siswa tidak ditemukan.' };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_SISWA);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_SISWA.ID_SISWA]) === String(payload.idSiswa)) {
      sheet.deleteRow(i + 1);
      
      // Hapus juga dari distribusi ruang jika sudah terdaftar
      removeStudentFromDistribution(payload.idSiswa);
      return { success: true, message: 'Data siswa berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Siswa tidak ditemukan.' };
}

function apiImportStudents(payload) {
  if (!payload || !payload.students || !payload.students.length) {
    return { success: false, message: 'Tidak ada data siswa yang diimpor.' };
  }
  
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.DATA_SISWA);
    const settings = apiGetSettings().settings || {};
    const prefix = settings.PrefixNoPeserta || 'UJN-2026-';
    
    let currentLastNum = getMaxNumberInCol(sheet, COL_SISWA.ID_SISWA, 'SISWA-');
    let currentPesertaNum = getMaxNumberInCol(sheet, COL_SISWA.NO_PESERTA, prefix);
    
    const rowsToAdd = [];
    payload.students.forEach(function(s) {
      currentLastNum++;
      currentPesertaNum++;
      const idSiswa = 'SISWA-' + padZero(currentLastNum, 4);
      const noPeserta = s.noPeserta ? s.noPeserta : (prefix + padZero(currentPesertaNum, 4));
      const foto = s.fotoUrl ? formatToLh3Url(s.fotoUrl) : DEFAULT_LOGO_URL;
      
      let cleanNisn = String(s.nisn || '').trim();
      if (cleanNisn && /^\d+$/.test(cleanNisn) && cleanNisn.length < 10) {
        cleanNisn = cleanNisn.padStart(10, '0');
      }
      const cleanNis = String(s.nis || '').trim();

      rowsToAdd.push([
        idSiswa,
        cleanNisn,
        cleanNis,
        s.namaSiswa || 'Siswa',
        s.kelas || 'Umum',
        s.jenisKelamin || 'Laki-laki',
        noPeserta,
        foto
      ]);
    });
    
    if (rowsToAdd.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 2, rowsToAdd.length, 2).setNumberFormat('@');
      sheet.getRange(startRow, 1, rowsToAdd.length, 8).setValues(rowsToAdd);
    }
    
    return { success: true, message: rowsToAdd.length + ' data siswa berhasil diimpor ke sistem (format NISN 10 digit dipertahankan).' };
  } finally {
    lock.releaseLock();
  }
}

// ==============================================================================
// 5. MODUL MASTER DATA PANITIA & PENGAWAS
// ==============================================================================

function apiGetPanitia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_PANITIA);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
  
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const result = [];
  
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    result.push({
      idPanitia: String(r[COL_PANITIA.ID_PANITIA] || ''),
      nip: String(r[COL_PANITIA.NIP] || ''),
      namaPanitia: String(r[COL_PANITIA.NAMA_PANITIA] || ''),
      jabatan: String(r[COL_PANITIA.JABATAN] || ''),
      noHp: String(r[COL_PANITIA.NO_HP] || ''),
      fotoUrl: formatToLh3Url(String(r[COL_PANITIA.FOTO_URL] || DEFAULT_LOGO_URL))
    });
  }
  return { success: true, data: result };
}

function apiSavePanitia(payload) {
  if (!payload || !payload.namaPanitia || !payload.jabatan) {
    return { success: false, message: 'Nama panitia dan jabatan wajib diisi.' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_PANITIA);
  const foto = payload.fotoUrl ? formatToLh3Url(payload.fotoUrl) : DEFAULT_LOGO_URL;
  
  if (payload.idPanitia) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL_PANITIA.ID_PANITIA]) === String(payload.idPanitia)) {
        sheet.getRange(i + 1, COL_PANITIA.NIP + 1).setValue(payload.nip || '');
        sheet.getRange(i + 1, COL_PANITIA.NAMA_PANITIA + 1).setValue(payload.namaPanitia);
        sheet.getRange(i + 1, COL_PANITIA.JABATAN + 1).setValue(payload.jabatan);
        sheet.getRange(i + 1, COL_PANITIA.NO_HP + 1).setValue(payload.noHp || '');
        sheet.getRange(i + 1, COL_PANITIA.FOTO_URL + 1).setValue(foto);
        return { success: true, message: 'Data panitia berhasil diperbarui.' };
      }
    }
  }
  
  const newId = generateNextId(sheet, COL_PANITIA.ID_PANITIA, 'PAN-', 3);
  sheet.appendRow([newId, payload.nip || '', payload.namaPanitia, payload.jabatan, payload.noHp || '', foto]);
  return { success: true, message: 'Panitia baru berhasil ditambahkan.' };
}

function apiDeletePanitia(payload) {
  if (!payload || !payload.idPanitia) return { success: false, message: 'ID Panitia tidak valid.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_PANITIA);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_PANITIA.ID_PANITIA]) === String(payload.idPanitia)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Data panitia berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Panitia tidak ditemukan.' };
}

function apiGetPengawas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_PENGAWAS);
  if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };
  
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const result = [];
  
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    result.push({
      idPengawas: String(r[COL_PENGAWAS.ID_PENGAWAS] || ''),
      nip: String(r[COL_PENGAWAS.NIP] || ''),
      namaPengawas: String(r[COL_PENGAWAS.NAMA_PENGAWAS] || ''),
      asalSekolah: String(r[COL_PENGAWAS.ASAL_SEKOLAH] || ''),
      ruangTugas: String(r[COL_PENGAWAS.RUANG_TUGAS] || ''),
      sesiTugas: String(r[COL_PENGAWAS.SESI_TUGAS] || ''),
      fotoUrl: formatToLh3Url(String(r[COL_PENGAWAS.FOTO_URL] || DEFAULT_LOGO_URL))
    });
  }
  return { success: true, data: result };
}

function apiSavePengawas(payload) {
  if (!payload || !payload.namaPengawas) {
    return { success: false, message: 'Nama pengawas wajib diisi.' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_PENGAWAS);
  const foto = payload.fotoUrl ? formatToLh3Url(payload.fotoUrl) : DEFAULT_LOGO_URL;
  
  if (payload.idPengawas) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL_PENGAWAS.ID_PENGAWAS]) === String(payload.idPengawas)) {
        sheet.getRange(i + 1, COL_PENGAWAS.NIP + 1).setValue(payload.nip || '');
        sheet.getRange(i + 1, COL_PENGAWAS.NAMA_PENGAWAS + 1).setValue(payload.namaPengawas);
        sheet.getRange(i + 1, COL_PENGAWAS.ASAL_SEKOLAH + 1).setValue(payload.asalSekolah || '');
        sheet.getRange(i + 1, COL_PENGAWAS.RUANG_TUGAS + 1).setValue(payload.ruangTugas || '');
        sheet.getRange(i + 1, COL_PENGAWAS.SESI_TUGAS + 1).setValue(payload.sesiTugas || '');
        sheet.getRange(i + 1, COL_PENGAWAS.FOTO_URL + 1).setValue(foto);
        return { success: true, message: 'Data pengawas berhasil diperbarui.' };
      }
    }
  }
  
  const newId = generateNextId(sheet, COL_PENGAWAS.ID_PENGAWAS, 'PGW-', 3);
  sheet.appendRow([newId, payload.nip || '', payload.namaPengawas, payload.asalSekolah || '', payload.ruangTugas || '', payload.sesiTugas || '', foto]);
  return { success: true, message: 'Pengawas baru berhasil ditambahkan.' };
}

function apiDeletePengawas(payload) {
  if (!payload || !payload.idPengawas) return { success: false, message: 'ID Pengawas tidak valid.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_PENGAWAS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_PENGAWAS.ID_PENGAWAS]) === String(payload.idPengawas)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Data pengawas berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Pengawas tidak ditemukan.' };
}

// ==============================================================================
// 6. MESIN GENERATOR PEMBAGIAN RUANG UJIAN (3 METODE & SESI FLEKSIBEL)
// ==============================================================================

function apiGenerateRuang(payload) {
  payload = payload || {};
  const mode = payload.mode || 'kelas'; // 'kelas' | 'acak' | 'manual'
  const kapasitas = parseInt(payload.kapasitasRuang, 10) || 20;
  const sistemSesi = payload.sistemSesi || 'Tunggal'; // 'Tunggal' | 'Multi'
  const jumlahSesi = parseInt(payload.jumlahSesi, 10) || (sistemSesi === 'Multi' ? 2 : 1);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const siswaSheet = ss.getSheetByName(SHEETS.DATA_SISWA);
  if (!siswaSheet || siswaSheet.getLastRow() <= 1) {
    return { success: false, message: 'Belum ada data siswa di database untuk dibagi ke ruang ujian.' };
  }
  
  const rawStudents = siswaSheet.getRange(2, 1, siswaSheet.getLastRow() - 1, siswaSheet.getLastColumn()).getValues();
  let students = [];
  
  for (let i = 0; i < rawStudents.length; i++) {
    const r = rawStudents[i];
    students.push({
      idSiswa: String(r[COL_SISWA.ID_SISWA]),
      nisn: String(r[COL_SISWA.NISN]),
      noPeserta: String(r[COL_SISWA.NO_PESERTA]),
      namaSiswa: String(r[COL_SISWA.NAMA_SISWA]),
      kelas: String(r[COL_SISWA.KELAS])
    });
  }
  
  // Algoritma Penataan Berdasarkan Mode
  if (mode === 'kelas') {
    // Urutkan berdasarkan Kelas ASC, lalu Nama Siswa ASC
    students.sort(function(a, b) {
      if (a.kelas === b.kelas) {
        return a.namaSiswa.localeCompare(b.namaSiswa);
      }
      return a.kelas.localeCompare(b.kelas);
    });
  } else if (mode === 'acak') {
    // Acak Silang Antar Kelas (Round-Robin Interleaving):
    // Kelompokkan per kelas, lalu ambil bergantian 1 per 1 agar siswa sekelas terpisah
    const classGroups = {};
    students.forEach(function(s) {
      if (!classGroups[s.kelas]) classGroups[s.kelas] = [];
      classGroups[s.kelas].push(s);
    });
    
    const interleaved = [];
    const keys = Object.keys(classGroups);
    let hasMore = true;
    let round = 0;
    
    while (hasMore) {
      hasMore = false;
      for (let k = 0; k < keys.length; k++) {
        const grp = classGroups[keys[k]];
        if (round < grp.length) {
          interleaved.push(grp[round]);
          hasMore = true;
        }
      }
      round++;
    }
    students = interleaved;
  }
  
  // Hitung Alokasi Ruang & Kursi
  const distSheet = ss.getSheetByName(SHEETS.DISTRIBUSI_RUANG);
  // Bersihkan data lama tanpa menghapus header (baris 1)
  if (distSheet.getLastRow() > 1) {
    distSheet.getRange(2, 1, distSheet.getLastRow() - 1, distSheet.getLastColumn()).clearContent();
  }
  
  const newDistributionRows = [];
  let distIdNum = 1;
  
  // Logika Pembagian Sesi
  const studentsPerSession = Math.ceil(students.length / jumlahSesi);
  
  for (let s = 1; s <= jumlahSesi; s++) {
    const sessionLabel = jumlahSesi > 1 ? ('Sesi ' + s) : 'Sesi 1';
    const sessionStudents = students.slice((s - 1) * studentsPerSession, s * studentsPerSession);
    
    let currentRoomIndex = 1;
    let currentSeatNumber = 1;
    
    for (let j = 0; j < sessionStudents.length; j++) {
      const st = sessionStudents[j];
      const roomName = 'Ruang ' + padZero(currentRoomIndex, 2);
      const idDist = 'DST-' + padZero(distIdNum++, 4);
      
      newDistributionRows.push([
        idDist,
        st.idSiswa,
        st.noPeserta,
        st.namaSiswa,
        st.kelas,
        roomName,
        currentSeatNumber,
        sessionLabel
      ]);
      
      currentSeatNumber++;
      if (currentSeatNumber > kapasitas) {
        currentSeatNumber = 1;
        currentRoomIndex++;
      }
    }
  }
  
  if (newDistributionRows.length > 0) {
    distSheet.getRange(2, 1, newDistributionRows.length, 8).setValues(newDistributionRows);
  }
  
  // Simpan pengaturan kapasitas & sesi yang dipakai
  apiSaveSettings({
    KapasitasRuang: kapasitas,
    SistemSesi: sistemSesi
  });
  
  return {
    success: true,
    message: 'Pembagian ' + students.length + ' siswa berhasil digenerate menggunakan mode: ' + mode.toUpperCase() + '.',
    totalSiswa: students.length,
    totalBaris: newDistributionRows.length
  };
}

function apiGetDistribusiRuang(payload) {
  payload = payload || {};
  const filterRuang = payload.ruang ? String(payload.ruang).trim() : '';
  const filterSesi = payload.sesi ? String(payload.sesi).trim() : '';
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DISTRIBUSI_RUANG);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, data: [], rooms: [], sessions: [] };
  }
  
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const roomsMap = {};
  const sessionsMap = {};
  const filtered = [];
  
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const namaRuang = String(r[COL_DISTRIBUSI.NAMA_RUANG] || '');
    const sesi = String(r[COL_DISTRIBUSI.SESI] || '');
    
    if (namaRuang) roomsMap[namaRuang] = true;
    if (sesi) sessionsMap[sesi] = true;
    
    if (filterRuang && namaRuang !== filterRuang) continue;
    if (filterSesi && sesi !== filterSesi) continue;
    
    filtered.push({
      idDistribusi: String(r[COL_DISTRIBUSI.ID_DISTRIBUSI]),
      idSiswa: String(r[COL_DISTRIBUSI.ID_SISWA]),
      noPeserta: String(r[COL_DISTRIBUSI.NO_PESERTA]),
      namaSiswa: String(r[COL_DISTRIBUSI.NAMA_SISWA]),
      kelas: String(r[COL_DISTRIBUSI.KELAS]),
      namaRuang: namaRuang,
      noMeja: parseInt(r[COL_DISTRIBUSI.NO_MEJA], 10) || 1,
      sesi: sesi
    });
  }
  
  return {
    success: true,
    data: filtered,
    rooms: Object.keys(roomsMap).sort(),
    sessions: Object.keys(sessionsMap).sort()
  };
}

function apiUpdateDistribusiKursi(payload) {
  if (!payload || !payload.idDistribusi || !payload.namaRuang || !payload.noMeja) {
    return { success: false, message: 'Data penyesuaian kursi tidak lengkap.' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DISTRIBUSI_RUANG);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL_DISTRIBUSI.ID_DISTRIBUSI]) === String(payload.idDistribusi)) {
      sheet.getRange(i + 1, COL_DISTRIBUSI.NAMA_RUANG + 1).setValue(payload.namaRuang);
      sheet.getRange(i + 1, COL_DISTRIBUSI.NO_MEJA + 1).setValue(payload.noMeja);
      if (payload.sesi) {
        sheet.getRange(i + 1, COL_DISTRIBUSI.SESI + 1).setValue(payload.sesi);
      }
      return { success: true, message: 'Posisi kursi peserta berhasil disesuaikan.' };
    }
  }
  return { success: false, message: 'Data distribusi tidak ditemukan.' };
}

// ==============================================================================
// 7. MODUL CETAK DOKUMEN ADMINISTRASI UJIAN (DATA FETCHER)
// ==============================================================================

function apiGetDataCetak(payload) {
  payload = payload || {};
  const jenis = payload.jenis; // 'kartuPeserta', 'nomorMeja', 'kartuPanitia', 'kartuPengawas', 'daftarHadir', 'beritaAcara', 'denahRuang'
  const filterRuang = payload.ruang ? String(payload.ruang).trim() : '';
  const filterSesi = payload.sesi ? String(payload.sesi).trim() : '';
  const filterKelas = payload.kelas ? String(payload.kelas).trim() : '';
  
  const settings = apiGetSettings().settings || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (jenis === 'kartuPanitia') {
    const panitiaRes = apiGetPanitia();
    return {
      success: true,
      jenis: jenis,
      settings: settings,
      items: panitiaRes.data || []
    };
  }
  
  if (jenis === 'kartuPengawas') {
    const pengawasRes = apiGetPengawas();
    return {
      success: true,
      jenis: jenis,
      settings: settings,
      items: pengawasRes.data || []
    };
  }
  
  // Ambil data siswa & distribusinya
  const distSheet = ss.getSheetByName(SHEETS.DISTRIBUSI_RUANG);
  const siswaSheet = ss.getSheetByName(SHEETS.DATA_SISWA);
  
  // Map foto, NIS, dan Jenis Kelamin siswa by idSiswa
  const fotoMap = {};
  const nisMap = {};
  const jkMap = {};
  if (siswaSheet && siswaSheet.getLastRow() > 1) {
    const sData = siswaSheet.getRange(2, 1, siswaSheet.getLastRow() - 1, siswaSheet.getLastColumn()).getValues();
    for (let i = 0; i < sData.length; i++) {
      const idS = String(sData[i][COL_SISWA.ID_SISWA]);
      fotoMap[idS] = formatToLh3Url(String(sData[i][COL_SISWA.FOTO_URL] || DEFAULT_LOGO_URL));
      nisMap[idS] = String(sData[i][COL_SISWA.NIS] || sData[i][COL_SISWA.NISN] || '-');
      jkMap[idS] = String(sData[i][COL_SISWA.JENIS_KELAMIN] || 'Laki-laki');
    }
  }
  
  const items = [];
  if (distSheet && distSheet.getLastRow() > 1) {
    const dData = distSheet.getRange(2, 1, distSheet.getLastRow() - 1, distSheet.getLastColumn()).getValues();
    for (let j = 0; j < dData.length; j++) {
      const row = dData[j];
      const idSiswa = String(row[COL_DISTRIBUSI.ID_SISWA]);
      const noPeserta = String(row[COL_DISTRIBUSI.NO_PESERTA]);
      const namaSiswa = String(row[COL_DISTRIBUSI.NAMA_SISWA]);
      const kelas = String(row[COL_DISTRIBUSI.KELAS]);
      const namaRuang = String(row[COL_DISTRIBUSI.NAMA_RUANG]);
      const noMeja = parseInt(row[COL_DISTRIBUSI.NO_MEJA], 10) || 1;
      const sesi = String(row[COL_DISTRIBUSI.SESI] || 'Sesi 1');
      
      if (filterRuang && namaRuang !== filterRuang) continue;
      if (filterSesi && sesi !== filterSesi) continue;
      if (filterKelas && kelas !== filterKelas) continue;
      
      items.push({
        idSiswa: idSiswa,
        noPeserta: noPeserta,
        namaSiswa: namaSiswa,
        kelas: kelas,
        namaRuang: namaRuang,
        noMeja: noMeja,
        sesi: sesi,
        nis: nisMap[idSiswa] || '-',
        jenisKelamin: jkMap[idSiswa] || 'Laki-laki',
        fotoUrl: fotoMap[idSiswa] || DEFAULT_LOGO_URL
      });
    }
  }
  
  // Jika items masih kosong (misal belum klik Pembagian Ruang), sediakan data dari DataSiswa agar pratinjau live tetap terlihat
  if (items.length === 0 && siswaSheet && siswaSheet.getLastRow() > 1) {
    const sData = siswaSheet.getRange(2, 1, siswaSheet.getLastRow() - 1, siswaSheet.getLastColumn()).getValues();
    for (let k = 0; k < sData.length; k++) {
      const r = sData[k];
      const idS = String(r[COL_SISWA.ID_SISWA]);
      const kl = String(r[COL_SISWA.KELAS]);
      if (filterKelas && kl !== filterKelas) continue;
      
      items.push({
        idSiswa: idS,
        noPeserta: String(r[COL_SISWA.NO_PESERTA] || ('UJN-' + (k + 1))),
        namaSiswa: String(r[COL_SISWA.NAMA_SISWA] || 'Siswa ' + (k + 1)),
        kelas: kl || 'Umum',
        namaRuang: filterRuang || 'Ruang 01',
        noMeja: k + 1,
        sesi: filterSesi || 'Sesi 1',
        nis: String(r[COL_SISWA.NIS] || r[COL_SISWA.NISN] || '-'),
        jenisKelamin: String(r[COL_SISWA.JENIS_KELAMIN] || 'Laki-laki'),
        fotoUrl: formatToLh3Url(String(r[COL_SISWA.FOTO_URL] || DEFAULT_LOGO_URL))
      });
    }
  }
  
  return {
    success: true,
    jenis: jenis,
    settings: settings,
    items: items
  };
}

// ==============================================================================
// 8. MODUL MEDIA & GOOGLE DRIVE (UPLOAD & BATCH MATCHING)
// ==============================================================================

function apiUploadPhoto(payload) {
  if (!payload || !payload.base64Data || !payload.fileName) {
    return { success: false, message: 'File gambar tidak valid.' };
  }
  
  const category = payload.folderType || 'FotoSiswa';
  const isLogo = (category === 'logoDinas' || category === 'logoSekolah');
  const contentType = payload.mimeType || 'image/jpeg';
  
  try {
    let targetFolder = null;
    const parentFolderId = DEFAULT_DRIVE_FOLDER_ID;
    
    // 1. Coba akses folder induk eksternal
    try {
      if (parentFolderId) {
        targetFolder = DriveApp.getFolderById(parentFolderId);
      }
    } catch (e) {
      Logger.log('Folder eksternal tidak dapat diakses: ' + e.toString());
      targetFolder = null;
    }
    
    // 2. Jika folder eksternal ditolak, gunakan folder "Sipandu_Media" di root Google Drive pengguna
    if (!targetFolder) {
      try {
        const root = DriveApp.getRootFolder();
        const existing = root.getFoldersByName('Sipandu_Media');
        if (existing.hasNext()) {
          targetFolder = existing.next();
        } else {
          targetFolder = root.createFolder('Sipandu_Media');
        }
      } catch (e) {
        Logger.log('Gagal membuat folder di root: ' + e.toString());
      }
    }
    
    // 3. Simpan berkas jika targetFolder tersedia
    if (targetFolder) {
      let subfolder = targetFolder;
      try {
        const subIter = targetFolder.getFoldersByName(category);
        if (subIter.hasNext()) {
          subfolder = subIter.next();
        } else {
          subfolder = targetFolder.createFolder(category);
        }
      } catch (e) {
        subfolder = targetFolder;
      }
      
      const decodedBytes = Utilities.base64Decode(payload.base64Data);
      const blob = Utilities.newBlob(decodedBytes, contentType, payload.fileName);
      const file = subfolder.createFile(blob);
      
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        Logger.log('Set sharing notice: ' + e.toString());
      }
      
      const fileId = file.getId();
      const lh3Url = 'https://lh3.googleusercontent.com/d/' + fileId;
      
      return {
        success: true,
        message: 'Foto berhasil disimpan ke Google Drive!',
        fileId: fileId,
        url: lh3Url
      };
    }
    
    throw new Error('DriveApp tidak dapat mengakses penyimpanan folder.');
  } catch (err) {
    Logger.log('Error upload: ' + err.toString());
    // Fallback: Jika DriveApp dibatasi, simpan logo/gambar langsung sebagai Base64 Data-URI
    if (isLogo || payload.base64Data.length < 2000000) {
      const dataUri = 'data:' + contentType + ';base64,' + payload.base64Data;
      return {
        success: true,
        message: 'Foto disimpan sebagai gambar lokal (Fallback Data-URI).',
        url: dataUri
      };
    }
    return { success: false, message: 'Gagal mengunggah foto: ' + err.toString() };
  }
}

function apiBatchMatchPhotos(payload) {
  payload = payload || {};
  const kategori = payload.kategori || 'siswa'; // 'siswa' | 'panitia' | 'pengawas'
  
  try {
    const rawSearchRoots = [];
    try {
      if (DEFAULT_DRIVE_FOLDER_ID) {
        rawSearchRoots.push(DriveApp.getFolderById(DEFAULT_DRIVE_FOLDER_ID));
      }
    } catch (e) {}
    
    try {
      const rootFolders = DriveApp.getRootFolder().getFoldersByName('Sipandu_Media');
      if (rootFolders.hasNext()) rawSearchRoots.push(rootFolders.next());
    } catch (e) {}
    
    if (rawSearchRoots.length === 0) {
      rawSearchRoots.push(DriveApp.getRootFolder());
    }
    
    // Kumpulkan seluruh subfolder rekursif (termasuk "Foto Siswa", "Foto Panitia", "Foto Pengawas", dll.)
    const searchFolders = [];
    function collectSubfolders(parent, list, depth) {
      if (!parent || depth > 3) return;
      list.push(parent);
      try {
        const sub = parent.getFolders();
        while (sub.hasNext()) {
          collectSubfolders(sub.next(), list, depth + 1);
        }
      } catch (e) {}
    }

    rawSearchRoots.forEach(function(rFolder) {
      collectSubfolders(rFolder, searchFolders, 0);
    });
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let targetSheet, mapCol, urlCol;
    
    if (kategori === 'siswa') {
      targetSheet = ss.getSheetByName(SHEETS.DATA_SISWA);
      mapCol = COL_SISWA.NISN; // Format: NISN.jpg
      urlCol = COL_SISWA.FOTO_URL;
    } else if (kategori === 'panitia') {
      targetSheet = ss.getSheetByName(SHEETS.DATA_PANITIA);
      mapCol = COL_PANITIA.NIP; // Format: NIP.jpg
      urlCol = COL_PANITIA.FOTO_URL;
    } else if (kategori === 'pengawas') {
      targetSheet = ss.getSheetByName(SHEETS.DATA_PENGAWAS);
      mapCol = COL_PENGAWAS.NIP; // Format: NIP.jpg
      urlCol = COL_PENGAWAS.FOTO_URL;
    }
    
    if (!targetSheet || targetSheet.getLastRow() <= 1) {
      return { success: false, message: 'Data ' + kategori + ' kosong.' };
    }
    
    const data = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).getValues();
    const entityMap = {};
    for (let i = 0; i < data.length; i++) {
      const rawKey = String(data[i][mapCol]).trim();
      if (rawKey) {
        entityMap[rawKey] = i + 2;
        // Jamin pencocokan dua arah: format 10 digit (0086951614) dan format polos (86951614)
        if (kategori === 'siswa' && /^\d+$/.test(rawKey)) {
          if (rawKey.length < 10) {
            entityMap[rawKey.padStart(10, '0')] = i + 2;
          } else if (rawKey.startsWith('0')) {
            entityMap[rawKey.replace(/^0+/, '')] = i + 2;
          }
        }
      }
      if (kategori === 'siswa') {
        const noPes = String(data[i][COL_SISWA.NO_PESERTA]).trim();
        if (noPes) entityMap[noPes] = i + 2;
      }
    }
    
    let matchCount = 0;
    const visitedFileIds = {};
    
    searchFolders.forEach(function(folder) {
      try {
        const files = folder.searchFiles('mimeType contains "image/"');
        while (files.hasNext()) {
          const f = files.next();
          const fId = f.getId();
          if (visitedFileIds[fId]) continue;
          visitedFileIds[fId] = true;
          
          const baseName = f.getName().replace(/\.[^/.]+$/, '').trim();
          const targetRow = entityMap[baseName];
          if (targetRow) {
            try {
              f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } catch (e) {}
            const lh3Url = 'https://lh3.googleusercontent.com/d/' + fId;
            targetSheet.getRange(targetRow, urlCol + 1).setValue(lh3Url);
            matchCount++;
          }
        }
      } catch (e) {
        Logger.log('Search folder warning: ' + e.toString());
      }
    });
    
    const keyLabel = (kategori === 'siswa') ? 'NISN.jpg' : 'NIP.jpg';
    return {
      success: true,
      message: 'Pencocokan massal foto ' + kategori.toUpperCase() + ' (' + keyLabel + ') selesai di seluruh folder & subfolder Drive. ' + matchCount + ' foto berhasil dipetakan.'
    };
  } catch (err) {
    return { success: false, message: 'Gagal melakukan pemetaan massal: ' + err.toString() };
  }
}

function apiImportPanitia(payload) {
  if (!payload || !payload.panitia || !payload.panitia.length) {
    return { success: false, message: 'Tidak ada data panitia yang diimpor.' };
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_PANITIA);
  let currentLastNum = getMaxNumberInCol(sheet, COL_PANITIA.ID_PANITIA, 'PAN-');
  const rowsToAdd = [];
  
  payload.panitia.forEach(function(p) {
    currentLastNum++;
    const idPan = 'PAN-' + padZero(currentLastNum, 3);
    rowsToAdd.push([
      idPan,
      p.nip || '',
      p.namaPanitia || 'Panitia',
      p.jabatan || 'Anggota',
      p.noHp || '',
      p.fotoUrl ? formatToLh3Url(p.fotoUrl) : DEFAULT_LOGO_URL
    ]);
  });
  
  if (rowsToAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 6).setValues(rowsToAdd);
  }
  return { success: true, message: rowsToAdd.length + ' data panitia berhasil diimpor ke sistem.' };
}

function apiImportPengawas(payload) {
  if (!payload || !payload.pengawas || !payload.pengawas.length) {
    return { success: false, message: 'Tidak ada data pengawas yang diimpor.' };
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DATA_PENGAWAS);
  let currentLastNum = getMaxNumberInCol(sheet, COL_PENGAWAS.ID_PENGAWAS, 'PGW-');
  const rowsToAdd = [];
  
  payload.pengawas.forEach(function(pg) {
    currentLastNum++;
    const idPgw = 'PGW-' + padZero(currentLastNum, 3);
    rowsToAdd.push([
      idPgw,
      pg.nip || '',
      pg.namaPengawas || 'Pengawas',
      pg.asalSekolah || '',
      pg.ruangTugas || '',
      pg.sesiTugas || '',
      pg.fotoUrl ? formatToLh3Url(pg.fotoUrl) : DEFAULT_LOGO_URL
    ]);
  });
  
  if (rowsToAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
  }
  return { success: true, message: rowsToAdd.length + ' data pengawas berhasil diimpor ke sistem.' };
}

function apiResetDatabase(payload) {
  if (!payload || payload.confirmation !== 'RESET') {
    return { success: false, message: 'Konfirmasi tidak sesuai. Harap ketik "RESET" dengan huruf besar.' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Kosongkan tab data transaksi
  const sheetsToClear = [SHEETS.DATA_SISWA, SHEETS.DATA_PANITIA, SHEETS.DATA_PENGAWAS, SHEETS.DISTRIBUSI_RUANG];
  sheetsToClear.forEach(function(sName) {
    const s = ss.getSheetByName(sName);
    if (s && s.getLastRow() > 1) {
      s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
    }
  });
  
  // Reset Pengaturan ke Default Pabrik
  const pSheet = ss.getSheetByName(SHEETS.PENGATURAN);
  if (pSheet && pSheet.getLastRow() > 1) {
    pSheet.getRange(2, 1, pSheet.getLastRow() - 1, pSheet.getLastColumn()).clearContent();
  }
  
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
  pSheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  
  return { success: true, message: 'Database Sipandu berhasil direset ke pengaturan pabrik!' };
}

// ==============================================================================
// 9. MODUL BACKUP & RESTORE DATABASE (JSON)
// ==============================================================================

function apiBackupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backup = {
    version: '1.0.0',
    timestamp: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'),
    appName: 'Sipandu',
    sheets: {}
  };
  
  for (const key in SHEETS) {
    const sName = SHEETS[key];
    const sheet = ss.getSheetByName(sName);
    if (sheet && sheet.getLastRow() > 0) {
      backup.sheets[sName] = sheet.getDataRange().getValues();
    }
  }
  
  const jsonStr = JSON.stringify(backup, null, 2);
  const fileName = 'backup_sipandu_' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmm') + '.json';
  
  return {
    success: true,
    jsonString: jsonStr,
    fileName: fileName
  };
}

function apiRestoreDatabase(payload) {
  if (!payload || !payload.jsonString) {
    return { success: false, message: 'Data cadangan JSON tidak boleh kosong.' };
  }
  
  try {
    const backup = JSON.parse(payload.jsonString);
    if (!backup.sheets) {
      return { success: false, message: 'Format berkas backup tidak valid.' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    for (const sName in backup.sheets) {
      const rows = backup.sheets[sName];
      if (rows && rows.length > 0) {
        let sheet = ss.getSheetByName(sName);
        if (!sheet) {
          sheet = ss.insertSheet(sName);
        } else {
          sheet.clearContents();
        }
        sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      }
    }
    
    return { success: true, message: 'Database Sipandu berhasil dipulihkan dari berkas backup.' };
  } catch (err) {
    return { success: false, message: 'Gagal memulihkan database: ' + err.toString() };
  }
}

// ==============================================================================
// 10. FUNGSI PEMBANTU (HELPER UTILITIES)
// ==============================================================================

/**
 * Mengubah tautan Google Drive apa saja menjadi tautan CDN lh3 direct view
 */
function formatToLh3Url(url) {
  if (!url) return DEFAULT_LOGO_URL;
  url = String(url).trim();
  
  // Jika sudah format base64 Data-URI, pertahankan
  if (url.indexOf('data:image') === 0) {
    return url;
  }
  
  // Jika sudah format lh3, kembalikan langsung
  if (url.indexOf('lh3.googleusercontent.com/d/') !== -1) {
    return url;
  }
  
  // Cari pola ID file drive
  let fileId = '';
  const match1 = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match1 && match1[1]) {
    fileId = match1[1];
  } else {
    const match2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match2 && match2[1]) {
      fileId = match2[1];
    } else if (url.length > 25 && url.indexOf('/') === -1 && url.indexOf('.') === -1) {
      fileId = url;
    }
  }
  
  if (fileId) {
    return 'https://lh3.googleusercontent.com/d/' + fileId;
  }
  
  return url;
}

/**
 * Menghasilkan kode unik sekuensial anti-tabrakan (PREFIX-XXXX)
 */
function generateNextId(sheet, colIndex, prefix, padLength) {
  const maxNum = getMaxNumberInCol(sheet, colIndex, prefix);
  return prefix + padZero(maxNum + 1, padLength);
}

function getMaxNumberInCol(sheet, colIndex, prefix) {
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  
  const values = sheet.getRange(2, colIndex + 1, sheet.getLastRow() - 1, 1).getValues();
  let max = 0;
  
  for (let i = 0; i < values.length; i++) {
    const val = String(values[i][0] || '');
    if (val.indexOf(prefix) === 0) {
      const numPart = parseInt(val.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > max) {
        max = numPart;
      }
    }
  }
  return max;
}

function padZero(num, size) {
  let s = String(num);
  while (s.length < size) s = '0' + s;
  return s;
}

function removeStudentFromDistribution(idSiswa) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DISTRIBUSI_RUANG);
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][COL_DISTRIBUSI.ID_SISWA]) === String(idSiswa)) {
      sheet.deleteRow(i + 1);
    }
  }
}
