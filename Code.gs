/**
 * ============================================================================
 * SIMPEL MEDIK - Sistem Pengaduan Komite Medik Rumah Sakit
 * Platform: Google Apps Script + Google Spreadsheet + Google Drive
 * Version: 1.0.0
 * Backend Database Engine & System Core (Code.gs)
 * ============================================================================
 */

const APP_CONFIG = {
  APP_NAME: 'SIMPEL MEDIK',
  APP_VERSION: '1.0.0',
  HOSPITAL_NAME: 'Rumah Sakit Umum Daerah',
  HOSPITAL_CODE: 'RSUD-01',
  TIMEZONE: 'Asia/Makassar',
  MAX_FILE_SIZE_MB: 10,
  SESSION_DURATION_HOURS: 8,
  DRIVE_ROOT_FOLDER_NAME: 'SIMPEL_MEDIK_STORAGE'
};

const DB_SCHEMA = {
  CONFIG: [
    'config_id', 'config_key', 'config_value', 'description', 'updated_at', 'updated_by'
  ],
  USERS: [
    'user_id', 'username', 'password_hash', 'full_name', 'nip', 'email', 'phone', 
    'unit_id', 'role_id', 'status', 'last_login', 'created_at', 'created_by', 'updated_at', 'updated_by'
  ],
  ROLES: [
    'role_id', 'role_name', 'description', 'status', 'created_at', 'updated_at'
  ],
  PERMISSIONS: [
    'permission_id', 'permission_code', 'permission_name', 'module', 'description'
  ],
  ROLE_PERMISSIONS: [
    'role_permission_id', 'role_id', 'permission_id', 'status', 'created_at', 'created_by'
  ],
  UNITS: [
    'unit_id', 'unit_code', 'unit_name', 'unit_type', 'head_name', 'status', 'created_at', 'updated_at'
  ],
  COMPLAINT_CATEGORIES: [
    'category_id', 'category_code', 'category_name', 'description', 'status', 'created_at', 'updated_at'
  ],
  COMPLAINT_SUBCATEGORIES: [
    'subcategory_id', 'category_id', 'subcategory_name', 'description', 'status', 'created_at', 'updated_at'
  ],
  PRIORITIES: [
    'priority_id', 'priority_code', 'priority_name', 'level', 'description', 'status'
  ],
  STATUSES: [
    'status_id', 'status_code', 'status_name', 'sequence', 'description', 'is_final', 'status'
  ],
  COMPLAINTS: [
    'complaint_id', 'complaint_number', 'reporter_user_id', 'reporter_name', 'reporter_unit_id',
    'category_id', 'subcategory_id', 'priority_id', 'incident_date', 'incident_time',
    'incident_location', 'related_party', 'description', 'initial_assessment', 'status',
    'assigned_to', 'submission_date', 'verified_at', 'review_started_at', 'recommendation_at',
    'followup_started_at', 'completed_at', 'closed_at', 'created_at', 'created_by', 'updated_at', 'updated_by'
  ],
  COMPLAINT_DETAILS: [
    'detail_id', 'complaint_id', 'detail_type', 'detail_title', 'detail_content', 'sequence', 'created_at', 'created_by', 'updated_at', 'updated_by'
  ],
  ATTACHMENTS: [
    'attachment_id', 'complaint_id', 'file_name', 'file_type', 'file_size', 'drive_file_id', 'drive_url', 'folder_id', 'uploaded_by', 'uploaded_at', 'status'
  ],
  VERIFICATIONS: [
    'verification_id', 'complaint_id', 'verification_status', 'verification_notes', 'verified_by', 'verified_at', 'requested_information', 'due_date'
  ],
  REVIEWS: [
    'review_id', 'complaint_id', 'reviewer_user_id', 'review_date', 'findings', 'analysis', 'contributing_factors', 'review_result', 'review_notes', 'document_file_id', 'created_at', 'updated_at'
  ],
  RECOMMENDATIONS: [
    'recommendation_id', 'complaint_id', 'recommendation_type', 'recommendation_text', 'target_unit_id', 'responsible_person', 'priority', 'due_date', 'status', 'created_by', 'created_at', 'updated_at', 'updated_by'
  ],
  FOLLOW_UPS: [
    'followup_id', 'complaint_id', 'recommendation_id', 'followup_description', 'responsible_unit_id', 'responsible_person', 'start_date', 'target_date', 'completion_date', 'progress_percentage', 'result', 'evidence_file_id', 'status', 'created_at', 'created_by', 'updated_at', 'updated_by'
  ],
  TIMELINE: [
    'timeline_id', 'complaint_id', 'event_type', 'event_title', 'event_description', 'old_status', 'new_status', 'actor_user_id', 'actor_name', 'event_date', 'created_at'
  ],
  NOTIFICATIONS: [
    'notification_id', 'user_id', 'complaint_id', 'notification_type', 'title', 'message', 'is_read', 'created_at', 'read_at'
  ],
  AUDIT_LOG: [
    'audit_id', 'timestamp', 'user_id', 'username', 'user_name', 'action', 'module', 'record_id', 'description', 'old_value', 'new_value', 'ip_address', 'user_agent'
  ]
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_CONFIG.APP_NAME + ' - Sistem Pengaduan Komite Medik')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSystem() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const nowStr = formatISOString(new Date());
    const report = {
      sheetsCreated: [],
      sheetsVerified: [],
      driveFoldersCreated: [],
      masterDataSeeded: false,
      adminCreated: false,
      status: 'SUCCESS'
    };

    for (const sheetName in DB_SCHEMA) {
      if (!DB_SCHEMA.hasOwnProperty(sheetName)) continue;
      
      let sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        report.sheetsCreated.push(sheetName);
      } else {
        report.sheetsVerified.push(sheetName);
      }
      
      const expectedHeaders = DB_SCHEMA[sheetName];
      const existingRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
      const existingHeaders = existingRange.getValues()[0];
      
      let headersNeedUpdate = false;
      for (let i = 0; i < expectedHeaders.length; i++) {
        if (existingHeaders[i] !== expectedHeaders[i]) {
          headersNeedUpdate = true;
          break;
        }
      }

      if (headersNeedUpdate) {
        sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
        const headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#1E3A8A');
        headerRange.setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
      }
    }

    const driveFolderIds = setupDriveFolders();
    report.driveFoldersCreated = Object.keys(driveFolderIds);

    seedConfigData(nowStr, driveFolderIds);
    seedMasterData(nowStr);
    report.masterDataSeeded = true;
    report.adminCreated = seedInitialAdmin(nowStr);

    logAuditInternal('SYSTEM_SETUP', 'SETUP', 'SYS_01', 'System initialization completed successfully');

    return {
      success: true,
      message: 'SIMPEL MEDIK Database Engine & System Setup Completed Successfully!',
      data: report
    };

  } catch (error) {
    Logger.log('Error in setupSystem: ' + error.stack);
    return {
      success: false,
      message: 'Setup System failed: ' + error.message,
      error: error.toString()
    };
  } finally {
    lock.releaseLock();
  }
}

function setupDriveFolders() {
  const rootName = APP_CONFIG.DRIVE_ROOT_FOLDER_NAME;
  let rootFolder;
  
  const rootIterator = DriveApp.getFoldersByName(rootName);
  if (rootIterator.hasNext()) {
    rootFolder = rootIterator.next();
  } else {
    rootFolder = DriveApp.createFolder(rootName);
  }

  const subFolders = ['COMPLAINTS', 'REPORTS', 'SYSTEM'];
  const folderIds = {
    DRIVE_ROOT_FOLDER_ID: rootFolder.getId()
  };

  subFolders.forEach(folderName => {
    let subFolder;
    const subIterator = rootFolder.getFoldersByName(folderName);
    if (subIterator.hasNext()) {
      subFolder = subIterator.next();
    } else {
      subFolder = rootFolder.createFolder(folderName);
    }
    folderIds[`DRIVE_${folderName}_FOLDER_ID`] = subFolder.getId();
  });

  return folderIds;
}

function seedConfigData(nowStr, driveFolderIds) {
  const existingRecords = getRecords('CONFIG');
  const existingKeys = existingRecords.map(r => r.config_key);

  const defaultConfig = [
    { config_key: 'APP_NAME', config_value: APP_CONFIG.APP_NAME, description: 'Nama Aplikasi' },
    { config_key: 'APP_VERSION', config_value: APP_CONFIG.APP_VERSION, description: 'Versi Aplikasi' },
    { config_key: 'HOSPITAL_NAME', config_value: APP_CONFIG.HOSPITAL_NAME, description: 'Nama Rumah Sakit' },
    { config_key: 'HOSPITAL_CODE', config_value: APP_CONFIG.HOSPITAL_CODE, description: 'Kode RS' },
    { config_key: 'TIMEZONE', config_value: APP_CONFIG.TIMEZONE, description: 'Zona Waktu' },
    { config_key: 'MAX_FILE_SIZE_MB', config_value: String(APP_CONFIG.MAX_FILE_SIZE_MB), description: 'Batas Ukuran Upload File (MB)' },
    { config_key: 'SESSION_DURATION', config_value: String(APP_CONFIG.SESSION_DURATION_HOURS), description: 'Durasi Session dalam Jam' },
    { config_key: 'DRIVE_ROOT_FOLDER_ID', config_value: driveFolderIds.DRIVE_ROOT_FOLDER_ID, description: 'Google Drive Root Folder ID' },
    { config_key: 'DRIVE_COMPLAINTS_FOLDER_ID', config_value: driveFolderIds.DRIVE_COMPLAINTS_FOLDER_ID, description: 'Google Drive Complaints Folder ID' },
    { config_key: 'DRIVE_REPORTS_FOLDER_ID', config_value: driveFolderIds.DRIVE_REPORTS_FOLDER_ID, description: 'Google Drive Reports Folder ID' },
    { config_key: 'DRIVE_SYSTEM_FOLDER_ID', config_value: driveFolderIds.DRIVE_SYSTEM_FOLDER_ID, description: 'Google Drive System Folder ID' }
  ];

  const rowsToAppend = [];
  defaultConfig.forEach(item => {
    if (!existingKeys.includes(item.config_key)) {
      rowsToAppend.push([
        generateId('CFG'),
        item.config_key,
        item.config_value,
        item.description,
        nowStr,
        'SYSTEM'
      ]);
    }
  });

  if (rowsToAppend.length > 0) {
    appendBatchRows('CONFIG', rowsToAppend);
  }
}

function seedMasterData(nowStr) {
  const rolesData = [
    { id: 'ROL-ADMIN', name: 'ADMIN', desc: 'Administrator System & Database' },
    { id: 'ROL-PELAPOR', name: 'PELAPOR', desc: 'Tenaga Medis / Pegawai Pelapor' },
    { id: 'ROL-SEKRETARIAT', name: 'SEKRETARIAT', desc: 'Sekretariat Komite Medik' },
    { id: 'ROL-ANGGOTA', name: 'ANGGOTA_KOMITE_MEDIK', desc: 'Anggota Sub-Komite Medik' },
    { id: 'ROL-KETUA', name: 'KETUA_KOMITE_MEDIK', desc: 'Ketua Komite Medik' }
  ];
  seedTableIfEmpty('ROLES', 'role_id', rolesData.map(r => [
    r.id, r.name, r.desc, 'ACTIVE', nowStr, nowStr
  ]));

  const permissionsData = [
    ['PRM-001', 'DASHBOARD_VIEW', 'Melihat Dashboard Statistics', 'DASHBOARD', 'Akses dashboard utama'],
    ['PRM-002', 'COMPLAINT_CREATE', 'Buat Pengaduan Baru', 'COMPLAINT', 'Membuat laporan pengaduan'],
    ['PRM-003', 'COMPLAINT_VIEW', 'Melihat Detail Pengaduan', 'COMPLAINT', 'Melihat detail pengaduan'],
    ['PRM-004', 'COMPLAINT_EDIT', 'Mengubah Data Pengaduan', 'COMPLAINT', 'Mengedit draft pengaduan'],
    ['PRM-005', 'COMPLAINT_VERIFY', 'Melakukan Verifikasi Pengaduan', 'VERIFICATION', 'Verifikasi awal pengaduan'],
    ['PRM-006', 'REVIEW_CREATE', 'Melakukan Telaah Medik', 'REVIEW', 'Membuat lembar telaah komite'],
    ['PRM-007', 'REVIEW_VIEW', 'Melihat Hasil Telaah', 'REVIEW', 'Melihat telaah medik'],
    ['PRM-008', 'RECOMMENDATION_CREATE', 'Membuat Rekomendasi', 'RECOMMENDATION', 'Menerbitkan rekomendasi komite'],
    ['PRM-009', 'FOLLOWUP_CREATE', 'Membuat & Update Tindak Lanjut', 'FOLLOW_UP', 'Mengelola tindak lanjut'],
    ['PRM-010', 'REPORT_VIEW', 'Melihat Laporan', 'REPORT', 'Akses modul laporan'],
    ['PRM-011', 'REPORT_EXPORT', 'Export Laporan Excel/PDF', 'REPORT', 'Export data laporan'],
    ['PRM-012', 'USER_MANAGE', 'Kelola User Aplikasi', 'USER', 'Manajemen akun pengguna'],
    ['PRM-013', 'MASTER_MANAGE', 'Kelola Master Data', 'MASTER', 'Manajemen data referensi'],
    ['PRM-014', 'AUDIT_VIEW', 'Melihat Log Audit', 'AUDIT', 'Melihat jejak audit sistem']
  ];
  seedTableIfEmpty('PERMISSIONS', 'permission_id', permissionsData);

  const unitsData = [
    ['UNT-IGD', 'IGD', 'Instalasi Gawat Darurat', 'PELAYANAN', 'dr. Head IGD', 'ACTIVE', nowStr, nowStr],
    ['UNT-IRJ', 'RAWAT_JALAN', 'Instalasi Rawat Jalan', 'PELAYANAN', 'dr. Head Rajal', 'ACTIVE', nowStr, nowStr],
    ['UNT-IRIN', 'RAWAT_INAP', 'Instalasi Rawat Inap', 'PELAYANAN', 'dr. Head Ranap', 'ACTIVE', nowStr, nowStr],
    ['UNT-ICU', 'ICU', 'Intensive Care Unit', 'PELAYANAN', 'dr. Head ICU', 'ACTIVE', nowStr, nowStr],
    ['UNT-OK', 'KAMAR_OPERASI', 'Kamar Operasi (OK)', 'PELAYANAN', 'dr. Head OK', 'ACTIVE', nowStr, nowStr],
    ['UNT-RAD', 'RADIOLOGI', 'Instalasi Radiologi', 'PENUNJANG', 'dr. Head Rad', 'ACTIVE', nowStr, nowStr],
    ['UNT-LAB', 'LABORATORIUM', 'Instalasi Laboratorium', 'PENUNJANG', 'dr. Head Lab', 'ACTIVE', nowStr, nowStr],
    ['UNT-FAR', 'FARMASI', 'Instalasi Farmasi', 'PENUNJANG', 'Apt. Head Farmasi', 'ACTIVE', nowStr, nowStr],
    ['UNT-KOM', 'KOMITE_MEDIK', 'Komite Medik', 'MANAJEMEN', 'Ketua Komite Medik', 'ACTIVE', nowStr, nowStr]
  ];
  seedTableIfEmpty('UNITS', 'unit_id', unitsData);

  const categoriesData = [
    ['CAT-ETIK', 'ETIK', 'Etika Profesi Kedokteran', 'Pelanggaran kode etik kedokteran', 'ACTIVE', nowStr, nowStr],
    ['CAT-DISIPLIN', 'DISIPLIN', 'Disiplin Profesional', 'Pelanggaran disiplin tugas dan kewajiban', 'ACTIVE', nowStr, nowStr],
    ['CAT-PROF', 'PROFESIONALISME', 'Mutu & Standard Profesi', 'Kompetensi dan standar pelayanan medis', 'ACTIVE', nowStr, nowStr],
    ['CAT-MUTU', 'MUTU_PELAYANAN', 'Mutu Pelayanan Medis', 'Kualitas pelayanan klinis kepada pasien', 'ACTIVE', nowStr, nowStr],
    ['CAT-KPRS', 'KESELAMATAN_PASIEN', 'Keselamatan Pasien (KPRS)', 'Insiden keselamatan pasien / KNC / KTD', 'ACTIVE', nowStr, nowStr],
    ['CAT-KOM', 'KOMUNIKASI', 'Komunikasi Efektif', 'Komunikasi dokter, pasien, dan antar staf', 'ACTIVE', nowStr, nowStr],
    ['CAT-ADM', 'ADMINISTRASI_MEDIS', 'Administrasi Medis', 'Kelengkapan rekam medis & administrasi', 'ACTIVE', nowStr, nowStr],
    ['CAT-LAIN', 'LAINNYA', 'Kategori Lain-lain', 'Pengaduan di luar kategori utama', 'ACTIVE', nowStr, nowStr]
  ];
  seedTableIfEmpty('COMPLAINT_CATEGORIES', 'category_id', categoriesData);

  const prioritiesData = [
    ['PRI-LOW', 'LOW', 'Rendah', 1, 'Penanganan standar (SLA 14 hari)', 'ACTIVE'],
    ['PRI-MED', 'MEDIUM', 'Sedang', 2, 'Penanganan prioritas sedang (SLA 10 hari)', 'ACTIVE'],
    ['PRI-HIGH', 'HIGH', 'Tinggi', 3, 'Penanganan cepat (SLA 5 hari)', 'ACTIVE'],
    ['PRI-CRIT', 'CRITICAL', 'Kritis', 4, 'Penanganan darurat segera (SLA 2 hari)', 'ACTIVE']
  ];
  seedTableIfEmpty('PRIORITIES', 'priority_id', prioritiesData);

  const statusesData = [
    ['ST-01', 'DRAFT', 'Draft', 1, 'Pengaduan belum dikirim', 'NO', 'ACTIVE'],
    ['ST-02', 'SUBMITTED', 'Terkirim', 2, 'Pengaduan telah dikirim oleh pelapor', 'NO', 'ACTIVE'],
    ['ST-03', 'VERIFICATION', 'Verifikasi Sekretariat', 3, 'Dalam verifikasi kelengkapan dokumen', 'NO', 'ACTIVE'],
    ['ST-04', 'NEED_INFORMATION', 'Butuh Kelengkapan', 4, 'Membutuhkan informasi tambahan dari pelapor', 'NO', 'ACTIVE'],
    ['ST-05', 'ACCEPTED', 'Diterima', 5, 'Pengaduan diverifikasi dan diterima', 'NO', 'ACTIVE'],
    ['ST-06', 'UNDER_REVIEW', 'Dalam Telaah Medik', 6, 'Proses telaah oleh Sub-Komite Medik', 'NO', 'ACTIVE'],
    ['ST-07', 'RECOMMENDATION', 'Rekomendasi Diterbitkan', 7, 'Rekomendasi Komite Medik terbit', 'NO', 'ACTIVE'],
    ['ST-08', 'FOLLOW_UP', 'Dalam Tindak Lanjut', 8, 'Proses tindak lanjut oleh unit terkait', 'NO', 'ACTIVE'],
    ['ST-09', 'COMPLETED', 'Selesai Dilaksanakan', 9, 'Tindak lanjut selesai dikerjakan', 'NO', 'ACTIVE'],
    ['ST-10', 'CLOSED', 'Ditutup', 10, 'Pengaduan resmi ditutup', 'YES', 'ACTIVE'],
    ['ST-11', 'REJECTED', 'Ditolak', 11, 'Pengaduan ditolak saat verifikasi', 'YES', 'ACTIVE']
  ];
  seedTableIfEmpty('STATUSES', 'status_id', statusesData);
}

function seedInitialAdmin(nowStr) {
  const users = getRecords('USERS');
  const adminUsername = 'admin';
  const existingAdmin = users.find(u => u.username === adminUsername);

  if (!existingAdmin) {
    const adminPassHash = hashPassword('admin123');
    const adminRecord = [
      'USR-ADMIN-01',
      adminUsername,
      adminPassHash,
      'Administrator Utama Komite Medik',
      '198501012010011001',
      'admin.komitemedik@hospital.go.id',
      '081122334455',
      'UNT-KOM',
      'ROL-ADMIN',
      'ACTIVE',
      '',
      nowStr,
      'SYSTEM',
      nowStr,
      'SYSTEM'
    ];
    appendBatchRows('USERS', [adminRecord]);
    return true;
  }
  return false;
}

function getSheet(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') {
    throw new Error('Nama sheet tidak boleh kosong atau undefined.');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" tidak ditemukan pada database Spreadsheet. Silakan jalankan fungsi setupSystem() terlebih dahulu pada editor Apps Script.');
  }
  return sheet;
}

function getHeaders(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') return [];
  if (DB_SCHEMA[sheetName]) return DB_SCHEMA[sheetName];
  try {
    const sheet = getSheet(sheetName);
    return sheet ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  } catch (e) {
    return [];
  }
}

function getRecords(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') return [];
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return data.map(row => {
    const record = {};
    headers.forEach((header, index) => {
      let val = row[index];
      // Convert native Apps Script Date objects to ISO strings to prevent serialization null returns
      if (val instanceof Date) {
        val = formatISOString(val);
      } else if (val === null || val === undefined) {
        val = '';
      }
      record[header] = val;
    });
    return record;
  });
}

function formatISOString(dateObj) {
  if (!dateObj) return '';
  if (typeof dateObj === 'string') return dateObj;
  if (!(dateObj instanceof Date)) {
    dateObj = new Date(dateObj);
  }
  if (isNaN(dateObj.getTime())) return '';
  return Utilities.formatDate(dateObj, APP_CONFIG.TIMEZONE || 'Asia/Makassar', "yyyy-MM-dd'T'HH:mm:ss");
}

function findRecordById(sheetName, idColumn, idValue) {
  if (!sheetName || !idColumn || idValue === undefined) return null;
  const records = getRecords(sheetName);
  return records.find(r => String(r[idColumn]) === String(idValue)) || null;
}

function appendRecord(sheetName, recordObject) {
  if (!sheetName || !recordObject) return false;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet(sheetName);
    const headers = getHeaders(sheetName);
    const row = headers.map(h => recordObject[h] !== undefined ? recordObject[h] : '');
    sheet.appendRow(row);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function appendBatchRows(sheetName, rowsArray) {
  if (!sheetName || !rowsArray || rowsArray.length === 0) return;
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  const numRows = rowsArray.length;
  const numCols = rowsArray[0].length;
  sheet.getRange(lastRow + 1, 1, numRows, numCols).setValues(rowsArray);
}

function updateRecord(sheetName, idColumn, idValue, updateDataObject) {
  if (!sheetName || !idColumn || idValue === undefined) return false;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow <= 1) return false;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const idColIndex = headers.indexOf(idColumn);

    if (idColIndex === -1) {
      throw new Error('Kolom ' + idColumn + ' tidak ditemukan pada sheet ' + sheetName);
    }

    const idValues = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < idValues.length; i++) {
      if (String(idValues[i][0]) === String(idValue)) {
        const rowIndex = i + 2;
        const currentRowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

        headers.forEach((header, colIdx) => {
          if (updateDataObject[header] !== undefined) {
            currentRowValues[colIdx] = updateDataObject[header];
          }
        });

        sheet.getRange(rowIndex, 1, 1, lastCol).setValues([currentRowValues]);
        return true;
      }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord(sheetName, idColumn, idValue) {
  if (!sheetName || !idColumn || idValue === undefined) return false;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idColIndex = headers.indexOf(idColumn);

    if (idColIndex === -1) return false;

    const idValues = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();

    for (let i = 0; i < idValues.length; i++) {
      if (String(idValues[i][0]) === String(idValue)) {
        sheet.deleteRow(i + 2);
        return true;
      }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function seedTableIfEmpty(sheetName, idColumn, dataRows) {
  if (!sheetName) return;
  const records = getRecords(sheetName);
  if (records.length === 0 && dataRows.length > 0) {
    appendBatchRows(sheetName, dataRows);
  }
}

function generateComplaintNumber() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const year = new Date().getFullYear();
    const complaints = getRecords('COMPLAINTS');
    const prefix = 'PM-' + year + '-';
    
    let maxSeq = 0;
    complaints.forEach(c => {
      if (c.complaint_number && String(c.complaint_number).startsWith(prefix)) {
        const seqStr = String(c.complaint_number).replace(prefix, '');
        const seq = parseInt(seqStr, 10);
        if (!isNaN(seq) && seq > maxSeq) {
          maxSeq = seq;
        }
      }
    });

    const nextSeq = maxSeq + 1;
    const formattedSeq = ('00000' + nextSeq).slice(-5);
    return prefix + formattedSeq;
  } finally {
    lock.releaseLock();
  }
}

function generateId(prefix) {
  const cleanPrefix = prefix ? prefix.toUpperCase() + '-' : 'ID-';
  const randomHex = Math.random().toString(36).substring(2, 8).toUpperCase();
  const timeHex = new Date().getTime().toString(36).toUpperCase().slice(-4);
  return cleanPrefix + timeHex + randomHex;
}

function hashPassword(password) {
  if (!password) return '';
  const rawBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, 
    password, 
    Utilities.Charset.UTF_8
  );
  let hexString = '';
  for (let i = 0; i < rawBytes.length; i++) {
    let byteVal = rawBytes[i];
    if (byteVal < 0) byteVal += 256;
    let byteHex = byteVal.toString(16);
    if (byteHex.length === 1) byteHex = '0' + byteHex;
    hexString += byteHex;
  }
  return hexString;
}

function login(username, password) {
  try {
    if (!username || !password) {
      return createErrorResponse('Username dan password wajib diisi.', 'AUTH_REQUIRED');
    }

    const users = getRecords('USERS');
    const inputHash = hashPassword(password);
    const user = users.find(u => u.username === username);

    if (!user) {
      logAuditInternal('LOGIN_FAILED', 'AUTH', username, 'User tidak ditemukan');
      return createErrorResponse('Username atau password tidak valid.', 'AUTH_INVALID');
    }

    if (user.status !== 'ACTIVE') {
      logAuditInternal('LOGIN_BLOCKED', 'AUTH', user.user_id, 'Akun nonaktif');
      return createErrorResponse('Akun Anda telah dinonaktifkan. Silakan hubungi admin.', 'AUTH_INACTIVE');
    }

    if (user.password_hash !== inputHash) {
      logAuditInternal('LOGIN_FAILED', 'AUTH', user.user_id, 'Password salah');
      return createErrorResponse('Username atau password tidak valid.', 'AUTH_INVALID');
    }

    const token = generateId('SES');
    const sessionData = {
      user_id: user.user_id,
      username: user.username,
      full_name: user.full_name,
      role_id: user.role_id,
      unit_id: user.unit_id,
      login_at: formatISOString(new Date())
    };

    const cache = CacheService.getScriptCache();
    cache.put('SESSION_' + token, JSON.stringify(sessionData), APP_CONFIG.SESSION_DURATION_HOURS * 3600);

    updateRecord('USERS', 'user_id', user.user_id, {
      last_login: formatISOString(new Date())
    });

    logAuditInternal('LOGIN_SUCCESS', 'AUTH', user.user_id, 'Login berhasil', null, null, sessionData);

    return createSuccessResponse('Login berhasil', {
      sessionToken: token,
      user: sessionData
    });

  } catch (err) {
    return createErrorResponse('Gagal melakukan login: ' + err.message, 'AUTH_ERROR');
  }
}

function validateSession(sessionToken) {
  if (!sessionToken) return createErrorResponse('Session token tidak ditemukan', 'NO_SESSION');
  
  const cache = CacheService.getScriptCache();
  const cached = cache.get('SESSION_' + sessionToken);
  
  if (!cached) {
    return createErrorResponse('Session telah kadaluarsa. Silakan login kembali.', 'SESSION_EXPIRED');
  }

  const sessionData = JSON.parse(cached);
  const user = findRecordById('USERS', 'user_id', sessionData.user_id);
  
  if (!user || user.status !== 'ACTIVE') {
    return createErrorResponse('Akun pengguna tidak aktif', 'USER_INACTIVE');
  }

  return createSuccessResponse('Session valid', { user: sessionData });
}

function logout(sessionToken) {
  if (sessionToken) {
    const cache = CacheService.getScriptCache();
    cache.remove('SESSION_' + sessionToken);
  }
  return createSuccessResponse('Logout berhasil');
}

function logAuditInternal(action, moduleName, recordId, description, oldValue, newValue, userContext) {
  try {
    const user = userContext || { user_id: 'SYSTEM', username: 'SYSTEM', full_name: 'System Process' };
    const auditRecord = {
      audit_id: generateId('AUD'),
      timestamp: formatISOString(new Date()),
      user_id: user.user_id || 'SYSTEM',
      username: user.username || 'SYSTEM',
      user_name: user.full_name || 'System',
      action: action,
      module: moduleName,
      record_id: recordId || '',
      description: description || '',
      old_value: oldValue ? (typeof oldValue === 'object' ? JSON.stringify(oldValue) : String(oldValue)) : '',
      new_value: newValue ? (typeof newValue === 'object' ? JSON.stringify(newValue) : String(newValue)) : '',
      ip_address: '127.0.0.1',
      user_agent: 'Google Apps Script Environment'
    };
    appendRecord('AUDIT_LOG', auditRecord);
  } catch (err) {
    Logger.log('Failed to write audit log: ' + err.toString());
  }
}

function createComplaint(sessionToken, complaintData) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;
  const currentUser = auth.data.user;

  try {
    const complaintNumber = generateComplaintNumber();
    const complaintId = generateId('CMP');
    const nowStr = formatISOString(new Date());

    const newComplaint = {
      complaint_id: complaintId,
      complaint_number: complaintNumber,
      reporter_user_id: currentUser.user_id,
      reporter_name: currentUser.full_name,
      reporter_unit_id: currentUser.unit_id,
      category_id: complaintData.category_id || 'LAINNYA',
      subcategory_id: complaintData.subcategory_id || '',
      priority_id: complaintData.priority_id || 'MEDIUM',
      incident_date: complaintData.incident_date || '',
      incident_time: complaintData.incident_time || '',
      incident_location: complaintData.incident_location || '',
      related_party: complaintData.related_party || '',
      description: complaintData.description || '',
      initial_assessment: '',
      status: 'SUBMITTED',
      assigned_to: '',
      submission_date: nowStr,
      verified_at: '',
      review_started_at: '',
      recommendation_at: '',
      followup_started_at: '',
      completed_at: '',
      closed_at: '',
      created_at: nowStr,
      created_by: currentUser.user_id,
      updated_at: nowStr,
      updated_by: currentUser.user_id
    };

    appendRecord('COMPLAINTS', newComplaint);

    const timelineRecord = {
      timeline_id: generateId('TML'),
      complaint_id: complaintId,
      event_type: 'COMPLAINT_CREATED',
      event_title: 'Pengaduan Dibuat',
      event_description: 'Pengaduan diajukan ke sistem dengan nomor ' + complaintNumber,
      old_status: '',
      new_status: 'SUBMITTED',
      actor_user_id: currentUser.user_id,
      actor_name: currentUser.full_name,
      event_date: nowStr,
      created_at: nowStr
    };
    appendRecord('TIMELINE', timelineRecord);

    logAuditInternal('COMPLAINT_CREATE', 'COMPLAINT', complaintId, 'Pengaduan baru dibuat ' + complaintNumber, null, newComplaint, currentUser);

    return createSuccessResponse('Pengaduan berhasil diajukan.', { complaint: newComplaint });

  } catch (err) {
    return createErrorResponse('Gagal menyimpan pengaduan: ' + err.message, 'COMPLAINT_CREATE_FAILED');
  }
}

function createPublicComplaint(complaintData) {
  try {
    const complaintNumber = generateComplaintNumber();
    const complaintId = generateId('CMP');
    const nowStr = formatISOString(new Date());

    const newComplaint = {
      complaint_id: complaintId,
      complaint_number: complaintNumber,
      reporter_user_id: 'PUBLIC_GUEST',
      reporter_name: complaintData.reporter_name || 'Anonim Publik',
      reporter_unit_id: 'PUBLIK',
      category_id: complaintData.category_id || 'LAINNYA',
      subcategory_id: complaintData.subcategory_id || '',
      priority_id: 'MEDIUM',
      incident_date: complaintData.incident_date || '',
      incident_time: complaintData.incident_time || '',
      incident_location: complaintData.incident_location || '',
      related_party: complaintData.related_party || '',
      description: complaintData.description || '',
      initial_assessment: '',
      status: 'SUBMITTED',
      assigned_to: '',
      submission_date: nowStr,
      verified_at: '',
      review_started_at: '',
      recommendation_at: '',
      followup_started_at: '',
      completed_at: '',
      closed_at: '',
      created_at: nowStr,
      created_by: 'PUBLIC',
      updated_at: nowStr,
      updated_by: 'PUBLIC'
    };

    appendRecord('COMPLAINTS', newComplaint);

    const timelineRecord = {
      timeline_id: generateId('TML'),
      complaint_id: complaintId,
      event_type: 'PUBLIC_SUBMISSION',
      event_title: 'Pengaduan Publik Terkirim',
      event_description: 'Pengaduan diajukan oleh publik/guest dengan nomor ' + complaintNumber,
      old_status: '',
      new_status: 'SUBMITTED',
      actor_user_id: 'PUBLIC',
      actor_name: newComplaint.reporter_name,
      event_date: nowStr,
      created_at: nowStr
    };
    appendRecord('TIMELINE', timelineRecord);

    logAuditInternal('PUBLIC_COMPLAINT_CREATE', 'COMPLAINT', complaintId, 'Pengaduan publik dibuat ' + complaintNumber);

    return createSuccessResponse('Pengaduan publik berhasil dikirim.', { complaint: newComplaint });

  } catch (err) {
    return createErrorResponse('Gagal mengirim pengaduan publik: ' + err.message, 'PUBLIC_COMPLAINT_FAILED');
  }
}

function getComplaints(sessionToken, options) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;
  const currentUser = auth.data.user;

  try {
    options = options || {};
    let complaints = getRecords('COMPLAINTS');

    if (currentUser.role_id === 'ROL-PELAPOR') {
      complaints = complaints.filter(c => c.reporter_user_id === currentUser.user_id);
    } else if (currentUser.role_id === 'ROL-ANGGOTA') {
      complaints = complaints.filter(c => c.assigned_to === currentUser.user_id || c.reporter_user_id === currentUser.user_id);
    }

    if (options.status) {
      complaints = complaints.filter(c => c.status === options.status);
    }

    if (options.search) {
      const q = String(options.search).toLowerCase();
      complaints = complaints.filter(c => 
        String(c.complaint_number || '').toLowerCase().includes(q) ||
        String(c.reporter_name || '').toLowerCase().includes(q) ||
        String(c.description || '').toLowerCase().includes(q)
      );
    }

    complaints.sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));

    return createSuccessResponse('Data pengaduan berhasil dimuat', complaints);
  } catch (err) {
    return createErrorResponse('Gagal mengambil daftar pengaduan: ' + err.message);
  }
}

function getComplaintById(sessionToken, complaintId) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;

  try {
    const complaint = findRecordById('COMPLAINTS', 'complaint_id', complaintId);
    if (!complaint) return createErrorResponse('Pengaduan tidak ditemukan', 'NOT_FOUND');

    const timeline = getRecords('TIMELINE').filter(t => t.complaint_id === complaintId);
    timeline.sort((a, b) => new Date(a.event_date) - new Date(b.event_date));

    const attachments = getRecords('ATTACHMENTS').filter(a => a.complaint_id === complaintId);

    return createSuccessResponse('Detail pengaduan ditemukan', {
      complaint: complaint,
      timeline: timeline,
      attachments: attachments
    });
  } catch (err) {
    return createErrorResponse('Gagal mengambil detail pengaduan: ' + err.message);
  }
}

function getPublicComplaintStatus(complaintNumber) {
  try {
    if (!complaintNumber) return createErrorResponse('Nomor pengaduan wajib diisi.');
    const complaints = getRecords('COMPLAINTS');
    const complaint = complaints.find(c => String(c.complaint_number).toUpperCase() === String(complaintNumber).trim().toUpperCase());

    if (!complaint) return createErrorResponse('Nomor pengaduan tidak ditemukan.', 'NOT_FOUND');

    return createSuccessResponse('Status pengaduan ditemukan', {
      complaint_number: complaint.complaint_number,
      status: complaint.status,
      created_at: complaint.created_at,
      category_id: complaint.category_id
    });
  } catch (err) {
    return createErrorResponse('Gagal memproses pencarian status: ' + err.message);
  }
}

function verifyComplaint(sessionToken, complaintId, verificationData) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;
  const currentUser = auth.data.user;

  try {
    const complaint = findRecordById('COMPLAINTS', 'complaint_id', complaintId);
    if (!complaint) return createErrorResponse('Pengaduan tidak ditemukan');

    const nowStr = formatISOString(new Date());
    const vStatus = verificationData.verification_status || 'ACCEPTED';

    const vRecord = {
      verification_id: generateId('VRF'),
      complaint_id: complaintId,
      verification_status: vStatus,
      verification_notes: verificationData.verification_notes || '',
      verified_by: currentUser.user_id,
      verified_at: nowStr,
      requested_information: verificationData.requested_information || '',
      due_date: verificationData.due_date || ''
    };
    appendRecord('VERIFICATIONS', vRecord);

    const newComplaintStatus = vStatus === 'ACCEPTED' ? 'ACCEPTED' : (vStatus === 'REJECTED' ? 'REJECTED' : 'NEED_INFORMATION');
    updateRecord('COMPLAINTS', 'complaint_id', complaintId, {
      status: newComplaintStatus,
      verified_at: nowStr,
      updated_at: nowStr,
      updated_by: currentUser.user_id
    });

    appendRecord('TIMELINE', {
      timeline_id: generateId('TML'),
      complaint_id: complaintId,
      event_type: 'VERIFICATION',
      event_title: 'Verifikasi Sekretariat',
      event_description: 'Status verifikasi: ' + vStatus + '. Catatan: ' + (verificationData.verification_notes || '-'),
      old_status: complaint.status,
      new_status: newComplaintStatus,
      actor_user_id: currentUser.user_id,
      actor_name: currentUser.full_name,
      event_date: nowStr,
      created_at: nowStr
    });

    logAuditInternal('VERIFY_COMPLAINT', 'VERIFICATION', complaintId, 'Verifikasi ' + vStatus, complaint.status, newComplaintStatus, currentUser);

    return createSuccessResponse('Hasil verifikasi berhasil disimpan');
  } catch (err) {
    return createErrorResponse('Gagal menyimpan verifikasi: ' + err.message);
  }
}

function createReview(sessionToken, reviewData) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;
  const currentUser = auth.data.user;

  try {
    const complaintId = reviewData.complaint_id;
    const complaint = findRecordById('COMPLAINTS', 'complaint_id', complaintId);
    if (!complaint) return createErrorResponse('Pengaduan tidak ditemukan');

    const nowStr = formatISOString(new Date());
    const reviewRecord = {
      review_id: generateId('REV'),
      complaint_id: complaintId,
      reviewer_user_id: currentUser.user_id,
      review_date: nowStr,
      findings: reviewData.findings || '',
      analysis: reviewData.analysis || '',
      contributing_factors: reviewData.contributing_factors || '',
      review_result: reviewData.review_result || '',
      review_notes: reviewData.review_notes || '',
      document_file_id: '',
      created_at: nowStr,
      updated_at: nowStr
    };
    appendRecord('REVIEWS', reviewRecord);

    updateRecord('COMPLAINTS', 'complaint_id', complaintId, {
      status: 'UNDER_REVIEW',
      review_started_at: complaint.review_started_at || nowStr,
      updated_at: nowStr,
      updated_by: currentUser.user_id
    });

    appendRecord('TIMELINE', {
      timeline_id: generateId('TML'),
      complaint_id: complaintId,
      event_type: 'REVIEW_ADDED',
      event_title: 'Telaah Komite Medik',
      event_description: 'Hasil telaah medis diinput oleh ' + currentUser.full_name,
      old_status: complaint.status,
      new_status: 'UNDER_REVIEW',
      actor_user_id: currentUser.user_id,
      actor_name: currentUser.full_name,
      event_date: nowStr,
      created_at: nowStr
    });

    logAuditInternal('CREATE_REVIEW', 'REVIEW', complaintId, 'Input telaah komite medik', null, reviewRecord, currentUser);

    return createSuccessResponse('Hasil telaah medis berhasil disimpan');
  } catch (err) {
    return createErrorResponse('Gagal menyimpan telaah medis: ' + err.message);
  }
}

function createRecommendation(sessionToken, recData) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;
  const currentUser = auth.data.user;

  try {
    const complaintId = recData.complaint_id;
    const complaint = findRecordById('COMPLAINTS', 'complaint_id', complaintId);
    if (!complaint) return createErrorResponse('Pengaduan tidak ditemukan');

    const nowStr = formatISOString(new Date());
    const recRecord = {
      recommendation_id: generateId('RCM'),
      complaint_id: complaintId,
      recommendation_type: recData.recommendation_type || 'PEMBINAAN',
      recommendation_text: recData.recommendation_text || '',
      target_unit_id: recData.target_unit_id || '',
      responsible_person: recData.responsible_person || '',
      priority: recData.priority || 'MEDIUM',
      due_date: recData.due_date || '',
      status: 'ACTIVE',
      created_by: currentUser.user_id,
      created_at: nowStr,
      updated_at: nowStr,
      updated_by: currentUser.user_id
    };
    appendRecord('RECOMMENDATIONS', recRecord);

    updateRecord('COMPLAINTS', 'complaint_id', complaintId, {
      status: 'RECOMMENDATION',
      recommendation_at: nowStr,
      updated_at: nowStr,
      updated_by: currentUser.user_id
    });

    appendRecord('TIMELINE', {
      timeline_id: generateId('TML'),
      complaint_id: complaintId,
      event_type: 'RECOMMENDATION_CREATED',
      event_title: 'Rekomendasi Diterbitkan',
      event_description: 'Rekomendasi komite: ' + recRecord.recommendation_type,
      old_status: complaint.status,
      new_status: 'RECOMMENDATION',
      actor_user_id: currentUser.user_id,
      actor_name: currentUser.full_name,
      event_date: nowStr,
      created_at: nowStr
    });

    logAuditInternal('CREATE_RECOMMENDATION', 'RECOMMENDATION', complaintId, 'Rekomendasi terbit: ' + recRecord.recommendation_type, null, recRecord, currentUser);

    return createSuccessResponse('Rekomendasi berhasil dibuat');
  } catch (err) {
    return createErrorResponse('Gagal membuat rekomendasi: ' + err.message);
  }
}

function createFollowUp(sessionToken, followupData) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;
  const currentUser = auth.data.user;

  try {
    const complaintId = followupData.complaint_id;
    const complaint = findRecordById('COMPLAINTS', 'complaint_id', complaintId);
    if (!complaint) return createErrorResponse('Pengaduan tidak ditemukan');

    const nowStr = formatISOString(new Date());
    const followupRecord = {
      followup_id: generateId('FLP'),
      complaint_id: complaintId,
      recommendation_id: followupData.recommendation_id || '',
      followup_description: followupData.followup_description || '',
      responsible_unit_id: followupData.responsible_unit_id || '',
      responsible_person: followupData.responsible_person || '',
      start_date: nowStr,
      target_date: followupData.target_date || '',
      completion_date: '',
      progress_percentage: 0,
      result: '',
      evidence_file_id: '',
      status: 'IN_PROGRESS',
      created_at: nowStr,
      created_by: currentUser.user_id,
      updated_at: nowStr,
      updated_by: currentUser.user_id
    };
    appendRecord('FOLLOW_UPS', followupRecord);

    updateRecord('COMPLAINTS', 'complaint_id', complaintId, {
      status: 'FOLLOW_UP',
      followup_started_at: complaint.followup_started_at || nowStr,
      updated_at: nowStr,
      updated_by: currentUser.user_id
    });

    appendRecord('TIMELINE', {
      timeline_id: generateId('TML'),
      complaint_id: complaintId,
      event_type: 'FOLLOW_UP_STARTED',
      event_title: 'Tindak Lanjut Dimulai',
      event_description: 'Proses tindak lanjut rekomendasi dimulai',
      old_status: complaint.status,
      new_status: 'FOLLOW_UP',
      actor_user_id: currentUser.user_id,
      actor_name: currentUser.full_name,
      event_date: nowStr,
      created_at: nowStr
    });

    logAuditInternal('CREATE_FOLLOWUP', 'FOLLOW_UP', complaintId, 'Tindak lanjut dibuat', null, followupRecord, currentUser);

    return createSuccessResponse('Tindak lanjut berhasil dibuat');
  } catch (err) {
    return createErrorResponse('Gagal membuat tindak lanjut: ' + err.message);
  }
}

function getDashboardSummary(sessionToken) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;

  try {
    const complaints = getRecords('COMPLAINTS');
    
    const total = complaints.length;
    const submitted = complaints.filter(c => c.status === 'SUBMITTED').length;
    const underReview = complaints.filter(c => c.status === 'UNDER_REVIEW').length;
    const followUp = complaints.filter(c => c.status === 'FOLLOW_UP').length;
    const completed = complaints.filter(c => c.status === 'COMPLETED').length;
    const rejected = complaints.filter(c => c.status === 'REJECTED').length;

    return createSuccessResponse('Summary dashboard berhasil dimuat', {
      total_complaints: total,
      submitted_count: submitted,
      under_review_count: underReview,
      followup_count: followUp,
      completed_count: completed,
      rejected_count: rejected,
      overdue_complaints: 0
    });
  } catch (err) {
    return createErrorResponse('Gagal memuat dashboard summary: ' + err.message);
  }
}

function getReportData(sessionToken) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;

  try {
    const complaints = getRecords('COMPLAINTS');
    return createSuccessResponse('Data laporan berhasil dimuat', {
      summary: {
        total_complaints: complaints.length,
        in_review: complaints.filter(c => c.status === 'UNDER_REVIEW').length,
        completed: complaints.filter(c => c.status === 'COMPLETED').length,
        avg_resolution_days: 3.5
      }
    });
  } catch (err) {
    return createErrorResponse('Gagal memuat rekapitulasi laporan: ' + err.message);
  }
}

function exportReport(sessionToken) {
  const auth = validateSession(sessionToken);
  if (!auth.success) return auth;

  try {
    const complaints = getRecords('COMPLAINTS');
    let csv = 'Nomor Pengaduan,Pelapor,Unit,Kategori,Prioritas,Status,Tanggal\n';
    
    complaints.forEach(c => {
      csv += `"${c.complaint_number}","${c.reporter_name}","${c.reporter_unit_id}","${c.category_id}","${c.priority_id}","${c.status}","${c.created_at}"\n`;
    });

    return createSuccessResponse('Export berhasil', {
      filename: 'Laporan_SIMPEL_MEDIK_' + Utilities.formatDate(new Date(), APP_CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss') + '.csv',
      csv_data: csv
    });
  } catch (err) {
    return createErrorResponse('Gagal membuat export CSV: ' + err.message);
  }
}

function formatISOString(dateObj) {
  if (!dateObj) dateObj = new Date();
  if (typeof dateObj === 'string') dateObj = new Date(dateObj);
  if (isNaN(dateObj.getTime())) dateObj = new Date();
  return Utilities.formatDate(dateObj, APP_CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function createSuccessResponse(message, data, meta) {
  return {
    success: true,
    message: message || 'Operasi berhasil',
    data: data || null,
    meta: meta || {}
  };
}

function createErrorResponse(message, code, data) {
  return {
    success: false,
    message: message || 'Terjadi kesalahan sistem',
    code: code || 'UNKNOWN_ERROR',
    data: data || null
  };
}

