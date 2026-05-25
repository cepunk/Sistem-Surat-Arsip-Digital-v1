// Konfigurasi Global
const SHEET_NAME = "Database_Utama";
const USERS_SHEET_NAME = "Users";

// Definisi Struktur Kolom Database Surat (0-indexed)
const COLUMNS = {
  ID: 0,
  NOMOR_SURAT: 1,
  TANGGAL_SURAT: 2,
  PENGIRIM: 3,
  PENERIMA: 4,
  PERIHAL: 5,
  KATEGORI: 6, // "Surat Masuk" atau "Surat Keluar"
  STATUS: 7,     // "Draft", "Proses", "Arsip"
  LAMPIRAN_URL: 8,
  CATATAN: 9,
  CREATED_AT: 10,
  UPDATED_AT: 11
};

// Header Kolom Surat untuk inisialisasi otomatis
const HEADER_LIST = [
  "ID", "Nomor_Surat", "Tanggal_Surat", "Pengirim", "Penerima", 
  "Perihal", "Kategori", "Status", "Lampiran_URL", "Catatan", 
  "Created_At", "Updated_At"
];

// Header Kolom Users
const USERS_HEADER_LIST = [
  "Username", "Password", "FullName", "Role"
];

/**
 * Fungsi utama untuk melayani tampilan web (Frontend).
 * Merender file Index.html.
 */
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ArsipKita - Sistem Surat & Arsip Digital')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput("Error ketika memuat aplikasi: " + error.message);
  }
}

/**
 * Fungsi koneksi ke Spreadsheet.
 * Mengambil sheet Database_Utama atau membuatnya otomatis jika belum ada.
 */
function getDatabaseSheet() {
  try {
    let ss;
    const properties = PropertiesService.getScriptProperties();
    const customId = properties.getProperty("CUSTOM_SPREADSHEET_ID");
    
    if (customId && customId.trim() !== "") {
      ss = SpreadsheetApp.openById(customId.trim());
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADER_LIST);
      sheet.getRange(1, 1, 1, HEADER_LIST.length)
           .setFontWeight("bold")
           .setBackground("#1e3a8a") 
           .setFontColor("#ffffff");
    }
    
    if (sheet.getLastRow() === 1) {
      seedSampleData(sheet);
    }
    
    return sheet;
  } catch (error) {
    throw new Error("Gagal terhubung ke database: " + error.message);
  }
}

/**
 * Mendapatkan sheet Users atau membuatnya jika belum ada.
 */
function getUsersSheet() {
  try {
    let ss;
    const properties = PropertiesService.getScriptProperties();
    const customId = properties.getProperty("CUSTOM_SPREADSHEET_ID");
    
    if (customId && customId.trim() !== "") {
      ss = SpreadsheetApp.openById(customId.trim());
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(USERS_SHEET_NAME);
      sheet.appendRow(USERS_HEADER_LIST);
      sheet.getRange(1, 1, 1, USERS_HEADER_LIST.length)
           .setFontWeight("bold")
           .setBackground("#1e3a8a") 
           .setFontColor("#ffffff");
      
      // Seed Akun Default (Admin & User)
      sheet.appendRow(["admin", "admin123", "Administrator Utama", "Admin"]);
      sheet.appendRow(["user", "user123", "Staff Pengarsipan", "User"]);
    }
    
    return sheet;
  } catch (error) {
    throw new Error("Gagal memproses database pengguna: " + error.message);
  }
}

/**
 * Fungsi Otentikasi Pengguna untuk Login.
 */
function authenticateUser(username, password) {
  try {
    const sheet = getUsersSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: false, error: "Database pengguna kosong." };
    }
    
    const values = sheet.getRange(2, 1, lastRow - 1, USERS_HEADER_LIST.length).getValues();
    
    for (let i = 0; i < values.length; i++) {
      const dbUsername = String(values[i][0]).trim().toLowerCase();
      const dbPassword = String(values[i][1]).trim();
      const dbFullName = String(values[i][2]);
      const dbRole = String(values[i][3]);
      
      if (dbUsername === username.trim().toLowerCase() && dbPassword === password) {
        return {
          success: true,
          user: {
            username: dbUsername,
            fullName: dbFullName,
            role: dbRole
          }
        };
      }
    }
    
    return { success: false, error: "Username atau Password salah!" };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function seedSampleData(sheet) {
  const sampleRows = [
    [
      "SRT-GAS-001",
      "021/SET/X/2026",
      "2026-05-10",
      "PT. Solusi Teknologi Nusantara",
      "Bagian TI & Infrastruktur",
      "Penawaran Kerjasama Pengadaan Server Kantor Cabang Baru",
      "Surat Masuk",
      "Proses",
      "https://drive.google.com/file/d/demo-file-1",
      "Harap dikoordinasikan dengan tim IT Support untuk analisis spesifikasi teknis hardware.",
      new Date(),
      new Date()
    ],
    [
      "SRT-GAS-002",
      "005/12-PEM/2026",
      "2026-05-12",
      "Dinas Komunikasi & Informatika Daerah",
      "Direktur Utama",
      "Undangan Rapat Koordinasi Penerapan Transformasi Arsip Digital 2026",
      "Surat Masuk",
      "Arsip",
      "https://drive.google.com/file/d/demo-file-2",
      "Selesai didelegasikan kepada Bpk. Ahmad (Kabag Umum) untuk menghadiri rapat.",
      new Date(),
      new Date()
    ],
    [
      "SRT-GAS-003",
      "OUT/094/HRD/V/2026",
      "2026-05-15",
      "Divisi HRD & GA",
      "Seluruh Staff & Karyawan",
      "Memo Internal Kebijakan Penyesuaian Jam Kerja Selama Periode Renovasi Kantor",
      "Surat Keluar",
      "Draft",
      "",
      "Draf masih menunggu approval dari Direktur Operasional sebelum diumumkan.",
      new Date(),
      new Date()
    ],
    [
      "SRT-GAS-004",
      "045.2/281/SET/2026",
      "2026-05-18",
      "Sekretariat Perusahaan",
      "PT. Bank Mandiri (Persero) Tbk",
      "Surat Permohonan Audiensi Pengenalan Sistem Integrasi Layanan Payroll",
      "Surat Keluar",
      "Arsip",
      "https://drive.google.com/file/d/demo-file-3",
      "Arsip surat keluar fisik telah disimpan di Loker A1 Lantai 2.",
      new Date(),
      new Date()
    ]
  ];
  sheet.getRange(2, 1, sampleRows.length, HEADER_LIST.length).setValues(sampleRows);
}

function createSurat(data) {
  try {
    const sheet = getDatabaseSheet();
    const timestamp = new Date();
    const newId = "SRT-" + generateUUID();
    
    const newRow = [
      newId,
      data.nomorSurat,
      data.tanggalSurat, 
      data.pengirim,
      data.penerima,
      data.perihal,
      data.kategori,
      data.status,
      data.lampiranUrl || "",
      data.catatan || "",
      timestamp,
      timestamp
    ];
    
    sheet.appendRow(newRow);
    return { success: true, message: "Surat berhasil disimpan dengan ID: " + newId };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function getAllSurat() {
  try {
    const sheet = getDatabaseSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, data: [] }; 
    }
    
    const values = sheet.getRange(2, 1, lastRow - 1, HEADER_LIST.length).getValues();
    const listSurat = values.map(row => mapRowToObject(row));
    
    // Inisialisasi Users Sheet secara senyap saat meload agar tabel Users dijamin ada
    getUsersSheet();

    return { success: true, data: listSurat };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function updateSurat(id, updatedData) {
  try {
    const sheet = getDatabaseSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: "Database kosong." };
    
    const range = sheet.getRange(2, 1, lastRow - 1, HEADER_LIST.length);
    const values = range.getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (values[i][COLUMNS.ID] === id) {
        const rowNum = i + 2; 
        const timestamp = new Date();
        
        if (updatedData.nomorSurat !== undefined) sheet.getRange(rowNum, COLUMNS.NOMOR_SURAT + 1).setValue(updatedData.nomorSurat);
        if (updatedData.tanggalSurat !== undefined) sheet.getRange(rowNum, COLUMNS.TANGGAL_SURAT + 1).setValue(updatedData.tanggalSurat);
        if (updatedData.pengirim !== undefined) sheet.getRange(rowNum, COLUMNS.PENGIRIM + 1).setValue(updatedData.pengirim);
        if (updatedData.penerima !== undefined) sheet.getRange(rowNum, COLUMNS.PENERIMA + 1).setValue(updatedData.penerima);
        if (updatedData.perihal !== undefined) sheet.getRange(rowNum, COLUMNS.PERIHAL + 1).setValue(updatedData.perihal);
        if (updatedData.kategori !== undefined) sheet.getRange(rowNum, COLUMNS.KATEGORI + 1).setValue(updatedData.kategori);
        if (updatedData.status !== undefined) sheet.getRange(rowNum, COLUMNS.STATUS + 1).setValue(updatedData.status);
        if (updatedData.lampiranUrl !== undefined) sheet.getRange(rowNum, COLUMNS.LAMPIRAN_URL + 1).setValue(updatedData.lampiranUrl);
        if (updatedData.catatan !== undefined) sheet.getRange(rowNum, COLUMNS.CATATAN + 1).setValue(updatedData.catatan);
        
        sheet.getRange(rowNum, COLUMNS.UPDATED_AT + 1).setValue(timestamp);
        
        return { success: true, message: "Data surat berhasil diperbarui." };
      }
    }
    return { success: false, error: "Surat tidak ditemukan." };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function deleteSurat(id) {
  try {
    const sheet = getDatabaseSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: "Database kosong." };
    
    const range = sheet.getRange(2, 1, lastRow - 1, HEADER_LIST.length);
    const values = range.getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (values[i][COLUMNS.ID] === id) {
        sheet.deleteRow(i + 2);
        return { success: true, message: "Surat berhasil dihapus." };
      }
    }
    return { success: false, error: "Surat tidak ditemukan." };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function getCustomSpreadsheetConfig() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const customId = properties.getProperty("CUSTOM_SPREADSHEET_ID") || "";
    let activeId = "";
    
    try {
      activeId = SpreadsheetApp.getActiveSpreadsheet().getId();
    } catch(e) {
      activeId = "Tidak dapat mengambil ID";
    }
    
    let currentName = "Spreadsheet Aktif (Bawaan)";
    if (customId !== "") {
      try {
        currentName = SpreadsheetApp.openById(customId).getName();
      } catch (e) {
        currentName = "Error: Tidak dapat mengakses file.";
      }
    }

    return {
      success: true,
      customId: customId,
      activeId: activeId,
      currentName: currentName
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function setCustomSpreadsheetId(id) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    if (!id || id.trim() === "") {
      properties.deleteProperty("CUSTOM_SPREADSHEET_ID");
      return { success: true, message: "Berhasil dikembalikan ke Spreadsheet bawaan." };
    }
    
    const targetId = id.trim();
    let tempSs;
    
    try {
      tempSs = SpreadsheetApp.openById(targetId);
    } catch (e) {
      return { success: false, error: "Spreadsheet tidak ditemukan atau Anda tidak memiliki akses edit." };
    }
    
    let sheet = tempSs.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = tempSs.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADER_LIST);
      sheet.getRange(1, 1, 1, HEADER_LIST.length)
           .setFontWeight("bold")
           .setBackground("#1e3a8a")
           .setFontColor("#ffffff");
    }
    
    if (sheet.getLastRow() === 1) {
      seedSampleData(sheet);
    }
    
    properties.setProperty("CUSTOM_SPREADSHEET_ID", targetId);
    return { success: true, message: "Koneksi sukses dihubungkan ke: " + tempSs.getName() };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function mapRowToObject(row) {
  return {
    id: row[COLUMNS.ID],
    nomorSurat: row[COLUMNS.NOMOR_SURAT],
    tanggalSurat: formatCellDate(row[COLUMNS.TANGGAL_SURAT]),
    pengirim: row[COLUMNS.PENGIRIM],
    penerima: row[COLUMNS.PENERIMA],
    perihal: row[COLUMNS.PERIHAL],
    kategori: row[COLUMNS.KATEGORI],
    status: row[COLUMNS.STATUS],
    lampiranUrl: row[COLUMNS.LAMPIRAN_URL],
    catatan: row[COLUMNS.CATATAN],
    createdAt: formatCellDate(row[COLUMNS.CREATED_AT]),
    updatedAt: formatCellDate(row[COLUMNS.UPDATED_AT])
  };
}

function formatCellDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const offset = val.getTimezoneOffset();
    const localDate = new Date(val.getTime() - (offset * 60 * 1000));
    return localDate.toISOString().split('T')[0];
  }
  return String(val);
}

function generateUUID() {
  const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let result = '';
  for (let i = 0; i < 6; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}
