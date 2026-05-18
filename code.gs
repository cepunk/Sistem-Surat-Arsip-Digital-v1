// Konfigurasi Global
const SHEET_NAME = "Database_Utama";

// Definisi Struktur Kolom Database (0-indexed)
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

// Header Kolom untuk inisialisasi otomatis
const HEADER_LIST = [
  "ID", "Nomor_Surat", "Tanggal_Surat", "Pengirim", "Penerima", 
  "Perihal", "Kategori", "Status", "Lampiran_URL", "Catatan", 
  "Created_At", "Updated_At"
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
 * Jika sheet baru dibuat/kosong, otomatis menyuntikkan data sampel (Seeder).
 */
function getDatabaseSheet() {
  try {
    let ss;
    const properties = PropertiesService.getScriptProperties();
    const customId = properties.getProperty("CUSTOM_SPREADSHEET_ID");
    
    // Gunakan Spreadsheet kustom jika dikonfigurasi, jika tidak gunakan yang aktif tempat script berada
    if (customId && customId.trim() !== "") {
      ss = SpreadsheetApp.openById(customId.trim());
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    // Buat otomatis sheet jika belum ada
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADER_LIST);
      sheet.getRange(1, 1, 1, HEADER_LIST.length)
           .setFontWeight("bold")
           .setBackground("#1e3a8a") // Brand Dark Blue
           .setFontColor("#ffffff");
    }
    
    // Auto-Seed Data Sampel jika sheet masih kosong (hanya ada baris header)
    if (sheet.getLastRow() === 1) {
      seedSampleData(sheet);
    }
    
    return sheet;
  } catch (error) {
    throw new Error("Gagal terhubung ke database: " + error.message);
  }
}

/**
 * Menyuntikkan data sampel dinas/kantor agar aplikasi langsung terisi saat pertama dideploy.
 */
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
  
  // Masukkan data ke sheet
  sheet.getRange(2, 1, sampleRows.length, HEADER_LIST.length).setValues(sampleRows);
}

/**
 * CREATE: Menyimpan data surat baru ke dalam sheet.
 * @param {Object} data - Objek surat dari form frontend.
 */
function createSurat(data) {
  try {
    const sheet = getDatabaseSheet();
    const timestamp = new Date();
    const newId = "SRT-" + generateUUID();
    
    const newRow = [
      newId,
      data.nomorSurat,
      data.tanggalSurat, // Format string "YYYY-MM-DD"
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

/**
 * READ: Mengambil seluruh data surat dari database.
 * Dioptimalkan dengan mengubah tipe objek Tanggal menjadi String ISO primitif agar tidak merusak serialisasi GAS.
 */
function getAllSurat() {
  try {
    const sheet = getDatabaseSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, data: [] }; // Kembalikan array kosong jika hanya ada header
    }
    
    const values = sheet.getRange(2, 1, lastRow - 1, HEADER_LIST.length).getValues();
    const listSurat = values.map(row => mapRowToObject(row));
    
    return {
      success: true,
      data: listSurat
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * READ SINGLE: Mengambil satu data surat berdasarkan ID.
 */
function getSuratById(id) {
  try {
    const sheet = getDatabaseSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return { success: false, error: "Data tidak ditemukan." };
    
    const values = sheet.getRange(2, 1, lastRow - 1, HEADER_LIST.length).getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (values[i][COLUMNS.ID] === id) {
        return {
          success: true,
          data: mapRowToObject(values[i])
        };
      }
    }
    
    return { success: false, error: "Surat dengan ID tersebut tidak ditemukan." };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * UPDATE: Memperbarui data surat yang sudah ada berdasarkan ID.
 */
function updateSurat(id, updatedData) {
  try {
    const sheet = getDatabaseSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return { success: false, error: "Database kosong." };
    
    const range = sheet.getRange(2, 1, lastRow - 1, HEADER_LIST.length);
    const values = range.getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (values[i][COLUMNS.ID] === id) {
        const rowNum = i + 2; // Offset header dan indeks 0
        const timestamp = new Date();
        
        // Memperbarui sel secara spesifik
        if (updatedData.nomorSurat !== undefined) sheet.getRange(rowNum, COLUMNS.NOMOR_SURAT + 1).setValue(updatedData.nomorSurat);
        if (updatedData.tanggalSurat !== undefined) sheet.getRange(rowNum, COLUMNS.TANGGAL_SURAT + 1).setValue(updatedData.tanggalSurat);
        if (updatedData.pengirim !== undefined) sheet.getRange(rowNum, COLUMNS.PENGIRIM + 1).setValue(updatedData.pengirim);
        if (updatedData.penerima !== undefined) sheet.getRange(rowNum, COLUMNS.PENERIMA + 1).setValue(updatedData.penerima);
        if (updatedData.perihal !== undefined) sheet.getRange(rowNum, COLUMNS.PERIHAL + 1).setValue(updatedData.perihal);
        if (updatedData.kategori !== undefined) sheet.getRange(rowNum, COLUMNS.KATEGORI + 1).setValue(updatedData.kategori);
        if (updatedData.status !== undefined) sheet.getRange(rowNum, COLUMNS.STATUS + 1).setValue(updatedData.status);
        if (updatedData.lampiranUrl !== undefined) sheet.getRange(rowNum, COLUMNS.LAMPIRAN_URL + 1).setValue(updatedData.lampiranUrl);
        if (updatedData.catatan !== undefined) sheet.getRange(rowNum, COLUMNS.CATATAN + 1).setValue(updatedData.catatan);
        
        // Selalu update waktu pembaruan terakhir
        sheet.getRange(rowNum, COLUMNS.UPDATED_AT + 1).setValue(timestamp);
        
        return {
          success: true,
          message: "Data surat berhasil diperbarui."
        };
      }
    }
    
    return { success: false, error: "Surat tidak ditemukan." };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * DELETE: Menghapus data surat berdasarkan ID.
 */
function deleteSurat(id) {
  try {
    const sheet = getDatabaseSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return { success: false, error: "Database kosong." };
    
    const range = sheet.getRange(2, 1, lastRow - 1, HEADER_LIST.length);
    const values = range.getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (values[i][COLUMNS.ID] === id) {
        const rowNum = i + 2; 
        sheet.deleteRow(rowNum);
        
        return {
          success: true,
          message: "Surat berhasil dihapus dari sistem."
        };
      }
    }
    
    return { success: false, error: "Surat tidak ditemukan." };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Mendapatkan konfigurasi ID Spreadsheet saat ini dan status koneksinya.
 */
function getCustomSpreadsheetConfig() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const customId = properties.getProperty("CUSTOM_SPREADSHEET_ID") || "";
    let activeId = "";
    
    try {
      activeId = SpreadsheetApp.getActiveSpreadsheet().getId();
    } catch(e) {
      activeId = "Tidak dapat mengambil ID spreadsheet internal";
    }
    
    let currentName = "Spreadsheet Aktif (Bawaan)";
    if (customId !== "") {
      try {
        currentName = SpreadsheetApp.openById(customId).getName();
      } catch (e) {
        currentName = "Error: Tidak dapat mengakses file (Periksa hak akses Anda / ID salah)";
      }
    } else {
      try {
        currentName = SpreadsheetApp.getActiveSpreadsheet().getName();
      } catch (e) {}
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

/**
 * Menyimpan ID Spreadsheet baru secara manual dan melakukan inisialisasi awal.
 * @param {string} id - ID Spreadsheet tujuan.
 */
function setCustomSpreadsheetId(id) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    // Skenario Reset ke Spreadsheet Bawaan
    if (!id || id.trim() === "") {
      properties.deleteProperty("CUSTOM_SPREADSHEET_ID");
      return { 
        success: true, 
        message: "Berhasil dikembalikan ke Spreadsheet bawaan." 
      };
    }
    
    const targetId = id.trim();
    let tempSs;
    
    // Validasi apakah ID valid dan bisa dibuka oleh script
    try {
      tempSs = SpreadsheetApp.openById(targetId);
    } catch (e) {
      return { 
        success: false, 
        error: "Spreadsheet tidak ditemukan atau Anda tidak memiliki akses edit. Pastikan ID benar dan akun Anda memiliki izin akses." 
      };
    }
    
    // Pastikan sheet database otomatis terbuat di spreadsheet target baru
    let sheet = tempSs.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = tempSs.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADER_LIST);
      sheet.getRange(1, 1, 1, HEADER_LIST.length)
           .setFontWeight("bold")
           .setBackground("#1e3a8a")
           .setFontColor("#ffffff");
    }
    
    // Auto-seed data juga jika spreadsheet kustom yang dihubungkan ternyata kosong
    if (sheet.getLastRow() === 1) {
      seedSampleData(sheet);
    }
    
    // Simpan ID baru di script properties
    properties.setProperty("CUSTOM_SPREADSHEET_ID", targetId);
    
    return { 
      success: true, 
      message: "Koneksi berhasil! Database terhubung ke: " + tempSs.getName() 
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Menerjemahkan satu baris data dari spreadsheet ke bentuk objek siap kirim.
 */
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

/**
 * Format sel tanggal aman. Mengubah Date Object menjadi format teks ISO YYYY-MM-DD
 * untuk menghindari kegagalan serialisasi data di google.script.run.
 */
function formatCellDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    // Format tanggal ke YYYY-MM-DD lokal
    const offset = val.getTimezoneOffset();
    const localDate = new Date(val.getTime() - (offset * 60 * 1000));
    return localDate.toISOString().split('T')[0];
  }
  return String(val);
}

/**
 * Membuat UUID sederhana dan unik untuk ID surat.
 */
function generateUUID() {
  const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let result = '';
  for (let i = 0; i < 6; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}
