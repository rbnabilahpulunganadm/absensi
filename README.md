Cara Menggunakan:
Buka Google Sheets Anda dan buka script editor melalui Tools > Script editor

Ganti semua kode yang ada dengan kode di bawah

Simpan project dengan nama yang sesuai (misal: "Sistem Absensi")

Lakukan deployment dengan mengklik Deploy > New deployment

Pilih Web app sebagai jenis deployment, atur siapa yang dapat mengakses, lalu klik Deploy

Salin URL yang dihasilkan dan tempel di kode HTML Anda (ganti nilai SCRIPT_URL)


---
---


// Konfigurasi
const SHEET_NAME = "Data Absensi";
const TIMEZONE = "Asia/Jakarta";

// Fungsi utama untuk menangani POST request
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    
    var sheet = getSheet();
    
    if (action === 'absen_datang') {
      return handleAbsenDatang(sheet, data);
    } 
    else if (action === 'absen_pulang') {
      return handleAbsenPulang(sheet, data);
    }
    else if (action === 'fix_user_data') {
      return handleFixUserData(sheet, data);
    }
    
    throw new Error("Action tidak valid: " + action);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi untuk menangani GET request
function doGet(e) {
  try {
    var action = e.parameter.action;
    
    if (action === 'get_attendance') {
      var nama = e.parameter.nama;
      var sheet = getSheet();
      return handleGetAttendance(sheet, nama);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error", 
      "message": "Action tidak valid"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi untuk mendapatkan sheet
function getSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    // Buat header jika sheet baru
    const headers = [
      "ID", "Tanggal", "Nama", "Shift", 
      "Waktu Datang", "Lokasi Datang", "Keterlambatan", 
      "Waktu Pulang", "Lokasi Pulang", "Lembur", 
      "Total Jam Kerja", "Latitude", "Longitude",
      "Status", "Terakhir Diperbarui"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  
  return sheet;
}

// Fungsi untuk menangani absen datang
function handleAbsenDatang(sheet, data) {
  // Generate unique ID
  var id = Utilities.getUuid();
  
  // Dapatkan waktu saat ini
  var now = new Date();
  
  // Append new row for attendance
  sheet.appendRow([
    id,
    data.tanggal,
    data.nama,
    data.shift,
    data.waktu,
    data.lokasi,
    data.keterlambatan,
    '', // Waktu Pulang
    '', // Lokasi Pulang
    '', // Lembur
    '', // Total Jam Kerja
    data.latitude,
    data.longitude,
    'Hadir', // Status
    Utilities.formatDate(now, TIMEZONE, "dd/MM/yyyy HH:mm:ss") // Terakhir Diperbarui
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "message": "Data absen datang berhasil disimpan",
    "id": id
  })).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi untuk menangani absen pulang
function handleAbsenPulang(sheet, data) {
  var allData = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  // Jika menggunakan ID untuk mencari
  if (data.absen_datang_id && data.absen_datang_id.startsWith('manual-')) {
    // Mode manual - buat entry baru
    return handleManualAbsenPulang(sheet, data);
  }
  
  // Find the row with matching ID
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.absen_datang_id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    // Coba cari berdasarkan tanggal, nama, dan shift
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][1] === data.tanggal && 
          allData[i][2] === data.nama && 
          allData[i][3] === data.shift &&
          allData[i][7] === '') { // Pastikan belum absen pulang
        rowIndex = i + 1;
        break;
      }
    }
  }
  
  if (rowIndex !== -1) {
    // Calculate working hours
    var waktuDatangStr = sheet.getRange(rowIndex, 5).getValue();
    var totalJamKerja = "-";

    if(waktuDatangStr && typeof waktuDatangStr === 'string') {
      try {
        var [jamDatang, menitDatang, detikDatang] = waktuDatangStr.split(':').map(Number);
        var [jamPulang, menitPulang, detikPulang] = data.waktu.split(':').map(Number);
        
        var waktuMulai = new Date(0, 0, 0, jamDatang, menitDatang, detikDatang || 0);
        var waktuSelesai = new Date(0, 0, 0, jamPulang, menitPulang, detikPulang || 0);
        
        var diffMs = waktuSelesai - waktuMulai;
        if (diffMs < 0) {
           waktuSelesai.setDate(waktuSelesai.getDate() + 1);
           diffMs = waktuSelesai - waktuMulai;
        }
        
        var diffHours = Math.floor(diffMs / 3600000);
        var diffMins = Math.floor((diffMs % 3600000) / 60000);
        totalJamKerja = diffHours + " jam " + diffMins + " menit";
      } catch(err) {
        console.error("Error calculating working hours:", err);
        totalJamKerja = "Gagal hitung";
      }
    }
    
    // Dapatkan waktu saat ini
    var now = new Date();
    
    // Update the row with pulang data
    sheet.getRange(rowIndex, 8, 1, 6).setValues([[
      data.waktu, 
      data.lokasi, 
      data.lembur, 
      totalJamKerja,
      data.latitude,
      data.longitude
    ]]);
    
    // Update status dan waktu terakhir diperbarui
    sheet.getRange(rowIndex, 14).setValue('Pulang');
    sheet.getRange(rowIndex, 15).setValue(Utilities.formatDate(now, TIMEZONE, "dd/MM/yyyy HH:mm:ss"));
    
    return ContentService.createTextOutput(JSON.stringify({
      "result": "success", 
      "message": "Data absen pulang berhasil disimpan"
    })).setMimeType(ContentService.MimeType.JSON);
  } else {
    throw new Error("Data absen datang tidak ditemukan. Silakan gunakan absen manual jika perlu.");
  }
}

// Fungsi untuk menangani absen pulang manual
function handleManualAbsenPulang(sheet, data) {
  // Generate unique ID
  var id = Utilities.getUuid();
  
  // Dapatkan waktu saat ini
  var now = new Date();
  
  // Buat entry baru untuk absen pulang manual
  sheet.appendRow([
    id,
    data.tanggal,
    data.nama,
    data.shift,
    '00:00:00', // Waktu Datang (default)
    'Absen manual - data tidak ditemukan', // Lokasi Datang
    'Tidak tercatat', // Keterlambatan
    data.waktu, // Waktu Pulang
    data.lokasi, // Lokasi Pulang
    data.lembur, // Lembur
    'Tidak dapat dihitung', // Total Jam Kerja
    data.latitude,
    data.longitude,
    'Pulang (Manual)', // Status
    Utilities.formatDate(now, TIMEZONE, "dd/MM/yyyy HH:mm:ss") // Terakhir Diperbarui
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "message": "Data absen pulang manual berhasil disimpan",
    "id": id
  })).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi untuk mendapatkan data absensi
function handleGetAttendance(sheet, nama) {
  var allData = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < allData.length; i++) {
    // Cek apakah nama cocok dan belum absen pulang (kolom 8/Waktu Pulang kosong)
    if (allData[i][2] === nama && (allData[i][7] === '' || allData[i][7] === null)) {
      result.push({
        id: allData[i][0],
        tanggal: allData[i][1],
        nama: allData[i][2],
        shift: allData[i][3],
        waktu: allData[i][4],
        lokasi: allData[i][5],
        keterlambatan: allData[i][6]
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "data": result
  })).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi untuk memperbaiki data user
function handleFixUserData(sheet, data) {
  var namaLama = data.nama_lama;
  var namaBaru = data.nama_baru;
  
  var allData = sheet.getDataRange().getValues();
  var updated = 0;
  
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][2] === namaLama) { // Kolom nama adalah index 2
      sheet.getRange(i+1, 3).setValue(namaBaru); // Update kolom nama
      updated++;
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "message": "Berhasil memperbaiki " + updated + " data",
    "updated": updated
  })).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi tambahan untuk membersihkan data duplikat
function cleanDuplicateData() {
  var sheet = getSheet();
  var allData = sheet.getDataRange().getValues();
  var uniqueData = {};
  var rowsToDelete = [];
  
  // Identifikasi data duplikat (berdasarkan tanggal, nama, dan shift)
  for (var i = 1; i < allData.length; i++) {
    var key = allData[i][1] + "|" + allData[i][2] + "|" + allData[i][3];
    
    if (uniqueData[key]) {
      rowsToDelete.push(i + 1);
    } else {
      uniqueData[key] = true;
    }
  }
  
  // Hapus data duplikat (dari bawah ke atas)
  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var j = 0; j < rowsToDelete.length; j++) {
    sheet.deleteRow(rowsToDelete[j]);
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "message": "Berhasil menghapus " + rowsToDelete.length + " data duplikat"
  })).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi untuk mencadangkan data
function backupData() {
  var sheet = getSheet();
  var allData = sheet.getDataRange().getValues();
  var backupSheetName = "Backup " + Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
  
  var backupSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(backupSheetName);
  backupSheet.getRange(1, 1, allData.length, allData[0].length).setValues(allData);
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "message": "Backup data berhasil dibuat: " + backupSheetName
  })).setMimeType(ContentService.MimeType.JSON);
}
