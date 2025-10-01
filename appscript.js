// Konfigurasi
const SHEET_NAME = "ABSENSIBIDAN";
const TIMEZONE = "Asia/Jakarta";
const ADMIN_PASSWORD = "112211";

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
    else if (action === 'get_all_attendance') {
      return handleGetAllAttendance(sheet, data);
    }
    else if (action === 'update_attendance') {
      return handleUpdateAttendance(sheet, data);
    }
    else if (action === 'delete_attendance') {
      return handleDeleteAttendance(sheet, data);
    }
    else if (action === 'get_statistics') {
      return handleGetStatistics(sheet, data);
    }
    else if (action === 'get_weekly_stats') {
      return handleGetWeeklyStats(sheet, data);
    }
    else if (action === 'get_monthly_stats') {
      return handleGetMonthlyStats(sheet, data);
    }
    else if (action === 'get_yearly_stats') {
      return handleGetYearlyStats(sheet, data);
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
    else if (action === 'get_all_attendance') {
      var sheet = getSheet();
      return handleGetAllAttendance(sheet, {});
    }
    else if (action === 'get_statistics') {
      var sheet = getSheet();
      return handleGetStatistics(sheet, {});
    }
    else if (action === 'get_weekly_stats') {
      var sheet = getSheet();
      return handleGetWeeklyStats(sheet, {});
    }
    else if (action === 'get_monthly_stats') {
      var sheet = getSheet();
      return handleGetMonthlyStats(sheet, {});
    }
    else if (action === 'get_yearly_stats') {
      var sheet = getSheet();
      return handleGetYearlyStats(sheet, {});
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
    
    // Set lebar kolom
    sheet.setColumnWidth(1, 250); // ID
    sheet.setColumnWidth(2, 120); // Tanggal
    sheet.setColumnWidth(3, 150); // Nama
    sheet.setColumnWidth(4, 80);  // Shift
    sheet.setColumnWidth(5, 100); // Waktu Datang
    sheet.setColumnWidth(6, 250); // Lokasi Datang
    sheet.setColumnWidth(7, 120); // Keterlambatan
    sheet.setColumnWidth(8, 100); // Waktu Pulang
    sheet.setColumnWidth(9, 250); // Lokasi Pulang
    sheet.setColumnWidth(10, 100); // Lembur
    sheet.setColumnWidth(11, 120); // Total Jam Kerja
    sheet.setColumnWidth(12, 100); // Latitude
    sheet.setColumnWidth(13, 100); // Longitude
    sheet.setColumnWidth(14, 100); // Status
    sheet.setColumnWidth(15, 150); // Terakhir Diperbarui
  }
  
  return sheet;
}

// Fungsi untuk menangani absen datang
function handleAbsenDatang(sheet, data) {
  // Generate unique ID
  var id = Utilities.getUuid();
  
  // Dapatkan waktu saat ini
  var now = new Date();
  
  // Format tanggal konsisten
  var tanggal = formatTanggalUntukSheet(data.tanggal);
  
  // Append new row for attendance
  sheet.appendRow([
    id,
    tanggal,
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
    // Format tanggal konsisten
    var tanggal = formatTanggalUntukSheet(data.tanggal);
    
    // Coba cari berdasarkan tanggal, nama, dan shift
    for (var i = 1; i < allData.length; i++) {
      var rowTanggal = allData[i][1];
      if (typeof rowTanggal === 'object') {
        rowTanggal = Utilities.formatDate(rowTanggal, TIMEZONE, "dd/MM/yyyy");
      }
      
      if (rowTanggal === tanggal && 
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
  
  // Format tanggal konsisten
  var tanggal = formatTanggalUntukSheet(data.tanggal);
  
  // Buat entry baru untuk absen pulang manual
  sheet.appendRow([
    id,
    tanggal,
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
        tanggal: formatTanggalDariSheet(allData[i][1]),
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

// Fungsi untuk mendapatkan semua data absensi
function handleGetAllAttendance(sheet, data) {
  var allData = sheet.getDataRange().getValues();
  var result = [];
  
  // Filter parameters
  var namaFilter = data.nama || '';
  var tanggalFilter = data.tanggal || '';
  var shiftFilter = data.shift || '';
  
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    var rowTanggal = formatTanggalDariSheet(row[1]);
    
    // Apply filters
    if (namaFilter && row[2] !== namaFilter) continue;
    if (tanggalFilter && rowTanggal !== tanggalFilter) continue;
    if (shiftFilter && row[3] !== shiftFilter) continue;
    
    result.push({
      id: row[0],
      tanggal: rowTanggal,
      nama: row[2],
      shift: row[3],
      waktu_datang: row[4],
      lokasi_datang: row[5],
      keterlambatan: row[6],
      waktu_pulang: row[7],
      lokasi_pulang: row[8],
      lembur: row[9],
      total_jam_kerja: row[10],
      status: row[13],
      terakhir_diperbarui: row[14]
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "data": result
  })).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi untuk memperbarui data absensi
function handleUpdateAttendance(sheet, data) {
  var allData = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  // Find the row with matching ID
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex !== -1) {
    // Format tanggal konsisten
    var tanggal = formatTanggalUntukSheet(data.tanggal);
    
    // Update the row
    sheet.getRange(rowIndex, 2, 1, 10).setValues([[
      tanggal,
      data.nama,
      data.shift,
      data.waktu_datang,
      data.lokasi_datang,
      data.keterlambatan,
      data.waktu_pulang,
      data.lokasi_pulang,
      data.lembur,
      data.total_jam_kerja
    ]]);
    
    // Update waktu terakhir diperbarui
    var now = new Date();
    sheet.getRange(rowIndex, 15).setValue(Utilities.formatDate(now, TIMEZONE, "dd/MM/yyyy HH:mm:ss"));
    
    return ContentService.createTextOutput(JSON.stringify({
      "result": "success", 
      "message": "Data absensi berhasil diperbarui"
    })).setMimeType(ContentService.MimeType.JSON);
  } else {
    throw new Error("Data absensi tidak ditemukan.");
  }
}

// Fungsi untuk menghapus data absensi
function handleDeleteAttendance(sheet, data) {
  var allData = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  // Find the row with matching ID
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex);
    
    return ContentService.createTextOutput(JSON.stringify({
      "result": "success", 
      "message": "Data absensi berhasil dihapus"
    })).setMimeType(ContentService.MimeType.JSON);
  } else {
    throw new Error("Data absensi tidak ditemukan.");
  }
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

// Fungsi untuk mendapatkan statistik umum
function handleGetStatistics(sheet, data) {
  var allData = sheet.getDataRange().getValues();
  var statistics = {
    total_absensi: 0,
    per_nama: {},
    per_shift: {},
    per_tanggal: {},
    rata_rata_keterlambatan: 0,
    total_keterlambatan: 0
  };
  
  var totalKeterlambatanMenit = 0;
  var jumlahKeterlambatan = 0;
  
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    statistics.total_absensi++;
    
    // Statistik per nama
    var nama = row[2];
    if (!statistics.per_nama[nama]) {
      statistics.per_nama[nama] = {
        total: 0,
        hadir: 0,
        pulang: 0,
        keterlambatan: 0
      };
    }
    statistics.per_nama[nama].total++;
    if (row[13] === 'Hadir' || row[13] === 'Pulang') statistics.per_nama[nama].hadir++;
    if (row[13] === 'Pulang') statistics.per_nama[nama].pulang++;
    
    // Hitung keterlambatan dalam menit
    if (row[6] && row[6] !== 'Tepat Waktu' && row[6] !== 'Tidak tercatat') {
      var keterlambatanStr = row[6];
      var matches = keterlambatanStr.match(/(\d+) menit/);
      if (matches && matches[1]) {
        var menit = parseInt(matches[1]);
        totalKeterlambatanMenit += menit;
        jumlahKeterlambatan++;
        statistics.per_nama[nama].keterlambatan += menit;
      }
    }
    
    // Statistik per shift
    var shift = row[3];
    if (!statistics.per_shift[shift]) {
      statistics.per_shift[shift] = 0;
    }
    statistics.per_shift[shift]++;
    
    // Statistik per tanggal
    var tanggal = formatTanggalDariSheet(row[1]);
    if (!statistics.per_tanggal[tanggal]) {
      statistics.per_tanggal[tanggal] = 0;
    }
    statistics.per_tanggal[tanggal]++;
  }
  
  // Hitung rata-rata keterlambatan
  if (jumlahKeterlambatan > 0) {
    statistics.rata_rata_keterlambatan = Math.round(totalKeterlambatanMenit / jumlahKeterlambatan);
    statistics.total_keterlambatan = totalKeterlambatanMenit;
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "statistics": statistics
  })).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi untuk mendapatkan statistik mingguan
function handleGetWeeklyStats(sheet, data) {
  var allData = sheet.getDataRange().getValues();
  var weeklyStats = {
    days: {},
    shift_comparison: { Pagi: 0, Sore: 0 },
    nama_comparison: {}
  };
  
  // Inisialisasi hari dalam seminggu
  var days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  days.forEach(day => {
    weeklyStats.days[day] = { Pagi: 0, Sore: 0, total: 0 };
  });
  
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    var tanggal = row[1];
    var nama = row[2];
    var shift = row[3];
    
    if (tanggal && typeof tanggal === 'object') {
      var dayOfWeek = days[tanggal.getDay()];
      
      // Hitung per hari
      weeklyStats.days[dayOfWeek][shift]++;
      weeklyStats.days[dayOfWeek].total++;
      
      // Hitung per shift
      weeklyStats.shift_comparison[shift]++;
      
      // Hitung per nama
      if (!weeklyStats.nama_comparison[nama]) {
        weeklyStats.nama_comparison[nama] = { Pagi: 0, Sore: 0, total: 0 };
      }
      weeklyStats.nama_comparison[nama][shift]++;
      weeklyStats.nama_comparison[nama].total++;
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "weekly_stats": weeklyStats
  })).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi untuk mendapatkan statistik bulanan
function handleGetMonthlyStats(sheet, data) {
  var allData = sheet.getDataRange().getValues();
  var monthlyStats = {
    weeks: {},
    shift_comparison: { Pagi: 0, Sore: 0 },
    nama_comparison: {}
  };
  
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    var tanggal = row[1];
    var nama = row[2];
    var shift = row[3];
    
    if (tanggal && typeof tanggal === 'object') {
      var weekNumber = getWeekNumber(tanggal);
      var month = tanggal.getMonth() + 1;
      var year = tanggal.getFullYear();
      var weekKey = `Minggu ${weekNumber} (${month}/${year})`;
      
      // Inisialisasi minggu jika belum ada
      if (!monthlyStats.weeks[weekKey]) {
        monthlyStats.weeks[weekKey] = { Pagi: 0, Sore: 0, total: 0 };
      }
      
      // Hitung per minggu
      monthlyStats.weeks[weekKey][shift]++;
      monthlyStats.weeks[weekKey].total++;
      
      // Hitung per shift
      monthlyStats.shift_comparison[shift]++;
      
      // Hitung per nama
      if (!monthlyStats.nama_comparison[nama]) {
        monthlyStats.nama_comparison[nama] = { Pagi: 0, Sore: 0, total: 0 };
      }
      monthlyStats.nama_comparison[nama][shift]++;
      monthlyStats.nama_comparison[nama].total++;
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "monthly_stats": monthlyStats
  })).setMimeType(ContentService.MimeType.JSON);
}

// Fungsi untuk mendapatkan statistik tahunan
function handleGetYearlyStats(sheet, data) {
  var allData = sheet.getDataRange().getValues();
  var yearlyStats = {
    months: {},
    shift_comparison: { Pagi: 0, Sore: 0 },
    nama_comparison: {}
  };
  
  var monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                   "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  
  // Inisialisasi bulan
  monthNames.forEach(month => {
    yearlyStats.months[month] = { Pagi: 0, Sore: 0, total: 0 };
  });
  
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    var tanggal = row[1];
    var nama = row[2];
    var shift = row[3];
    
    if (tanggal && typeof tanggal === 'object') {
      var monthIndex = tanggal.getMonth();
      var monthName = monthNames[monthIndex];
      
      // Hitung per bulan
      yearlyStats.months[monthName][shift]++;
      yearlyStats.months[monthName].total++;
      
      // Hitung per shift
      yearlyStats.shift_comparison[shift]++;
      
      // Hitung per nama
      if (!yearlyStats.nama_comparison[nama]) {
        yearlyStats.nama_comparison[nama] = { Pagi: 0, Sore: 0, total: 0 };
      }
      yearlyStats.nama_comparison[nama][shift]++;
      yearlyStats.nama_comparison[nama].total++;
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    "result": "success", 
    "yearly_stats": yearlyStats
  })).setMimeType(ContentService.MimeType.JSON);
}

// Helper function untuk mendapatkan nomor minggu
function getWeekNumber(date) {
  var firstDayOfYear = new Date(date.getFullYear(), 0, 1);
  var pastDaysOfYear = (date - firstDayOfYear) / 86400000;
  return Math.ceil((pastDaysOfYear + firstDayOfYear.getDay() + 1) / 7);
}

// Fungsi untuk memformat tanggal dari aplikasi ke format sheet
function formatTanggalUntukSheet(tanggalStr) {
  try {
    // Format: "1 Januari 2024" -> Date object
    var parts = tanggalStr.split(' ');
    var day = parseInt(parts[0]);
    var month = getMonthNumber(parts[1]);
    var year = parseInt(parts[2]);
    
    var date = new Date(year, month, day);
    return date;
  } catch (e) {
    // Jika gagal, gunakan tanggal hari ini
    return new Date();
  }
}

// Fungsi untuk memformat tanggal dari sheet ke format aplikasi
function formatTanggalDariSheet(tanggal) {
  if (typeof tanggal === 'object') {
    // Jika sudah Date object
    return Utilities.formatDate(tanggal, TIMEZONE, "d MMMM yyyy");
  } else if (typeof tanggal === 'string') {
    // Jika sudah string, coba parse
    try {
      var date = new Date(tanggal);
      return Utilities.formatDate(date, TIMEZONE, "d MMMM yyyy");
    } catch (e) {
      return tanggal;
    }
  }
  return tanggal;
}

// Helper function untuk mendapatkan nomor bulan dari nama bulan
function getMonthNumber(monthName) {
  var months = {
    'Januari': 0, 'Februari': 1, 'Maret': 2, 'April': 3,
    'Mei': 4, 'Juni': 5, 'Juli': 6, 'Agustus': 7,
    'September': 8, 'Oktober': 9, 'November': 10, 'Desember': 11
  };
  return months[monthName] || 0;
}

// Fungsi tambahan untuk membersihkan data duplikat
function cleanDuplicateData() {
  var sheet = getSheet();
  var allData = sheet.getDataRange().getValues();
  var uniqueData = {};
  var rowsToDelete = [];
  
  // Identifikasi data duplikat (berdasarkan tanggal, nama, dan shift)
  for (var i = 1; i < allData.length; i++) {
    var key = formatTanggalDariSheet(allData[i][1]) + "|" + allData[i][2] + "|" + allData[i][3];
    
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