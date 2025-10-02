// Konfigurasi
const SHEET_NAME = "ABSENSIBIDAN";
const TIMEZONE = "Asia/Jakarta";

// Fungsi utama untuk menangani POST request
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const sheet = getSheet();
    
    switch (action) {
      case 'absen_datang':
        return handleAbsenDatang(sheet, data);
      case 'absen_pulang':
        return handleAbsenPulang(sheet, data);
      case 'fix_user_data':
        return handleFixUserData(sheet, data);
      case 'update_attendance':
        return handleUpdateAttendance(sheet, data);
      case 'delete_attendance':
        return handleDeleteAttendance(sheet, data);
      default:
        throw new Error("Action tidak valid: " + action);
    }
  } catch (error) {
    Logger.log("Error in doPost: " + error.toString() + "\nStack: " + error.stack);
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi untuk menangani GET request
function doGet(e) {
  try {
    const action = e.parameter.action;
    const sheet = getSheet();
    
    switch(action) {
      case 'get_attendance':
        const nama = e.parameter.nama;
        return handleGetAttendance(sheet, nama);
      case 'get_all_attendance':
        return handleGetAllAttendance(sheet);
      case 'get_statistics':
        return handleGetStatistics(sheet);
      case 'get_stats_weekly':
        return handleGetWeeklyStats(sheet);
      case 'get_stats_monthly':
        return handleGetMonthlyStats(sheet);
      case 'get_stats_yearly':
        return handleGetYearlyStats(sheet);
      default:
        throw new Error("Action GET tidak valid: " + action);
    }
    
  } catch (error) {
    Logger.log("Error in doGet: " + error.toString() + "\nStack: " + error.stack);
    return ContentService.createTextOutput(JSON.stringify({
      "result": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Fungsi untuk mendapatkan sheet dan memastikan header ada
function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    const headers = [
      "ID", "Tanggal", "Nama", "Shift", 
      "Waktu Datang", "Lokasi Datang", "Keterlambatan", 
      "Waktu Pulang", "Lokasi Pulang", "Lembur", 
      "Total Jam Kerja", "Latitude Datang", "Longitude Datang",
      "Status", "Terakhir Diperbarui", "Latitude Pulang", "Longitude Pulang"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setColumnWidths(1, 17, 150);
    sheet.getRange("B2:B").setNumberFormat("yyyy-mm-dd"); // Format tanggal ISO
  }
  return sheet;
}

// Menangani absen datang
function handleAbsenDatang(sheet, data) {
  const id = Utilities.getUuid();
  const now = new Date();
  
  sheet.appendRow([
    id, now, data.nama, data.shift,
    data.waktu, data.lokasi, data.keterlambatan,
    '', '', '', '', // Kolom pulang dikosongkan
    data.latitude, data.longitude,
    'Hadir', now, '', ''
  ]);
  
  return createSuccessResponse("Absen datang berhasil disimpan", { id: id });
}

// Menangani absen pulang
function handleAbsenPulang(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  const rowIndex = allData.findIndex(row => row[0] === data.absen_datang_id) + 1;
  
  if (rowIndex > 0) {
    const waktuDatangStr = sheet.getRange(rowIndex, 5).getValue();
    const totalJamKerja = calculateWorkingHours(waktuDatangStr, data.waktu);
    
    sheet.getRange(rowIndex, 8, 1, 4).setValues([[data.waktu, data.lokasi, data.lembur, totalJamKerja]]);
    sheet.getRange(rowIndex, 14).setValue('Pulang');
    sheet.getRange(rowIndex, 15).setValue(new Date());
    sheet.getRange(rowIndex, 16, 1, 2).setValues([[data.latitude, data.longitude]]);

    return createSuccessResponse("Absen pulang berhasil disimpan");
  } else {
    throw new Error("Data absen datang tidak ditemukan. Harap refresh halaman.");
  }
}

// Menangani update data oleh admin
function handleUpdateAttendance(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  const rowIndex = allData.findIndex(row => row[0] === data.id) + 1;

  if (rowIndex > 0) {
    const tanggal = new Date(data.tanggal);
    
    sheet.getRange(rowIndex, 2).setValue(tanggal);
    sheet.getRange(rowIndex, 3).setValue(data.nama);
    sheet.getRange(rowIndex, 4).setValue(data.shift);
    sheet.getRange(rowIndex, 5).setValue(data.waktu_datang);
    sheet.getRange(rowIndex, 7).setValue(data.keterlambatan);
    sheet.getRange(rowIndex, 8).setValue(data.waktu_pulang);
    sheet.getRange(rowIndex, 10).setValue(data.lembur);
    sheet.getRange(rowIndex, 11).setValue(data.total_jam_kerja);
    sheet.getRange(rowIndex, 15).setValue(new Date());

    return createSuccessResponse("Data absensi berhasil diperbarui");
  } else {
    throw new Error("Data absensi dengan ID tersebut tidak ditemukan.");
  }
}

// Menangani hapus data
function handleDeleteAttendance(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  const rowIndex = allData.findIndex(row => row[0] === data.id) + 1;
  
  if (rowIndex > 0) {
    sheet.deleteRow(rowIndex);
    return createSuccessResponse("Data absensi berhasil dihapus");
  } else {
    throw new Error("Data absensi tidak ditemukan.");
  }
}

// Memperbaiki nama user
function handleFixUserData(sheet, data) {
  const namaLama = data.nama_lama;
  const namaBaru = data.nama_baru;
  if (!namaLama || !namaBaru) throw new Error("Nama lama dan baru harus diisi.");

  const range = sheet.getRange("C2:C" + sheet.getLastRow());
  const allData = range.getValues();
  let updated = 0;
  
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][0].toString().trim().toLowerCase() === namaLama.trim().toLowerCase()) {
      sheet.getRange(i + 2, 3).setValue(namaBaru);
      updated++;
    }
  }
  
  return createSuccessResponse("Berhasil memperbaiki " + updated + " data", { updated: updated });
}

// Mengambil data untuk daftar absen pulang
function handleGetAttendance(sheet, nama) {
  const allData = sheet.getDataRange().getValues();
  const result = allData.slice(1)
    .filter(row => row[2] === nama && row[7] === '')
    .map(row => ({
      id: row[0],
      tanggal: formatTanggalDariSheet(row[1]),
      nama: row[2],
      shift: row[3],
      waktu: row[4],
      lokasi: row[5],
    }));
  return createSuccessResponse("Data berhasil diambil", { data: result });
}

// Mengambil semua data untuk admin
function handleGetAllAttendance(sheet) {
  const allData = sheet.getDataRange().getValues();
  const result = allData.slice(1).map(row => ({
    id: row[0],
    tanggal: formatTanggalDariSheet(row[1]),
    tanggal_iso: Utilities.formatDate(new Date(row[1]), TIMEZONE, "yyyy-MM-dd"),
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
    terakhir_diperbarui: formatTanggalDariSheet(row[14], true)
  })).sort((a, b) => new Date(b.terakhir_diperbarui) - new Date(a.terakhir_diperbarui));

  return createSuccessResponse("Semua data berhasil diambil", { data: result });
}


// --- FUNGSI STATISTIK ---

function handleGetStatistics(sheet) {
  const allData = sheet.getDataRange().getValues().slice(1);
  const now = new Date();
  const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastDayOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  let totalKeterlambatan = 0;
  let jumlahTerlambat = 0;
  let totalAbsensiBulanIni = 0;

  allData.forEach(row => {
    const tanggalAbsen = new Date(row[1]);
    if (tanggalAbsen >= firstDayOfMonth && tanggalAbsen <= lastDayOfMonth) {
      totalAbsensiBulanIni++;
    }

    const keterlambatan = row[6]; // Kolom Keterlambatan
    if (keterlambatan && keterlambatan.toString().includes('menit')) {
      const minutes = parseInt(keterlambatan.toString().replace(' menit', ''));
      if (!isNaN(minutes) && minutes > 0) {
        totalKeterlambatan += minutes;
        jumlahTerlambat++;
      }
    }
  });
  
  const rataRataKeterlambatan = jumlahTerlambat > 0 ? Math.round(totalKeterlambatan / jumlahTerlambat) : 0;

  const stats = {
    total_absensi_bulan_ini: totalAbsensiBulanIni,
    total_keterlambatan: totalKeterlambatan,
    rata_rata_keterlambatan: rataRataKeterlambatan
  };

  return createSuccessResponse("Statistik umum berhasil dimuat", { statistics: stats });
}

function handleGetWeeklyStats(sheet) {
  const allData = sheet.getDataRange().getValues().slice(1);
  const now = new Date();
  const startOfWeek = new Date(now.setDate(now.getDate() - now.getDay())); // Sunday
  const endOfWeek = new Date(now.setDate(now.getDate() - now.getDay() + 6)); // Saturday

  const stats = {
    daily_attendance: { 'Minggu': 0, 'Senin': 0, 'Selasa': 0, 'Rabu': 0, 'Kamis': 0, 'Jumat': 0, 'Sabtu': 0 },
    shift_comparison: { 'Pagi': 0, 'Sore': 0 }
  };
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];

  allData.forEach(row => {
    const tanggalAbsen = new Date(row[1]);
    if (tanggalAbsen >= startOfWeek && tanggalAbsen <= endOfWeek) {
      const dayName = days[tanggalAbsen.getDay()];
      stats.daily_attendance[dayName]++;
      
      const shift = row[3];
      if (stats.shift_comparison.hasOwnProperty(shift)) {
        stats.shift_comparison[shift]++;
      }
    }
  });

  return createSuccessResponse("Statistik mingguan berhasil dimuat", { stats: stats });
}


function handleGetMonthlyStats(sheet) {
    const allData = sheet.getDataRange().getValues().slice(1);
    const now = new Date();
    const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const lastDayOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);

    const stats = {
        user_attendance: {}
    };

    allData.forEach(row => {
        const tanggalAbsen = new Date(row[1]);
        if (tanggalAbsen >= firstDayOfMonth && tanggalAbsen <= lastDayOfMonth) {
            const nama = row[2];
            if (nama) {
              stats.user_attendance[nama] = (stats.user_attendance[nama] || 0) + 1;
            }
        }
    });

    return createSuccessResponse("Statistik bulanan berhasil dimuat", { stats: stats });
}


function handleGetYearlyStats(sheet) {
    const allData = sheet.getDataRange().getValues().slice(1);
    const now = new Date();
    const currentYear = now.getFullYear();

    const stats = {
        monthly_attendance: {} // {1: 20, 2: 30, ...}
    };

    for (let i = 1; i <= 12; i++) {
        stats.monthly_attendance[i] = 0;
    }

    allData.forEach(row => {
        const tanggalAbsen = new Date(row[1]);
        if (tanggalAbsen.getFullYear() === currentYear) {
            const month = tanggalAbsen.getMonth() + 1; // 1-12
            stats.monthly_attendance[month]++;
        }
    });
    
    return createSuccessResponse("Statistik tahunan berhasil dimuat", { stats: stats });
}


// --- FUNGSI HELPER ---

function createSuccessResponse(message, additionalData = {}) {
  const response = { "result": "success", "message": message, ...additionalData };
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

function calculateWorkingHours(startTime, endTime) {
  if (!startTime || !endTime) return "-";
  try {
    const start = new Date("1970-01-01T" + startTime);
    const end = new Date("1970-01-01T" + endTime);
    
    let diffMs = end - start;
    if (diffMs < 0) diffMs += 24 * 60 * 60 * 1000;
    
    const diffHours = Math.floor(diffMs / 3600000);
    const diffMins = Math.floor((diffMs % 3600000) / 60000);
    return diffHours + " jam " + diffMins + " menit";
  } catch (e) {
    return "Gagal hitung";
  }
}

function formatTanggalDariSheet(tanggal, withTime = false) {
  if (tanggal instanceof Date) {
    const format = withTime ? "d MMMM yyyy HH:mm:ss" : "d MMMM yyyy";
    return Utilities.formatDate(tanggal, TIMEZONE, format);
  }
  return tanggal;
}
