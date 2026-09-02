// Konfigurasi Utama
const PASSWORD_AKSES = "admin1234"; // Password Khusus Petugas Input Data
const FOLDER_DRIVE_ID = "1Iv6DP7jP4SDYhUMSZnt50y4vSO5zlVDu"; 

function doGet(e) {
  if (e && e.parameter && e.parameter.p === 'pemeriksa') {
    return HtmlService.createTemplateFromFile('Pemeriksa')
        .evaluate()
        .setTitle('Monitoring & Pemeriksaan E-Ternak')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Aplikasi E-Ternak')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function periksaPassword(inputPassword) {
  return inputPassword === PASSWORD_AKSES;
}

// OPTIMALISASI: Tarik seluruh data tabel 1x saja untuk statistik & cek duplikasi
function getAplikasiData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { stats: { total: 0, kerbau: 0, sapi: 0, kuda: 0 }, values: [] };
    }
    
    // Tarik semua data dari kolom A sampai I sekaligus dalam 1 perintah (Sangat Cepat)
    var range = sheet.getRange(2, 1, lastRow - 1, 9);
    var values = range.getValues();
    
    var kerbau = 0, sapi = 0, kuda = 0;
    
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var jenis = row[2] ? row[2].toString().trim().toLowerCase() : "";
      
      if (jenis === "kerbau") kerbau++;
      else if (jenis === "sapi") sapi++;
      else if (jenis === "kuda") kuda++;
    }
    
    return {
      stats: { total: values.length, kerbau: kerbau, sapi: sapi, kuda: kuda },
      values: values
    };
  } catch (error) {
    return { stats: { total: 0, kerbau: 0, sapi: 0, kuda: 0 }, values: [] };
  }
}

// Fungsi Cek Duplikasi menggunakan data memori (Tanpa buka Sheet berulang)
function cekKkmtAda(nomorKkmt) {
  try {
    var dataApp = getAplikasiData();
    var values = dataApp.values;
    if (values.length === 0) return false;
    
    var targetKkmt = nomorKkmt.toString().trim();

    for (var i = 0; i < values.length; i++) {
      var val = values[i][1] ? values[i][1].toString().replace("'", "").trim() : "";
      if (val === targetKkmt) {
        return true; 
      }
    }
    return false;
  } catch (e) {
    return false;
  }
}

// Mengambil statistik secara cepat dari helper
function getLivestockStats() {
  return getAplikasiData().stats;
}

function getDaftarTernak() {
  try {
    var dataApp = getAplikasiData();
    var values = dataApp.values;
    var result = [];
    
    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      var tglRaw = row[0];
      var tglFormatted = "";
      
      if (tglRaw instanceof Date) {
        tglFormatted = Utilities.formatDate(tglRaw, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      } else {
        tglFormatted = tglRaw ? tglRaw.toString() : "-";
      }

      result.push({
        tanggal: tglFormatted,
        nomorKkmt: row[1] ? row[1].toString().replace("'", "") : "-",
        jenisTernak: row[2] || "-",
        jenisKelamin: row[3] || "-",
        namaPemilik: row[4] || "-",
        nikPemilik: row[5] ? row[5].toString().replace("'", "") : "-",
        asalTernak: row[6] || "-",
        namaPengantar: row[7] || "-",
        fileUrl: row[8] || ""
      });
    }
    
    return result;
  } catch (error) {
    return [];
  }
}

// Fungsi Simpan Data Baru
function simpanDataTernak(data, fileData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    if (cekKkmtAda(data.nomorKkmt)) {
      return { status: "DUPLIKAT", message: "Gagal menyimpan! Nomor KKMT '" + data.nomorKkmt + "' sudah terdaftar di sistem." };
    }

    var fileUrl = "";

    if (fileData && fileData.base64) {
      var folder = DriveApp.getFolderById(FOLDER_DRIVE_ID);
      var contentType = fileData.type;
      var extension = contentType.split('/')[1] || 'dat';
      
      var decodedData = Utilities.base64Decode(fileData.base64.split(',')[1]);
      var blob = Utilities.newBlob(decodedData, contentType, "Ternak_" + data.nomorKkmt + "_" + new Date().getTime() + "." + extension);
      
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    var formatKkmt = "'" + data.nomorKkmt;
    var formatNik = "'" + data.nikPemilik;

    sheet.appendRow([
      new Date(),
      formatKkmt,
      data.jenisTernak,
      data.jenisKelamin,
      data.namaPemilik,
      formatNik,
      data.asalTernak,
      data.namaPengantar,
      fileUrl
    ]);

    return { status: "SUKSES", message: "Data berhasil disimpan!" };
  } catch (error) {
    return { status: "GAGAL", message: error.toString() };
  }
}
