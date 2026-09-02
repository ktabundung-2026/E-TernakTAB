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

// FUNGSI BARU: Cek apakah Nomor KKMT sudah pernah terdaftar
function cekKkmtAda(nomorKkmt) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;

    var dataKkmt = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    var targetKkmt = nomorKkmt.toString().trim();

    for (var i = 0; i < dataKkmt.length; i++) {
      var val = dataKkmt[i][0] ? dataKkmt[i][0].toString().replace("'", "").trim() : "";
      if (val === targetKkmt) {
        return true; // Ditemukan nomor KKMT yang sama
      }
    }
    return false;
  } catch (e) {
    return false;
  }
}

function getLivestockStats() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { total: 0, kerbau: 0, sapi: 0, kuda: 0 };
    }
    
    var data = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
    var total = data.length;
    var kerbau = 0, sapi = 0, kuda = 0;
    
    for (var i = 0; i < total; i++) {
      var val = data[i][0];
      if (!val) continue;
      var jenis = val.toString().trim().toLowerCase();
      if (jenis === "kerbau") kerbau++;
      else if (jenis === "sapi") sapi++;
      else if (jenis === "kuda") kuda++;
    }
    
    return { total: total, kerbau: kerbau, sapi: sapi, kuda: kuda };
  } catch (error) {
    return { total: 0, kerbau: 0, sapi: 0, kuda: 0 };
  }
}

function getDaftarTernak() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    var range = sheet.getRange(2, 1, lastRow - 1, 9);
    var values = range.getValues();
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

// Fungsi Simpan Data Baru dengan Proteksi Duplikasi
function simpanDataTernak(data, fileData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Proteksi Sisi Server: Cegah KKMT Duplikat
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