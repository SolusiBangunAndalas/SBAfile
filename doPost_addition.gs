// ============================================================
//  doPost — Terima data dari Web App Form
//  Tambahkan fungsi ini ke script Sheets yang sudah ada,
//  lalu Deploy ulang sebagai Web App dengan:
//  - Execute as: Me
//  - Who has access: Anyone
// ============================================================

function doPost(e) {
  try {
    // Parse JSON dari web form
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action || "";

    // Route berdasarkan action
    if (action === "submitRisk") {
      return handleSubmitRisk(payload);
    } else if (action === "getHostFisik") {
      return handleGetHostFisik();
    } else {
      return response({status:"error", message:"Unknown action: " + action});
    }
  } catch(err) {
    return response({status:"error", message:"Server error: " + err.toString()});
  }
}

// GET request — untuk getHostFisik dari web form (fetch via URL?action=getHostFisik)
function doGet(e) {
  try {
    var action = e.parameter.action || "";
    if (action === "getHostFisik") {
      return handleGetHostFisik();
    }
    return response({status:"ok", message:"IT Risk Management API is running"});
  } catch(err) {
    return response({status:"error", message:err.toString()});
  }
}


// ── HANDLER: Submit Risk Assessment ──────────────────────────
function handleSubmitRisk(payload) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  // Kolom yang akan ditulis — urutan sesuai header
  var fields = [
    "timestamp", "dept", "namaPengisi", "jabatan", "emailPengisi",
    "tglIsi", "lokasi", "namaAset", "idAset", "jenisAset",
    // Physical
    "vendor", "model", "serialNumber", "tahunBeli", "osPhysical",
    // VM
    "platform", "hostFisik", "osVM", "snapshot", "jumlahHost",
    // Jaringan
    "vlan", "ipAddress", "hostname", "picTeknis", "kontakPIC", "vendorSupport",
    // Teknis
    "patch", "koneksi", "akses", "backup", "monitoring", "antivirus", "auditTerakhir",
    // Ancaman
    "ancamanAll", "ancamanUtama", "insiden", "keteranganInsiden",
    // Dampak
    "dampakProd", "rto", "nilaiAset", "sopDarurat",
    // Penutup
    "rekomitigasi", "infoTambahan"
  ];

  // Buat header jika belum ada
  var lastCol = sheet.getLastColumn();
  var hasHeaders = lastCol > 0;

  if (!hasHeaders || sheet.getLastRow() === 0) {
    var headerRow = fields.concat(["Likelihood (L)", "Impact (I)", "Risk Score (L×I)", "Risk Level", "Status Mitigasi", "Tanggal Scoring"]);
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow])
      .setFontWeight("bold").setBackground("#2F5496").setFontColor("#FFFFFF");
  }

  // Tulis data ke baris baru
  var newRow = fields.map(function(f) {
    return payload[f] !== undefined ? payload[f] : "";
  });
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

  // Hitung risk score
  hitungScoreUntukBaris(sheet, nextRow, true);

  // Update dropdown host fisik jika jenis = Server Physical
  if ((payload.jenisAset || "").indexOf("Server Physical") > -1) {
    try { updateDropdownHost(); } catch(e2) {}
  }

  // Ambil skor yang baru dihitung untuk dikembalikan ke web form
  var headers2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var scoreCol = -1, levelCol = -1;
  for (var i = 0; i < headers2.length; i++) {
    if (headers2[i] === "Risk Score (L×I)") scoreCol = i + 1;
    if (headers2[i] === "Risk Level")       levelCol = i + 1;
  }
  var score = scoreCol > -1 ? sheet.getRange(nextRow, scoreCol).getValue() : 0;
  var level = levelCol > -1 ? sheet.getRange(nextRow, levelCol).getValue() : "-";

  return response({
    status: "ok",
    message: "Data berhasil disimpan",
    score: score,
    level: level,
    row: nextRow
  });
}


// ── HANDLER: Get Host Fisik untuk dropdown VM ─────────────────
function handleGetHostFisik() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return response({status:"ok", hosts:[]});
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  function colIdx(keyword) {
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().toLowerCase()
          .indexOf(keyword.toLowerCase()) > -1) return i;
    }
    return -1;
  }

  var idxNama  = colIdx("namaAset");
  var idxJenis = colIdx("jenisAset");

  // Fallback: cari kolom nama dan jenis aset
  if (idxNama === -1)  idxNama  = colIdx("nama aset");
  if (idxJenis === -1) idxJenis = colIdx("jenis aset");

  if (idxNama === -1 || idxJenis === -1) {
    return response({status:"ok", hosts:[], note:"Kolom tidak ditemukan"});
  }

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var hosts = [];

  data.forEach(function(row) {
    var jenis = (row[idxJenis] || "").toString();
    var nama  = (row[idxNama]  || "").toString().trim();
    if (jenis.indexOf("Server Physical") > -1 && nama && hosts.indexOf(nama) === -1) {
      hosts.push(nama);
    }
  });

  return response({status:"ok", hosts: hosts.sort()});
}


// ── Helper: JSON response dengan CORS header ──────────────────
function response(obj) {
  var output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
