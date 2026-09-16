/**
 * SISTEM INFORMASI DATA ISPU SEKOLAH DASAR 2026
 * Backend Script (Google Apps Script / Code.gs)
 */

const SHEET_NAME = "Data_ISPU_SD";

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Setup Headers with clean property names
    const headers = [
      "ID",
      "Tanggal",
      "NPSN",
      "Nama_Sekolah",
      "Kecamatan",
      "Kualitas_Udara_PM25",
      "Kualitas_Udara_PM10",
      "Usulan_Pelaksanaan_Pembelajaran",
      "Keterangan_Kondisi_Sekolah",
      "Bukti_Dukung"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return responseJSON({ status: "success", data: [] });
    }
    
    const headers = data[0];
    const rows = data.slice(1);
    
    const result = rows.map((row, index) => {
      let obj = {};
      headers.forEach((header, colIndex) => {
        let val = row[colIndex];
        if (header === "Tanggal" && val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        // Support both old header (PM2.5) and new header (PM25)
        let key = header === "Kualitas_Udara_PM2.5" ? "Kualitas_Udara_PM25" : header;
        obj[key] = val;
      });
      if (!obj["ID"]) obj["ID"] = index + 1;
      return obj;
    });
    
    return responseJSON({ status: "success", data: result });
  } catch (error) {
    return responseJSON({ status: "error", message: error.toString() });
  }
}

function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;
    const payload = contents.data;
    const pm25Val = payload.Kualitas_Udara_PM25 || payload["Kualitas_Udara_PM2.5"] || 0;
    
    if (action === 'CREATE') {
      const newId = new Date().getTime().toString();
      const newRow = [
        newId,
        payload.Tanggal || "",
        payload.NPSN || "",
        payload.Nama_Sekolah || "",
        payload.Kecamatan || "",
        pm25Val,
        payload.Kualitas_Udara_PM10 || 0,
        payload.Usulan_Pelaksanaan_Pembelajaran || "",
        payload.Keterangan_Kondisi_Sekolah || "",
        payload.Bukti_Dukung || ""
      ];
      sheet.appendRow(newRow);
      return responseJSON({ status: "success", message: "Data berhasil ditambahkan!", id: newId });
    }
    
    if (action === 'UPDATE') {
      const data = sheet.getDataRange().getValues();
      const targetId = payload.ID;
      let rowIndex = -1;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(targetId)) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex !== -1) {
        const updateRow = [
          targetId,
          payload.Tanggal || "",
          payload.NPSN || "",
          payload.Nama_Sekolah || "",
          payload.Kecamatan || "",
          pm25Val,
          payload.Kualitas_Udara_PM10 || 0,
          payload.Usulan_Pelaksanaan_Pembelajaran || "",
          payload.Keterangan_Kondisi_Sekolah || "",
          payload.Bukti_Dukung || ""
        ];
        sheet.getRange(rowIndex, 1, 1, updateRow.length).setValues([updateRow]);
        return responseJSON({ status: "success", message: "Data berhasil diperbarui!" });
      } else {
        return responseJSON({ status: "error", message: "ID data tidak ditemukan!" });
      }
    }
    
    if (action === 'DELETE') {
      const data = sheet.getDataRange().getValues();
      const targetId = payload.ID;
      let rowIndex = -1;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(targetId)) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex !== -1) {
        sheet.deleteRow(rowIndex);
        return responseJSON({ status: "success", message: "Data berhasil dihapus!" });
      } else {
        return responseJSON({ status: "error", message: "ID data tidak ditemukan!" });
      }
    }
    
    return responseJSON({ status: "error", message: "Aksi tidak dikenal!" });
  } catch (error) {
    return responseJSON({ status: "error", message: error.toString() });
  }
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}