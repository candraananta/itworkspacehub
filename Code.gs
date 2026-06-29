/**
 * Senior Fullstack Developer - IT Tools System Backend Hub
 * Database Engine: Google Sheets (Strict Locking, Auto-Indexing, Auto-Patch & Universal Soft-Delete)
 */

const SHEETS = {
  INCIDENTS:     "incidents",
  LOCATIONS:     "locations",
  CATEGORIES:    "categories",
  VLANS:         "vlans",
  RACKS:         "racks",
  KPI_PROJECTS:  "kpi_projects",
  KPI_TRAININGS: "kpi_trainings",
  WIKI:          "knowledge_base",
  MASTER_ITEMS:  "master_items",
  PASSWORDS:     "passwords",
  MASTER_USERS:  "master_users",       // NEW
  LICENSES:      "office_licenses"     // NEW
};

/**
 * Inisialisasi Otomatis Skema Database & Auto-Patch Kolom Baru
 */
function setupDatabase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const createSheetIfNotExist = (name, headers, headerBg, fontColor = "#ffffff") => {
      let sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
        sheet.getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setFontWeight("bold")
          .setBackground(headerBg)
          .setFontColor(fontColor);
        sheet.setFrozenRows(1);
      }
    };

    createSheetIfNotExist(SHEETS.INCIDENTS,     ["ID","Tanggal","Location","Category","Description","Status","Photo Before","Photo After","Timestamp"],                            "#000000");
    createSheetIfNotExist(SHEETS.LOCATIONS,     ["ID","Location Name","Status"],                                                                                                   "#334155");
    createSheetIfNotExist(SHEETS.CATEGORIES,    ["ID","Category Name","Status"],                                                                                                   "#334155");
    createSheetIfNotExist(SHEETS.VLANS,         ["ID","VLAN ID","VLAN Name","Network","Gateway","Switch Target","Location","Notes","Status","Timestamp"],                          "#1e3a8a");
    createSheetIfNotExist(SHEETS.RACKS,         ["ID","Rack Name","Location","Size","Photos","Devices","Status","Timestamp"],                                                       "#111827");
    createSheetIfNotExist(SHEETS.KPI_PROJECTS,  ["ID","Tanggal","Project Name","Location","Description","Photos","Status","Timestamp"],                                            "#7c3aed");
    createSheetIfNotExist(SHEETS.KPI_TRAININGS, ["ID","Tanggal","Judul Training","Topic","Method","Location","Feedback","Photos","Status","Timestamp"],                            "#db2777");
    createSheetIfNotExist(SHEETS.MASTER_ITEMS,  ["ID","Item Name","Category","Brand","Serial Number","est_price","Status","Timestamp"],                                            "#ea580c");
    createSheetIfNotExist(SHEETS.PASSWORDS,     ["ID","App Name","URL","Username","Password","Category","Location","Notes","Timestamp","Is Active"],                              "#0d9488");

    // NEW: Master Users
    createSheetIfNotExist(SHEETS.MASTER_USERS,  ["ID","Full Name","Department","Email","Phone","Status","Timestamp"],                                                              "#1e40af");

    // NEW: Office Licenses
    createSheetIfNotExist(SHEETS.LICENSES, [
      "ID","Part Number","Product Name","Qty","Product Key",
      "Unique Key","Tanggal Aktivasi","Hostname","User",
      "Item ID","Laptop Brand Type","Akun Admin","Status","Timestamp"
    ], "#1d4ed8");

    // Auto-Patching untuk Modul WIKI (Menambah Kolom PDF_Attachment tanpa menghapus sheet)
    let wikiSheet = ss.getSheetByName(SHEETS.WIKI);
    if (!wikiSheet) {
      wikiSheet = ss.insertSheet(SHEETS.WIKI);
      wikiSheet.getRange(1, 1, 1, 7)
        .setValues([["ID","Title","Category","Steps_JSON","PDF_Attachment","Status","Timestamp"]])
        .setFontWeight("bold").setBackground("#059669").setFontColor("#ffffff");
      wikiSheet.setFrozenRows(1);
    } else {
      let headersRange = wikiSheet.getRange(1, 1, 1, wikiSheet.getLastColumn());
      let headers = headersRange.getValues()[0];
      if (headers.indexOf("PDF_Attachment") === -1) {
        let statusIdx = headers.indexOf("Status") + 1;
        if (statusIdx > 0) {
          wikiSheet.insertColumnBefore(statusIdx);
          wikiSheet.getRange(1, statusIdx).setValue("PDF_Attachment")
            .setFontWeight("bold").setBackground("#059669").setFontColor("#ffffff");
        }
      }
    }

    return "✅ SINKRONISASI DATABASE SUKSES! (Semua tabel aman, auto-patch aktif, modul baru terdaftar)";
  } catch (error) {
    return "❌ GAGAL MENYUSUN DATABASE: " + error.toString();
  }
}

function doGet() {
  try {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('IT Tools ERP - Enterprise Ecosystem')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput("Critical UI Loader Error: " + error.toString());
  }
}

function getSheetData(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    const headers = rows[0];
    const formattedData = [];
    const statusColIndex = headers.indexOf("Status");

    for (let i = 1; i < rows.length; i++) {
      if (statusColIndex !== -1 && rows[i][statusColIndex] === "DELETED") continue;
      const obj = {};
      headers.forEach((header, colIndex) => {
        let val = rows[i][colIndex];
        if (val instanceof Date) {
          obj[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        } else if (header === 'Photos' || header === 'Devices' || header === 'Steps_JSON') {
          try { obj[header] = val ? JSON.parse(val) : (header === 'Devices' ? {} : []); }
          catch(e) { obj[header] = header === 'Devices' ? {} : []; }
        } else if (header === 'Is Active') {
          obj[header] = (val === true || val === 1 || String(val).toUpperCase() === 'TRUE');
        } else {
          obj[header] = val;
        }
      });
      formattedData.push(obj);
    }
    return formattedData;
  } catch (error) { return []; }
}

function getInitialData() {
  try {
    return {
      success:      true,
      incidents:    getSheetData(SHEETS.INCIDENTS),
      locations:    getSheetData(SHEETS.LOCATIONS),
      categories:   getSheetData(SHEETS.CATEGORIES),
      vlans:        getSheetData(SHEETS.VLANS),
      racks:        getSheetData(SHEETS.RACKS),
      kpiProjects:  getSheetData(SHEETS.KPI_PROJECTS),
      kpiTrainings: getSheetData(SHEETS.KPI_TRAININGS),
      wiki:         getSheetData(SHEETS.WIKI),
      masterItems:  getSheetData(SHEETS.MASTER_ITEMS),
      passwords:    getSheetData(SHEETS.PASSWORDS),
      masterUsers:  getSheetData(SHEETS.MASTER_USERS),   // NEW
      licenses:     getSheetData(SHEETS.LICENSES)         // NEW
    };
  } catch (error) { return { success: false, message: error.toString() }; }
}

// ================= TRANSACTIONAL CRUD ENGINE =================
function createIncident(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INCIDENTS);
    const rows = sheet.getDataRange().getValues();
    let nextId = "INC-0001";
    if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.includes("INC-")) nextId = "INC-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
    sheet.appendRow([nextId, data.date ? new Date(data.date) : new Date(), data.location, data.category, data.description || "", "Waiting", data.photoBefore || "", "", new Date()]);
    return { success: true, message: `Sukses membuat log kendala teknis ${nextId}.` };
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

function updateIncident(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INCIDENTS);
    const rows = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) { if (rows[i][0] === data.id) { rowIndex = i + 1; break; } }
    if (rowIndex === -1) return { success: false, message: "ID Incident tidak valid." };
    sheet.getRange(rowIndex, 2).setValue(data.date ? new Date(data.date) : new Date());
    sheet.getRange(rowIndex, 3).setValue(data.location);
    sheet.getRange(rowIndex, 4).setValue(data.category);
    sheet.getRange(rowIndex, 5).setValue(data.description || "");
    sheet.getRange(rowIndex, 7).setValue(data.photoBefore || "");
    return { success: true, message: `Log incident ${data.id} berhasil diperbarui.` };
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

function closeIncident(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INCIDENTS);
    const rows = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < rows.length; i++) { if (rows[i][0] === data.id) { rowIndex = i + 1; break; } }
    if (rowIndex === -1) return { success: false, message: "ID tidak ditemukan." };
    sheet.getRange(rowIndex, 6).setValue("Closed");
    sheet.getRange(rowIndex, 8).setValue(data.photoAfter || "");
    return { success: true, message: `Log incident ${data.id} ditutup (Selesai Penanganan).` };
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

// ================= INFRASTRUCTURE ENGINE =================
function saveVlanRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VLANS);
    const rows = sheet.getDataRange().getValues();
    const isEdit = data.id !== null && data.id !== undefined && data.id !== '';
    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const rIdx = i + 1;
          sheet.getRange(rIdx, 2).setValue(data.vlanId); sheet.getRange(rIdx, 3).setValue(data.name);
          sheet.getRange(rIdx, 4).setValue(data.network); sheet.getRange(rIdx, 5).setValue(data.gateway);
          sheet.getRange(rIdx, 6).setValue(data.switchName); sheet.getRange(rIdx, 7).setValue(data.location);
          sheet.getRange(rIdx, 8).setValue(data.notes || "");
          return { success: true, message: `Topologi VLAN ${data.vlanId} diperbarui.` };
        }
      }
    } else {
      let nextId = "VL-0001";
      if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.includes("VL-")) nextId = "VL-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
      sheet.appendRow([nextId, data.vlanId, data.name, data.network, data.gateway, data.switchName, data.location, data.notes || "", "ACTIVE", new Date()]);
      return { success: true, message: `VLAN ID ${data.vlanId} deployed successfully.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

function saveRackRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RACKS);
    const rows = sheet.getDataRange().getValues();
    const isEdit = data.id !== null && data.id !== undefined && data.id !== '';
    const photosStr = JSON.stringify(data.photos || []);
    const devicesStr = JSON.stringify(data.devices || {});
    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const rIdx = i + 1;
          sheet.getRange(rIdx, 2).setValue(data.name); sheet.getRange(rIdx, 3).setValue(data.location);
          sheet.getRange(rIdx, 4).setValue(data.size); sheet.getRange(rIdx, 5).setValue(photosStr);
          sheet.getRange(rIdx, 6).setValue(devicesStr);
          return { success: true, message: `Arsitektur Server Rack ${data.name} disimpan.` };
        }
      }
    } else {
      let nextId = "RCK-0001";
      if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.includes("RCK-")) nextId = "RCK-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
      sheet.appendRow([nextId, data.name, data.location, data.size, photosStr, devicesStr, "ACTIVE", new Date()]);
      return { success: true, message: `Registrasi fisik Server Rack ${data.name} berhasil.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

// ================= KPI OPERATIONS ENGINE =================
function saveKpiProjectRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KPI_PROJECTS);
    const rows = sheet.getDataRange().getValues();
    const isEdit = data.id !== null && data.id !== undefined && data.id !== '';
    const photosStr = JSON.stringify(data.photos || []);
    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const rIdx = i + 1;
          sheet.getRange(rIdx, 2).setValue(data.date ? new Date(data.date) : new Date());
          sheet.getRange(rIdx, 3).setValue(data.projectName); sheet.getRange(rIdx, 4).setValue(data.location);
          sheet.getRange(rIdx, 5).setValue(data.description || ""); sheet.getRange(rIdx, 6).setValue(photosStr);
          return { success: true, message: `Project log '${data.projectName}' diperbarui.` };
        }
      }
    } else {
      let nextId = "KP-0001";
      if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.includes("KP-")) nextId = "KP-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
      sheet.appendRow([nextId, data.date ? new Date(data.date) : new Date(), data.projectName, data.location, data.description || "", photosStr, "ACTIVE", new Date()]);
      return { success: true, message: `KPI Delivery Project '${data.projectName}' terdata.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

function saveKpiTrainingRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KPI_TRAININGS);
    const rows = sheet.getDataRange().getValues();
    const isEdit = data.id !== null && data.id !== undefined && data.id !== '';
    const photosStr = JSON.stringify(data.photos || []);
    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const rIdx = i + 1;
          sheet.getRange(rIdx, 2).setValue(data.date ? new Date(data.date) : new Date());
          sheet.getRange(rIdx, 3).setValue(data.title); sheet.getRange(rIdx, 4).setValue(data.topic);
          sheet.getRange(rIdx, 5).setValue(data.method); sheet.getRange(rIdx, 6).setValue(data.location);
          sheet.getRange(rIdx, 7).setValue(data.feedback || ""); sheet.getRange(rIdx, 8).setValue(photosStr);
          return { success: true, message: `Data Training '${data.title}' berhasil diperbarui.` };
        }
      }
    } else {
      let nextId = "TR-0001";
      if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.includes("TR-")) nextId = "TR-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
      sheet.appendRow([nextId, data.date ? new Date(data.date) : new Date(), data.title, data.topic, data.method, data.location, data.feedback || "", photosStr, "ACTIVE", new Date()]);
      return { success: true, message: `Log Training Record '${data.title}' berhasil dimasukkan.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

// ================= WIKI RESOURCE ENGINE (WITH PDF) =================
function saveWikiRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WIKI);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const pdfColIndex = headers.indexOf("PDF_Attachment") + 1;
    const isEdit = data.id !== null && data.id !== undefined && data.id !== '';
    const stepsStr = JSON.stringify(data.steps || []);
    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const rIdx = i + 1;
          sheet.getRange(rIdx, 2).setValue(data.title);
          sheet.getRange(rIdx, 3).setValue(data.category);
          sheet.getRange(rIdx, 4).setValue(stepsStr);
          if (pdfColIndex > 0) sheet.getRange(rIdx, pdfColIndex).setValue(data.pdf || "");
          return { success: true, message: `Artikel Wiki '${data.title}' berhasil di-update.` };
        }
      }
    } else {
      let nextId = "WK-0001";
      if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.includes("WK-")) nextId = "WK-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
      let newRow = Array(headers.length).fill("");
      newRow[0] = nextId; newRow[1] = data.title; newRow[2] = data.category; newRow[3] = stepsStr;
      if (pdfColIndex > 0) newRow[pdfColIndex - 1] = data.pdf || "";
      newRow[headers.indexOf("Status")] = "ACTIVE";
      newRow[headers.indexOf("Timestamp")] = new Date();
      sheet.appendRow(newRow);
      return { success: true, message: `Knowledge base Wiki '${data.title}' berhasil dipublikasikan.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

// ================= DATA MANAGEMENT SYSTEM ENGINE =================
function saveMasterRow(sheetName, data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const rows = sheet.getDataRange().getValues();
    const isEdit = data.idField !== null && data.idField !== undefined && data.idField !== '';
    const targetValue = data.valueField.trim();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][1].toString().toLowerCase() === targetValue.toLowerCase() && rows[i][2] !== "DELETED") {
        if (!isEdit || (isEdit && rows[i][0] !== data.idField)) return { success: false, message: `Entitas master '${targetValue}' sudah aktif!` };
      }
    }
    if (isEdit) {
      for (let i = 1; i < rows.length; i++) { if (rows[i][0] === data.idField) { sheet.getRange(i+1, 2).setValue(targetValue); return { success: true, message: "Data master diperbarui." }; } }
    } else {
      let prefix = sheetName === SHEETS.LOCATIONS ? "LC-" : "CT-"; let nextId = prefix + "0001";
      if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.includes("-")) nextId = prefix + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
      sheet.appendRow([nextId, targetValue, "ACTIVE"]);
      return { success: true, message: `Master data ${nextId} ditambahkan.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

function executeDeleteRecord(sheetName, idToDelete) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Koneksi database terputus." };
    const data = sheet.getDataRange().getValues();
    const statusColIdx = data[0].indexOf("Status") + 1;
    if (statusColIdx === 0) return { success: false, message: "Format tabel salah." };
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === idToDelete.toString()) {
        sheet.getRange(i + 1, statusColIdx).setValue("DELETED");
        return { success: true, message: `Data ${idToDelete} berhasil diarsip (Soft Deleted).` };
      }
    }
    return { success: false, message: "Data tidak eksis di server." };
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

// ================= PASSWORD VAULT ENGINE =================
function savePasswordRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PASSWORDS);
    if (!sheet) return { success: false, message: "Sheet passwords tidak ditemukan." };
    const rows = sheet.getDataRange().getValues();
    const isEdit = data.id !== null && data.id !== undefined && data.id !== '';
    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const rIdx = i + 1;
          sheet.getRange(rIdx, 2).setValue(data.appName); sheet.getRange(rIdx, 3).setValue(data.url || "");
          sheet.getRange(rIdx, 4).setValue(data.username || ""); sheet.getRange(rIdx, 5).setValue(data.password);
          sheet.getRange(rIdx, 6).setValue(data.category || ""); sheet.getRange(rIdx, 7).setValue(data.location || "");
          sheet.getRange(rIdx, 8).setValue(data.notes || ""); sheet.getRange(rIdx, 10).setValue(data.isActive === true || data.isActive === 'true');
          return { success: true, message: `Credential '${data.appName}' berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID Credential tidak ditemukan." };
    } else {
      let nextId = "PW-0001";
      if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.toString().includes("PW-")) nextId = "PW-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
      sheet.appendRow([nextId, data.appName, data.url || "", data.username || "", data.password, data.category || "", data.location || "", data.notes || "", new Date(), data.isActive === true || data.isActive === 'true' || data.isActive === undefined ? true : false]);
      return { success: true, message: `Credential '${data.appName}' (${nextId}) berhasil disimpan ke vault.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

function saveMasterItemRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MASTER_ITEMS);
    if (!sheet) return { success: false, message: "Sheet master_items tidak ditemukan." };
    const rows = sheet.getDataRange().getValues();
    const isEdit = data.id !== null && data.id !== undefined && data.id !== '';
    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const rIdx = i + 1;
          sheet.getRange(rIdx, 2).setValue(data.itemName); sheet.getRange(rIdx, 3).setValue(data.category || "");
          sheet.getRange(rIdx, 4).setValue(data.brand || ""); sheet.getRange(rIdx, 5).setValue(data.serialNumber || "");
          sheet.getRange(rIdx, 6).setValue(data.estPrice !== '' ? Number(data.estPrice) : "");
          return { success: true, message: `Item '${data.itemName}' berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID Item tidak ditemukan." };
    } else {
      let nextId = "ITM-0001";
      if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.toString().includes("ITM-")) nextId = "ITM-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
      sheet.appendRow([nextId, data.itemName, data.category || "", data.brand || "", data.serialNumber || "", data.estPrice !== '' ? Number(data.estPrice) : "", "ACTIVE", new Date()]);
      return { success: true, message: `Item baru '${data.itemName}' (${nextId}) berhasil ditambahkan.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

// ================= MASTER USERS ENGINE (NEW) =================
function saveMasterUserRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MASTER_USERS);
    if (!sheet) return { success: false, message: "Sheet master_users tidak ditemukan. Jalankan setupDatabase()." };
    const rows = sheet.getDataRange().getValues();
    const isEdit = data.id !== null && data.id !== undefined && data.id !== '';
    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const rIdx = i + 1;
          sheet.getRange(rIdx, 2).setValue(data.fullName);
          sheet.getRange(rIdx, 3).setValue(data.department || "");
          sheet.getRange(rIdx, 4).setValue(data.email || "");
          sheet.getRange(rIdx, 5).setValue(data.phone || "");
          return { success: true, message: `User '${data.fullName}' berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID User tidak ditemukan." };
    } else {
      let nextId = "USR-0001";
      if (rows.length > 1) { const lastId = rows[rows.length-1][0]; if (lastId && lastId.toString().includes("USR-")) nextId = "USR-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0'); }
      sheet.appendRow([nextId, data.fullName, data.department || "", data.email || "", data.phone || "", "ACTIVE", new Date()]);
      return { success: true, message: `User '${data.fullName}' (${nextId}) berhasil ditambahkan.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}

// ================= OFFICE LICENSE ENGINE (NEW) =================
function saveLicenseRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LICENSES);
    if (!sheet) return { success: false, message: "Sheet office_licenses tidak ditemukan. Jalankan setupDatabase()." };
    const rows = sheet.getDataRange().getValues();
    const isEdit = data.id !== null && data.id !== undefined && data.id !== '';

    // Duplicate check: Unique Key harus unik
    if (data.uniqueKey && data.uniqueKey.trim() !== '') {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][5].toString().toLowerCase() === data.uniqueKey.trim().toLowerCase() && rows[i][12] !== 'DELETED') {
          if (!isEdit || (isEdit && rows[i][0] !== data.id)) {
            return { success: false, message: `Unique Key '${data.uniqueKey}' sudah terdaftar di sistem!` };
          }
        }
      }
    }

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const rIdx = i + 1;
          sheet.getRange(rIdx, 2).setValue(data.partNumber || "");
          sheet.getRange(rIdx, 3).setValue(data.productName || "");
          sheet.getRange(rIdx, 4).setValue(data.qty ? Number(data.qty) : 1);
          sheet.getRange(rIdx, 5).setValue(data.productKey || "");
          sheet.getRange(rIdx, 6).setValue(data.uniqueKey || "");
          sheet.getRange(rIdx, 7).setValue(data.activationDate ? new Date(data.activationDate) : "");
          sheet.getRange(rIdx, 8).setValue(data.hostname || "");
          sheet.getRange(rIdx, 9).setValue(data.user || "");
          sheet.getRange(rIdx, 10).setValue(data.itemId || "");
          sheet.getRange(rIdx, 11).setValue(data.laptopBrandType || "");
          sheet.getRange(rIdx, 12).setValue(data.akunAdmin || "");
          return { success: true, message: `License '${data.productName}' berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID License tidak ditemukan." };
    } else {
      let nextId = "LIC-0001";
      if (rows.length > 1) {
        const lastId = rows[rows.length-1][0];
        if (lastId && lastId.toString().includes("LIC-")) {
          nextId = "LIC-" + String(parseInt(lastId.split("-")[1], 10) + 1).padStart(4, '0');
        }
      }
      sheet.appendRow([
        nextId,
        data.partNumber    || "",
        data.productName   || "",
        data.qty           ? Number(data.qty) : 1,
        data.productKey    || "",
        data.uniqueKey     || "",
        data.activationDate ? new Date(data.activationDate) : "",
        data.hostname      || "",
        data.user          || "",
        data.itemId        || "",
        data.laptopBrandType || "",
        data.akunAdmin     || "",
        "ACTIVE",
        new Date()
      ]);
      return { success: true, message: `License '${data.productName}' (${nextId}) berhasil didaftarkan.` };
    }
  } catch (error) { return { success: false, message: error.toString() }; } finally { lock.releaseLock(); }
}