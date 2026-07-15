/**
 * ============================================================
 *  IT Tools ERP — Code_Main.gs
 *  Entry Point · Constants · DB Setup · Initial Data Loader
 * ============================================================
 */

// ── Sheet name constants (dipakai di semua module) ──────────
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
  MASTER_USERS:  "master_users",
  LICENSES:      "office_licenses",
  CHECKLISTS:    "checklist_forms"
};

// ── doGet: serve WebApp ──────────────────────────────────────
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('IT Tools ERP - Enterprise Ecosystem')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(
      "<h3>Critical UI Loader Error:</h3><pre>" + error.toString() + "</pre>"
    );
  }
}

// ── include() helper untuk HtmlService.createTemplateFromFile ──
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Universal sheet reader (dipakai semua module) ────────────
function getSheetData(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];

    const headers        = rows[0];
    const formattedData  = [];
    const statusColIndex = headers.indexOf("Status");

    for (let i = 1; i < rows.length; i++) {
      if (statusColIndex !== -1 && rows[i][statusColIndex] === "DELETED") continue;
      const obj = {};
      headers.forEach((header, colIndex) => {
        let val = rows[i][colIndex];
        if (val instanceof Date) {
          obj[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        } else if (header === 'Photos' || header === 'Devices' || header === 'Steps_JSON' || header === 'Items_JSON') {
          try   { obj[header] = val ? JSON.parse(val) : (header === 'Devices' ? {} : []); }
          catch (e) { obj[header] = header === 'Devices' ? {} : []; }
        } else if (header === 'Is Active') {
          obj[header] = (val === true || val === 1 || String(val).toUpperCase() === 'TRUE');
        } else {
          obj[header] = val;
        }
      });
      formattedData.push(obj);
    }
    return formattedData;
  } catch (error) {
    Logger.log('getSheetData error [' + sheetName + ']: ' + error.toString());
    return [];
  }
}

// ── getInitialData: dipanggil dari frontend via google.script.run ──
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
      masterUsers:  getSheetData(SHEETS.MASTER_USERS),
      licenses:     getSheetData(SHEETS.LICENSES),
      checklists:   getSheetData(SHEETS.CHECKLISTS)
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ── ID generator helper (dipakai di semua Code_*.gs) ─────────
function generateId(sheet, prefix) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return prefix + "0001";
  const lastId = rows[rows.length - 1][0].toString();
  if (!lastId.includes(prefix.replace('-',''))) return prefix + "0001";
  const num = parseInt(lastId.split("-").pop(), 10);
  return prefix + String((isNaN(num) ? 0 : num) + 1).padStart(4, '0');
}

// ── setupDatabase: buat semua sheet jika belum ada (auto-patch) ──
function setupDatabase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const ensure = (name, headers, bgColor) => {
      let sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
        sheet.getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setFontWeight("bold")
          .setBackground(bgColor)
          .setFontColor("#ffffff");
        sheet.setFrozenRows(1);
      }
      return sheet;
    };

    ensure(SHEETS.INCIDENTS,     ["ID","Tanggal","Location","Category","Description","Status","Photo Before","Photo After","Timestamp"],                            "#000000");
    ensure(SHEETS.LOCATIONS,     ["ID","Location Name","Status"],                                                                                                   "#334155");
    ensure(SHEETS.CATEGORIES,    ["ID","Category Name","Status"],                                                                                                   "#334155");
    ensure(SHEETS.VLANS,         ["ID","VLAN ID","VLAN Name","Network","Gateway","Switch Target","Location","Notes","Status","Timestamp"],                          "#1e3a8a");
    ensure(SHEETS.RACKS,         ["ID","Rack Name","Location","Size","Photos","Devices","Status","Timestamp"],                                                       "#111827");
    ensure(SHEETS.KPI_PROJECTS,  ["ID","Tanggal","Project Name","Location","Description","Photos","Status","Timestamp"],                                            "#7c3aed");
    ensure(SHEETS.KPI_TRAININGS, ["ID","Tanggal","Judul Training","Topic","Method","Location","Feedback","Photos","Status","Timestamp"],                            "#db2777");
    ensure(SHEETS.MASTER_ITEMS,  ["ID","Item Name","Category","Brand","Serial Number","est_price","Status","Timestamp"],                                            "#ea580c");
    ensure(SHEETS.PASSWORDS,     ["ID","App Name","URL","Username","Password","Category","Location","Notes","Timestamp","Is Active"],                              "#0d9488");
    ensure(SHEETS.MASTER_USERS,  ["ID","Full Name","Department","Email","Phone","Status","Timestamp"],                                                              "#1e40af");
    ensure(SHEETS.LICENSES,      ["ID","Part Number","Product Name","Qty","Product Key","Unique Key","Tanggal Aktivasi","Hostname","User","Item ID","Laptop Brand Type","Akun Admin","Status","Timestamp"], "#1d4ed8");
    ensure(SHEETS.CHECKLISTS,    ["ID","Form Code","Judul","Tanggal","Site","Items_JSON","Status","Timestamp"],                                                     "#15803d");

    // Auto-patch: Wiki — tambah kolom PDF_Attachment bila belum ada
    let wikiSheet = ss.getSheetByName(SHEETS.WIKI);
    if (!wikiSheet) {
      wikiSheet = ss.insertSheet(SHEETS.WIKI);
      wikiSheet.getRange(1, 1, 1, 7)
        .setValues([["ID","Title","Category","Steps_JSON","PDF_Attachment","Status","Timestamp"]])
        .setFontWeight("bold").setBackground("#059669").setFontColor("#ffffff");
      wikiSheet.setFrozenRows(1);
    } else {
      const headers = wikiSheet.getRange(1, 1, 1, wikiSheet.getLastColumn()).getValues()[0];
      if (headers.indexOf("PDF_Attachment") === -1) {
        const statusIdx = headers.indexOf("Status") + 1;
        if (statusIdx > 0) {
          wikiSheet.insertColumnBefore(statusIdx);
          wikiSheet.getRange(1, statusIdx).setValue("PDF_Attachment")
            .setFontWeight("bold").setBackground("#059669").setFontColor("#ffffff");
        }
      }
    }

    return "✅ SINKRONISASI DATABASE SUKSES! (Auto-patch aktif, semua tabel aman)";
  } catch (error) {
    return "❌ GAGAL: " + error.toString();
  }
}