/**
 * ============================================================
 *  IT Tools ERP — Code_Knowledge.gs
 *  Wiki Knowledge Base · SOP & Troubleshooting Docs
 * ============================================================
 */

function saveWikiRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.WIKI);
    const rows     = sheet.getDataRange().getValues();
    const headers  = rows[0];
    const isEdit   = !!(data.id);
    const stepsStr = JSON.stringify(data.steps || []);

    // Cari index kolom PDF_Attachment (sudah di-patch oleh setupDatabase)
    const pdfCol = headers.indexOf("PDF_Attachment") + 1; // 1-based

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 2).setValue(data.title);
          sheet.getRange(r, 3).setValue(data.category);
          sheet.getRange(r, 4).setValue(stepsStr);
          if (pdfCol > 0) sheet.getRange(r, pdfCol).setValue(data.pdf || "");
          return { success: true, message: `Artikel Wiki "${data.title}" berhasil di-update.` };
        }
      }
      return { success: false, message: "ID Wiki tidak ditemukan." };
    }

    // Insert baru dengan array agar kolom PDF_Attachment tetap sejajar
    const nextId  = generateId(sheet, "WK-");
    const newRow  = Array(headers.length).fill("");
    newRow[headers.indexOf("ID")]        = nextId;
    newRow[headers.indexOf("Title")]     = data.title;
    newRow[headers.indexOf("Category")]  = data.category;
    newRow[headers.indexOf("Steps_JSON")] = stepsStr;
    if (pdfCol > 0) newRow[pdfCol - 1]  = data.pdf || "";
    newRow[headers.indexOf("Status")]    = "ACTIVE";
    newRow[headers.indexOf("Timestamp")] = new Date();
    sheet.appendRow(newRow);

    return { success: true, message: `Wiki "${data.title}" berhasil dipublikasikan.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}