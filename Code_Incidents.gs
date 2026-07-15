/**
 * ============================================================
 *  IT Tools ERP — Code_Incidents.gs
 *  Incident Reports · Create · Update · Close
 * ============================================================
 */

function createIncident(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INCIDENTS);
    const nextId = generateId(sheet, "INC-");
    sheet.appendRow([
      nextId,
      data.date ? new Date(data.date) : new Date(),
      data.location,
      data.category,
      data.description || "",
      "Waiting",
      data.photoBefore || "",
      "",
      new Date()
    ]);
    return { success: true, message: `Log kendala teknis ${nextId} berhasil dibuat.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateIncident(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INCIDENTS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        const r = i + 1;
        sheet.getRange(r, 2).setValue(data.date ? new Date(data.date) : new Date());
        sheet.getRange(r, 3).setValue(data.location);
        sheet.getRange(r, 4).setValue(data.category);
        sheet.getRange(r, 5).setValue(data.description || "");
        sheet.getRange(r, 7).setValue(data.photoBefore || "");
        return { success: true, message: `Incident ${data.id} berhasil diperbarui.` };
      }
    }
    return { success: false, message: "ID Incident tidak ditemukan." };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function closeIncident(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.INCIDENTS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        sheet.getRange(i + 1, 6).setValue("Closed");
        sheet.getRange(i + 1, 8).setValue(data.photoAfter || "");
        return { success: true, message: `Incident ${data.id} ditutup (Selesai Penanganan).` };
      }
    }
    return { success: false, message: "ID tidak ditemukan." };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}