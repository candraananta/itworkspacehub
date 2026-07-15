/**
 * ============================================================
 *  IT Tools ERP — Code_KPI.gs
 *  KPI Management · Quality Delivery · Training Records
 * ============================================================
 */

// ── KPI Projects ─────────────────────────────────────────────
function saveKpiProjectRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KPI_PROJECTS);
    const rows      = sheet.getDataRange().getValues();
    const isEdit    = !!(data.id);
    const photosStr = JSON.stringify(data.photos || []);

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 2).setValue(data.date ? new Date(data.date) : new Date());
          sheet.getRange(r, 3).setValue(data.projectName);
          sheet.getRange(r, 4).setValue(data.location);
          sheet.getRange(r, 5).setValue(data.description || "");
          sheet.getRange(r, 6).setValue(photosStr);
          return { success: true, message: `Project log "${data.projectName}" diperbarui.` };
        }
      }
      return { success: false, message: "ID Project tidak ditemukan." };
    }

    const nextId = generateId(sheet, "KP-");
    sheet.appendRow([
      nextId,
      data.date ? new Date(data.date) : new Date(),
      data.projectName, data.location,
      data.description || "", photosStr,
      "ACTIVE", new Date()
    ]);
    return { success: true, message: `KPI Delivery Project "${data.projectName}" berhasil direkam.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── KPI Trainings ─────────────────────────────────────────────
function saveKpiTrainingRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.KPI_TRAININGS);
    const rows      = sheet.getDataRange().getValues();
    const isEdit    = !!(data.id);
    const photosStr = JSON.stringify(data.photos || []);

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 2).setValue(data.date ? new Date(data.date) : new Date());
          sheet.getRange(r, 3).setValue(data.title);
          sheet.getRange(r, 4).setValue(data.topic);
          sheet.getRange(r, 5).setValue(data.method);
          sheet.getRange(r, 6).setValue(data.location);
          sheet.getRange(r, 7).setValue(data.feedback || "");
          sheet.getRange(r, 8).setValue(photosStr);
          return { success: true, message: `Data Training "${data.title}" berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID Training tidak ditemukan." };
    }

    const nextId = generateId(sheet, "TR-");
    sheet.appendRow([
      nextId,
      data.date ? new Date(data.date) : new Date(),
      data.title, data.topic, data.method, data.location,
      data.feedback || "", photosStr,
      "ACTIVE", new Date()
    ]);
    return { success: true, message: `Training Record "${data.title}" berhasil dimasukkan.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}