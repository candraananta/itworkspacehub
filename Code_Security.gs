/**
 * ============================================================
 *  IT Tools ERP — Code_Security.gs
 *  Security · Password Vault Engine
 * ============================================================
 */

function savePasswordRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PASSWORDS);
    if (!sheet) return { success: false, message: "Sheet passwords tidak ditemukan." };
    const rows   = sheet.getDataRange().getValues();
    const isEdit = !!(data.id);

    // Normalize isActive
    const isActive = (data.isActive === true || data.isActive === 'true' || data.isActive === undefined);

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 2).setValue(data.appName);
          sheet.getRange(r, 3).setValue(data.url      || "");
          sheet.getRange(r, 4).setValue(data.username || "");
          sheet.getRange(r, 5).setValue(data.password);
          sheet.getRange(r, 6).setValue(data.category || "");
          sheet.getRange(r, 7).setValue(data.location || "");
          sheet.getRange(r, 8).setValue(data.notes    || "");
          sheet.getRange(r, 10).setValue(isActive);
          return { success: true, message: `Credential "${data.appName}" berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID Credential tidak ditemukan." };
    }

    const nextId = generateId(sheet, "PW-");
    sheet.appendRow([
      nextId, data.appName, data.url || "", data.username || "",
      data.password, data.category || "", data.location || "",
      data.notes || "", new Date(), isActive
    ]);
    return { success: true, message: `Credential "${data.appName}" (${nextId}) berhasil disimpan ke vault.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}