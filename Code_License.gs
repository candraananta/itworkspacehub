/**
 * ============================================================
 *  IT Tools ERP — Code_Licenses.gs
 *  Office 2024 License Manager · Activation Tracking
 * ============================================================
 */

function saveLicenseRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LICENSES);
    if (!sheet) return { success: false, message: "Sheet office_licenses tidak ditemukan. Jalankan setupDatabase()." };
    const rows   = sheet.getDataRange().getValues();
    const isEdit = !!(data.id);

    // ── Unique Key duplicate guard ───────────────────────────
    if (data.uniqueKey && data.uniqueKey.trim() !== '') {
      for (let i = 1; i < rows.length; i++) {
        if (
          rows[i][5].toString().toLowerCase() === data.uniqueKey.trim().toLowerCase() &&
          rows[i][12] !== 'DELETED'
        ) {
          if (!isEdit || (isEdit && rows[i][0] !== data.id)) {
            return { success: false, message: `Unique Key "${data.uniqueKey}" sudah terdaftar!` };
          }
        }
      }
    }

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 2).setValue(data.partNumber     || "");
          sheet.getRange(r, 3).setValue(data.productName    || "");
          sheet.getRange(r, 4).setValue(data.qty ? Number(data.qty) : 1);
          sheet.getRange(r, 5).setValue(data.productKey     || "");
          sheet.getRange(r, 6).setValue(data.uniqueKey      || "");
          sheet.getRange(r, 7).setValue(data.activationDate ? new Date(data.activationDate) : "");
          sheet.getRange(r, 8).setValue(data.hostname       || "");
          sheet.getRange(r, 9).setValue(data.user           || "");
          sheet.getRange(r, 10).setValue(data.itemId        || "");
          sheet.getRange(r, 11).setValue(data.laptopBrandType || "");
          sheet.getRange(r, 12).setValue(data.akunAdmin     || "");
          return { success: true, message: `License "${data.productName}" berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID License tidak ditemukan." };
    }

    const nextId = generateId(sheet, "LIC-");
    sheet.appendRow([
      nextId,
      data.partNumber      || "",
      data.productName     || "",
      data.qty ? Number(data.qty) : 1,
      data.productKey      || "",
      data.uniqueKey       || "",
      data.activationDate  ? new Date(data.activationDate) : "",
      data.hostname        || "",
      data.user            || "",
      data.itemId          || "",
      data.laptopBrandType || "",
      data.akunAdmin       || "",
      "ACTIVE",
      new Date()
    ]);
    return { success: true, message: `License "${data.productName}" (${nextId}) berhasil didaftarkan.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}