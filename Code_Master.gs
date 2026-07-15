/**
 * ============================================================
 *  IT Tools ERP — Code_Master.gs
 *  Master Data · Locations · Categories · Items · Users
 *  Universal Soft-Delete Engine
 * ============================================================
 */

// ── Locations & Categories (generic) ─────────────────────────
function saveMasterRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheetName   = data.targetSheet;
    const sheet       = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: `Sheet "${sheetName}" tidak ditemukan.` };

    const rows        = sheet.getDataRange().getValues();
    const isEdit      = !!(data.idField);
    const targetValue = data.valueField.trim();

    // Duplicate guard
    for (let i = 1; i < rows.length; i++) {
      if (
        rows[i][1].toString().toLowerCase() === targetValue.toLowerCase() &&
        rows[i][2] !== "DELETED"
      ) {
        if (!isEdit || (isEdit && rows[i][0] !== data.idField))
          return { success: false, message: `Entitas "${targetValue}" sudah aktif!` };
      }
    }

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.idField) {
          sheet.getRange(i + 1, 2).setValue(targetValue);
          return { success: true, message: "Data master berhasil diperbarui." };
        }
      }
      return { success: false, message: "ID tidak ditemukan." };
    }

    // Auto-prefix berdasarkan sheetName
    const prefix = sheetName === SHEETS.LOCATIONS ? "LC-" : "CT-";
    const nextId = generateId(sheet, prefix);
    sheet.appendRow([nextId, targetValue, "ACTIVE"]);
    return { success: true, message: `Master ${nextId} ditambahkan.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── Master Items ──────────────────────────────────────────────
function saveMasterItemRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MASTER_ITEMS);
    if (!sheet) return { success: false, message: "Sheet master_items tidak ditemukan." };
    const rows   = sheet.getDataRange().getValues();
    const isEdit = !!(data.id);

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 2).setValue(data.itemName);
          sheet.getRange(r, 3).setValue(data.category     || "");
          sheet.getRange(r, 4).setValue(data.brand        || "");
          sheet.getRange(r, 5).setValue(data.serialNumber || "");
          sheet.getRange(r, 6).setValue(data.estPrice !== '' && data.estPrice != null ? Number(data.estPrice) : "");
          return { success: true, message: `Item "${data.itemName}" berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID Item tidak ditemukan." };
    }

    const nextId = generateId(sheet, "ITM-");
    sheet.appendRow([
      nextId, data.itemName, data.category || "", data.brand || "",
      data.serialNumber || "",
      data.estPrice !== '' && data.estPrice != null ? Number(data.estPrice) : "",
      "ACTIVE", new Date()
    ]);
    return { success: true, message: `Item "${data.itemName}" (${nextId}) berhasil ditambahkan.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── Master Users ──────────────────────────────────────────────
function saveMasterUserRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MASTER_USERS);
    if (!sheet) return { success: false, message: "Sheet master_users tidak ditemukan. Jalankan setupDatabase()." };
    const rows   = sheet.getDataRange().getValues();
    const isEdit = !!(data.id);

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 2).setValue(data.fullName);
          sheet.getRange(r, 3).setValue(data.department || "");
          sheet.getRange(r, 4).setValue(data.email      || "");
          sheet.getRange(r, 5).setValue(data.phone      || "");
          return { success: true, message: `User "${data.fullName}" berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID User tidak ditemukan." };
    }

    const nextId = generateId(sheet, "USR-");
    sheet.appendRow([
      nextId, data.fullName, data.department || "",
      data.email || "", data.phone || "",
      "ACTIVE", new Date()
    ]);
    return { success: true, message: `User "${data.fullName}" (${nextId}) berhasil ditambahkan.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── Universal Soft-Delete (dipakai semua modul) ───────────────
function executeDeleteRecord(sheetName, idToDelete) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: "Sheet tidak ditemukan." };

    const data        = sheet.getDataRange().getValues();
    const statusColIdx = data[0].indexOf("Status") + 1;
    if (statusColIdx === 0) return { success: false, message: "Kolom Status tidak ada di tabel ini." };

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === idToDelete.toString()) {
        sheet.getRange(i + 1, statusColIdx).setValue("DELETED");
        return { success: true, message: `Data ${idToDelete} berhasil diarsip (Soft Deleted).` };
      }
    }
    return { success: false, message: "Data tidak ditemukan di server." };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}