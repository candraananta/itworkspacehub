/**
 * ============================================================
 *  IT Tools ERP — Code_Checklists.gs
 *  Audit & Compliance · Checklist Perangkat IT
 * ============================================================
 *  Form Code Pattern:  AKR-{LOC}-NT-{SEQ}
 *   AKR = Audit Checklist (fixed)
 *   LOC = 3 huruf dari nama Site (uppercase)
 *   NT  = Network Tools (fixed)
 *   SEQ = 3-digit sequential per site
 */

function _generateChecklistFormCode_(site) {
  const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CHECKLISTS);
  const rows     = sheet.getDataRange().getValues();
  const siteCode = (site || "GEN")
    .replace(/[^A-Za-z]/g, '')
    .substring(0, 3)
    .toUpperCase() || "GEN";

  let maxSeq = 0;
  const marker = '-' + siteCode + '-NT-';
  for (let i = 1; i < rows.length; i++) {
    const fc = rows[i][1].toString();
    if (fc.includes(marker)) {
      const seq = parseInt(fc.split('-').pop(), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }
  return `AKR-${siteCode}-NT-${String(maxSeq + 1).padStart(3, '0')}`;
}

function saveChecklistRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CHECKLISTS);
    if (!sheet) return { success: false, message: "Sheet checklist_forms tidak ditemukan. Jalankan setupDatabase()." };
    const rows     = sheet.getDataRange().getValues();
    const isEdit   = !!(data.id);
    const itemsStr = JSON.stringify(data.items || []);

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 3).setValue(data.judul   || "");
          sheet.getRange(r, 4).setValue(data.tanggal ? new Date(data.tanggal) : new Date());
          sheet.getRange(r, 5).setValue(data.site    || "");
          sheet.getRange(r, 6).setValue(itemsStr);
          return { success: true, message: `Form Checklist "${data.judul}" berhasil diperbarui.` };
        }
      }
      return { success: false, message: "ID Checklist tidak ditemukan." };
    }

    const nextId   = generateId(sheet, "CHK-");
    const formCode = _generateChecklistFormCode_(data.site);
    sheet.appendRow([
      nextId, formCode, data.judul || "",
      data.tanggal ? new Date(data.tanggal) : new Date(),
      data.site || "", itemsStr,
      "ACTIVE", new Date()
    ]);
    return {
      success:  true,
      message:  `Form Checklist "${formCode}" berhasil dibuat.`,
      formCode: formCode
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}