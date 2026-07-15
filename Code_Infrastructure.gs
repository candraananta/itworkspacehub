/**
 * ============================================================
 *  IT Tools ERP — Code_Infrastructure.gs
 *  Infrastructure · VLAN Setup · Rack Management
 * ============================================================
 */

// ── VLAN ─────────────────────────────────────────────────────
function saveVlanRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.VLANS);
    const rows   = sheet.getDataRange().getValues();
    const isEdit = !!(data.id);

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 2).setValue(data.vlanId);
          sheet.getRange(r, 3).setValue(data.name);
          sheet.getRange(r, 4).setValue(data.network);
          sheet.getRange(r, 5).setValue(data.gateway);
          sheet.getRange(r, 6).setValue(data.switchName);
          sheet.getRange(r, 7).setValue(data.location);
          sheet.getRange(r, 8).setValue(data.notes || "");
          return { success: true, message: `Topologi VLAN ${data.vlanId} diperbarui.` };
        }
      }
      return { success: false, message: "ID VLAN tidak ditemukan." };
    }

    const nextId = generateId(sheet, "VL-");
    sheet.appendRow([
      nextId, data.vlanId, data.name, data.network,
      data.gateway, data.switchName, data.location,
      data.notes || "", "ACTIVE", new Date()
    ]);
    return { success: true, message: `VLAN ID ${data.vlanId} berhasil di-deploy.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── Rack ─────────────────────────────────────────────────────
function saveRackRow(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet      = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.RACKS);
    const rows       = sheet.getDataRange().getValues();
    const isEdit     = !!(data.id);
    const photosStr  = JSON.stringify(data.photos  || []);
    const devicesStr = JSON.stringify(data.devices || {});

    if (isEdit) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
          const r = i + 1;
          sheet.getRange(r, 2).setValue(data.name);
          sheet.getRange(r, 3).setValue(data.location);
          sheet.getRange(r, 4).setValue(data.size);
          sheet.getRange(r, 5).setValue(photosStr);
          sheet.getRange(r, 6).setValue(devicesStr);
          return { success: true, message: `Arsitektur Server Rack "${data.name}" disimpan.` };
        }
      }
      return { success: false, message: "ID Rack tidak ditemukan." };
    }

    const nextId = generateId(sheet, "RCK-");
    sheet.appendRow([
      nextId, data.name, data.location, data.size,
      photosStr, devicesStr, "ACTIVE", new Date()
    ]);
    return { success: true, message: `Registrasi Server Rack "${data.name}" berhasil.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}