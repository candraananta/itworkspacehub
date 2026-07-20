/**
 * ============================================================
 *  IT Tools ERP — Code_ChecklistReport.gs
 *  Checklist Perangkat IT · Report · Export Sheets · PDF
 * ============================================================
 *
 *  Public functions:
 *    getChecklistReport(filters)         → data agregasi + detail
 *    exportChecklistToSheets(checklistId, signConfig) → URL Sheets
 *    exportChecklistToPdf(checklistId, signConfig)    → base64 PDF
 */

// ── Helper: hitung total row dalam satu checklist ────────────
function _countRows_(items) {
  return (items || []).reduce((sum, cat) => sum + (cat.rows || []).length, 0);
}

// ── Helper: hitung row dengan status tertentu ─────────────────
function _countByStatus_(items, status) {
  return (items || []).reduce((sum, cat) =>
    sum + (cat.rows || []).filter(r => r.statusBaik === status).length, 0);
}

// ═══════════════════════════════════════════════════════════════
//  getChecklistReport(filters)
//  filters: { site, status, month, year }
// ═══════════════════════════════════════════════════════════════
function getChecklistReport(filters) {
  try {
    filters = filters || {};
    const all = getSheetData(SHEETS.CHECKLISTS);

    // ── Apply filters ────────────────────────────────────────
    let data = all.filter(c => {
      if (filters.site && c.Site !== filters.site) return false;
      if (filters.status && c.Status !== filters.status) return false;
      if (filters.year || filters.month) {
        const d = c.Tanggal ? new Date(c.Tanggal) : null;
        if (!d) return false;
        if (filters.year  && d.getFullYear().toString() !== filters.year.toString()) return false;
        if (filters.month && (d.getMonth() + 1).toString() !== filters.month.toString()) return false;
      }
      return true;
    });

    // ── Summary stats ────────────────────────────────────────
    const totalForms   = data.length;
    let totalParams    = 0, totalBaik = 0, totalTidakBaik = 0, totalBelum = 0;

    data.forEach(c => {
      const items = c.Items_JSON || [];
      items.forEach(cat => {
        (cat.rows || []).forEach(r => {
          totalParams++;
          if (r.statusBaik === 'Baik')        totalBaik++;
          else if (r.statusBaik === 'Tidak Baik') totalTidakBaik++;
          else                                    totalBelum++;
        });
      });
    });

    const healthRate = totalParams > 0 ? Math.round((totalBaik / totalParams) * 100) : 0;

    // ── By Site ──────────────────────────────────────────────
    const siteMap = {};
    data.forEach(c => {
      const s = c.Site || 'Unknown';
      if (!siteMap[s]) siteMap[s] = { forms: 0, baik: 0, tidakBaik: 0, belum: 0, total: 0 };
      siteMap[s].forms++;
      const items = c.Items_JSON || [];
      items.forEach(cat => {
        (cat.rows || []).forEach(r => {
          siteMap[s].total++;
          if (r.statusBaik === 'Baik')         siteMap[s].baik++;
          else if (r.statusBaik === 'Tidak Baik') siteMap[s].tidakBaik++;
          else                                    siteMap[s].belum++;
        });
      });
    });
    const bySite = Object.entries(siteMap)
      .sort((a, b) => b[1].forms - a[1].forms)
      .map(([site, v]) => ({ site, ...v, healthRate: v.total > 0 ? Math.round((v.baik / v.total) * 100) : 0 }));

    // ── By Category (aggregate semua form yang difilter) ─────
    const catMap = {};
    data.forEach(c => {
      (c.Items_JSON || []).forEach(cat => {
        const cn = cat.categoryName || 'Unknown';
        if (!catMap[cn]) catMap[cn] = { baik: 0, tidakBaik: 0, belum: 0, total: 0 };
        (cat.rows || []).forEach(r => {
          catMap[cn].total++;
          if (r.statusBaik === 'Baik')         catMap[cn].baik++;
          else if (r.statusBaik === 'Tidak Baik') catMap[cn].tidakBaik++;
          else                                    catMap[cn].belum++;
        });
      });
    });
    const byCategory = Object.entries(catMap)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([cat, v]) => ({ cat, ...v, healthRate: v.total > 0 ? Math.round((v.baik / v.total) * 100) : 0 }));

    // ── Issues: parameter "Tidak Baik" dari semua form ───────
    const issues = [];
    data.forEach(c => {
      (c.Items_JSON || []).forEach(cat => {
        (cat.rows || []).forEach(r => {
          if (r.statusBaik === 'Tidak Baik') {
            issues.push({
              formCode:    c['Form Code'] || c.ID,
              site:        c.Site || '-',
              tanggal:     (c.Tanggal || '').toString().substring(0, 10),
              category:    cat.categoryName || '-',
              parameter:   r.parameter || '-',
              standarLimit: r.standarLimit || '-',
              potensiResiko: r.potensiResiko || '-',
              keterangan:  r.keterangan || '-'
            });
          }
        });
      });
    });

    // ── List ringkasan form ───────────────────────────────────
    const formList = data
      .sort((a, b) => String(b.ID).localeCompare(String(a.ID)))
      .map(c => ({
        id:        c.ID,
        formCode:  c['Form Code'] || c.ID,
        judul:     c.Judul || '-',
        site:      c.Site || '-',
        tanggal:   (c.Tanggal || '').toString().substring(0, 10),
        status:    c.Status || '-',
        totalParam: _countRows_(c.Items_JSON),
        baik:      _countByStatus_(c.Items_JSON, 'Baik'),
        tidakBaik: _countByStatus_(c.Items_JSON, 'Tidak Baik'),
        healthRate: _countRows_(c.Items_JSON) > 0
          ? Math.round((_countByStatus_(c.Items_JSON, 'Baik') / _countRows_(c.Items_JSON)) * 100) : 0
      }));

    // ── Available years ───────────────────────────────────────
    const years = [...new Set(
      all.map(c => { const d = c.Tanggal ? new Date(c.Tanggal) : null; return d ? d.getFullYear().toString() : null; })
        .filter(Boolean)
    )].sort().reverse();

    // ── Available sites ───────────────────────────────────────
    const sites = [...new Set(all.map(c => c.Site).filter(Boolean))].sort();

    return {
      success: true,
      summary: { totalForms, totalParams, totalBaik, totalTidakBaik, totalBelum, healthRate },
      bySite, byCategory, issues, formList, years, sites,
      appliedFilters: filters,
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
//  _buildChecklistSheet_(ss, chk, signConfig)
//  Buat satu sheet berisi form checklist lengkap + signature
//  Dipakai oleh exportChecklistToSheets & exportChecklistToPdf
// ═══════════════════════════════════════════════════════════════
function _buildChecklistSheet_(ss, chk, signConfig) {
  const sign       = signConfig || {};
  const itName     = sign.itName     || 'IT Department';
  const itTitle    = sign.itTitle    || 'IT Staff';
  const customName = sign.customName || '';
  const customTitle= sign.customTitle|| 'Dept. Head / Supervisor';

  const BLUE   = '#1A3A66';
  const LBLUE  = '#2B6CB0';
  const WHITE  = '#FFFFFF';
  const HEADER = '#EBF4FF';
  const GREEN  = '#D1FAE5';
  const RED    = '#FEE2E2';
  const AMBER  = '#FEF3C7';
  const GRAY   = '#F8FAFC';

  // Nama sheet = Form Code (max 31 karakter, karakter khusus diganti)
  const sheetName = (chk['Form Code'] || chk.ID || 'Checklist').replace(/[\\\/\?\*\[\]]/g, '-').substring(0, 31);
  const sheet     = ss.insertSheet(sheetName);

  let row = 1;

  // ── HEADER ──────────────────────────────────────────────────
  sheet.getRange(row, 1, 1, 8).merge()
    .setValue('FORM CHECKLIST PERANGKAT IT')
    .setBackground(BLUE).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 36);
  row++;

  // ── Info header ──────────────────────────────────────────────
  const infoData = [
    ['No. Form',   chk['Form Code'] || chk.ID, '', '', 'Tanggal', (chk.Tanggal || '').toString().substring(0, 10), '', ''],
    ['Judul',      chk.Judul || '-',            '', '', 'Site',    chk.Site || '-',                                 '', ''],
    ['Status',     chk.Status || 'ACTIVE',      '', '', '',        '',                                              '', '']
  ];
  infoData.forEach(r => {
    sheet.getRange(row, 1, 1, 8).setValues([r]);
    sheet.getRange(row, 1).setFontWeight('bold').setBackground(HEADER);
    sheet.getRange(row, 5).setFontWeight('bold').setBackground(HEADER);
    sheet.getRange(row, 2, 1, 3).merge();
    sheet.getRange(row, 6, 1, 3).merge();
    row++;
  });
  row++;

  // ── TABLE HEADER ─────────────────────────────────────────────
  const colHeaders = ['No.', 'Kategori Pemeriksaan', 'Parameter', 'Standar Limit', 'Potensi Resiko', 'Frekuensi', 'Status', 'Keterangan'];
  const colWidths  = [35,    180,                     160,         140,              200,               80,          90,       160];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.getRange(row, 1, 1, 8).setValues([colHeaders])
    .setBackground(LBLUE).setFontColor(WHITE)
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  sheet.setRowHeight(row, 30);
  row++;

  // ── CHECKLIST ROWS ───────────────────────────────────────────
  const items    = chk.Items_JSON || [];
  let   paramNum = 1;
  let   startRow = row;

  items.forEach(cat => {
    const catStartRow = row;
    const catRows     = cat.rows || [];

    catRows.forEach((r, idx) => {
      const statusBg = r.statusBaik === 'Baik' ? GREEN : r.statusBaik === 'Tidak Baik' ? RED : AMBER;
      const rowData  = [
        paramNum,
        idx === 0 ? (cat.categoryName || '') : '',   // Kategori: hanya baris pertama
        r.parameter    || '',
        r.standarLimit || '',
        r.potensiResiko|| '',
        r.frekuensi    || 'MINGGUAN',
        r.statusBaik   || '—',
        r.keterangan   || ''
      ];
      sheet.getRange(row, 1, 1, 8).setValues([rowData]);
      sheet.getRange(row, 1).setHorizontalAlignment('center');
      sheet.getRange(row, 6).setHorizontalAlignment('center');
      sheet.getRange(row, 7).setHorizontalAlignment('center').setBackground(statusBg).setFontWeight('bold');
      sheet.getRange(row, 1, 1, 8).setVerticalAlignment('middle').setWrap(true).setBorder(true,true,true,true,true,true,'#E2E8F0',SpreadsheetApp.BorderStyle.SOLID);
      sheet.setRowHeight(row, 42);
      paramNum++;
      row++;
    });

    // Merge sel Kategori untuk baris yang sama
    if (catRows.length > 1) {
      sheet.getRange(catStartRow, 2, catRows.length, 1).merge()
        .setFontWeight('bold').setBackground('#EFF6FF')
        .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    } else {
      sheet.getRange(catStartRow, 2).setFontWeight('bold').setBackground('#EFF6FF').setVerticalAlignment('middle');
    }

    // Batas antar kategori
    sheet.getRange(catStartRow, 1, catRows.length, 8)
      .setBorder(true, true, true, true, null, null, '#2B6CB0', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  row++;

  // ── SUMMARY BARIS ────────────────────────────────────────────
  const totalParam  = _countRows_(items);
  const totalBaik   = _countByStatus_(items, 'Baik');
  const totalTidak  = _countByStatus_(items, 'Tidak Baik');
  const totalBelum  = totalParam - totalBaik - totalTidak;
  const healthRate  = totalParam > 0 ? Math.round((totalBaik / totalParam) * 100) : 0;

  sheet.getRange(row, 1, 1, 8).merge()
    .setValue('RINGKASAN PEMERIKSAAN')
    .setBackground(LBLUE).setFontColor(WHITE).setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center');
  row++;

  const summaryRows = [
    ['Total Parameter', totalParam,  '', 'Parameter Baik',     totalBaik,  '', 'Health Rate', healthRate + '%'],
    ['',                '',          '', 'Parameter Tidak Baik', totalTidak, '', 'Belum Diperiksa', totalBelum]
  ];
  summaryRows.forEach(sr => {
    sheet.getRange(row, 1, 1, 8).setValues([sr]);
    [1, 4, 7].forEach(c => sheet.getRange(row, c).setFontWeight('bold').setBackground(HEADER));
    row++;
  });

  // Color health rate
  const hrBg = healthRate >= 80 ? GREEN : healthRate >= 50 ? AMBER : RED;
  sheet.getRange(row - 2, 8).setBackground(hrBg).setFontWeight('bold');

  row += 2;

  // ── SIGNATURE SECTION ─────────────────────────────────────────
  // Judul signature
  sheet.getRange(row, 1, 1, 8).merge()
    .setValue('LEMBAR PERSETUJUAN & TANDA TANGAN')
    .setBackground(BLUE).setFontColor(WHITE).setFontWeight('bold')
    .setFontSize(11).setHorizontalAlignment('center');
  row++;

  // Sub-header dua kolom (4 kolom masing-masing)
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('IT Department')
    .setBackground(LBLUE).setFontColor(WHITE).setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center');
  sheet.getRange(row, 5, 1, 4).merge()
    .setValue(customTitle)
    .setBackground(LBLUE).setFontColor(WHITE).setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center');
  row++;

  // Keterangan nama / jabatan
  const signLabel = [
    ['Nama', itName, '', '', 'Nama', customName || '......................', '', ''],
    ['Jabatan', itTitle, '', '', 'Jabatan', customTitle, '', '']
  ];
  signLabel.forEach(sl => {
    sheet.getRange(row, 1, 1, 8).setValues([sl]);
    sheet.getRange(row, 1).setFontWeight('bold').setBackground(HEADER);
    sheet.getRange(row, 5).setFontWeight('bold').setBackground(HEADER);
    sheet.getRange(row, 2, 1, 3).merge();
    sheet.getRange(row, 6, 1, 3).merge();
    row++;
  });

  // Ruang tanda tangan (tinggi kosong)
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('Tanda Tangan:')
    .setFontSize(9).setFontColor('#94A3B8').setVerticalAlignment('top');
  sheet.getRange(row, 5, 1, 4).merge()
    .setValue('Tanda Tangan:')
    .setFontSize(9).setFontColor('#94A3B8').setVerticalAlignment('top');
  sheet.setRowHeight(row, 80);
  sheet.getRange(row, 1, 1, 4).setBorder(true,true,true,true,false,false,'#CBD5E1',SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row, 5, 1, 4).setBorder(true,true,true,true,false,false,'#CBD5E1',SpreadsheetApp.BorderStyle.SOLID);
  row++;

  // Baris tanggal ttd
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy');
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('Tanggal: ' + todayStr)
    .setFontSize(9).setHorizontalAlignment('center')
    .setBorder(true,true,true,true,false,false,'#CBD5E1',SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(row, 5, 1, 4).merge()
    .setValue('Tanggal: ................................')
    .setFontSize(9).setHorizontalAlignment('center')
    .setBorder(true,true,true,true,false,false,'#CBD5E1',SpreadsheetApp.BorderStyle.SOLID);
  row++;

  // ── Footer ───────────────────────────────────────────────────
  row++;
  sheet.getRange(row, 1, 1, 8).merge()
    .setValue('Dicetak dari IT Tools ERP · ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ' · Form: ' + (chk['Form Code'] || chk.ID))
    .setFontSize(8).setFontColor('#94A3B8').setFontStyle('italic').setHorizontalAlignment('center');

  SpreadsheetApp.flush();
  return sheet;
}

// ═══════════════════════════════════════════════════════════════
//  exportChecklistToSheets(checklistId, signConfig)
//  signConfig: { itName, itTitle, customName, customTitle }
// ═══════════════════════════════════════════════════════════════
function exportChecklistToSheets(checklistId, signConfig) {
  try {
    // Ambil data checklist
    const allData = getSheetData(SHEETS.CHECKLISTS);
    const chk     = allData.find(c => c.ID === checklistId || String(c.ID) === String(checklistId));
    if (!chk) return { success: false, message: 'Data checklist tidak ditemukan: ' + checklistId };

    const fileName = 'Checklist IT - ' + (chk['Form Code'] || chk.ID) + ' - ' + (chk.Site || '') + ' - ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');

    const ss = SpreadsheetApp.create(fileName);
    // Hapus sheet default
    const defSheet = ss.getSheets()[0];
    _buildChecklistSheet_(ss, chk, signConfig);
    ss.deleteSheet(defSheet);
    SpreadsheetApp.flush();

    return { success: true, url: ss.getUrl(), fileName, message: 'Export Excel checklist berhasil!' };
  } catch (e) {
    return { success: false, message: 'Export gagal: ' + e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
//  exportChecklistToPdf(checklistId, signConfig)
// ═══════════════════════════════════════════════════════════════
function exportChecklistToPdf(checklistId, signConfig) {
  try {
    const sheetResult = exportChecklistToSheets(checklistId, signConfig);
    if (!sheetResult.success) return sheetResult;

    const ssId  = sheetResult.url.match(/\/d\/([^\/]+)\//)[1];
    const ss    = SpreadsheetApp.openById(ssId);
    const shId  = ss.getSheets()[0].getSheetId();

    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${ssId}/export` +
      `?format=pdf&size=A4&portrait=true&fitw=true&gridlines=false` +
      `&printtitle=false&sheetnames=false&pagenum=UNDEFINED&gid=${shId}`;

    const token    = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      return { success: true, sheetsUrl: sheetResult.url, fileName: sheetResult.fileName, pdfFailed: true, message: 'PDF gagal, tersedia di Google Sheets.' };
    }

    const bytes  = response.getBlob().getBytes();
    const base64 = Utilities.base64Encode(bytes);
    const pdfFile= DriveApp.createFile(Utilities.newBlob(bytes, 'application/pdf', sheetResult.fileName + '.pdf'));

    return {
      success:   true,
      base64,
      mimeType:  'application/pdf',
      fileName:  sheetResult.fileName + '.pdf',
      sheetsUrl: sheetResult.url,
      driveUrl:  pdfFile.getUrl(),
      message:   'Export PDF checklist berhasil!'
    };
  } catch (e) {
    return { success: false, message: 'Export PDF gagal: ' + e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
//  exportChecklistBatchToPdf(checklistIds, signConfig)
//  Export beberapa form sekaligus dalam satu Spreadsheet (multi-sheet)
// ═══════════════════════════════════════════════════════════════
function exportChecklistBatchToPdf(checklistIds, signConfig) {
  try {
    if (!checklistIds || checklistIds.length === 0) return { success: false, message: 'Tidak ada checklist yang dipilih.' };

    const allData = getSheetData(SHEETS.CHECKLISTS);
    const fileName = 'Checklist IT Batch - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
    const ss       = SpreadsheetApp.create(fileName);
    const defSheet = ss.getSheets()[0];
    let   built    = 0;

    checklistIds.forEach(id => {
      const chk = allData.find(c => String(c.ID) === String(id));
      if (chk) { _buildChecklistSheet_(ss, chk, signConfig); built++; }
    });

    if (built === 0) { ss.deleteSheet(defSheet); return { success: false, message: 'Tidak ada data valid.' }; }
    ss.deleteSheet(defSheet);
    SpreadsheetApp.flush();

    return { success: true, url: ss.getUrl(), fileName, message: `${built} form berhasil diekspor.` };
  } catch (e) {
    return { success: false, message: 'Batch export gagal: ' + e.toString() };
  }
}