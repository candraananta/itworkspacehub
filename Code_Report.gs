/**
 * ============================================================
 *  IT Tools ERP — Code_Report.gs
 *  Incident Report · Aggregation · Export Excel · Export PDF
 * ============================================================
 *
 *  Fungsi yang tersedia (dipanggil dari frontend):
 *    getIncidentReport(filters)  → data agregasi untuk chart & tabel
 *    exportReportToSheets(filters) → buat Google Spreadsheet baru, return URL
 *    exportReportToPdf(filters)    → buat Spreadsheet → export PDF, return base64
 */

// ── Helper: parse tanggal dari berbagai format GAS ───────────
function _parseDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const s = val.toString().trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ── Helper: format tanggal → "YYYY-MM" ───────────────────────
function _toYearMonth_(val) {
  const d = _parseDate_(val);
  if (!d) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

// ── Helper: format display "Januari 2025" ────────────────────
function _formatMonthLabel_(ym) {
  if (!ym) return '-';
  const months = ['Januari','Februari','Maret','April','Mei','Juni',
                  'Juli','Agustus','September','Oktober','November','Desember'];
  const [y, m] = ym.split('-');
  return `${months[parseInt(m, 10) - 1]} ${y}`;
}

// ═══════════════════════════════════════════════════════════════
//  getIncidentReport(filters)
//  filters: { month, year, location, category }
//  Return: summary stats + chart data + table rows
// ═══════════════════════════════════════════════════════════════
function getIncidentReport(filters) {
  try {
    filters = filters || {};
    const allRows = getSheetData(SHEETS.INCIDENTS);

    // ── Apply filters ────────────────────────────────────────
    let rows = allRows.filter(r => {
      const ym = _toYearMonth_(r.Tanggal);
      if (filters.year && filters.month) {
        const target = `${filters.year}-${String(filters.month).padStart(2, '0')}`;
        if (ym !== target) return false;
      } else if (filters.year) {
        if (!ym || !ym.startsWith(filters.year)) return false;
      }
      if (filters.location && r.Location !== filters.location) return false;
      if (filters.category && r.Category !== filters.category) return false;
      return true;
    });

    // ── Sort by date asc ─────────────────────────────────────
    rows.sort((a, b) => {
      const da = _parseDate_(a.Tanggal), db = _parseDate_(b.Tanggal);
      return (da || 0) - (db || 0);
    });

    const total   = rows.length;
    const waiting = rows.filter(r => r.Status === 'Waiting').length;
    const closed  = rows.filter(r => r.Status === 'Closed').length;
    const resolveRate = total > 0 ? Math.round((closed / total) * 100) : 0;

    // ── Chart 1: Tren per bulan ──────────────────────────────
    const monthMap = {};
    rows.forEach(r => {
      const ym = _toYearMonth_(r.Tanggal) || 'Unknown';
      if (!monthMap[ym]) monthMap[ym] = { total: 0, waiting: 0, closed: 0 };
      monthMap[ym].total++;
      if (r.Status === 'Waiting') monthMap[ym].waiting++;
      if (r.Status === 'Closed')  monthMap[ym].closed++;
    });
    const trendData = Object.keys(monthMap).sort().map(ym => ({
      label:   _formatMonthLabel_(ym),
      ym:      ym,
      total:   monthMap[ym].total,
      waiting: monthMap[ym].waiting,
      closed:  monthMap[ym].closed
    }));

    // ── Chart 2: By Location ─────────────────────────────────
    const locMap = {};
    rows.forEach(r => {
      const loc = r.Location || 'Unknown';
      if (!locMap[loc]) locMap[loc] = { total: 0, waiting: 0, closed: 0 };
      locMap[loc].total++;
      if (r.Status === 'Waiting') locMap[loc].waiting++;
      if (r.Status === 'Closed')  locMap[loc].closed++;
    });
    const byLocation = Object.entries(locMap)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([loc, v]) => ({ label: loc, ...v }));

    // ── Chart 3: By Category ─────────────────────────────────
    const catMap = {};
    rows.forEach(r => {
      const cat = r.Category || 'Unknown';
      if (!catMap[cat]) catMap[cat] = { total: 0, waiting: 0, closed: 0 };
      catMap[cat].total++;
      if (r.Status === 'Waiting') catMap[cat].waiting++;
      if (r.Status === 'Closed')  catMap[cat].closed++;
    });
    const byCategory = Object.entries(catMap)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([cat, v]) => ({ label: cat, ...v }));

    // ── Chart 4: Waiting vs Closed (donut) ───────────────────
    const statusDist = [
      { label: 'Closed',  value: closed,  color: '#059669' },
      { label: 'Waiting', value: waiting, color: '#D97706' }
    ];

    // ── Table: detail rows untuk export ──────────────────────
    const tableRows = rows.map(r => ({
      id:          r.ID,
      tanggal:     (r.Tanggal || '').toString().substring(0, 10),
      location:    r.Location  || '-',
      category:    r.Category  || '-',
      description: r.Description || '-',
      status:      r.Status    || '-'
    }));

    // ── Available years untuk filter dropdown ─────────────────
    const years = [...new Set(
      allRows
        .map(r => _toYearMonth_(r.Tanggal))
        .filter(Boolean)
        .map(ym => ym.split('-')[0])
    )].sort().reverse();

    return {
      success:     true,
      summary:     { total, waiting, closed, resolveRate },
      trendData,
      byLocation,
      byCategory,
      statusDist,
      tableRows,
      years,
      appliedFilters: filters,
      generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
//  exportReportToSheets(filters)
//  Buat Google Spreadsheet baru → isi dengan data report
//  Return: { success, url, fileName }
// ═══════════════════════════════════════════════════════════════
function exportReportToSheets(filters) {
  try {
    const report = getIncidentReport(filters);
    if (!report.success) return report;

    const f = filters || {};
    const periodLabel = f.year
      ? (f.month ? _formatMonthLabel_(`${f.year}-${String(f.month).padStart(2,'0')}`) : `Tahun ${f.year}`)
      : 'Semua Periode';
    const locLabel  = f.location || 'Semua Lokasi';
    const catLabel  = f.category || 'Semua Kategori';
    const fileName  = `Laporan Insiden IT - ${periodLabel} - ${report.generatedAt.replace(/\//g,'-').replace(' ','_').replace(':','')}`;

    const ss    = SpreadsheetApp.create(fileName);
    const sheet = ss.getActiveSheet();
    sheet.setName('Laporan Insiden');

    // ── HEADER SECTION ───────────────────────────────────────
    const BLUE     = '#1A3A66';
    const LBLUE    = '#2B6CB0';
    const HDRFG    = '#FFFFFF';
    const ACCBG    = '#EBF4FF';
    const GREENBG  = '#D1FAE5';
    const AMBERBG  = '#FEF3C7';

    // Title
    sheet.getRange('A1:F1').merge()
      .setValue('LAPORAN INSIDEN IT — ' + periodLabel.toUpperCase())
      .setBackground(BLUE).setFontColor(HDRFG)
      .setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center');

    // Filter info
    sheet.getRange('A2').setValue('Lokasi').setFontWeight('bold');
    sheet.getRange('B2').setValue(locLabel);
    sheet.getRange('C2').setValue('Kategori').setFontWeight('bold');
    sheet.getRange('D2').setValue(catLabel);
    sheet.getRange('E2').setValue('Dibuat').setFontWeight('bold');
    sheet.getRange('F2').setValue(report.generatedAt);

    // ── SUMMARY CARDS ────────────────────────────────────────
    sheet.getRange('A4:F4').merge()
      .setValue('RINGKASAN EKSEKUTIF')
      .setBackground(LBLUE).setFontColor(HDRFG)
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center');

    const summaryHeaders = [['Total Insiden', 'Menunggu (Waiting)', 'Selesai (Closed)', 'Resolve Rate', 'Periode', 'Sumber Data']];
    sheet.getRange('A5:F5').setValues(summaryHeaders)
      .setBackground(ACCBG).setFontWeight('bold').setHorizontalAlignment('center');

    sheet.getRange('A6:F6').setValues([[
      report.summary.total,
      report.summary.waiting,
      report.summary.closed,
      report.summary.resolveRate + '%',
      periodLabel,
      'Google Sheets (IT Tools ERP)'
    ]]).setHorizontalAlignment('center').setFontSize(12).setFontWeight('bold');

    // Color code summary
    sheet.getRange('B6').setBackground(report.summary.waiting > 0 ? AMBERBG : GREENBG);
    sheet.getRange('C6').setBackground(GREENBG);
    sheet.getRange('D6').setBackground(report.summary.resolveRate >= 80 ? GREENBG : AMBERBG);

    // ── TREND PER BULAN ──────────────────────────────────────
    let row = 8;
    sheet.getRange(`A${row}:F${row}`).merge()
      .setValue('TREN INSIDEN PER BULAN')
      .setBackground(LBLUE).setFontColor(HDRFG)
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center');
    row++;

    sheet.getRange(`A${row}:D${row}`)
      .setValues([['Bulan', 'Total', 'Waiting', 'Closed']])
      .setBackground(ACCBG).setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    if (report.trendData.length > 0) {
      const trendRows = report.trendData.map(t => [t.label, t.total, t.waiting, t.closed]);
      sheet.getRange(row, 1, trendRows.length, 4).setValues(trendRows).setHorizontalAlignment('center');
      row += trendRows.length;
    } else {
      sheet.getRange(`A${row}`).setValue('Tidak ada data untuk periode ini.').setFontColor('#94A3B8').setFontStyle('italic');
      row++;
    }

    // ── BY LOCATION ──────────────────────────────────────────
    row += 2;
    sheet.getRange(`A${row}:D${row}`).merge()
      .setValue('INSIDEN PER LOKASI')
      .setBackground(LBLUE).setFontColor(HDRFG)
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center');
    row++;

    sheet.getRange(`A${row}:D${row}`)
      .setValues([['Lokasi', 'Total', 'Waiting', 'Closed']])
      .setBackground(ACCBG).setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    if (report.byLocation.length > 0) {
      const locRows = report.byLocation.map(l => [l.label, l.total, l.waiting, l.closed]);
      sheet.getRange(row, 1, locRows.length, 4).setValues(locRows).setHorizontalAlignment('center');
      row += locRows.length;
    }

    // ── BY CATEGORY ──────────────────────────────────────────
    row += 2;
    sheet.getRange(`A${row}:D${row}`).merge()
      .setValue('INSIDEN PER KATEGORI')
      .setBackground(LBLUE).setFontColor(HDRFG)
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center');
    row++;

    sheet.getRange(`A${row}:D${row}`)
      .setValues([['Kategori', 'Total', 'Waiting', 'Closed']])
      .setBackground(ACCBG).setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    if (report.byCategory.length > 0) {
      const catRows = report.byCategory.map(c => [c.label, c.total, c.waiting, c.closed]);
      sheet.getRange(row, 1, catRows.length, 4).setValues(catRows).setHorizontalAlignment('center');
      row += catRows.length;
    }

    // ── DETAIL TABLE ─────────────────────────────────────────
    row += 2;
    sheet.getRange(`A${row}:F${row}`).merge()
      .setValue('DATA DETAIL INSIDEN')
      .setBackground(BLUE).setFontColor(HDRFG)
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center');
    row++;

    sheet.getRange(`A${row}:F${row}`)
      .setValues([['ID Log', 'Tanggal', 'Lokasi', 'Kategori', 'Deskripsi', 'Status']])
      .setBackground(LBLUE).setFontColor(HDRFG)
      .setFontWeight('bold').setHorizontalAlignment('center');
    row++;

    if (report.tableRows.length > 0) {
      const detailRows = report.tableRows.map(r => [
        r.id, r.tanggal, r.location, r.category, r.description, r.status
      ]);
      const detailRange = sheet.getRange(row, 1, detailRows.length, 6);
      detailRange.setValues(detailRows);

      // Warna baris berdasarkan status
      detailRows.forEach((dr, i) => {
        const bg = dr[5] === 'Closed' ? '#F0FDF4' : '#FFFBEB';
        sheet.getRange(row + i, 1, 1, 6).setBackground(bg);
      });
      row += detailRows.length;
    }

    // ── Auto-resize & formatting ──────────────────────────────
    sheet.setColumnWidths(1, 6, 150);
    sheet.setColumnWidth(5, 300); // Deskripsi lebih lebar
    sheet.setRowHeights(1, row, 22);

    // ── Footer ───────────────────────────────────────────────
    row += 2;
    sheet.getRange(`A${row}:F${row}`).merge()
      .setValue(`Dicetak dari IT Tools ERP • ${report.generatedAt} • Data: ${report.summary.total} insiden`)
      .setFontColor('#94A3B8').setFontSize(9).setFontStyle('italic')
      .setHorizontalAlignment('center');

    SpreadsheetApp.flush();

    return {
      success:  true,
      url:      ss.getUrl(),
      fileName: fileName,
      message:  `Export Excel berhasil! ${report.summary.total} insiden diekspor.`
    };

  } catch (e) {
    return { success: false, message: 'Export Excel gagal: ' + e.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
//  exportReportToPdf(filters)
//  Buat Spreadsheet sementara → export PDF blob → return base64
//  Return: { success, base64, fileName, mimeType }
// ═══════════════════════════════════════════════════════════════
function exportReportToPdf(filters) {
  try {
    // Buat spreadsheet dulu (reuse exportReportToSheets)
    const sheetResult = exportReportToSheets(filters);
    if (!sheetResult.success) return sheetResult;

    // Ambil spreadsheet yang baru dibuat
    const ssId  = sheetResult.url.match(/\/d\/([^\/]+)\//)[1];
    const ss    = SpreadsheetApp.openById(ssId);
    const shId  = ss.getActiveSheet().getSheetId();

    // Export ke PDF via Drive API URL
    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${ssId}/export` +
      `?format=pdf&size=A4&portrait=true&fitw=true&gridlines=false` +
      `&printtitle=false&sheetnames=false&pagenum=UNDEFINED` +
      `&gid=${shId}`;

    const token    = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      // Kalau PDF gagal, tetap kembalikan URL Sheets
      return {
        success:    true,
        sheetsUrl:  sheetResult.url,
        fileName:   sheetResult.fileName,
        pdfFailed:  true,
        message:    'PDF export gagal (quota/permission), file tersedia di Google Sheets.'
      };
    }

    const blob     = response.getBlob();
    const bytes    = blob.getBytes();
    const base64   = Utilities.base64Encode(bytes);

    // Simpan PDF ke Drive (folder root)
    const pdfFile  = DriveApp.createFile(
      Utilities.newBlob(bytes, 'application/pdf', sheetResult.fileName + '.pdf')
    );

    return {
      success:   true,
      base64:    base64,
      mimeType:  'application/pdf',
      fileName:  sheetResult.fileName + '.pdf',
      sheetsUrl: sheetResult.url,
      driveUrl:  pdfFile.getUrl(),
      message:   'Export PDF berhasil!'
    };

  } catch (e) {
    return { success: false, message: 'Export PDF gagal: ' + e.toString() };
  }
}