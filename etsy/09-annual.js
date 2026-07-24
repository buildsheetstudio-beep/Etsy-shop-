'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS, EXPENSE_CATS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SH = sheetMap['Annual Summary'];
const S  = "'Annual Summary'";

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 56 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'ROWS', startIndex: 3, endIndex: 200 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths — A=label, B-M=months, N=total
  [[0,200],[13,110]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  for (let c = 1; c <= 12; c++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SH, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } });
  }

  // Title A1:N1
  reqs.push({ mergeCells: { range: gridRange(SH, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 0, 1, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Subheader A2:N2
  reqs.push({ mergeCells: { range: gridRange(SH, 1, 2, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(SH, 1, 2, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Column header row 3
  reqs.push({ repeatCell: { range: gridRange(SH, 2, 3, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  function sectionHdr(r, color) {
    reqs.push({ mergeCells: { range: gridRange(SH, r, r+1, 0, 14), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 14), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }
  function colHdr(r) {
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 14), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }
  function dataBlock(r1, r2) {
    for (let r = r1; r < r2; r++) {
      const bg = r % 2 === 0 ? C.panel : C.altRow;
      reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 14), cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.mainText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
    }
  }
  function totalRow(r, color) {
    reqs.push({ repeatCell: { range: gridRange(SH, r, r+1, 0, 14), cell: { userEnteredFormat: {
      backgroundColor: hex(color),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  }

  // Income section: row 4 hdr, 5 col hdr, 6-13 data, 14 total
  sectionHdr(3, C.success);
  colHdr(4);
  dataBlock(5, 15);
  totalRow(15, C.success);

  // Expense section: row 17 hdr, 18 col hdr, 19-48 data, 49 total
  sectionHdr(16, C.attention);
  colHdr(17);
  dataBlock(18, 48);
  totalRow(48, C.attention);

  // Summary section: row 50 hdr, 51-55 data
  sectionHdr(49, C.primary);
  dataBlock(50, 57);

  // Currency format
  for (let c = 1; c <= 13; c++) {
    reqs.push({ repeatCell: { range: gridRange(SH, 4, 57, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' }});
  }

  // Percent format in summary rows 53-54 (margin)
  for (let c = 1; c <= 13; c++) {
    reqs.push({ repeatCell: { range: gridRange(SH, 52, 54, c, c+1), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } }, fields: 'userEnteredFormat.numberFormat' }});
  }

  // Borders
  reqs.push({ updateBorders: { range: gridRange(SH, 2, 57, 0, 14),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Freeze rows 3 only (merged header row spans all cols — can't also freeze cols)
  reqs.push({ updateSheetProperties: { properties: { sheetId: SH, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'annual-format');

  // Values
  const data = [];
  const yr = 2026;
  data.push({ range: `${S}!A1`, values: [['📆 Annual Summary — ${yr} Full Year by Month']] });
  data.push({ range: `${S}!A2`, values: [['All 12 months · January–December ${yr} · Willow Paper Studio']] });

  // Column headers
  const headerRow = ['Metric', ...MONTHS.map(m => m.slice(0,3)), 'Full Year'];
  data.push({ range: `${S}!A3:N3`, values: [headerRow] });

  // Helper: SUMPRODUCT for month filter
  function monthSum(sheet, dateCol, valCol, mo) {
    return `=SUMPRODUCT((MONTH(${sheet}!${dateCol}6:${dateCol}500)=${mo})*(YEAR(${sheet}!${dateCol}6:${dateCol}500)=${yr})*${sheet}!${valCol}6:${valCol}500)`;
  }

  // Income section
  data.push({ range: `${S}!A4`, values: [['  💰 Revenue by Source']] });
  data.push({ range: `${S}!A5:N5`, values: [['#','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Full Year']] });

  const SOURCES = ['Etsy Order','Etsy Shipping Collected','Off-Etsy Sale','Custom Order','Other Business Income'];
  SOURCES.forEach((src, i) => {
    const r = 6 + i;
    const monthCols = MONTHS.map((_, mi) =>
      `=SUMPRODUCT((MONTH('Income & Orders'!$B$6:$B$500)=${mi+1})*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*('Income & Orders'!$F$6:$F$500="${src}")*'Income & Orders'!$H$6:$H$500)`
    );
    data.push({ range: `${S}!A${r}:N${r}`, values: [[src, ...monthCols, `=SUM(B${r}:M${r})`]] });
  });
  // Income total row 16
  const incomeTotals = MONTHS.map((_, mi) =>
    `=SUMPRODUCT((MONTH('Income & Orders'!$B$6:$B$500)=${mi+1})*(YEAR('Income & Orders'!$B$6:$B$500)=${yr})*'Income & Orders'!$H$6:$H$500)`
  );
  data.push({ range: `${S}!A16:N16`, values: [['TOTAL', ...incomeTotals, '=SUM(B16:M16)']] });

  // Expense section
  data.push({ range: `${S}!A17`, values: [['  💸 Expenses by Category']] });
  data.push({ range: `${S}!A18:N18`, values: [['Category','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Full Year']] });

  const KEY_CATS = ['Etsy Transaction Fees','Etsy Listing Fees','Etsy Payment Processing Fees','Etsy Ads','Offsite Ads','Software & Subscriptions','Design Assets','Print-on-Demand Costs','Contractors','Photography','Education & Training','Banking Fees','Other Business Expense'];
  KEY_CATS.forEach((cat, i) => {
    const r = 19 + i;
    const monthCols = MONTHS.map((_, mi) =>
      `=SUMPRODUCT((MONTH('Expense Log'!$A$6:$A$500)=${mi+1})*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*('Expense Log'!$E$6:$E$500="${cat}")*'Expense Log'!$H$6:$H$500)`
    );
    data.push({ range: `${S}!A${r}:N${r}`, values: [[cat, ...monthCols, `=SUM(B${r}:M${r})`]] });
  });
  // Expense total row 49
  const expTotals = MONTHS.map((_, mi) =>
    `=SUMPRODUCT((MONTH('Expense Log'!$A$6:$A$500)=${mi+1})*(YEAR('Expense Log'!$A$6:$A$500)=${yr})*'Expense Log'!$H$6:$H$500)`
  );
  data.push({ range: `${S}!A49:N49`, values: [['TOTAL', ...expTotals, '=SUM(B49:M49)']] });

  // Summary section
  data.push({ range: `${S}!A50`, values: [['  📊 Monthly P&L Summary']] });
  data.push({ range: `${S}!A51:N51`, values: [['Metric','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Full Year']] });

  data.push({ range: `${S}!A52:N52`, values: [['Gross Revenue', ...MONTHS.map((_,mi)=>`=B${16}:M${16}`.includes('M') ? `=B16` : `=${String.fromCharCode(66+mi)}16`), '=SUM(B52:M52)']] });
  // Better approach: just reference row 16 directly by column
  const monthLetters = MONTHS.map((_, i) => String.fromCharCode(66+i)); // B through M
  data.push({ range: `${S}!A52:N52`, values: [['Gross Revenue', ...monthLetters.map(l => `=${l}16`), '=N16']] });
  data.push({ range: `${S}!A53:N53`, values: [['Total Expenses', ...monthLetters.map(l => `=${l}49`), '=N49']] });
  data.push({ range: `${S}!A54:N54`, values: [['Net Profit', ...monthLetters.map(l => `=${l}52-${l}53`), '=N52-N53']] });
  data.push({ range: `${S}!A55:N55`, values: [['Profit Margin %', ...monthLetters.map(l => `=IFERROR(${l}54/${l}52,0)`), '=IFERROR(N54/N52,0)']] });
  data.push({ range: `${S}!A56:N56`, values: [['Revenue Goal', ...MONTHS.map(_ => 7083.33), 85000]] });
  data.push({ range: `${S}!A57:N57`, values: [['vs Goal', ...monthLetters.map(l => `=${l}52-${l}56`), '=N52-N56']] });

  await valuesBatchUpdate(id, data, 'annual-values');
  console.log('Annual Summary complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
