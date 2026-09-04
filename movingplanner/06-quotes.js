'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const QT = sheetMap['🚛 Moving Company Quotes'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  values.push({ range: "'🚛 Moving Company Quotes'!A1", values: [['🚛 All-in-One Moving Planner — Moving Company Quotes']] });

  // ── Recommendation summary (rows 2–4) ────────────────────────────────────
  values.push({ range: "'🚛 Moving Company Quotes'!A2", values: [['Best Price', 'Best Rated', 'Recommended']] });
  values.push({ range: "'🚛 Moving Company Quotes'!A3", values: [[
    '=IFERROR(INDEX(B6:B8,MATCH(MIN(E6:E8),E6:E8,0)),"—")',
    '=IFERROR(INDEX(B6:B8,MATCH(MAX(H6:H8),H6:H8,0)),"—")',
    '=IFERROR(INDEX(B6:B8,MATCH(MIN(E6:E8),E6:E8,0)),"—")',
  ]]});
  values.push({ range: "'🚛 Moving Company Quotes'!A4", values: [[
    '=IFERROR("$"&TEXT(MIN(E6:E8),"#,##0.00"),"—")',
    '=IFERROR(MAX(H6:H8)&" stars","—")',
    '=IFERROR("Lowest price: $"&TEXT(MIN(E6:E8),"#,##0.00"),"—")',
  ]]});

  // ── Column headers (row 5) ────────────────────────────────────────────────
  values.push({ range: "'🚛 Moving Company Quotes'!A5", values: [[
    '#', 'Company Name', 'Phone', 'Website', 'Quote ($)',
    'Includes Insurance?', 'Availability', 'Rating (1–5)',
    'Deposit Required', 'Notes',
  ]]});

  // ── 3 sample company quotes ───────────────────────────────────────────────
  const quotes = [
    [1, 'Premier Removals', '1300 555 001', 'HYPERLINK("https://example.com","premieremovals.com")', 1750, true,  '2025-08-30', 4.8, 200, 'Local and interstate. 2 men + truck.'],
    [2, 'EasyMove Co.',     '1800 555 002', 'HYPERLINK("https://example.com","easymoveco.com")',     1950, true,  '2025-08-30', 4.2, 300, 'Full packing service available.'],
    [3, 'Budget Movers',    '03 5555 0123', 'HYPERLINK("https://example.com","budgetmovers.com")',   1450, false, '2025-08-29', 3.9, 150, 'No insurance — check our own.'],
  ];

  const quoteRows = quotes.map(q => [
    q[0], q[1], q[2], `=${q[3]}`, q[4], q[5], q[6], q[7], q[8], q[9],
  ]);
  values.push({ range: "'🚛 Moving Company Quotes'!A6", values: quoteRows });

  await valuesBatchUpdate(id, values, 'quotes-values');

  // ── Merge title ───────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(QT, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(QT, 0, 1, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.darkBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Recommendation header row 2 ───────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(QT, 1, 2, 0, 3),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.antiqueBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  // Recommendation value rows
  reqs.push({
    repeatCell: {
      range: gridRange(QT, 2, 3, 0, 3),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(QT, 3, 4, 0, 3),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.warmCream),
          textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10, italic: true },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Column headers row 5 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(QT, 4, 5, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Data rows ─────────────────────────────────────────────────────────────
  for (let r = 0; r < 3; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(QT, 5 + r, 6 + r, 0, 10),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // Currency for Quote column (E)
  reqs.push({
    repeatCell: {
      range: gridRange(QT, 5, 8, 4, 5),
      cell: {
        userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });
  // Currency for Deposit (col I)
  reqs.push({
    repeatCell: {
      range: gridRange(QT, 5, 8, 8, 9),
      cell: {
        userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [35, 160, 110, 160, 90, 110, 100, 80, 100, 200];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: QT, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: QT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 40 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: QT, dimension: 'ROWS', startIndex: 1, endIndex: 5 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: QT, dimension: 'ROWS', startIndex: 5, endIndex: 8 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze rows 1–5 ───────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: QT, gridProperties: { frozenRowCount: 5 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'quotes-format');
  console.log('Moving Company Quotes complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
