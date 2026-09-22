'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const GBD = sheetMap['💰 Grocery Budget'];

// A=WeekOf B=Budget C=ActualSpend D=CostPerMeal E=Variance F=Notes
// Row 8 (index 5) is the current week (2026-07-13); rows 3-7 are historical
const rows = [
  ['2026-06-08',180,162.40,null,null,'Good week, mostly pantry'],
  ['2026-06-15',180,195.30,null,null,'Stocked up on protein'],
  ['2026-06-22',180,171.20,null,null,'Used meal prep well'],
  ['2026-06-29',200,188.70,null,null,'Summer produce splurge'],
  ['2026-07-06',200,143.50,null,null,'Minimal shopping — good prep'],
  ['2026-07-13',200,null,null,null,'Current week'],
];

const headers = ['Week Of','Budget','Actual Spend','Cost Per Meal','Variance','Notes'];

(async () => {
  const dataRows = rows.map((r, i) => {
    const row = i + 3;
    const isCurrent = i === 5; // last row = current week
    const actualSpend = isCurrent
      ? `=IFERROR(SUMIFS('🛒 Grocery List'!$E:$E,'🛒 Grocery List'!$G:$G,TRUE),0)`
      : r[2];
    return [
      r[0], r[1], actualSpend,
      `=IFERROR(C${row}/21,0)`,
      `=IFERROR(B${row}-C${row},0)`,
      r[5],
    ];
  });

  const data = [
    { range: "'💰 Grocery Budget'!A1", values: [['💰 Grocery Budget']] },
    { range: "'💰 Grocery Budget'!A2", values: [headers] },
    { range: "'💰 Grocery Budget'!A3", values: dataRows },
  ];
  await valuesBatchUpdate(id, data, 'gbd-values');

  const reqs = [];

  // Merge A1:F1
  reqs.push({ mergeCells: { range: gridRange(GBD, 0, 1, 0, 6), mergeType: 'MERGE_ALL' } });

  // Title
  reqs.push({ repeatCell: {
    range: gridRange(GBD, 0, 1, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmOrange), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Header
  reqs.push({ repeatCell: {
    range: gridRange(GBD, 1, 2, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Alt rows
  for (let r = 0; r < 6; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(GBD, 2 + r, 3 + r, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  // Formula cols C, D, E: formulaBg
  reqs.push({ repeatCell: {
    range: gridRange(GBD, 2, 8, 2, 5),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg) } },
    fields: 'userEnteredFormat.backgroundColor',
  }});
  // Currency format B, C, D, E
  const currFmt = { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' };
  [1, 2, 3, 4].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(GBD, 2, 8, c, c + 1),
    cell: { userEnteredFormat: currFmt },
    fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
  }}));
  // Date format col A
  reqs.push({ repeatCell: {
    range: gridRange(GBD, 2, 8, 0, 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
  }});

  // Highlight current week row (row index 7, 0-based) with lightAmber
  reqs.push({ repeatCell: {
    range: gridRange(GBD, 7, 8, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.lightAmber), textFormat: { bold: true } } },
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});

  const widths = [110, 90, 110, 110, 100, 200];
  widths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: GBD, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GBD, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 34 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'gbd-format');
  console.log('Grocery Budget complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
