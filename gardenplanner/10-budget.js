'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const BUD = sheetMap['Garden Budget'];

// 10 budget expense rows
const EXPENSE_DATA = [
  // Date, Category, Item, Vendor, Budgeted, Actual, Notes
  ['2025-02-15', 'Seeds & Starts',       'Tomato & pepper seedlings (12-pack)',  'Local Nursery',     24.00, 22.50, 'Bought 2 varieties'],
  ['2025-02-20', 'Seeds & Starts',       'Seed packet assortment (spring)',       'Johnny Seeds',      35.00, 31.80, '12 varieties incl. beans, lettuce, kale'],
  ['2025-03-01', 'Soil & Amendments',    'Premium potting mix (4 bags × 40qt)',   'Home Depot',        60.00, 54.96, '$13.74/bag on sale'],
  ['2025-03-10', 'Soil & Amendments',    'Compost (2 cu ft bags × 3)',            'Ace Hardware',      30.00, 27.00, 'Used for raised bed top-dress'],
  ['2025-03-20', 'Fertilizer',           'Tomato-tone organic fertilizer (4lb)',  'Garden Center',     18.00, 16.99, 'Also good for peppers'],
  ['2025-04-01', 'Irrigation',           'Drip irrigation kit + timer',           'Amazon',            80.00, 74.99, 'Auto-timer saves water; payback in 1 season'],
  ['2025-04-15', 'Structures & Supports','Tomato cages × 6 (54" heavy duty)',     'Tractor Supply',    54.00, 47.94, '$7.99 each; reusable'],
  ['2025-04-20', 'Tools & Equipment',    'Hori hori knife + kneeling pad',        'Lee Valley',        65.00, 58.00, 'Essential hand tool kit'],
  ['2025-05-05', 'Pest Control',         'Neem oil + insecticidal soap concentrate','Amazon',           28.00, 24.99, 'Organic pest control staples'],
  ['2025-05-20', 'Seeds & Starts',       'Fall seed packets (kale, spinach, beet)','Baker Creek',      20.00, 17.50, 'Planning ahead for fall succession'],
];

(async () => {
  const reqs = [];
  const values = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(BUD, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Garden Budget!A1', values: [['💰 GARDEN BUDGET — Season 2025']] });
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 0, 1, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.soil),
          textFormat: { foregroundColor: hex(C.cream), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BUD, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});

  // ── Budget Input block (rows 2-3) ─────────────────────────────────────────
  values.push({ range: 'Garden Budget!A2', values: [
    ['Total Season Budget ($):', 500, '', 'Total Spent:', `=SUM(F4:F${3+EXPENSE_DATA.length})`, '', 'Remaining:', `=B2-E2`, '', 'vs Budget:'],
  ]});
  values.push({ range: 'Garden Budget!J2', values: [
    [`=IFERROR(IF(H2>=0,"Under Budget 🎉","Over Budget ⚠️"),"")`],
  ]});

  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 1, 2, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.parchment),
          textFormat: { fontSize: 11 },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  // Bold labels
  [0, 3, 6, 9].forEach(col => {
    reqs.push({
      repeatCell: {
        range: gridRange(BUD, 1, 2, col, col + 1),
        cell: { userEnteredFormat: { textFormat: { bold: true } } },
        fields: 'userEnteredFormat.textFormat.bold',
      },
    });
  });
  // Currency for budget cells
  [1, 4, 7].forEach(col => {
    reqs.push({
      repeatCell: {
        range: gridRange(BUD, 1, 2, col, col + 1),
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }, textFormat: { bold: true, fontSize: 12 } } },
        fields: 'userEnteredFormat(numberFormat,textFormat)',
      },
    });
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BUD, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  }});

  // ── Header row ────────────────────────────────────────────────────────────
  const headers = ['Date', 'Category', 'Item Description', 'Vendor', 'Budgeted ($)', 'Actual ($)', 'Difference', 'Notes'];
  values.push({ range: 'Garden Budget!A3', values: [headers] });
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 2, 3, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.soilLight),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BUD, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // ── Expense data rows ─────────────────────────────────────────────────────
  const dataWithFormulas = EXPENSE_DATA.map((row, i) => {
    const [date, cat, item, vendor, budgeted, actual, notes] = row;
    const rowNum = i + 4;
    return [date, cat, item, vendor, budgeted, actual, `=E${rowNum}-F${rowNum}`, notes];
  });
  values.push({ range: 'Garden Budget!A4', values: dataWithFormulas });

  // Format data rows
  for (let i = 0; i < EXPENSE_DATA.length; i++) {
    const bg = i % 2 === 0 ? C.cream2 : C.parchment;
    reqs.push({
      repeatCell: {
        range: gridRange(BUD, 3 + i, 4 + i, 0, 8),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // Date format col A
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 3, 3 + EXPENSE_DATA.length, 0, 1),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Currency cols E, F, G
  [4, 5, 6].forEach(col => {
    reqs.push({
      repeatCell: {
        range: gridRange(BUD, 3, 3 + EXPENSE_DATA.length, col, col + 1),
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
      },
    });
  });

  // Center category and vendor
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 3, 3 + EXPENSE_DATA.length, 1, 4),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat.horizontalAlignment',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BUD, dimension: 'ROWS', startIndex: 3, endIndex: 3 + EXPENSE_DATA.length },
    properties: { pixelSize: 26 }, fields: 'pixelSize',
  }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [110, 140, 240, 140, 100, 100, 90, 200];
  colWidths.forEach((px, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: BUD, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // ── Borders ───────────────────────────────────────────────────────────────
  reqs.push({
    updateBorders: {
      range: gridRange(BUD, 2, 3 + EXPENSE_DATA.length, 0, 8),
      innerHorizontal: { style: 'SOLID', color: hex(C.lightGray) },
      innerVertical: { style: 'SOLID', color: hex(C.lightGray) },
      bottom: { style: 'SOLID_MEDIUM', color: hex(C.soil) },
      top: { style: 'SOLID_MEDIUM', color: hex(C.soil) },
      left: { style: 'SOLID_MEDIUM', color: hex(C.soil) },
      right: { style: 'SOLID_MEDIUM', color: hex(C.soil) },
    },
  });

  // ── Category Summary block ────────────────────────────────────────────────
  const catRow = 3 + EXPENSE_DATA.length + 2;
  reqs.push({ mergeCells: { range: gridRange(BUD, catRow, catRow + 1, 0, 5), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Garden Budget!A${catRow + 1}`, values: [['📊 SPENDING BY CATEGORY']] });
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, catRow, catRow + 1, 0, 5),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.olive),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  values.push({ range: `Garden Budget!A${catRow + 2}`, values: [['Category', 'Budgeted', 'Actual', 'Difference', '% of Total']] });
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, catRow + 1, catRow + 2, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.soilLight), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  const lastData = 3 + EXPENSE_DATA.length;
  const categories = ['Seeds & Starts', 'Soil & Amendments', 'Tools & Equipment', 'Irrigation', 'Pest Control', 'Structures & Supports', 'Fertilizer', 'Other'];
  const catRows = categories.map(cat => [
    cat,
    `=IFERROR(SUMIF(B4:B${lastData},"${cat}",E4:E${lastData}),0)`,
    `=IFERROR(SUMIF(B4:B${lastData},"${cat}",F4:F${lastData}),0)`,
    `=IFERROR(B${catRow+3}-C${catRow+3},0)`,
    `=IFERROR(C${catRow+3}/SUM(C${catRow+3}:C${catRow+3+categories.length-1}),0)`,
  ]);
  // Fix formula refs
  catRows.forEach((row, i) => {
    const r = catRow + 3 + i;
    row[3] = `=IFERROR(B${r}-C${r},0)`;
    row[4] = `=IFERROR(C${r}/SUM($C$${catRow+3}:$C$${catRow+2+categories.length}),0)`;
  });
  values.push({ range: `Garden Budget!A${catRow + 3}`, values: catRows });

  reqs.push({
    repeatCell: {
      range: gridRange(BUD, catRow + 2, catRow + 2 + categories.length, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cream2), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  [1, 2, 3].forEach(col => {
    reqs.push({
      repeatCell: {
        range: gridRange(BUD, catRow + 2, catRow + 2 + categories.length, col, col + 1),
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
      },
    });
  });
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, catRow + 2, catRow + 2 + categories.length, 4, 5),
      cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // ── Totals row ────────────────────────────────────────────────────────────
  const totRow = catRow + 2 + categories.length;
  values.push({ range: `Garden Budget!A${totRow + 1}`, values: [[
    'TOTAL',
    `=SUM(B${catRow+3}:B${catRow+2+categories.length})`,
    `=SUM(C${catRow+3}:C${catRow+2+categories.length})`,
    `=B${totRow+1}-C${totRow+1}`,
    '100%',
  ]]});
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, totRow, totRow + 1, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.soil), textFormat: { foregroundColor: hex(C.cream), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  [1, 2, 3].forEach(col => {
    reqs.push({
      repeatCell: {
        range: gridRange(BUD, totRow, totRow + 1, col, col + 1),
        cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: BUD, gridProperties: { frozenRowCount: 3 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await valuesBatchUpdate(id, values, 'budget-values');
  await batchUpdate(id, reqs, 'budget-format');
  console.log('Garden Budget complete — 10 expenses + category summary');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
