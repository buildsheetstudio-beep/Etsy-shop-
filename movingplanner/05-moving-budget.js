'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const BUD = sheetMap['💰 Moving Budget'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  values.push({ range: "'💰 Moving Budget'!A1", values: [['💰 All-in-One Moving Planner — Moving Budget']] });

  // ── Input block (row 2) ───────────────────────────────────────────────────
  values.push({ range: "'💰 Moving Budget'!A2", values: [
    ['Total Budget', 'Total Spent', 'Remaining', 'Over Budget?'],
  ]});
  values.push({ range: "'💰 Moving Budget'!A3", values: [[
    5000,
    '=SUM(E6:E35)',
    '=A3-B3',
    '=IF(C3<0,"YES ⚠️","No ✓")',
  ]]});

  // ── Column headers (row 5) ────────────────────────────────────────────────
  values.push({ range: "'💰 Moving Budget'!A5", values: [[
    '#', 'Budget Category', 'Description', 'Budgeted ($)', 'Actual ($)', 'Variance ($)', 'Notes',
  ]]});

  // ── 10 expense rows ───────────────────────────────────────────────────────
  const expenses = [
    [1,  'Moving Company',         'Professional removalists — full truck',  1800, 0, '', 'See Quotes tab'],
    [2,  'Packing Supplies',       'Boxes, tape, bubble wrap, paper',         150, 0, '', ''],
    [3,  'Storage Unit',           '1 month storage — 3×3m unit',             200, 0, '', 'If needed'],
    [4,  'Cleaning Services',      'Bond clean — old home',                   400, 0, '', ''],
    [5,  'Repairs & Maintenance',  'Patch walls, touch-up paint — old home',  120, 0, '', ''],
    [6,  'New Furniture',          'Bookshelf and side tables for new home',  350, 0, '', ''],
    [7,  'Utility Connections',    'Electricity/gas/internet connection fees', 150, 0, '', ''],
    [8,  'Travel & Fuel',          'Fuel for car trip — Springfield to Riverside', 80, 0, '', ''],
    [9,  'Temporary Accommodation','1 night motel if needed',                 120, 0, '', ''],
    [10, 'Miscellaneous',          'Buffer for unexpected costs',             200, 0, '', ''],
  ];

  const expenseRows = expenses.map((e, i) => [
    e[0], e[1], e[2], e[3], e[4],
    `=D${6 + i}-E${6 + i}`,
    e[6],
  ]);
  values.push({ range: "'💰 Moving Budget'!A6", values: expenseRows });

  // ── Totals row ────────────────────────────────────────────────────────────
  values.push({ range: "'💰 Moving Budget'!A16", values: [[
    '', 'TOTALS', '', '=SUM(D6:D15)', '=SUM(E6:E15)', '=SUM(F6:F15)', '',
  ]]});

  // ── Category Summary header (col I–M, rows 2–3) ───────────────────────────
  values.push({ range: "'💰 Moving Budget'!I2", values: [[
    'Category', 'Budgeted', 'Actual', 'Variance', 'Spend Bar',
  ]]});

  const summaryCategories = [
    'Moving Company', 'Packing Supplies', 'Storage Unit', 'Cleaning Services',
    'Repairs & Maintenance', 'New Furniture', 'Utility Connections',
    'Travel & Fuel', 'Temporary Accommodation', 'Miscellaneous',
  ];

  const summaryRows = summaryCategories.map(cat => [
    cat,
    `=IFERROR(SUMIF($B$6:$B$15,"${cat}",$D$6:$D$15),0)`,
    `=IFERROR(SUMIF($B$6:$B$15,"${cat}",$E$6:$E$15),0)`,
    `=J3-K3`,
    `=IFERROR(SPARKLINE({K3,MAX(J3-K3,0)},{"charttype","bar";"color1","#2C4A35";"color2","#FBF5E6"}),"-")`,
  ]);
  // Fix formula references — they're relative, need to be per-row
  const summaryRowsFixed = summaryCategories.map((cat, i) => [
    cat,
    `=IFERROR(SUMIF($B$6:$B$15,"${cat}",$D$6:$D$15),0)`,
    `=IFERROR(SUMIF($B$6:$B$15,"${cat}",$E$6:$E$15),0)`,
    `=J${3 + i}-K${3 + i}`,
    `=IFERROR(SPARKLINE({K${3 + i},MAX(J${3 + i}-K${3 + i},0)},{"charttype","bar";"color1","#2C4A35";"color2","#FBF5E6"}),"-")`,
  ]);
  values.push({ range: "'💰 Moving Budget'!I3", values: summaryRowsFixed });

  // ── Summary totals ────────────────────────────────────────────────────────
  values.push({ range: "'💰 Moving Budget'!I13", values: [[
    'TOTAL', '=SUM(J3:J12)', '=SUM(K3:K12)', '=SUM(L3:L12)', '',
  ]]});

  await valuesBatchUpdate(id, values, 'budget-values');

  // ── Merge title ───────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(BUD, 0, 1, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 0, 1, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.forestGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Summary input block headers row 2 ────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 1, 2, 0, 4),
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

  // ── Summary values row 3 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 2, 3, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 2, 3, 1, 3),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.readOnlyBg),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 2, 3, 3, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.readOnlyBg),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Column headers row 5 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 4, 5, 0, 7),
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

  // ── Zebra stripe data rows ────────────────────────────────────────────────
  for (let r = 0; r < 10; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(BUD, 5 + r, 6 + r, 0, 7),
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

  // Currency format for budget/actual/variance columns
  for (let r = 0; r < 10; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(BUD, 5 + r, 6 + r, 3, 6),
        cell: {
          userEnteredFormat: {
            numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
          },
        },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  }

  // ── Totals row format ─────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 15, 16, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
    },
  });

  // ── Category Summary headers (col I–M) ───────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 1, 2, 8, 13),
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

  // ── Category summary data rows ────────────────────────────────────────────
  for (let r = 0; r < 10; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(BUD, 2 + r, 3 + r, 8, 12),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            horizontalAlignment: 'CENTER',
            numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
      },
    });
  }

  // ── Summary totals row ────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(BUD, 12, 13, 8, 13),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
    },
  });

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [35, 140, 240, 90, 90, 90, 160, 20, 140, 90, 90, 90, 150];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: BUD, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: BUD, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 40 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: BUD, dimension: 'ROWS', startIndex: 1, endIndex: 5 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: BUD, dimension: 'ROWS', startIndex: 5, endIndex: 16 },
      properties: { pixelSize: 24 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze ────────────────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: BUD, gridProperties: { frozenRowCount: 5 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'budget-format');
  console.log('Moving Budget complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
