'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BU = sheetMap['💰 Budget & Expenses'];

// Layout:
// Row 1: Title (A1:I1 merged)
// Row 2: Summary/KPI header
// Rows 3-14: Budget allocation table (cols B-F) + KPI block (cols H-I)
// Row 15: spacer
// Row 16: Expense log section header
// Row 17: Expense log col headers
// Rows 18-32: 15 expense entries
// Col A: row numbers (auto formula in allocation)
// Budget allocation cols: B=Category, C=Budgeted, D=Actual(SUMIF), E=Remaining, F=Sparkline
// KPI block cols: H=label, I=value (rows 3-8)
// Expense log cols: A=Date, B=Category, C=Item/Vendor, D=Budgeted, E=Actual, F=Variance, G=Paid?, H=Receipt Ref, I=Notes

const NCOLS = 9; // A-I

const CATEGORIES = [
  ['Venue',                    0],
  ['Food & Catering',        350],
  ['Cake',                   280],
  ['Drinks',                 230],
  ['Decorations & Florals',  600],
  ['Invitations & Stationery', 80],
  ['Party Favours',           90],
  ['Games & Prizes',         120],
  ['Photography',            200],
  ['Gift for Mum-to-Be',     150],
  ['Miscellaneous',          100],
  ['Buffer',                 100],
];

const EXPENSES = [
  // [date, category, item, budgeted, actual, receiptRef, notes]
  ['01 Jun 2025','Venue','Home venue — no cost',0,0,'',''],
  ['05 Jun 2025','Invitations & Stationery','Digital invites via Canva',80,45,'INV-001','Printed 5 copies'],
  ['10 Jun 2025','Food & Catering','Afternoon tea platter — The Pantry',350,285,'REC-001','Confirmed order'],
  ['10 Jun 2025','Cake','Custom nappy cake — Sweet Creations',280,280,'REC-002','Lilac & blush theme'],
  ['12 Jun 2025','Decorations & Florals','Balloon arch & florals — Bloom & Co',600,540,'REC-003','Paid deposit'],
  ['15 Jun 2025','Drinks','Sparkling, juices & mocktail kit',230,195,'REC-004',''],
  ['18 Jun 2025','Photography','1/2 day photographer — J. Mills Photo',200,200,'REC-005','Booked'],
  ['20 Jun 2025','Party Favours','Personalised seed packet favours x20',90,78,'REC-006',''],
  ['22 Jun 2025','Games & Prizes','Game prizes & ribbons',120,65,'REC-007','5 games budgeted'],
  ['25 Jun 2025','Gift for Mum-to-Be','Spa voucher + keepsake box',150,150,'REC-008','Wrapped'],
  ['28 Jun 2025','Miscellaneous','Name cards & menus — printed',100,42,'REC-009',''],
  ['01 Jul 2025','Decorations & Florals','Fresh florals top-up',0,60,'REC-010','Extra centrepieces'],
  ['02 Jul 2025','Food & Catering','Dietary extras — GF platter',0,35,'REC-011','For Lucy & Kate'],
  ['05 Jul 2025','Games & Prizes','Extra prize — last minute',0,18,'REC-012',''],
  ['06 Jul 2025','Miscellaneous','Ribbon & wrapping supplies',0,22,'REC-013',''],
];

(async () => {
  const reqs = [];

  // Column widths
  [60, 200, 110, 110, 110, 140, 40, 180, 110].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: BU, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BU, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BU, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BU, dimension: 'ROWS', startIndex: 3, endIndex: 15 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BU, dimension: 'ROWS', startIndex: 15, endIndex: 17 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BU, dimension: 'ROWS', startIndex: 17, endIndex: 33 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Title row A1:I1
  reqs.push({ mergeCells: { range: gridRange(BU, 0, 1, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BU, 0, 1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Budget allocation header row 2 (B2:F2)
  reqs.push({ mergeCells: { range: gridRange(BU, 1, 2, 1, 6), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BU, 1, 2, 1, 6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // KPI block header row 2 (H2:I2)
  reqs.push({ mergeCells: { range: gridRange(BU, 1, 2, 7, 9), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BU, 1, 2, 7, 9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Budget allocation table rows 3-14 (cols A-F)
  for (let r = 2; r < 14; r++) {
    const bg = (r % 2 === 0) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(BU, r, r + 1, 0, 6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Totals row 15 (row index 14) for allocation table
  reqs.push({ repeatCell: {
    range: gridRange(BU, 14, 15, 1, 6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});

  // Number format for currency cols C, D, E in allocation (rows 3-15)
  [2, 3, 4].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(BU, 2, 15, c, c + 1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT' }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    }});
  });

  // Borders for allocation table (rows 2-14, cols A-F)
  reqs.push({ updateBorders: {
    range: gridRange(BU, 1, 15, 1, 6),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // KPI block rows 3-8 (H-I, indices 7-9)
  for (let r = 2; r < 8; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(BU, r, r + 1, 7, 8),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.champagne),
        textFormat: { foregroundColor: hex(C.muted), bold: true, fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(BU, r, r + 1, 8, 9),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.champagne),
        textFormat: { foregroundColor: hex(C.text), bold: true, fontSize: 11 },
        verticalAlignment: 'MIDDLE', horizontalAlignment: 'RIGHT',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
    }});
  }
  reqs.push({ updateBorders: {
    range: gridRange(BU, 1, 8, 7, 9),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Expense log section header row 16 (index 15)
  reqs.push({ mergeCells: { range: gridRange(BU, 15, 16, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BU, 15, 16, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.roseLilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Expense log col headers row 17 (index 16)
  reqs.push({ repeatCell: {
    range: gridRange(BU, 16, 17, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Expense log data rows 18-32 (indices 17-31)
  for (let r = 17; r < 32; r++) {
    const bg = (r % 2 === 1) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(BU, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Currency format for expense cols D, E, F (indices 3,4,5) in expense rows
  [3, 4, 5].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(BU, 17, 32, c, c + 1),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT' }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    }});
  });

  // Date col A — center, text format
  reqs.push({ repeatCell: {
    range: gridRange(BU, 17, 32, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'TEXT' } }},
    fields: 'userEnteredFormat(horizontalAlignment,numberFormat)',
  }});

  // Paid? col G (index 6) — center + checkbox validation
  reqs.push({ repeatCell: {
    range: gridRange(BU, 17, 32, 6, 7),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', backgroundColor: hex(C.ivory) }},
    fields: 'userEnteredFormat(horizontalAlignment,backgroundColor)',
  }});
  reqs.push({ setDataValidation: {
    range: gridRange(BU, 17, 32, 6, 7),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
  }});

  // Borders for expense log
  reqs.push({ updateBorders: {
    range: gridRange(BU, 16, 32, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'budget-format');

  // Values
  const vals = [];
  vals.push({ range: "'💰 Budget & Expenses'!A1", values: [['  💰  Budget & Expenses']] });
  vals.push({ range: "'💰 Budget & Expenses'!B2", values: [['BUDGET ALLOCATION']] });
  vals.push({ range: "'💰 Budget & Expenses'!H2", values: [['BUDGET SUMMARY']] });

  // Allocation col headers
  vals.push({ range: "'💰 Budget & Expenses'!B3", values: [['CATEGORY','BUDGETED','ACTUAL','REMAINING','SPEND BAR']] });

  // Category rows with formulas
  const catRows = CATEGORIES.map(([cat, bud], i) => {
    const row = 4 + i;  // rows 4-15 (1-indexed)
    return [
      cat,
      bud,
      `=IFERROR(SUMIF($B$18:$B$200,B${row},$E$18:$E$200),0)`,
      `=IFERROR(C${row}-D${row},"")`,
      `=IFERROR(SPARKLINE(D${row}/C${row},{"charttype","bar";"color1",IF(C${row}-D${row}<0,"#9B4A5A","#9B7DB5");"color2","#E5D5DF";"max",1}),"")`,
    ];
  });
  vals.push({ range: "'💰 Budget & Expenses'!B4", values: catRows });

  // Totals row (row 16 = 1-indexed)
  vals.push({ range: "'💰 Budget & Expenses'!B16", values: [['TOTAL','=SUM(C4:C15)','=SUM(D4:D15)','=SUM(E4:E15)','']] });

  // KPI block (H3:I8 = rows 3-8 in 1-indexed)
  vals.push({ range: "'💰 Budget & Expenses'!H3", values: [
    ['Total Budget', '=C16'],
    ['Total Spent',  '=D16'],
    ['Remaining',    '=E16'],
    ['% Spent',      '=IFERROR(TEXT(D16/C16,"0%"),"—")'],
    ['# Paid Items', '=COUNTIF(G18:G200,TRUE)'],
    ['# Categories', `=${CATEGORIES.length}`],
  ]});

  // Expense log header
  vals.push({ range: "'💰 Budget & Expenses'!A16", values: [['  💳  Expense Log']] });

  // Expense log col headers (row 17)
  vals.push({ range: "'💰 Budget & Expenses'!A17", values: [['DATE','CATEGORY','ITEM / VENDOR','BUDGETED','ACTUAL','VARIANCE','PAID?','RECEIPT REF','NOTES']] });

  // Expense rows (starting row 18)
  const expRows = EXPENSES.map(([date, cat, item, budg, actual, ref, notes], i) => {
    const row = 18 + i;
    return [date, cat, item, budg, actual, `=IFERROR(D${row}-E${row},"")`, actual > 0 ? 'TRUE' : 'FALSE', ref, notes];
  });
  vals.push({ range: "'💰 Budget & Expenses'!A18", values: expRows });

  await valuesBatchUpdate(id, vals, 'budget-values');
  console.log('Budget & Expenses complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
