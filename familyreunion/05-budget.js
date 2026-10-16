'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BDG = sheetMap['💰 Budget & Expenses'];

const CATEGORIES = ['Venue','Accommodation','Catering & Food','Drinks & Ice','Activities & Games','Decorations','Printed Materials & Signage','Transport & Shuttle','Photography','Keepsakes & Gifts','Miscellaneous','Buffer'];
const BUDGETS    = [1200, 540, 600, 260, 145, 90, 65, 120, 80, 560, 55, 285]; // total = $4000

// [category, date, item, paidBy, budgeted, actual, paid]
const EXPENSES = [
  ['Venue',                    '2025-07-01','Lake Eildon Cabin hire — 3 nights','Henderson Family',900,900,true],
  ['Venue',                    '2025-07-01','Venue deposit','Henderson Family',300,300,true],
  ['Accommodation',            '2025-07-15','Overflow hotel rooms × 2','Henderson Family',360,340,true],
  ['Catering & Food',          '2025-07-20','BBQ meats and seafood','Henderson + Thompson',420,395,true],
  ['Catering & Food',          '2025-07-20','Fresh produce, salads, bread','Wilson Family',180,165,true],
  ['Drinks & Ice',             '2025-07-20','Beer, wine, soft drinks','Smith Family',220,210,true],
  ['Drinks & Ice',             '2025-07-20','Ice × 4 bags, cups','Henderson Family',40,38,true],
  ['Activities & Games',       '2025-07-25','Lawn games set (bocce, cricket)','Henderson Family',85,79,true],
  ['Activities & Games',       '2025-07-25','Prizes for family trivia × 5','Henderson Family',60,55,true],
  ['Decorations',              '2025-07-28','Bunting, table covers, centrepieces','Thompson Family',90,82,true],
  ['Printed Materials & Signage','2025-07-15','Welcome banner + name tags','Henderson Family',65,58,true],
  ['Photography',              '2025-08-01','Disposable cameras × 6 + shared album','Henderson Family',80,74,true],
  ['Keepsakes & Gifts',        '2025-08-10','Custom family reunion t-shirts × 28','Henderson Family',560,532,false],
  ['Transport & Shuttle',      '2025-08-12','Airport pickup runs × 3','Wilson Family',120,95,true],
  ['Miscellaneous',            '2025-08-14','Sunscreen, first aid, bug spray','Henderson Family',55,48,true],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'💰 Budget & Expenses'";
  const EXP = "'💰 Budget & Expenses'"; // same tab
  const NCOLS = 9; // A-I

  // Column widths
  [160,100,220,140,110,110,100,70,200].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: BDG, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });
  // Col J-K for KPI block
  [160,140].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: BDG, dimension: 'COLUMNS', startIndex: 9+i, endIndex: 10+i },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner
  reqs.push({ mergeCells: { range: gridRange(BDG,0,1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BDG,0,1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  💰  Budget & Expenses — Henderson Family Reunion 2025']] });

  // Row 1: Budget allocation section header
  reqs.push({ mergeCells: { range: gridRange(BDG,1,2,0,7), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BDG,1,2,0,7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['  📊  BUDGET ALLOCATION BY CATEGORY']] });

  // Row 2: Budget table headers (0-indexed row 2 = GSheets row 3)
  const budgetHdrs = ['CATEGORY','BUDGETED','ACTUAL SPENT','VARIANCE','% USED'];
  reqs.push({ repeatCell: {
    range: gridRange(BDG,2,3,0,5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [budgetHdrs] });

  // Budget rows (0-indexed 3-14 = GSheets rows 4-15)
  for (let i = 0; i < CATEGORIES.length; i++) {
    const ri  = 3 + i;
    const row = 4 + i;
    const bg  = i % 2 === 0 ? C.white : C.warmCream;
    reqs.push({ repeatCell: {
      range: gridRange(BDG, ri, ri+1, 0, 6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    // Currency format cols B-D
    reqs.push({ repeatCell: {
      range: gridRange(BDG, ri, ri+1, 1, 4),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    }});
    // % format col E
    reqs.push({ repeatCell: {
      range: gridRange(BDG, ri, ri+1, 4, 5),
      cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' }, horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    }});
    const cat     = CATEGORIES[i];
    const budgeted = BUDGETS[i];
    const actualFormula  = `=IFERROR(SUMIF(A18:A300,"${cat}",F18:F300),0)`;
    const varFormula     = `=IFERROR(C${row}-B${row},0)`;
    const pctUsedFormula = `=IFERROR(SPARKLINE(C${row}/B${row},{"charttype","bar";"color1",IF(E${row}<0,"${C.burnSienna}",IF(C${row}/B${row}>1,"${C.burnSienna}","${C.forestGreen}"));"max",1}),"")`;
    vals.push({ range: `${TAB}!A${row}`, values: [[cat, budgeted, actualFormula, varFormula]] });
    vals.push({ range: `${TAB}!E${row}`, values: [[pctUsedFormula]] });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 3, endIndex: 15 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Budget totals row (0-indexed 15 = GSheets row 16)
  reqs.push({ repeatCell: {
    range: gridRange(BDG,15,16,0,5),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 15, endIndex: 16 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A16`, values: [['TOTAL','=SUM(B4:B15)','=SUM(C4:C15)','=SUM(D4:D15)','']] });

  // ── KPI Block (cols J-K, rows 3-8) ──
  const kpiLabels = ['TOTAL BUDGET','TOTAL SPENT','REMAINING','% USED','COST / CONFIRMED GUEST','REMAINING / GUEST'];
  const kpiFormulas = [
    '=IFERROR(SUM(B4:B15),"—")',
    '=IFERROR(SUM(F18:F300),"—")',
    '=IFERROR(J4-J5,"—")',
    '=IFERROR(J5/J4,"—")',
    `=IFERROR(J5/COUNTIF('👨‍👩‍👧‍👦 Guest List & RSVPs'!H4:H200,"Confirmed"),"—")`,
    `=IFERROR(J6/COUNTIF('👨‍👩‍👧‍👦 Guest List & RSVPs'!H4:H200,"Confirmed"),"—")`,
  ];
  reqs.push({ repeatCell: {
    range: gridRange(BDG,2,9,9,11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.warmCream),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(BDG,2,3,9,11),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!J3`, values: [['KPI','VALUE']] });
  for (let i = 0; i < kpiLabels.length; i++) {
    const row = 4 + i;
    reqs.push({ repeatCell: {
      range: gridRange(BDG,3+i,4+i,10,11),
      cell: { userEnteredFormat: {
        textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.warmAmber) },
        horizontalAlignment: 'CENTER',
        numberFormat: i === 3 ? { type: 'PERCENT', pattern: '0%' } : (i >= 4 ? { type: 'CURRENCY', pattern: '$#,##0.00' } : { type: 'CURRENCY', pattern: '$#,##0.00' }),
      }},
      fields: 'userEnteredFormat(textFormat,horizontalAlignment,numberFormat)',
    }});
    vals.push({ range: `${TAB}!J${row}`, values: [[kpiLabels[i], kpiFormulas[i]]] });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 2, endIndex: 9 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Divider row 16 (0-indexed)
  reqs.push({ repeatCell: {
    range: gridRange(BDG,16,17,0,NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmAmber) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 16, endIndex: 17 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  // Expense log section header (0-indexed row 17 = GSheets row 18 — but we use row 17 header and 18 as data start... wait, let me recalculate)
  // EXPENSES is at GSheets rows 18+ per the spec — so 0-indexed 17 = GSheets 18 is data
  // Let's put expense header at 0-indexed 17 (GSheets 18) and data at 0-indexed 18+ (GSheets 19+)
  // But SUMIF refs A18:A300 so data must start at GSheets row 18
  // Let's put section header at GSheets 17 (0-indexed 16) — but 16 is the divider.
  // Move: divider at 0-indexed 15 (was totals row), section header at 16, col headers at 17, data at 18 (GSheets 18)

  // Revise: move expense section header to GSheets row 17 (0-indexed 16) and col headers to GSheets 18, data from GSheets 19
  // BUT the SUMIF references F18:F300 in the spec (GSheets 18 = data row 1).
  // Let's keep data starting GSheets 18. So:
  //   GSheets 16 = totals (0-indexed 15) ✓ already done
  //   GSheets 17 = divider (make thin amber) — let's call 0-indexed 16
  //   GSheets 18 = expense section header — 0-indexed 17
  //   GSheets 19 = col headers — 0-indexed 18... but data must start GSheets 18

  // Simplest fix: let data start at GSheets 19 (0-indexed 18) and change SUMIF to use F19:F300
  // Actually, re-read spec: SUMIF uses A18:A300 — let's just put col headers at row 17 and data at row 18 (GSheets)
  // That means: row 16 = section header, row 17 = col headers (0-indexed), data at 0-indexed 17 = GSheets 18 for headers...
  // Let me simplify: put everything 1 row lower
  // - 0-indexed 15 = budget totals (GSheets 16)
  // - 0-indexed 16 = thin divider (GSheets 17)
  // - 0-indexed 17 = expense section header (GSheets 18) — matches A18 in SUMIF labels? No, SUMIF is A18:A300 for data
  // I need data in A18+. So section header = row 17 (GSheets 18) can't work as section header.
  // SIMPLEST: put col headers at GSheets 17, data at GSheets 18+
  // So: 0-indexed 16 = expense section header (GSheets 17), 0-indexed 17 = col headers (GSheets 18 — but wait that conflicts with data)
  //
  // Actually just: expense section header = GSheets 17 (0-indexed 16)
  //                col headers = GSheets 18 (0-indexed 17) -- and the SUMIF formulas reference A18:A300 which picks up col headers...
  // So just change SUMIF to reference A19:A300. That's the cleanest approach.
  // The budget SUMIF was already written as `SUMIF(A18:A300,...,F18:F300)` above. But actual data will start at GSheets 19.
  // The SUMIF will still work correctly since the header text won't match any category string.

  // Expense section header (0-indexed 16 = GSheets 17)
  reqs.push({ mergeCells: { range: gridRange(BDG,16,17,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BDG,16,17,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 16, endIndex: 17 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A17`, values: [['  🧾  EXPENSE LOG']] });

  // Expense col headers (0-indexed 17 = GSheets 18)
  const expHdrs = ['CATEGORY','DATE','ITEM / VENDOR','PAID BY','BUDGETED','ACTUAL','VARIANCE','PAID?','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(BDG,17,18,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 17, endIndex: 18 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A18`, values: [expHdrs] });

  // Expense data rows (0-indexed 18+ = GSheets 19+)
  for (let i = 0; i < EXPENSES.length; i++) {
    const ri  = 18 + i;
    const row = 19 + i;
    const bg  = i % 2 === 0 ? C.white : C.warmCream;
    const [cat, date, item, paidBy, budgeted, actual, paid] = EXPENSES[i];
    reqs.push({ repeatCell: {
      range: gridRange(BDG, ri, ri+1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    // Currency format cols E, F, G
    reqs.push({ repeatCell: {
      range: gridRange(BDG, ri, ri+1, 4, 7),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }, horizontalAlignment: 'CENTER', backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment,backgroundColor)',
    }});
    // Date col B
    reqs.push({ repeatCell: {
      range: gridRange(BDG, ri, ri+1, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' }, backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
    // Center col H (paid)
    reqs.push({ repeatCell: {
      range: gridRange(BDG, ri, ri+1, 7, 8),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', backgroundColor: hex(bg) }},
      fields: 'userEnteredFormat(horizontalAlignment,backgroundColor)',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [[cat, date, item, paidBy, budgeted, actual, `=IFERROR(E${row}-F${row},"")`, paid ? 'TRUE' : 'FALSE', '']] });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BDG, dimension: 'ROWS', startIndex: 18, endIndex: 18+EXPENSES.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Checkbox for col H (paid) on expense rows
  reqs.push({ setDataValidation: {
    range: gridRange(BDG, 18, 18+EXPENSES.length, 7, 8),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true, strict: true },
  }});

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(BDG,17,18+EXPENSES.length,0,NCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});
  reqs.push({ updateBorders: {
    range: gridRange(BDG,2,16,0,5),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'budget-format');
  await valuesBatchUpdate(id, vals, 'budget-values');
  console.log('Budget & Expenses complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
