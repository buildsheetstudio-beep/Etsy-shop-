'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BUD = sheetMap['💰 Budget & Expenses'];

const CATS = [
  'Venue Hire','Food & Catering','Cake & Desserts','Drinks',
  'Decorations & Florals','Invitations & Stationery','Party Favours','Games & Prizes',
  'Photography','Gift for Bride','Entertainment','Miscellaneous',
];
const BUDGETS = [600,650,400,210,460,100,200,150,400,200,100,100];

const EXPENSES = [
  ['Venue Hire',       '2026-09-01','The Garden Room hire fee',         400,400],
  ['Venue Hire',       '2026-09-15','Room dressing / chair covers',      200,185],
  ['Food & Catering',  '2026-09-10','Afternoon tea catering deposit',    325,325],
  ['Food & Catering',  '2026-09-20','Final catering balance',            325,310],
  ['Cake & Desserts',  '2026-09-05','Custom bridal shower cake',         280,280],
  ['Cake & Desserts',  '2026-09-05','Macaron tower',                     120,115],
  ['Drinks',           '2026-09-12','Prosecco & cocktail mixers',        150,148],
  ['Drinks',           '2026-09-12','Non-alcoholic options',              60, 55],
  ['Decorations & Florals','2026-08-20','Centrepiece florals',           260,255],
  ['Decorations & Florals','2026-08-25','Table garlands & candles',      200,190],
  ['Invitations & Stationery','2026-08-01','Custom invitations (22 sets)',75, 78],
  ['Invitations & Stationery','2026-08-01','Menus & place cards',         25, 22],
  ['Party Favours',    '2026-09-01','Personalised candle favours (20)',  200,195],
  ['Games & Prizes',   '2026-08-28','Prize hamper & game supplies',      150,145],
  ['Photography',      '2026-07-15','Event photographer deposit',        400,400],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'💰 Budget & Expenses'";

  // Column widths: A=180, B=120, C=120, D=110, E=110, F=110, G=80, H=90, I=80, J=160
  const colW = [180,120,120,110,110,110,80,90,80,160];
  colW.forEach((w,i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: BUD, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  // Row 0: title
  reqs.push({ mergeCells: { range: gridRange(BUD,0,1,1,10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BUD,0,1,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(BUD,0,1,0,1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.dustyRose) }},
    fields: 'userEnteredFormat(backgroundColor)',
  }});
  vals.push({ range: `${TAB}!B1`, values: [['💰 Budget & Expenses']] });

  // Row 1: section header for allocation
  reqs.push({ mergeCells: { range: gridRange(BUD,1,2,0,10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BUD,1,2,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['BUDGET ALLOCATION']] });

  // Row 2: column headers for budget table
  const budHeaders = ['CATEGORY','BUDGETED','SPENT (AUTO)','REMAINING','% USED','SPARKLINE','','KPI SUMMARY',''];
  reqs.push({ repeatCell: {
    range: gridRange(BUD,2,3,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A3`, values: [budHeaders] });

  // Rows 3-14: budget categories (index 3..14)
  for (let i = 0; i < CATS.length; i++) {
    const row = 4 + i; // 1-indexed: rows 4..15
    const ri  = 3 + i; // 0-indexed: rows 3..14
    const bg  = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(BUD, ri, ri+1, 0, 6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    const bRow = `B${row}`;
    const cRow = `C${row}`;
    const dRow = `D${row}`;
    const eRow = `E${row}`;
    vals.push({ range: `${TAB}!A${row}`, values: [[
      CATS[i],
      BUDGETS[i],
      `=IFERROR(SUMIF($A$19:$A$200,A${row},$F$19:$F$200),0)`,
      `=IFERROR(${bRow}-${cRow},"")`,
      `=IFERROR(${cRow}/${bRow},"")`,
      `=SPARKLINE(${cRow}/${bRow},{"charttype","bar";"color1",IF(${dRow}<0,"#8B3A4A","#C9A96E");"max",1})`,
    ]]});
    // E column format as percentage
    reqs.push({ repeatCell: {
      range: gridRange(BUD, ri, ri+1, 4, 5),
      cell: { userEnteredFormat: {
        numberFormat: { type: 'PERCENT', pattern: '0%' },
        backgroundColor: hex(bg),
      }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
    // B,C,D as currency
    reqs.push({ repeatCell: {
      range: gridRange(BUD, ri, ri+1, 1, 4),
      cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
        backgroundColor: hex(bg),
      }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
  }

  // KPI block: H col (index 7), I col (index 8), rows 3-8 (0-indexed)
  const kpiLabels = ['Total Budget','Total Spent','Remaining','% Used','Largest Expense','Avg per Guest'];
  const kpiFormulas = [
    '=SUM(B4:B15)',
    '=SUM(C4:C15)',
    '=B18-C18',
    '=IFERROR(C18/B18,"")',
    '=IFERROR(MAX(F19:F200),"")',
    '=IFERROR(B18/COUNTA(\'👰 Guest List & Seating\'!A4:A23),"")',
  ];
  for (let i = 0; i < 6; i++) {
    const ri = 3 + i;
    reqs.push({ repeatCell: {
      range: gridRange(BUD, ri, ri+1, 7, 9),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.parchment),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        borders: { bottom: { style: 'SOLID', color: hex(C.border) }},
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,borders)',
    }});
    vals.push({ range: `${TAB}!H${4+i}`, values: [[kpiLabels[i], kpiFormulas[i]]] });
  }
  // KPI header
  reqs.push({ mergeCells: { range: gridRange(BUD,2,3,7,9), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${TAB}!H3`, values: [['KPI SUMMARY']] });
  reqs.push({ repeatCell: {
    range: gridRange(BUD,2,3,7,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.gold),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Format KPI row 6 (% Used) as percent
  reqs.push({ repeatCell: {
    range: gridRange(BUD, 6, 7, 8, 9),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0.0%' }}},
    fields: 'userEnteredFormat(numberFormat)',
  }});
  // Format KPI currency rows
  [0,1,2,4,5].forEach(ki => {
    reqs.push({ repeatCell: {
      range: gridRange(BUD, 3+ki, 4+ki, 8, 9),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }}},
      fields: 'userEnteredFormat(numberFormat)',
    }});
  });

  // Row 16: totals row (0-indexed row 15)
  reqs.push({ repeatCell: {
    range: gridRange(BUD,15,16,0,6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A16`, values: [['TOTAL','=SUM(B4:B15)','=SUM(C4:C15)','=IFERROR(B16-C16,"")','=IFERROR(C16/B16,"")','',]] });
  reqs.push({ repeatCell: {
    range: gridRange(BUD,15,16,1,4),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }}},
    fields: 'userEnteredFormat(numberFormat)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(BUD,15,16,4,5),
    cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' }}},
    fields: 'userEnteredFormat(numberFormat)',
  }});

  // Row 17: divider
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BUD, dimension: 'ROWS', startIndex: 16, endIndex: 17 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(BUD,16,17,0,10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Row 18: expense log section header
  reqs.push({ mergeCells: { range: gridRange(BUD,17,18,0,10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(BUD,17,18,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A18`, values: [['EXPENSE LOG']] });

  // Row 19: expense col headers (0-indexed row 18)
  const expHeaders = ['CATEGORY','DATE','ITEM / VENDOR','BUDGETED','ACTUAL','VARIANCE','PAID?','PAYMENT METHOD','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(BUD,18,19,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A19`, values: [expHeaders] });

  // Expense rows 20-34 (0-indexed 19-33)
  for (let i = 0; i < EXPENSES.length; i++) {
    const [cat, date, item, budgeted, actual] = EXPENSES[i];
    const row  = 20 + i; // 1-indexed
    const ri   = 19 + i; // 0-indexed
    const bg   = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(BUD, ri, ri+1, 0, 9),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    // Currency format for D,E,F cols (3,4,5)
    reqs.push({ repeatCell: {
      range: gridRange(BUD, ri, ri+1, 3, 6),
      cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
        backgroundColor: hex(bg),
      }},
      fields: 'userEnteredFormat(numberFormat,backgroundColor)',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [[
      cat, date, item, budgeted, actual,
      `=IFERROR(D${row}-E${row},"")`,
      'TRUE',
      'Bank Transfer',
      '',
    ]]});
  }

  // Row height for title
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: BUD, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'budget-format');
  await valuesBatchUpdate(id, vals, 'budget-values');
  console.log('Budget & Expenses complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
