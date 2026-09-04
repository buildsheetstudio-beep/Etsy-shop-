'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Retirement Expenses'];
const S = "'Retirement Expenses'";
const REF = "'Reference Data'";

// Column layout (0-indexed):
// A(0)=ExpID  B(1)=ExpenseName  C(2)=Category  D(3)=Owner
// E(4)=MonthlyAmt  F(5)=AnnualAmt  G(6)=RetirementChange%  H(7)=AdjustedMonthly
// I(8)=InflationRate  J(9)=Priority  K(10)=Status  L(11)=Notes

const COL_WIDTHS = [85, 220, 170, 70, 110, 110, 120, 130, 100, 90, 110, 200];

// 28 expenses covering essential, housing, healthcare, discretionary, insurance, one-time
const EXPENSES = [
  // Essential
  { name:'Groceries & Food',              cat:'Food / Groceries',             owner:'Joint', monthly:1100, retChg: 0.00, infl:0.03, pri:'High',   status:'Active',    notes:'Plan for modest increase in retirement' },
  { name:'Utilities (Electric/Gas/Water)',cat:'Utilities',                    owner:'Joint', monthly:320,  retChg: 0.10, infl:0.03, pri:'High',   status:'Active',    notes:'Slight increase as more time at home' },
  { name:'Mobile Phone (2 Lines)',        cat:'Essential Living',             owner:'Joint', monthly:160,  retChg: 0.00, infl:0.02, pri:'High',   status:'Active',    notes:'' },
  { name:'Internet & Streaming Bundle',  cat:'Subscriptions',               owner:'Joint', monthly:185,  retChg: 0.00, infl:0.03, pri:'Medium', status:'Active',    notes:'Internet + Netflix, Prime, Music' },
  // Housing
  { name:'Mortgage Payment',             cat:'Housing / Rent / Mortgage',   owner:'Joint', monthly:2450, retChg:-1.00, infl:0,    pri:'High',   status:'Active',    notes:'Paid off by 2037 — eliminated at retirement' },
  { name:'Property Taxes',               cat:'Housing / Rent / Mortgage',   owner:'Joint', monthly:650,  retChg: 0.15, infl:0.03, pri:'High',   status:'Active',    notes:'Estimate ~3% annual increase' },
  { name:'Homeowner\'s Insurance',       cat:'Insurance',                   owner:'Joint', monthly:175,  retChg: 0.05, infl:0.03, pri:'High',   status:'Active',    notes:'' },
  { name:'Home Maintenance & Repair',    cat:'Home Repair / Maintenance',   owner:'Joint', monthly:400,  retChg: 0.25, infl:0.03, pri:'High',   status:'Active',    notes:'Increase as home ages' },
  // Transportation
  { name:'Auto Loan — James',            cat:'Transportation',               owner:'P1',   monthly:485,  retChg:-1.00, infl:0,    pri:'Medium', status:'Active',    notes:'Paid off 2028; eliminated pre-retirement' },
  { name:'Auto Loan — Patricia',         cat:'Transportation',               owner:'P2',   monthly:395,  retChg:-1.00, infl:0,    pri:'Medium', status:'Active',    notes:'Paid off 2029; eliminated pre-retirement' },
  { name:'Auto Insurance (2 Cars)',      cat:'Transportation',               owner:'Joint', monthly:280,  retChg:-0.20, infl:0.02, pri:'High',   status:'Active',    notes:'Reduce to 1 car in retirement possible' },
  { name:'Gas & Fuel',                   cat:'Transportation',               owner:'Joint', monthly:220,  retChg:-0.30, infl:0.04, pri:'Medium', status:'Active',    notes:'Drive less in retirement' },
  { name:'Vehicle Maintenance',          cat:'Transportation',               owner:'Joint', monthly:150,  retChg:-0.10, infl:0.03, pri:'Medium', status:'Active',    notes:'' },
  // Healthcare
  { name:'Employer Health Insurance',    cat:'Healthcare / Medical',         owner:'Joint', monthly:0,    retChg: 0,    infl:0,    pri:'High',   status:'Active',    notes:'Employer-sponsored; $0 cost pre-retirement' },
  { name:'Health Insurance (Self-Pay)',  cat:'Healthcare / Medical',         owner:'Joint', monthly:0,    retChg: 0,    infl:0.055,pri:'High',   status:'Projected', notes:'~$1,850/mo at retirement; replace employer plan' },
  { name:'Medicare Premiums (P1)',       cat:'Healthcare / Medical',         owner:'P1',   monthly:0,    retChg: 0,    infl:0.04, pri:'High',   status:'Projected', notes:'Starts at 65; est. $185/mo Part B' },
  { name:'Medicare Premiums (P2)',       cat:'Healthcare / Medical',         owner:'P2',   monthly:0,    retChg: 0,    infl:0.04, pri:'High',   status:'Projected', notes:'Starts at 65; est. $185/mo Part B' },
  { name:'Dental & Vision',             cat:'Healthcare / Medical',         owner:'Joint', monthly:130,  retChg: 0.20, infl:0.04, pri:'Medium', status:'Active',    notes:'Costs tend to rise with age' },
  { name:'Medications & Supplements',   cat:'Healthcare / Medical',         owner:'Joint', monthly:225,  retChg: 0.50, infl:0.055,pri:'High',   status:'Active',    notes:'Expect significant increase in retirement' },
  { name:'Long-term Care Insurance',    cat:'Healthcare / Medical',         owner:'Joint', monthly:425,  retChg: 0,    infl:0.04, pri:'High',   status:'Active',    notes:'Both covered; policy locked-in at current rate' },
  // Insurance
  { name:'Life Insurance — James (Term)',cat:'Insurance',                   owner:'P1',   monthly:380,  retChg:-1.00, infl:0,    pri:'Medium', status:'Active',    notes:'20-yr term ends 2033; eliminate before retirement' },
  { name:'Life Insurance — Patricia (Term)',cat:'Insurance',               owner:'P2',   monthly:285,  retChg:-1.00, infl:0,    pri:'Medium', status:'Active',    notes:'20-yr term ends 2035' },
  // Discretionary
  { name:'Dining Out & Restaurants',    cat:'Dining Out',                  owner:'Joint', monthly:450,  retChg: 0.30, infl:0.03, pri:'Medium', status:'Active',    notes:'Expect to dine out more in retirement' },
  { name:'Entertainment & Events',      cat:'Entertainment',               owner:'Joint', monthly:280,  retChg: 0.40, infl:0.03, pri:'Low',    status:'Active',    notes:'' },
  { name:'Travel & Vacation',           cat:'Travel & Leisure',            owner:'Joint', monthly:800,  retChg: 0.50, infl:0.03, pri:'Medium', status:'Active',    notes:'Plan to travel more in early retirement' },
  { name:'Hobbies & Recreation',        cat:'Hobbies',                     owner:'Joint', monthly:350,  retChg: 0.40, infl:0.03, pri:'Medium', status:'Active',    notes:'Golf, gardening, photography' },
  { name:'Gifts, Donations & Charitable',cat:'Gifts & Charitable',         owner:'Joint', monthly:400,  retChg: 0.00, infl:0,    pri:'Low',    status:'Active',    notes:'' },
  { name:'Clothing & Personal Care',    cat:'Personal Care',               owner:'Joint', monthly:380,  retChg:-0.20, infl:0.02, pri:'Low',    status:'Active',    notes:'Expect reduction in retirement' },
];

(async () => {
  const vals = [];
  const fmt  = [];

  COL_WIDTHS.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  fmt.push({ repeatCell: { range: gridRange(SID,0,400,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ===== TITLE =====
  vals.push({ range: `${S}!A1`, values: [['RETIREMENT EXPENSES']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A2`, values: [['Monthly expense estimates for James & Patricia Whitmore. "Retirement Change %" adjusts amount at retirement. Yellow = editable.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,12), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrB), textFormat: { fontSize: 9, foregroundColor: hex(C.primaryText), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ===== SUMMARY CARDS =====
  const CARDS = [
    { label: 'Total Monthly (Active)',      formula: `=IFERROR(SUMIF(K7:K200,"Active",E7:E200),"—")`,          fmt: '"$"#,##0' },
    { label: 'Essential & Housing',         formula: `=IFERROR(SUMIFS(E7:E200,K7:K200,"Active",C7:C200,"Housing / Rent / Mortgage")+SUMIFS(E7:E200,K7:K200,"Active",C7:C200,"Food / Groceries"),"—")`, fmt: '"$"#,##0' },
    { label: 'Healthcare (Active)',         formula: `=IFERROR(SUMIFS(E7:E200,K7:K200,"Active",C7:C200,"Healthcare / Medical"),"—")`, fmt: '"$"#,##0' },
    { label: 'Discretionary (Active)',      formula: `=IFERROR(SUMIFS(E7:E200,K7:K200,"Active",C7:C200,"Travel & Leisure")+SUMIFS(E7:E200,K7:K200,"Active",C7:C200,"Dining Out")+SUMIFS(E7:E200,K7:K200,"Active",C7:C200,"Entertainment"),"—")`, fmt: '"$"#,##0' },
    { label: 'Total Annual (Active)',       formula: `=IFERROR(SUMIF(K7:K200,"Active",F7:F200),"—")`,          fmt: '"$"#,##0' },
  ];
  const cardSpans = [[0,2],[2,4],[4,6],[6,8],[8,12]];
  CARDS.forEach((card, ci) => {
    const [c1,c2] = cardSpans[ci];
    const col = String.fromCharCode(65+c1);
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB), textFormat: { bold: true, fontSize: 8, foregroundColor: hex('#D0D4E8'), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}3`, values: [[card.label]] });
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.attention), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: 'NUMBER', pattern: card.fmt },
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}4`, values: [[card.formula]] });
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ===== COLUMN HEADERS =====
  vals.push({ range: `${S}!A6`, values: [[
    'Exp ID','Expense Name','Category','Owner','Monthly Amount\n(Today\'s $)',
    'Annual Amount','Retirement\nChange %','Adjusted Monthly\n@ Retirement',
    'Inflation Rate','Priority','Status','Notes',
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    borders: { bottom: { style: 'SOLID_MEDIUM', color: hex(C.primary) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' }});

  // ===== DATA ROWS =====
  const DATA_START = 6;
  const catColor = cat =>
    cat.includes('Housing') || cat.includes('Essential') || cat.includes('Food') ? C.taxable :
    cat.includes('Healthcare') ? C.roth :
    cat.includes('Transportation') ? C.preTax :
    cat.includes('Insurance') ? C.taxDeferred :
    C.altRow;

  const priColor = pri =>
    pri === 'High' ? '#F8DCDC' : pri === 'Medium' ? '#FDF3DC' : '#FFFFFF';

  EXPENSES.forEach((exp, ei) => {
    const rIdx = DATA_START + ei;
    const rowNum = rIdx + 1;
    const isEven = ei % 2 === 0;
    const isProjected = exp.status === 'Projected';
    const rowBg = isProjected ? C.altRow : (isEven ? C.panel : '#F9F8F5');
    const ownerColor = exp.owner === 'P1' ? C.preTax : exp.owner === 'P2' ? C.roth : C.taxable;

    vals.push({ range: `${S}!A${rowNum}`, values: [[
      `=IF(B${rowNum}="","","EXP-"&TEXT(ROW()-6,"000"))`,
      exp.name, exp.cat, exp.owner, exp.monthly,
      `=IFERROR(E${rowNum}*12,"")`,
      exp.retChg,
      `=IFERROR(E${rowNum}*(1+G${rowNum}),"")`,
      exp.infl,
      exp.pri, exp.status, exp.notes,
    ]] });

    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,12), cell: { userEnteredFormat: {
      backgroundColor: hex(rowBg),
      textFormat: { fontSize: 9, foregroundColor: hex(isProjected ? C.secText : C.text), fontFamily: 'Arial', italic: isProjected },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});

    // ID
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,1), cell: { userEnteredFormat: {
      textFormat: { fontSize: 8, foregroundColor: hex(C.secText), italic: false }, padding: { left: 4 },
    }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.padding' }});

    // Name (bold)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,1,2), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 9, italic: false }, padding: { left: 4 },
    }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.padding' }});

    // Category color
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,2,3), cell: { userEnteredFormat: {
      backgroundColor: hex(catColor(exp.cat)), textFormat: { fontSize: 9 }, padding: { left: 4 },
    }}, fields: 'userEnteredFormat' }});

    // Owner
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,3,4), cell: { userEnteredFormat: {
      backgroundColor: hex(ownerColor), textFormat: { bold: true, fontSize: 9, italic: false }, horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});

    // Monthly (col E)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,4,5), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
      horizontalAlignment: 'RIGHT', textFormat: { bold: true, fontSize: 9, italic: false },
    }}, fields: 'userEnteredFormat' }});

    // Annual (col F, formula)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,5,6), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
      horizontalAlignment: 'RIGHT', textFormat: { bold: true, fontSize: 9, italic: false },
    }}, fields: 'userEnteredFormat' }});

    // Retirement Change % (col G)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,6,7), cell: { userEnteredFormat: {
      backgroundColor: exp.retChg < 0 ? hex('#E8F8ED') : exp.retChg > 0 ? hex('#FDF3DC') : hex(C.input),
      numberFormat: { type: 'PERCENT', pattern: '0%' }, horizontalAlignment: 'CENTER',
      textFormat: { fontSize: 9, italic: false },
    }}, fields: 'userEnteredFormat' }});

    // Adjusted Monthly (col H, formula)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,7,8), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
      horizontalAlignment: 'RIGHT', textFormat: { bold: true, fontSize: 9, italic: false },
    }}, fields: 'userEnteredFormat' }});

    // Inflation Rate (col I)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,8,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), numberFormat: { type: 'PERCENT', pattern: '0.0%' },
      horizontalAlignment: 'CENTER', textFormat: { fontSize: 9, italic: false },
    }}, fields: 'userEnteredFormat' }});

    // Priority (col J) — colored badge
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,9,10), cell: { userEnteredFormat: {
      backgroundColor: hex(priColor(exp.pri)), horizontalAlignment: 'CENTER',
      textFormat: { bold: true, fontSize: 9, italic: false },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ setDataValidation: { range: gridRange(SID,rIdx,rIdx+1,9,10), rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$A$101:$A$104` }] },
      showCustomUi: true, strict: true,
    }}});

    // Status (col K)
    fmt.push({ setDataValidation: { range: gridRange(SID,rIdx,rIdx+1,10,11), rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `='Reference Data'!$A$82:$A$84` }] },
      showCustomUi: true, strict: true,
    }}});

    fmt.push({ updateBorders: { range: gridRange(SID,rIdx,rIdx+1,0,12), bottom: { style: 'SOLID', color: hex(C.border) }}});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: rIdx, endIndex: rIdx+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  });

  // ===== TOTALS ROW =====
  const TOTALS_ROW = DATA_START + EXPENSES.length;
  const totR = TOTALS_ROW + 1;
  vals.push({ range: `${S}!A${totR}`, values: [[
    'TOTALS','','','',
    `=IFERROR(SUM(E7:E${totR-1}),"")`,
    `=IFERROR(SUM(F7:F${totR-1}),"")`,
    '','','','','','',
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID,TOTALS_ROW,TOTALS_ROW+1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
    borders: { top: { style: 'SOLID_MEDIUM', color: hex(C.primary) } },
  }}, fields: 'userEnteredFormat' }});
  [4,5].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID,TOTALS_ROW,TOTALS_ROW+1,ci,ci+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT',
    }}, fields: 'userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: TOTALS_ROW, endIndex: TOTALS_ROW+1 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Conditional: highlight high-priority expenses in data column A
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, DATA_START, DATA_START + EXPENSES.length, 9, 10)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'High' }] },
      format: { backgroundColor: hex('#F8DCDC'), textFormat: { bold: true } },
    },
  }, index: 0 }});

  await valuesBatchUpdate(id, vals, '06-expenses values');
  await batchUpdate(id, fmt, '06-expenses format');

  console.log(`✅ Retirement Expenses done — ${EXPENSES.length} expenses seeded.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
