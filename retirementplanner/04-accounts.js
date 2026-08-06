'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Retirement Accounts'];
const S = "'Retirement Accounts'";
const REF = "'Reference Data'";
const SETUP = "'Personal & Household Setup'";

// Column layout (0-indexed):
// A(0)=AccID  B(1)=Name  C(2)=Owner  D(3)=Type  E(4)=Institution
// F(5)=Balance  G(6)=AnnualContrib  H(7)=EmployerMatch  I(8)=TotalInflow
// J(9)=ExpReturn  K(10)=ProjBalance  L(11)=Status  M(12)=Notes

const COL_WIDTHS = [90, 210, 70, 150, 140, 120, 110, 110, 110, 90, 130, 110, 190];

// 18+ sample accounts
const ACCOUNTS = [
  // P1 = James Whitmore
  { name:'TechCorp 401(k)',            owner:'P1', type:'Pre-Tax 401(k)',       inst:'Fidelity',        bal:285000, contrib:15500, match:7750,  ret:0.07, status:'Active',     notes:'6% employer match on 6%' },
  { name:'Traditional IRA (Rollover)', owner:'P1', type:'Pre-Tax IRA',          inst:'Vanguard',        bal:142000, contrib:0,     match:0,     ret:0.07, status:'Active',     notes:'Consolidated from prior employer' },
  { name:'Roth IRA',                   owner:'P1', type:'Roth IRA',             inst:'Vanguard',        bal:92000,  contrib:7000,  match:0,     ret:0.07, status:'Active',     notes:'Max annual contribution' },
  { name:'Individual Brokerage',       owner:'P1', type:'Taxable Brokerage',    inst:'Charles Schwab',  bal:175000, contrib:12000, match:0,     ret:0.07, status:'Active',     notes:'Index funds + ETFs' },
  { name:'SEP-IRA (Consulting)',        owner:'P1', type:'Pre-Tax IRA',          inst:'Fidelity',        bal:45000,  contrib:8000,  match:0,     ret:0.07, status:'Active',     notes:'Self-employment income' },
  { name:'HSA Investment Account',     owner:'P1', type:'Pre-Tax IRA',          inst:'HealthEquity',    bal:28000,  contrib:3850,  match:0,     ret:0.06, status:'Active',     notes:'Invested portion; used for retirement healthcare' },
  { name:'Variable Annuity',           owner:'P1', type:'Tax-Deferred Annuity', inst:'Nationwide',      bal:65000,  contrib:3000,  match:0,     ret:0.05, status:'Active',     notes:'Surrender period ends 2028' },
  { name:'Prior Employer 401(k)',      owner:'P1', type:'Pre-Tax 401(k)',       inst:'Voya',            bal:0,      contrib:0,     match:0,     ret:0.07, status:'Rolled Over', notes:'Rolled into Vanguard IRA 2022' },
  // P2 = Patricia Whitmore
  { name:'Riverside Medical 403(b)',   owner:'P2', type:'Pre-Tax 403(b)',       inst:'TIAA',            bal:198000, contrib:12000, match:6000,  ret:0.07, status:'Active',     notes:'5% employer match' },
  { name:'Traditional IRA',           owner:'P2', type:'Pre-Tax IRA',          inst:'Fidelity',        bal:87000,  contrib:7000,  match:0,     ret:0.07, status:'Active',     notes:'' },
  { name:'Roth IRA',                   owner:'P2', type:'Roth IRA',             inst:'Fidelity',        bal:68000,  contrib:7000,  match:0,     ret:0.07, status:'Active',     notes:'Max annual contribution' },
  { name:'Roth 401(k)',               owner:'P2', type:'Roth 401(k)',           inst:'Empower',         bal:52000,  contrib:5000,  match:0,     ret:0.07, status:'Active',     notes:'' },
  { name:'457(b) Deferred Comp',      owner:'P2', type:'Pre-Tax 403(b)',       inst:'MissionSquare',   bal:78000,  contrib:8500,  match:3500,  ret:0.065,status:'Active',     notes:'Government 457(b) plan' },
  { name:'Individual Brokerage',      owner:'P2', type:'Taxable Brokerage',    inst:'TD Ameritrade',   bal:88000,  contrib:8000,  match:0,     ret:0.07, status:'Active',     notes:'' },
  { name:'HSA Investment Account',    owner:'P2', type:'Pre-Tax IRA',          inst:'Optum Bank',      bal:22000,  contrib:3850,  match:0,     ret:0.06, status:'Active',     notes:'Triple tax advantage' },
  { name:'Variable Annuity (Legacy)', owner:'P2', type:'Tax-Deferred Annuity', inst:'MetLife',         bal:35000,  contrib:0,     match:0,     ret:0.05, status:'Paused',     notes:'No new contributions; maturing' },
  // Joint
  { name:'Joint Brokerage',           owner:'Joint',type:'Taxable Brokerage',  inst:'Vanguard',        bal:245000, contrib:15000, match:0,     ret:0.07, status:'Active',     notes:'Dividend reinvestment on' },
  { name:'Prior Employer 403(b)',     owner:'P2', type:'Pre-Tax 403(b)',       inst:'Principal',       bal:0,      contrib:0,     match:0,     ret:0.07, status:'Rolled Over', notes:'Rolled into Fidelity IRA 2021' },
];

// FV formula: uses P1 or P2 years-to-retirement from Setup
function fvFormula(r) {
  const ytr = `IF(C${r}="P2",${SETUP}!E12,${SETUP}!B12)`;
  // FV of current balance + FV of annual annuity
  return `=IFERROR(IF(J${r}=0,F${r}+I${r}*(${ytr}),F${r}*(1+J${r})^(${ytr})+I${r}*((1+J${r})^(${ytr})-1)/J${r}),"")`;
}

(async () => {
  const vals = [];
  const fmt  = [];

  // Column widths
  COL_WIDTHS.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,500,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ===== ROW 1: TITLE =====
  vals.push({ range: `${S}!A1`, values: [['RETIREMENT ACCOUNTS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,13), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ===== ROW 2: SUBTITLE =====
  vals.push({ range: `${S}!A2`, values: [['Track all retirement accounts for James & Patricia Whitmore. Yellow = editable. Blue = formula.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,13), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrB), textFormat: { fontSize: 9, foregroundColor: hex(C.primaryText), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ===== ROWS 3-4: SUMMARY CARDS =====
  // 5 cards × 2 cols (A:B, C:D, E:F, G:H, I:K)
  const SUMMARY_CARDS = [
    { label: 'Total Balance (Active)',        formula: `=IFERROR(SUMIF(L7:L200,"Active",F7:F200),"—")`, fmt: '"$"#,##0' },
    { label: 'P1 Balance (James)',            formula: `=IFERROR(SUMPRODUCT((C7:C200="P1")*(F7:F200)),"—")`, fmt: '"$"#,##0' },
    { label: 'P2 Balance (Patricia)',         formula: `=IFERROR(SUMPRODUCT((C7:C200="P2")*(F7:F200)),"—")`, fmt: '"$"#,##0' },
    { label: 'Annual Contributions (Active)', formula: `=IFERROR(SUMIF(L7:L200,"Active",I7:I200),"—")`, fmt: '"$"#,##0' },
    { label: 'Projected Total at Retirement', formula: `=IFERROR(SUMIF(L7:L200,"Active",K7:K200),"—")`, fmt: '"$"#,##0' },
  ];

  // Card column spans: A:B, C:D, E:F, G:H, I:K (I-K = 3 cols for the last card)
  const cardColSpans = [[0,2],[2,4],[4,6],[6,8],[8,11]];

  SUMMARY_CARDS.forEach((card, ci) => {
    const [c1, c2] = cardColSpans[ci];
    const col = String.fromCharCode(65 + c1);
    // Label row (row 3 = 0-idx 2)
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 8, foregroundColor: hex('#D0D4E8'), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}3`, values: [[card.label]] });
    // Value row (row 4 = 0-idx 3)
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: 'NUMBER', pattern: card.fmt },
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}4`, values: [[card.formula]] });
  });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});
  // spacer row 5
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 },
    properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ===== ROW 6: COLUMN HEADERS =====
  const HEADERS = ['Acct ID','Account Name','Owner','Account Type','Institution',
    'Current Balance','Annual Contrib.','Employer Match','Total Inflow',
    'Exp. Return','Proj. Balance @\nRetirement','Status','Notes'];
  vals.push({ range: `${S}!A6`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    borders: { bottom: { style: 'SOLID_MEDIUM', color: hex(C.primary) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 },
    properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ===== FREEZE =====
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } },
    fields: 'gridProperties.frozenRowCount' }});

  // ===== DATA ROWS (rows 7-24, 0-idx 6-23) =====
  const DATA_START = 6; // 0-indexed

  ACCOUNTS.forEach((acc, ai) => {
    const rIdx = DATA_START + ai;
    const rowNum = rIdx + 1;
    const isEven = ai % 2 === 0;
    const isRolledOrClosed = acc.status === 'Rolled Over' || acc.status === 'Closed';

    // Owner-based color for col C
    const ownerColor = acc.owner === 'P1' ? C.preTax : acc.owner === 'P2' ? C.roth : C.taxable;
    // Account type color for col D
    const typeColor = acc.type.includes('Pre-Tax') ? C.preTax : acc.type.includes('Roth') ? C.roth :
                      acc.type.includes('Taxable') ? C.taxable : C.taxDeferred;

    const rowBg = isRolledOrClosed ? C.altRow : (isEven ? C.panel : '#F9F8F5');

    // Write data values
    vals.push({ range: `${S}!A${rowNum}`, values: [[
      `=IF(B${rowNum}="","","ACC-"&TEXT(ROW()-6,"000"))`,
      acc.name, acc.owner, acc.type, acc.inst,
      acc.bal, acc.contrib, acc.match,
      `=IFERROR(G${rowNum}+H${rowNum},"")`,
      acc.ret,
      fvFormula(rowNum),
      acc.status, acc.notes,
    ]] });

    // Row background
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,13), cell: { userEnteredFormat: {
      backgroundColor: hex(rowBg),
      textFormat: { fontSize: 9, foregroundColor: hex(isRolledOrClosed ? C.secText : C.text), fontFamily: 'Arial', italic: isRolledOrClosed },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});

    // Acct ID (col A) — formula, slightly muted
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,1), cell: { userEnteredFormat: {
      textFormat: { fontSize: 8, foregroundColor: hex(C.secText), italic: false },
      padding: { left: 4 },
    }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.padding' }});

    // Account Name (col B)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,1,2), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 9 },
      padding: { left: 4 },
    }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.padding' }});

    // Owner (col C)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,2,3), cell: { userEnteredFormat: {
      backgroundColor: hex(ownerColor), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text) },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});

    // Type (col D)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,3,4), cell: { userEnteredFormat: {
      backgroundColor: hex(typeColor), textFormat: { fontSize: 9 },
    }}, fields: 'userEnteredFormat' }});

    // Currency cols: F(5), G(6), H(7), I(8), K(10)
    [5,6,7,8,10].forEach(ci => {
      fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,ci,ci+1), cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
        horizontalAlignment: 'RIGHT',
        backgroundColor: hex(ci===8 || ci===10 ? C.formula : (ci===5 ? C.input : C.input)),
      }}, fields: 'userEnteredFormat' }});
    });
    // I and K are formulas
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,8,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), textFormat: { bold: true },
    }}, fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,10,11), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), textFormat: { bold: true },
    }}, fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat' }});

    // Return % (col J)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,9,10), cell: { userEnteredFormat: {
      numberFormat: { type: 'PERCENT', pattern: '0.0%' },
      horizontalAlignment: 'CENTER', backgroundColor: hex(C.input),
    }}, fields: 'userEnteredFormat' }});

    // Status (col L) — dropdown
    fmt.push({ setDataValidation: { range: gridRange(SID,rIdx,rIdx+1,11,12), rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$71:$A$74` }] },
      showCustomUi: true, strict: true,
    }}});

    // Row borders
    fmt.push({ updateBorders: { range: gridRange(SID,rIdx,rIdx+1,0,13), bottom: {
      style: 'SOLID', color: hex(C.border),
    }}});

    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: rIdx, endIndex: rIdx+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  });

  // ===== TOTALS ROW =====
  const TOTALS_ROW = DATA_START + ACCOUNTS.length; // 0-indexed = 24
  const totR = TOTALS_ROW + 1; // 1-indexed = 25
  const dataRange = `7:${DATA_START + ACCOUNTS.length}`;
  vals.push({ range: `${S}!A${totR}`, values: [[
    'TOTALS', '', '', '', '',
    `=IFERROR(SUM(F7:F${totR-1}),"")`,
    `=IFERROR(SUM(G7:G${totR-1}),"")`,
    `=IFERROR(SUM(H7:H${totR-1}),"")`,
    `=IFERROR(SUM(I7:I${totR-1}),"")`,
    '', // return
    `=IFERROR(SUM(K7:K${totR-1}),"")`,
    '', '',
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID,TOTALS_ROW,TOTALS_ROW+1,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
    borders: { top: { style: 'SOLID_MEDIUM', color: hex(C.primary) }, bottom: { style: 'SOLID', color: hex(C.border) } },
  }}, fields: 'userEnteredFormat' }});
  [5,6,7,8,10].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID,TOTALS_ROW,TOTALS_ROW+1,ci,ci+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT',
    }}, fields: 'userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: TOTALS_ROW, endIndex: TOTALS_ROW+1 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Conditional formatting: Active rows — no extra; Rolled Over / Closed — handled by italic/grey above
  // Highlight projected balance > $1M
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, DATA_START, DATA_START + ACCOUNTS.length, 10, 11)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1000000' }] },
      format: { backgroundColor: hex(C.success), textFormat: { bold: true } },
    },
  }, index: 0 }});

  await valuesBatchUpdate(id, vals, '04-accounts values');
  await batchUpdate(id, fmt, '04-accounts format');

  console.log('✅ Retirement Accounts done — 18 accounts seeded.');
  console.log('  Projected total at retirement in K column (FV formula).');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
