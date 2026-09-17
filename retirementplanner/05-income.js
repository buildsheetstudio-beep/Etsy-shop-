'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Retirement Income'];
const S = "'Retirement Income'";
const REF = "'Reference Data'";
const SETUP = "'Personal & Household Setup'";

// Column layout (0-indexed):
// A(0)=IncID  B(1)=Source  C(2)=Owner  D(3)=Type  E(4)=AnnualAmt
// F(5)=StartAge  G(6)=StartYear  H(7)=EndAge  I(8)=EndYear
// J(9)=InflAdj(checkbox)  K(10)=COLARate  L(11)=Status  M(12)=Notes

const COL_WIDTHS = [85, 200, 70, 160, 120, 80, 90, 80, 100, 80, 90, 110, 190];

// Income streams for James & Patricia Whitmore
const INCOME = [
  // Social Security
  { source:'Social Security',             owner:'P1', type:'Social Security',       amt:40800,  startAge:67, endAge:'', inflAdj:true,  cola:0.025, status:'Projected',  notes:'$3,400/mo est. at 67; COLA per SSA' },
  { source:'Social Security',             owner:'P2', type:'Social Security',       amt:33000,  startAge:67, endAge:'', inflAdj:true,  cola:0.025, status:'Projected',  notes:'$2,750/mo est. at 67' },
  // Pension
  { source:'Riverside Medical Pension',   owner:'P2', type:'Pension',               amt:14400,  startAge:65, endAge:'', inflAdj:false, cola:0,     status:'Projected',  notes:'$1,200/mo vested pension; no COLA' },
  // Part-time work (during and after retirement)
  { source:'Consulting Income',           owner:'P1', type:'Part-time Work',        amt:45000,  startAge:56, endAge:67, inflAdj:false, cola:0,     status:'Active',     notes:'Self-employment consulting; ends at retirement' },
  { source:'Post-Retirement Consulting',  owner:'P1', type:'Part-time Work',        amt:18000,  startAge:67, endAge:73, inflAdj:false, cola:0,     status:'Projected',  notes:'Light consulting work in early retirement' },
  { source:'Part-time Teaching (Post-Ret)',owner:'P2',type:'Part-time Work',        amt:24000,  startAge:65, endAge:72, inflAdj:false, cola:0,     status:'Projected',  notes:'Adjunct teaching after retirement' },
  // Rental
  { source:'Rental Property (Duplex)',    owner:'Joint',type:'Rental Income',       amt:21600,  startAge:0,  endAge:'', inflAdj:true,  cola:0.03,  status:'Active',     notes:'$1,800/mo net; annual rent increases ~3%' },
  // Annuity
  { source:'Nationwide Variable Annuity', owner:'P1', type:'Annuity',               amt:4800,   startAge:72, endAge:'', inflAdj:false, cola:0,     status:'Projected',  notes:'Annuitized payout estimate; confirm with insurer' },
  // Investment income
  { source:'Brokerage Dividends (Joint)', owner:'Joint',type:'Dividends / Interest',amt:12250,  startAge:0,  endAge:'', inflAdj:true,  cola:0.04,  status:'Active',     notes:'Dividend portfolio; ~4% growth annually' },
  { source:'Bond Interest',               owner:'Joint',type:'Dividends / Interest',amt:6800,   startAge:0,  endAge:'', inflAdj:false, cola:0,     status:'Active',     notes:'Government bonds; fixed coupon' },
  { source:'CD & Savings Interest',       owner:'P1', type:'Dividends / Interest',  amt:2400,   startAge:0,  endAge:67, inflAdj:false, cola:0,     status:'Active',     notes:'Short-term CDs; rolls off near retirement' },
  // Other
  { source:'Textbook Royalties',          owner:'P2', type:'Other',                 amt:4200,   startAge:0,  endAge:70, inflAdj:false, cola:0,     status:'Active',     notes:"Patricia's nursing textbook; declining over time" },
];

function startYrFormula(r) {
  // Start Year = YEAR(DOB) + start_age. If startAge = 0 = current/ongoing, use current year.
  return `=IFERROR(IF(F${r}=0,YEAR(TODAY()),YEAR(IF(C${r}="P2",${SETUP}!E8,${SETUP}!B8))+F${r}),"")`;
}
function endYrFormula(r) {
  return `=IFERROR(IF(H${r}="","Lifetime",YEAR(IF(C${r}="P2",${SETUP}!E8,${SETUP}!B8))+H${r}),"")`;
}

(async () => {
  const vals = [];
  const fmt  = [];

  COL_WIDTHS.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  fmt.push({ repeatCell: { range: gridRange(SID,0,400,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ===== TITLE =====
  vals.push({ range: `${S}!A1`, values: [['RETIREMENT INCOME']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,13), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ===== SUBTITLE =====
  vals.push({ range: `${S}!A2`, values: [['All income streams for James & Patricia Whitmore. Start/End Year calculated from dates of birth in Setup tab.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,13), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrB), textFormat: { fontSize: 9, foregroundColor: hex(C.primaryText), italic: true, fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ===== SUMMARY CARDS (rows 3-4) =====
  const CARDS = [
    { label: 'Total Active Annual Income',  formula: `=IFERROR(SUMIF(L7:L200,"Active",E7:E200),"—")`,          fmt: '"$"#,##0' },
    { label: 'Combined Social Security',    formula: `=IFERROR(SUMIF(D7:D200,"Social Security",E7:E200),"—")`,  fmt: '"$"#,##0' },
    { label: 'Pension & Annuity Income',    formula: `=IFERROR(SUMIFS(E7:E200,D7:D200,"Pension")+SUMIFS(E7:E200,D7:D200,"Annuity"),"—")`, fmt: '"$"#,##0' },
    { label: 'Employment Income (Active)',  formula: `=IFERROR(SUMIFS(E7:E200,D7:D200,"Part-time Work",L7:L200,"Active"),"—")`, fmt: '"$"#,##0' },
    { label: 'Investment & Rental (Active)',formula: `=IFERROR(SUMIFS(E7:E200,L7:L200,"Active",D7:D200,"Rental Income")+SUMIFS(E7:E200,L7:L200,"Active",D7:D200,"Dividends / Interest"),"—")`, fmt: '"$"#,##0' },
  ];
  const cardSpans = [[0,2],[2,4],[4,6],[6,8],[8,11]];
  CARDS.forEach((card, ci) => {
    const [c1, c2] = cardSpans[ci];
    const col = String.fromCharCode(65 + c1);
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB), textFormat: { bold: true, fontSize: 8, foregroundColor: hex('#D0D4E8'), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}3`, values: [[card.label]] });
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c1,c2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c1,c2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: 'NUMBER', pattern: card.fmt },
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    vals.push({ range: `${S}!${col}4`, values: [[card.formula]] });
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ===== COLUMN HEADERS (row 6) =====
  vals.push({ range: `${S}!A6`, values: [[
    'Inc ID','Income Source','Owner','Type','Annual Amount\n(Today\'s $)',
    'Start\nAge','Start Year','End Age','End Year',
    'Inflation\nAdj?','COLA /\nGrowth %','Status','Notes',
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    borders: { bottom: { style: 'SOLID_MEDIUM', color: hex(C.primary) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' }});

  // ===== DATA ROWS (starting row 7, 0-idx 6) =====
  const DATA_START = 6;
  // Type → color map
  const typeColor = t =>
    t === 'Social Security' ? C.info :
    t === 'Pension' ? C.preTax :
    t === 'Part-time Work' ? C.roth :
    t === 'Rental Income' ? C.taxable :
    t === 'Annuity' ? C.taxDeferred :
    C.altRow;

  INCOME.forEach((inc, ii) => {
    const rIdx = DATA_START + ii;
    const rowNum = rIdx + 1;
    const isEven = ii % 2 === 0;
    const rowBg = inc.status === 'Active' ? (isEven ? C.panel : '#F9F8F5') : C.altRow;
    const ownerColor = inc.owner === 'P1' ? C.preTax : inc.owner === 'P2' ? C.roth : C.taxable;

    vals.push({ range: `${S}!A${rowNum}`, values: [[
      `=IF(B${rowNum}="","","INC-"&TEXT(ROW()-6,"000"))`,
      inc.source, inc.owner, inc.type, inc.amt,
      inc.startAge === 0 ? 'Now' : inc.startAge,
      startYrFormula(rowNum),
      inc.endAge === '' ? '' : inc.endAge,
      endYrFormula(rowNum),
      inc.inflAdj,
      inc.cola,
      inc.status, inc.notes,
    ]] });

    // Base row format
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,13), cell: { userEnteredFormat: {
      backgroundColor: hex(rowBg),
      textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});

    // Acct ID
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,0,1), cell: { userEnteredFormat: {
      textFormat: { fontSize: 8, foregroundColor: hex(C.secText) }, padding: { left: 4 },
    }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.padding' }});

    // Source (bold)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,1,2), cell: { userEnteredFormat: {
      textFormat: { bold: true, fontSize: 9 }, padding: { left: 4 },
    }}, fields: 'userEnteredFormat.textFormat,userEnteredFormat.padding' }});

    // Owner color
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,2,3), cell: { userEnteredFormat: {
      backgroundColor: hex(ownerColor), textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});

    // Type color
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,3,4), cell: { userEnteredFormat: {
      backgroundColor: hex(typeColor(inc.type)), textFormat: { fontSize: 9 }, padding: { left: 4 },
    }}, fields: 'userEnteredFormat' }});

    // Annual Amount (currency)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,4,5), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' },
      horizontalAlignment: 'RIGHT', textFormat: { bold: true, fontSize: 9 },
    }}, fields: 'userEnteredFormat' }});

    // Start/End Age, Start/End Year
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,5,9), cell: { userEnteredFormat: {
      horizontalAlignment: 'CENTER', textFormat: { fontSize: 9 },
    }}, fields: 'userEnteredFormat.horizontalAlignment,userEnteredFormat.textFormat' }});
    // Start Year and End Year are formula cells
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,6,7), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), textFormat: { bold: true, fontSize: 9 },
    }}, fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,8,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula), textFormat: { bold: true, fontSize: 9 },
    }}, fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat' }});

    // Inflation Adj checkbox
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,9,10), cell: { userEnteredFormat: {
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat.horizontalAlignment' }});
    fmt.push({ setDataValidation: { range: gridRange(SID,rIdx,rIdx+1,9,10), rule: {
      condition: { type: 'BOOLEAN' }, strict: true, showCustomUi: true,
    }}});

    // COLA Rate (%)
    fmt.push({ repeatCell: { range: gridRange(SID,rIdx,rIdx+1,10,11), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), numberFormat: { type: 'PERCENT', pattern: '0.0%' },
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat' }});

    // Status dropdown
    fmt.push({ setDataValidation: { range: gridRange(SID,rIdx,rIdx+1,11,12), rule: {
      condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$77:$A$79` }] },
      showCustomUi: true, strict: true,
    }}});

    // Border
    fmt.push({ updateBorders: { range: gridRange(SID,rIdx,rIdx+1,0,13), bottom: { style: 'SOLID', color: hex(C.border) }}});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: rIdx, endIndex: rIdx+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  });

  // ===== TOTALS ROW =====
  const TOTALS_ROW = DATA_START + INCOME.length;
  const totR = TOTALS_ROW + 1;
  vals.push({ range: `${S}!A${totR}`, values: [[
    'TOTALS','','','',
    `=IFERROR(SUM(E7:E${totR-1}),"")`,
    '','','','','','','','',
  ]] });
  fmt.push({ repeatCell: { range: gridRange(SID,TOTALS_ROW,TOTALS_ROW+1,0,13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
    borders: { top: { style: 'SOLID_MEDIUM', color: hex(C.primary) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,TOTALS_ROW,TOTALS_ROW+1,4,5), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0' }, horizontalAlignment: 'RIGHT',
  }}, fields: 'userEnteredFormat.numberFormat,userEnteredFormat.horizontalAlignment' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: TOTALS_ROW, endIndex: TOTALS_ROW+1 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Conditional format: SS rows highlighted
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, DATA_START, DATA_START + INCOME.length, 0, 13)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Projected' }] },
      format: { textFormat: { italic: true, foregroundColor: hex(C.secText) } },
    },
  }, index: 0 }});

  await valuesBatchUpdate(id, vals, '05-income values');
  await batchUpdate(id, fmt, '05-income format');

  console.log('✅ Retirement Income done — 12 income streams seeded.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
