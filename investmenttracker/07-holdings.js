'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Holdings'];
const S = "'Holdings'";
const REF = "'Reference Data'";
const TXN = "'Transaction Log'";
const DIV = "'Dividend Income'";
const PRC = "'Price Updates'";
const ACT = "'Account Tracker'";

const HEADERS = [
  'Holding ID','Owner','Account ID','Account Name','Security ID','Ticker',
  'Security Name','Security Type','Asset Class','Quantity Held','Avg Cost / Unit',
  'Total Cost Basis','Current Price','Exchange Rate','Current Value',
  'Unrealized G/L','Unrealized G/L %','Dividends Received','Total Return Est.',
  'Portfolio Alloc %','Last Price Update','Notes'
];
const COL_W = [110,100,80,160,90,70,175,110,130,90,100,110,100,90,110,110,90,110,110,90,105,160];

const HDR  = 4;
const DATA = 5;

// [AccountID, SecurityID] — will be expanded in the script
const HOLDINGS = [
  // ACT-001 Walsh Family Taxable (Daniel Walsh)
  ['ACT-001','SEC-0001'],  // VTI
  ['ACT-001','SEC-0006'],  // SCHD
  ['ACT-001','SEC-0011'],  // GLD
  ['ACT-001','SEC-0012'],  // BITO — unrealized loss
  ['ACT-001','SEC-0018'],  // AAPL
  ['ACT-001','SEC-0019'],  // MSFT
  ['ACT-001','SEC-0020'],  // NVDA — unrealized gain
  ['ACT-001','SEC-0022'],  // V
  ['ACT-001','SEC-0026'],  // STXBND
  ['ACT-001','SEC-0027'],  // CBND
  ['ACT-001','SEC-0028'],  // CD5
  // ACT-002 Daniel's Roth IRA
  ['ACT-002','SEC-0001'],  // VTI
  ['ACT-002','SEC-0005'],  // QQQ
  ['ACT-002','SEC-0020'],  // NVDA
  // ACT-003 Daniel's 401(k)
  ['ACT-003','SEC-0013'],  // VTSAX
  // ACT-004 Emily's Taxable Brokerage
  ['ACT-004','SEC-0001'],  // VTI
  ['ACT-004','SEC-0002'],  // VXUS
  ['ACT-004','SEC-0007'],  // AGG
  ['ACT-004','SEC-0008'],  // EMB
  ['ACT-004','SEC-0009'],  // VWO
  ['ACT-004','SEC-0010'],  // IEFA
  ['ACT-004','SEC-0021'],  // JNJ
  ['ACT-004','SEC-0022'],  // V
  // ACT-005 Emily's Roth IRA
  ['ACT-005','SEC-0001'],  // VTI
  ['ACT-005','SEC-0002'],  // VXUS
  ['ACT-005','SEC-0006'],  // SCHD
  ['ACT-005','SEC-0009'],  // VWO
  ['ACT-005','SEC-0010'],  // IEFA
  // ACT-006 Emily's 403(b)
  ['ACT-006','SEC-0014'],  // VBTLX
  ['ACT-006','SEC-0016'],  // SWTSX
  ['ACT-006','SEC-0017'],  // VGHCX
  // ACT-007 Walsh Joint Brokerage
  ['ACT-007','SEC-0001'],  // VTI
  ['ACT-007','SEC-0002'],  // VXUS
  ['ACT-007','SEC-0003'],  // BND
  ['ACT-007','SEC-0004'],  // VNQ
  ['ACT-007','SEC-0009'],  // VWO
  ['ACT-007','SEC-0023'],  // SPG
  ['ACT-007','SEC-0024'],  // PLD
  ['ACT-007','SEC-0025'],  // AMT
  // ACT-008 Daniel's HSA
  ['ACT-008','SEC-0001'],  // VTI
  // ACT-009 Emily's HSA
  ['ACT-009','SEC-0001'],  // VTI
  ['ACT-009','SEC-0003'],  // BND
  // ACT-010 College Savings 529
  ['ACT-010','SEC-0013'],  // VTSAX
  ['ACT-010','SEC-0014'],  // VBTLX
  // ACT-011 Walsh Family Trust
  ['ACT-011','SEC-0001'],  // VTI
  ['ACT-011','SEC-0011'],  // GLD
  ['ACT-011','SEC-0018'],  // AAPL
  ['ACT-011','SEC-0019'],  // MSFT
  ['ACT-011','SEC-0025'],  // AMT
  // ACT-012 Daniel's Crypto
  ['ACT-012','SEC-0029'],  // BTC-USD
  ['ACT-012','SEC-0030'],  // ETH-USD
];

// Quantity Held formula
function qtyFormula(r) {
  const acctCol = 'C'; // col C = Account ID
  const secCol  = 'E'; // col E = Security ID
  const tL = TXN;
  return `=IFERROR(SUMPRODUCT((${tL}!$F$6:$F$5005=${acctCol}${r})*(${tL}!$H$6:$H$5005=${secCol}${r})*((${tL}!$L$6:$L$5005="Buy")+(${tL}!$L$6:$L$5005="Reinvestment")+(${tL}!$L$6:$L$5005="Transfer In"))*${tL}!$M$6:$M$5005)-SUMPRODUCT((${tL}!$F$6:$F$5005=${acctCol}${r})*(${tL}!$H$6:$H$5005=${secCol}${r})*((${tL}!$L$6:$L$5005="Sell")+(${tL}!$L$6:$L$5005="Transfer Out"))*${tL}!$M$6:$M$5005),0)`;
}

// Average Cost per Unit (buy+reinvestment+transferin weighted)
function avgCostFormula(r) {
  const tL = TXN;
  return `=IFERROR(SUMPRODUCT((${tL}!$F$6:$F$5005=C${r})*(${tL}!$H$6:$H$5005=E${r})*((${tL}!$L$6:$L$5005="Buy")+(${tL}!$L$6:$L$5005="Reinvestment")+(${tL}!$L$6:$L$5005="Transfer In"))*${tL}!$O$6:$O$5005)/SUMPRODUCT((${tL}!$F$6:$F$5005=C${r})*(${tL}!$H$6:$H$5005=E${r})*((${tL}!$L$6:$L$5005="Buy")+(${tL}!$L$6:$L$5005="Reinvestment")+(${tL}!$L$6:$L$5005="Transfer In"))*${tL}!$M$6:$M$5005),"")`;
}

// Dividends Received
function divFormula(r) {
  return `=IFERROR(SUMPRODUCT((${DIV}!$F$6:$F$2005=C${r})*(${DIV}!$G$6:$G$2005=E${r})*(${DIV}!$S$6:$S$2005="Received")*(${DIV}!$O$6:$O$2005)),0)`;
}

(async () => {
  const vals = [];
  const fmt  = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,1100,0,22), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Column widths
  COL_W.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Title banner
  fmt.push({ mergeCells: { range: gridRange(SID,0,2,0,22), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A1`, values: [['HOLDINGS\nCurrent quantity • cost basis • market value • gain/loss • allocation']] });
  fmt.push({ repeatCell: { range: gridRange(SID,0,2,0,22), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' }});

  // Disclaimer note on average cost
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,0,22), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A3`, values: [['⚠  Average cost per unit is a weighted-average organizational estimate and may differ from brokerage or tax-lot accounting. Verify with your brokerage for tax purposes.']] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,22), cell: { userEnteredFormat: {
    backgroundColor: hex(C.warning),
    textFormat: { fontSize: 8, foregroundColor: hex(C.text), fontFamily: 'Arial', italic: true },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Summary cards (row 4)
  const cards = [
    ['Total Holdings',    `=COUNTA($A$6:$A$1005)`],
    ['Total Portfolio Value', `=IFERROR(SUM($O$6:$O$1005),0)`],
    ['Total Cost Basis',  `=IFERROR(SUM($L$6:$L$1005),0)`],
    ['Unrealized G/L',    `=IFERROR(SUM($P$6:$P$1005),0)`],
  ];
  cards.forEach((card, i) => {
    const c = i * 5;
    if (c >= 20) return;
    fmt.push({ mergeCells: { range: gridRange(SID,3,4,c,c+4), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}4`, values: [[card[0]]] });
    fmt.push({ mergeCells: { range: gridRange(SID,4,5,c,c+4), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}5`, values: [[card[1]]] });
    fmt.push({ repeatCell: { range: gridRange(SID,3,4,c,c+4), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,4,5,c,c+4), cell: { userEnteredFormat: {
      backgroundColor: hex(C.highlight), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 5 },
      properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  });

  // Column headers
  vals.push({ range: `${S}!A${HDR+1}`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID,HDR,HDR+1,0,22), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: HDR, endIndex: HDR+1 },
    properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Data rows
  for (let i = 0; i < HOLDINGS.length; i++) {
    const r = DATA + i;
    const [actId, secId] = HOLDINGS[i];
    const rn = r + 1; // 1-indexed row number
    const bg = i % 2 === 0 ? C.panel : C.altRow;

    // A: Holding ID
    vals.push({ range: `${S}!A${rn}`, values: [[`=IF(OR(C${rn}="",E${rn}=""),"",C${rn}&"-"&E${rn})`]] });
    // B: Owner (from Account Tracker)
    vals.push({ range: `${S}!B${rn}`, values: [[`=IFERROR(VLOOKUP(C${rn},${ACT}!$A$6:$C$305,3,FALSE),"")`]] });
    // C: Account ID
    vals.push({ range: `${S}!C${rn}`, values: [[actId]] });
    // D: Account Name
    vals.push({ range: `${S}!D${rn}`, values: [[`=IFERROR(VLOOKUP(C${rn},${ACT}!$A$6:$B$305,2,FALSE),"")`]] });
    // E: Security ID
    vals.push({ range: `${S}!E${rn}`, values: [[secId]] });
    // F: Ticker
    vals.push({ range: `${S}!F${rn}`, values: [[`=IFERROR(VLOOKUP(E${rn},${PRC}!$A$6:$B$1005,2,FALSE),"")`]] });
    // G: Security Name
    vals.push({ range: `${S}!G${rn}`, values: [[`=IFERROR(VLOOKUP(E${rn},${PRC}!$A$6:$C$1005,3,FALSE),"")`]] });
    // H: Security Type
    vals.push({ range: `${S}!H${rn}`, values: [[`=IFERROR(VLOOKUP(E${rn},${PRC}!$A$6:$D$1005,4,FALSE),"")`]] });
    // I: Asset Class
    vals.push({ range: `${S}!I${rn}`, values: [[`=IFERROR(VLOOKUP(E${rn},${PRC}!$A$6:$E$1005,5,FALSE),"")`]] });
    // J: Quantity Held
    vals.push({ range: `${S}!J${rn}`, values: [[qtyFormula(rn)]] });
    // K: Avg Cost per Unit
    vals.push({ range: `${S}!K${rn}`, values: [[avgCostFormula(rn)]] });
    // L: Total Cost Basis
    vals.push({ range: `${S}!L${rn}`, values: [[`=IFERROR(J${rn}*K${rn},0)`]] });
    // M: Current Price
    vals.push({ range: `${S}!M${rn}`, values: [[`=IFERROR(VLOOKUP(E${rn},${PRC}!$A$6:$G$1005,7,FALSE),"")`]] });
    // N: Exchange Rate
    vals.push({ range: `${S}!N${rn}`, values: [[`=IFERROR(VLOOKUP(E${rn},${PRC}!$A$6:$N$1005,14,FALSE),1)`]] });
    // O: Current Value
    vals.push({ range: `${S}!O${rn}`, values: [[`=IFERROR(J${rn}*M${rn}*N${rn},0)`]] });
    // P: Unrealized G/L
    vals.push({ range: `${S}!P${rn}`, values: [[`=IFERROR(O${rn}-L${rn},0)`]] });
    // Q: Unrealized G/L %
    vals.push({ range: `${S}!Q${rn}`, values: [[`=IFERROR(P${rn}/L${rn},0)`]] });
    // R: Dividends Received
    vals.push({ range: `${S}!R${rn}`, values: [[divFormula(rn)]] });
    // S: Total Return Estimate
    vals.push({ range: `${S}!S${rn}`, values: [[`=IFERROR(P${rn}+R${rn},0)`]] });
    // T: Portfolio Allocation %
    vals.push({ range: `${S}!T${rn}`, values: [[`=IFERROR(O${rn}/SUM($O$6:$O$1005),0)`]] });
    // U: Last Price Update
    vals.push({ range: `${S}!U${rn}`, values: [[`=IFERROR(VLOOKUP(E${rn},${PRC}!$A$6:$K$1005,11,FALSE),"")`]] });
    // V: Notes
    vals.push({ range: `${S}!V${rn}`, values: [['']] });

    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,22), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  }

  // Number formats
  const dn = DATA; const de = DATA + HOLDINGS.length;
  // Quantity: 0.0000
  fmt.push({ repeatCell: { range: gridRange(SID,dn,de,9,10), cell: { userEnteredFormat: {
    numberFormat: { type: 'NUMBER', pattern: '0.0000' }
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Price/cost: $0.0000
  [10,12].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,dn,de,c,c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$0.0000' }
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Currency values: $#,##0.00
  [11,14,15,17,18].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,dn,de,c,c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' }
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Percentages: 0.0%
  [16,19].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,dn,de,c,c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'PERCENT', pattern: '0.0%' }
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Date col U (20)
  fmt.push({ repeatCell: { range: gridRange(SID,dn,de,20,21), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'MMM D, YYYY' }
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Exchange rate: 0.0000
  fmt.push({ repeatCell: { range: gridRange(SID,dn,de,13,14), cell: { userEnteredFormat: {
    numberFormat: { type: 'NUMBER', pattern: '0.0000' }
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // Formula tint for all computed columns
  [0,1,3,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,dn,de,c,c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });
  // Input tint for C (Account ID) and E (Security ID)
  [2,4].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID,dn,de,c,c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });

  // Conditional formatting
  // Positive unrealized G/L → sage text
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,dn,de,15,17)],
    booleanRule: { condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.01' }] },
      format: { textFormat: { foregroundColor: hex(C.success), bold: true } }
    }
  }, index: 0 }});
  // Negative G/L → muted coral
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,dn,de,15,17)],
    booleanRule: { condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '-0.01' }] },
      format: { textFormat: { foregroundColor: hex(C.attention), bold: true } }
    }
  }, index: 1 }});
  // Missing current price → amber
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,dn,de,12,13)],
    booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: '' }] },
      format: { backgroundColor: hex(C.warning) }
    }
  }, index: 2 }});
  // Negative quantity (data issue) → attention
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID,dn,de,9,10)],
    booleanRule: { condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
      format: { backgroundColor: hex(C.attention) }
    }
  }, index: 3 }});

  // Freeze rows 1:5 and cols A:D
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '07-holdings values');
  await batchUpdate(id, fmt, '07-holdings format');
  console.log(`✅ Holdings done. ${HOLDINGS.length} holdings defined.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
