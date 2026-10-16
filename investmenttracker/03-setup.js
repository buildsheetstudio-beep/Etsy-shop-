'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Portfolio Setup'];
const S = "'Portfolio Setup'";
const REF = "'Reference Data'";

// ── Row layout (0-indexed) ──────────────────────────────────────────
const R_TITLE  = 0;   // merged A1:N2
const R_SUB    = 1;
const R_IOHDR  = 3;   // "Investor / Household Information" header
const R_IO0    = 4;   // investor info fields start (A=label, B=value)
// 15 investor fields (rows 4-18)
const R_FHDR   = 20;  // Forecast Assumptions header
const R_F0     = 21;  // forecast fields (rows 21-33)
const R_ALHDR  = 35;  // Target Allocation header
const R_ALCOL  = 36;  // column headers
const R_AL0    = 37;  // allocation data rows 37-47 (11 rows)
const R_CTHDR  = 50;  // Dashboard Controls header
const R_CTCOL  = 51;
const R_CT0    = 52;  // 7 control rows
const R_CARDHDR= 61;  // Summary Cards header
const R_CARD   = 62;  // cards row
const R_DISC   = 68;  // disclaimer

(async () => {
  const vals = [];
  const fmt  = [];

  // ── Column widths ──────────────────────────────────────────────────
  [[0,200],[1,160],[2,30],[3,200],[4,160],[5,30],[6,200],[7,160],[8,30],[9,120],[10,120],[11,120],[12,120],[13,120]].forEach(([ci, px]) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // ── Background ────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ── Title banner ──────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID,R_TITLE,R_TITLE+2,0,14), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A1`, values: [['ULTIMATE INVESTMENT TRACKER\nPortfolio Setup & Planning Hub']] });
  fmt.push({ repeatCell: { range: gridRange(SID,R_TITLE,R_TITLE+2,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_TITLE, endIndex: R_TITLE+2 },
    properties: { pixelSize: 54 }, fields: 'pixelSize' }});

  // ── Freeze rows 1:3 ───────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 3 } }, fields: 'gridProperties.frozenRowCount' }});

  // ── Helper: section header ────────────────────────────────────────
  function sectionHdr(row, label, colSpan) {
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, colSpan), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${row+1}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, colSpan), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  }

  // ── Helper: input row ─────────────────────────────────────────────
  function inputRow(row, label, colOffset) {
    const lc = colOffset * 3;    // label column (0-indexed)
    const vc = colOffset * 3 + 1; // value column
    vals.push({ range: `${S}!${String.fromCharCode(65+lc)}${row+1}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, lc, lc+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, vc, vc+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input),
      textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT',
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  }

  // ── Section 1: Investor / Household Info (left panel, colOffset=0) ─
  sectionHdr(R_IOHDR, 'Investor / Household Information', 5);

  const ioFields = [
    ['Portfolio Name',             'Walsh Family Investment Portfolio'],
    ['Planning Type',              'Couple / Household'],
    ['Investor 1 Name',            'Daniel Walsh'],
    ['Investor 2 Name',            'Emily Walsh'],
    ['Primary Currency',           'USD'],
    ['Current Year',               2026],
    ['Investment Start Date',      '3/15/2019'],
    ['Current Household Income',   185000],
    ['Annual Investment Goal',     36000],
    ['Monthly Investment Goal',    3000],
    ['Target Portfolio Value',     2000000],
    ['Target Net Worth',           2500000],
    ['Target Annual Dividend Income', 25000],
    ['Investment Time Horizon (yrs)', 25],
    ['Notes',                      'Diversified long-term portfolio for retirement in 2049'],
  ];

  ioFields.forEach(([label, val], i) => {
    const r = R_IO0 + i;
    inputRow(r, label, 0);
    // push value into col B
    vals.push({ range: `${S}!B${r+1}`, values: [[val]] });
  });

  // ── Section 2: Forecast Assumptions (right panel, colOffset=3 → cols G-H) ─
  sectionHdr(R_IOHDR, 'Investor / Household Information', 5); // already set above
  sectionHdr(R_FHDR, 'Forecast Assumptions', 14);

  const fAssump = [
    ['Conservative Return %',       0.05],
    ['Expected Return %',           0.075],
    ['Optimistic Return %',         0.10],
    ['Annual Contribution Growth %',0.03],
    ['General Inflation %',         0.03],
    ['Annual Portfolio Fee %',      0.0015],
    ['Forecast Start Year',         2026],
    ['Forecast End Year',           2055],
    ['Include Dividends in Return?','TRUE'],  // checkbox
    ['Reinvest Dividends?',         'TRUE'],
    ['Include Inflation-Adjusted Values?','TRUE'],
    ['Include Fees?',               'TRUE'],
    ['Default Scenario',            'Expected'],
  ];

  fAssump.forEach(([label, val], i) => {
    const r = R_F0 + i;
    // left column (col A) = label, col B = value
    vals.push({ range: `${S}!A${r+1}`, values: [[label]] });
    vals.push({ range: `${S}!B${r+1}`, values: [[val]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 1, 2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE', borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  });

  // Checkbox validation for boolean forecast fields
  [[R_F0+8,'Include Dividends'],[R_F0+9,'Reinvest Dividends'],[R_F0+10,'Inflation-Adjusted'],[R_F0+11,'Include Fees']].forEach(([r]) => {
    fmt.push({ setDataValidation: { range: gridRange(SID, r, r+1, 1, 2),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } }});
  });

  // Planning Type dropdown
  fmt.push({ setDataValidation: { range: gridRange(SID, R_IO0+1, R_IO0+2, 1, 2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$199:$A$200` }] }, showCustomUi: true, strict: false } }});

  // Primary Currency dropdown
  fmt.push({ setDataValidation: { range: gridRange(SID, R_IO0+4, R_IO0+5, 1, 2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$164:$A$172` }] }, showCustomUi: true, strict: false } }});

  // Default Scenario dropdown
  fmt.push({ setDataValidation: { range: gridRange(SID, R_F0+12, R_F0+13, 1, 2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$179:$A$183` }] }, showCustomUi: true, strict: false } }});

  // Format percent/currency cells
  [[R_IO0+7,'$#,##0.00'],[R_IO0+8,'$#,##0.00'],[R_IO0+9,'$#,##0.00'],[R_IO0+10,'$#,##0.00'],
   [R_IO0+11,'$#,##0.00'],[R_IO0+12,'$#,##0.00']].forEach(([r, fmt2]) => {
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 1, 2), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: fmt2 }
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  [[R_F0,R_F0+3,'0.0%']].forEach(([r1, r2, patt]) => {
    fmt.push({ repeatCell: { range: gridRange(SID, r1, r2+1, 1, 2), cell: { userEnteredFormat: {
      numberFormat: { type: 'PERCENT', pattern: patt }
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  fmt.push({ repeatCell: { range: gridRange(SID, R_F0+4, R_F0+6, 1, 2), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.00%' }
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // ── Section 3: Target Allocation ──────────────────────────────────
  sectionHdr(R_ALHDR, 'Target Allocation  —  Per owner/household total must equal 100%. Flag if not.', 14);

  const alCols = ['Owner / Household','Asset Class','Target Alloc %','Min %','Max %','Notes'];
  vals.push({ range: `${S}!A${R_ALCOL+1}`, values: [alCols] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ALCOL, R_ALCOL+1, 0, alCols.length), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrC),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_ALCOL, endIndex: R_ALCOL+1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  const allocData = [
    // Joint Household 100%
    ['Joint Household','U.S. Stocks',     0.50, 0.40, 0.60, 'Core US equity exposure'],
    ['Joint Household','International Stocks', 0.15, 0.10, 0.25, ''],
    ['Joint Household','Emerging Markets', 0.05, 0.00, 0.10, ''],
    ['Joint Household','Bonds',           0.20, 0.10, 0.30, 'Interest rate diversifier'],
    ['Joint Household','REITs',           0.05, 0.00, 0.10, ''],
    ['Joint Household','Cash',            0.03, 0.02, 0.10, 'Settlement + liquidity'],
    ['Joint Household','Cryptocurrency',  0.02, 0.00, 0.05, 'Speculative, low weight'],
    // Daniel only — intentionally incomplete (90%) to show warning
    ['Daniel Walsh',   'U.S. Stocks',     0.60, 0.50, 0.70, ''],
    ['Daniel Walsh',   'Bonds',           0.15, 0.10, 0.25, ''],
    ['Daniel Walsh',   'Cash',            0.15, 0.05, 0.20, ''],
    // Emily — complete at 100%
    ['Emily Walsh',    'U.S. Stocks',     0.55, 0.45, 0.65, ''],
    ['Emily Walsh',    'International Stocks', 0.20, 0.10, 0.30, ''],
    ['Emily Walsh',    'Bonds',           0.20, 0.10, 0.30, ''],
    ['Emily Walsh',    'Cash',            0.05, 0.02, 0.10, ''],
  ];
  vals.push({ range: `${S}!A${R_AL0+1}`, values: allocData });

  // Format allocation % columns (C, D, E)
  fmt.push({ repeatCell: { range: gridRange(SID, R_AL0, R_AL0+allocData.length, 2, 5), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0%' },
    horizontalAlignment: 'CENTER',
    backgroundColor: hex(C.formula),
    textFormat: { fontSize: 9, fontFamily: 'Arial' },
  }}, fields: 'userEnteredFormat' }});

  // Style allocation rows
  for (let i = 0; i < allocData.length; i++) {
    const r = R_AL0 + i;
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 2), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 5, 6), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  }

  // Dropdown validation for allocation Owner and Asset Class
  fmt.push({ setDataValidation: { range: gridRange(SID, R_AL0, R_AL0+allocData.length, 0, 1),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$4:$A$6` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID, R_AL0, R_AL0+allocData.length, 1, 2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$33:$A$45` }] }, showCustomUi: true, strict: false } }});

  // ── Section 4: Dashboard Controls ─────────────────────────────────
  sectionHdr(R_CTHDR, 'Dashboard Controls', 14);

  const ctLabels = ['Selected Owner','Selected Account','Selected Year','Selected Month / Full Year','Selected Asset Class','Selected Scenario','Display Currency'];
  const ctVals   = ['Joint Household','All Accounts','2026','Full Year','All Classes','Expected','USD'];
  ctLabels.forEach((lbl, i) => {
    const r = R_CT0 + i;
    vals.push({ range: `${S}!A${r+1}`, values: [[lbl]] });
    vals.push({ range: `${S}!B${r+1}`, values: [[ctVals[i]]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 1, 2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
      borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r, endIndex: r+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  });

  // Control dropdowns
  fmt.push({ setDataValidation: { range: gridRange(SID, R_CT0, R_CT0+1, 1, 2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$4:$A$6` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID, R_CT0+2, R_CT0+3, 1, 2),
    rule: { condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: '2024' },{ userEnteredValue: '2025' },{ userEnteredValue: '2026' },{ userEnteredValue: '2027' }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID, R_CT0+5, R_CT0+6, 1, 2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$179:$A$183` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ setDataValidation: { range: gridRange(SID, R_CT0+6, R_CT0+7, 1, 2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$164:$A$172` }] }, showCustomUi: true, strict: false } }});

  // ── Section 5: Summary Cards ───────────────────────────────────────
  sectionHdr(R_CARDHDR, 'Summary Overview', 14);

  const cards = [
    ['Total Accounts',   `=COUNTA('Account Tracker'!$A$6:$A$305)`],
    ['Portfolio Value',  `=IFERROR(SUMPRODUCT(('Account Tracker'!$O$6:$O$305=TRUE)*('Account Tracker'!$K$6:$K$305)),0)`],
    ['Annual Goal',      `=B${R_IO0+9}`],
    ['YTD Contributions',`=IFERROR(SUMPRODUCT((YEAR('Contributions & Savings'!$B$6:$B$1505)=YEAR(TODAY()))*('Contributions & Savings'!$I$6:$I$1505)),0)`],
    ['Target Portfolio', `=B${R_IO0+11}`],
    ['Target Net Worth', `=B${R_IO0+12}`],
    ['Target Dividend',  `=B${R_IO0+13}`],
    ['Forecast Horizon', `=B${R_F0+7}-B${R_F0+6}`],
  ];

  const cardCols = [0,2,4,6,8,10,12]; // wider spread
  // Use rows R_CARD (label) and R_CARD+1 (value)
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARD, endIndex: R_CARD+4 },
    properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // 2 rows of 4 cards each
  for (let i = 0; i < 4; i++) {
    const r = R_CARD;
    const c = i * 3;
    fmt.push({ mergeCells: { range: gridRange(SID, r, r+1, c, c+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID, r+1, r+2, c, c+2), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}${r+1}`, values: [[cards[i][0]]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}${r+2}`, values: [[cards[i][1]]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, c, c+2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrA), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, r+1, r+2, c, c+2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.highlight), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  for (let i = 4; i < 8; i++) {
    const r = R_CARD + 2;
    const c = (i - 4) * 3;
    const cEnd = (i - 4 === 3) ? c + 2 : c + 2; // last card gets remaining cols
    fmt.push({ mergeCells: { range: gridRange(SID, r, r+1, c, c+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID, r+1, r+2, c, c+2), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}${r+1}`, values: [[cards[i][0]]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}${r+2}`, values: [[cards[i][1]]] });
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, c, c+2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, r+1, r+2, c, c+2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.highlight), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }

  // Format currency cards
  const moneyCardRows = [R_CARD+1, R_CARD+3]; // value rows
  for (const r of moneyCardRows) {
    for (let c = 3; c < 12; c += 3) { // skip Totals (col 0) which are counts
      fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, c, c+2), cell: { userEnteredFormat: {
        numberFormat: { type: 'CURRENCY', pattern: '$#,##0' }
      }}, fields: 'userEnteredFormat.numberFormat' }});
    }
  }

  // ── Section 6: Disclaimer ──────────────────────────────────────────
  sectionHdr(R_DISC, 'Important Disclaimer', 14);

  const discText = 'This workbook is an educational investment-tracking and planning tool only. It does not provide investment, tax, legal, accounting, insurance, or financial advice and does not recommend any security, account, allocation, contribution amount, or investment strategy. Market prices, returns, dividends, fees, taxes, inflation, and investment outcomes are uncertain. Replace all sample assumptions with your own information and verify all records and decisions with qualified professionals and official account statements.';
  fmt.push({ mergeCells: { range: gridRange(SID, R_DISC+1, R_DISC+5, 0, 14), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A${R_DISC+2}`, values: [[discText]] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_DISC+1, R_DISC+5, 0, 14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial', italic: true },
    verticalAlignment: 'MIDDLE', horizontalAlignment: 'LEFT', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_DISC+1, endIndex: R_DISC+5 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  await valuesBatchUpdate(id, vals, '03-setup values');
  await batchUpdate(id, fmt, '03-setup format');
  console.log('✅ Portfolio Setup done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
