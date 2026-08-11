'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID  = sheetMap['Portfolio Allocation'];
const S    = "'Portfolio Allocation'";
const SETUP = "'Portfolio Setup'";
const HOLD  = "'Holdings'";
const ACT   = "'Account Tracker'";
const REF   = "'Reference Data'";

// ── Row layout (0-indexed) ────────────────────────────────────────────
const R_TITLE   = 0;   // rows 0-1: 2-row merged title
const R_NOTE    = 2;
const R_CARDHDR = 3;   // card label row
const R_CARD    = 4;   // card value row
const R_CTRL    = 5;   // owner filter dropdown
const R_BLANK1  = 6;
const R_ACHDR   = 7;   // Asset Class section header
const R_ACOLS   = 8;   // Asset Class column headers
const R_AD0     = 9;   // Asset Class data: 0-indexed 9-21 (13 rows)
const R_ATOT    = 22;  // Asset Class total
const R_BLANK2  = 23;
const R_BAHDR   = 24;  // By Account header
const R_BACOLS  = 25;
const R_BAD0    = 26;  // Account data: 0-indexed 26-38 (13 rows)
const R_BATOT   = 39;  // Account total
const R_BLANK3  = 40;
const R_BIHDR   = 41;  // By Investor header
const R_BICOLS  = 42;
const R_BID0    = 43;  // Investor data: 0-indexed 43-45 (3 rows)
const R_BITOT   = 46;  // Investor total
const R_BLANK4  = 47;

// 1-indexed row numbers for formula strings
const AC1   = R_AD0  + 1;                   // 10  first asset class data row
const ACN   = R_AD0  + 13;                  // 22  last  asset class data row
const ACTOT = R_ATOT + 1;                   // 23  asset class total row
const BA1   = R_BAD0 + 1;                   // 27  first account data row
const BAN   = R_BAD0 + 13;                  // 39  last  account data row
const BATOT = R_BATOT + 1;                  // 40  account total row
const BI1   = R_BID0 + 1;                   // 44  first investor data row
const BIN   = R_BID0 + 3;                   // 46  last  investor data row
const BITOT = R_BITOT + 1;                  // 47  investor total row
const CTRL  = '$B$6';                        // filter dropdown cell (1-indexed B6)

const ASSET_CLASSES = [
  ['U.S. Stocks',           '#496A88'],
  ['International Stocks',  '#7A6680'],
  ['Emerging Markets',      '#C98474'],
  ['Bonds',                 '#8FA98C'],
  ['REITs',                 '#C6A15B'],
  ['Real Estate',           '#C6A15B'],
  ['Cash',                  '#8FB9B7'],
  ['Money Market',          '#8FB9B7'],
  ['Gold / Precious Metals','#D0B36A'],
  ['Commodities',           '#D0B36A'],
  ['Cryptocurrency',        '#74758F'],
  ['Alternatives',          '#A59482'],
  ['Other',                 '#A9A9A6'],
];

const ACCOUNT_IDS = [
  'ACT-001','ACT-002','ACT-003','ACT-004','ACT-005','ACT-006',
  'ACT-007','ACT-008','ACT-009','ACT-010','ACT-011','ACT-012','ACT-013',
];

const INVESTORS = ['Daniel Walsh', 'Emily Walsh', 'Joint Household'];

function colLetter(c) { return String.fromCharCode(65 + c); }

(async () => {
  const vals = [];
  const fmt  = [];

  // ── Background ────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID,0,500,0,22), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ── Column widths ─────────────────────────────────────────────────
  [
    [0,155],[1,130],[2,82],[3,82],[4,82],[5,110],[6,72],[7,115],[8,75],[9,75],[10,120],
    [11,16],
    [12,120],[13,120],[14,120],[15,120],[16,120],[17,120],[18,120],[19,120],[20,120],
  ].forEach(([ci,px]) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // ── Title banner (rows 0-1, merged A1:K2) ────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_TITLE, R_TITLE+2, 0, 11), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A1`, values: [['PORTFOLIO ALLOCATION\nCurrent vs Target  ·  By Asset Class  ·  By Account  ·  By Investor']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_TITLE, R_TITLE+2, 0, 11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_TITLE, endIndex: R_TITLE+2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' }});
  // Freeze title rows only (do NOT set frozenColumnCount — merged title spans all cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' }});

  // ── Note row ─────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_NOTE, R_NOTE+1, 0, 11), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A3`, values: [['All values pull live from Holdings & Account Tracker. Use "View Targets For" (row 6) to filter target allocations by investor or household.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_NOTE, R_NOTE+1, 0, 11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.info),
    textFormat: { fontSize: 8, italic: true, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_NOTE, endIndex: R_NOTE+1 },
    properties: { pixelSize: 20 }, fields: 'pixelSize' }});

  // ── Summary cards (rows 3-4) ──────────────────────────────────────
  const CARDS = [
    { label: 'Total Portfolio Value', val: `=IFERROR(SUM(${HOLD}!$O$6:$O$1005),0)`, type: 'CURRENCY', pat: '$#,##0.00' },
    { label: 'Active Asset Classes',  val: `=SUMPRODUCT(($B$${AC1}:$B$${ACN}>0)*1)`,  type: 'NUMBER',   pat: '#,##0' },
    { label: 'Total Unrealized G/L',  val: `=IFERROR(SUM(${HOLD}!$P$6:$P$1005),0)`,  type: 'CURRENCY', pat: '$#,##0.00' },
    { label: 'Overall Return %',      val: `=IFERROR(SUM(${HOLD}!$P$6:$P$1005)/SUM(${HOLD}!$L$6:$L$1005),0)`, type: 'PERCENT', pat: '0.0%' },
  ];
  const cardCols = [[0,3],[3,6],[6,9],[9,11]];
  CARDS.forEach((card, i) => {
    const [c0, c1] = cardCols[i];
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARDHDR, R_CARDHDR+1, c0, c1), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARD, R_CARD+1, c0, c1), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${colLetter(c0)}${R_CARDHDR+1}`, values: [[card.label]] });
    vals.push({ range: `${S}!${colLetter(c0)}${R_CARD+1}`, values: [[card.val]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARDHDR, R_CARDHDR+1, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARD, R_CARD+1, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.highlight),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: card.type, pattern: card.pat },
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARDHDR, endIndex: R_CARD+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ── Owner filter control (row 5) ──────────────────────────────────
  vals.push({ range: `${S}!A6`, values: [['View Targets For:']] });
  vals.push({ range: `${S}!B6`, values: [['Joint Household']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_CTRL, R_CTRL+1, 0, 1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.panel),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID, R_CTRL, R_CTRL+1, 1, 2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
    borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ setDataValidation: { range: gridRange(SID, R_CTRL, R_CTRL+1, 1, 2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$4:$A$6` }] }, showCustomUi: true, strict: false } }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CTRL, endIndex: R_CTRL+1 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Blank row 6
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_BLANK1, endIndex: R_BLANK1+1 },
    properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ── Helpers ───────────────────────────────────────────────────────
  function sectionHdr(row, label) {
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, 11), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${row+1}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, 11), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  }

  function colHdr(row, headers, nCols) {
    vals.push({ range: `${S}!A${row+1}`, values: [headers] });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, nCols), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrC),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 30 }, fields: 'pixelSize' }});
  }

  function blankRow(row) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 10 }, fields: 'pixelSize' }});
  }

  function totalRowStyle(row, label, nCols) {
    vals.push({ range: `${S}!A${row+1}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, nCols), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  }

  // ── SECTION 1: By Asset Class ─────────────────────────────────────
  sectionHdr(R_ACHDR, '  Portfolio Allocation by Asset Class');
  colHdr(R_ACOLS, [
    'Asset Class','Current Value','Portfolio %','Target %','Difference','Status',
    '# Holdings','Unrealized G/L','Min Target %','Max Target %','Notes'
  ], 11);

  // Target lookup template (avoids "" vs 0 ambiguity)
  const tgtLookup = (row, col) =>
    `IFERROR(IF(SUMPRODUCT((${SETUP}!$A$38:$A$51=${CTRL})*(${SETUP}!$B$38:$B$51=A${row}))=0,"",SUMPRODUCT((${SETUP}!$A$38:$A$51=${CTRL})*(${SETUP}!$B$38:$B$51=A${row})*${SETUP}!$${col}$38:$${col}$51)),"")`;

  for (let i = 0; i < ASSET_CLASSES.length; i++) {
    const r0 = R_AD0 + i;
    const r  = r0 + 1;            // 1-indexed row for formulas
    const [acName] = ASSET_CLASSES[i];
    const bg = i % 2 === 0 ? C.panel : C.altRow;

    vals.push({ range: `${S}!A${r}`, values: [[acName]] });
    vals.push({ range: `${S}!B${r}`, values: [[`=IFERROR(SUMPRODUCT((${HOLD}!$I$6:$I$1005=A${r})*(${HOLD}!$O$6:$O$1005)),0)`]] });
    vals.push({ range: `${S}!C${r}`, values: [[`=IFERROR(B${r}/SUM($B$${AC1}:$B$${ACN}),0)`]] });
    vals.push({ range: `${S}!D${r}`, values: [[`=${tgtLookup(r,'C')}`]] });
    vals.push({ range: `${S}!E${r}`, values: [[`=IFERROR(IF(D${r}="","",C${r}-D${r}),"")`]] });
    vals.push({ range: `${S}!F${r}`, values: [[`=IF(D${r}="","—",IF(C${r}-D${r}>0.03,"Overweight",IF(C${r}-D${r}<-0.03,"Underweight","On Target")))`]] });
    vals.push({ range: `${S}!G${r}`, values: [[`=IFERROR(SUMPRODUCT((${HOLD}!$I$6:$I$1005=A${r})*(${HOLD}!$J$6:$J$1005>0)*1),0)`]] });
    vals.push({ range: `${S}!H${r}`, values: [[`=IFERROR(SUMPRODUCT((${HOLD}!$I$6:$I$1005=A${r})*(${HOLD}!$P$6:$P$1005)),0)`]] });
    vals.push({ range: `${S}!I${r}`, values: [[`=${tgtLookup(r,'D')}`]] });
    vals.push({ range: `${S}!J${r}`, values: [[`=${tgtLookup(r,'E')}`]] });

    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 11), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { fontSize: 9, fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r0, endIndex: r0+1 },
      properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  }

  // Asset class number formats
  fmt.push({ repeatCell: { range: gridRange(SID, R_AD0, R_ATOT, 1, 2), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  [2,3,4,8,9].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_AD0, R_ATOT, c, c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'PERCENT', pattern: '0.0%' },
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  fmt.push({ repeatCell: { range: gridRange(SID, R_AD0, R_ATOT, 7, 8), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Formula tint on computed cols (everything except A = name)
  [1,2,3,4,5,6,7,8,9].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_AD0, R_ATOT, c, c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });
  // Input tint on asset class name col
  fmt.push({ repeatCell: { range: gridRange(SID, R_AD0, R_ATOT, 0, 1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Conditional formats: Status (col F = index 5)
  [
    ['TEXT_EQ', 'Overweight', C.attention],
    ['TEXT_EQ', 'Underweight', C.info],
    ['TEXT_EQ', 'On Target', C.success],
  ].forEach(([type, val, color]) => {
    fmt.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, R_AD0, R_ATOT, 5, 6)],
      booleanRule: {
        condition: { type, values: [{ userEnteredValue: val }] },
        format: { backgroundColor: hex(color) },
      },
    }, index: 0 }});
  });
  // Conditional formats: Difference (col E = index 4)
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_AD0, R_ATOT, 4, 5)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.03' }] },
      format: { textFormat: { foregroundColor: hex(C.attention), bold: true } },
    },
  }, index: 0 }});
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_AD0, R_ATOT, 4, 5)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '-0.03' }] },
      format: { textFormat: { foregroundColor: hex(C.info), bold: true } },
    },
  }, index: 0 }});

  // Asset class total row
  totalRowStyle(R_ATOT, 'TOTAL', 11);
  vals.push({ range: `${S}!B${ACTOT}`, values: [[`=SUM($B$${AC1}:$B$${ACN})`]] });
  vals.push({ range: `${S}!C${ACTOT}`, values: [[`=IFERROR(SUM($C$${AC1}:$C$${ACN}),0)`]] });
  vals.push({ range: `${S}!D${ACTOT}`, values: [[`=IFERROR(SUM($D$${AC1}:$D$${ACN}),"")`]] });
  vals.push({ range: `${S}!E${ACTOT}`, values: [[`=IFERROR(SUM($E$${AC1}:$E$${ACN}),"")`]] });
  vals.push({ range: `${S}!G${ACTOT}`, values: [[`=SUM($G$${AC1}:$G$${ACN})`]] });
  vals.push({ range: `${S}!H${ACTOT}`, values: [[`=SUM($H$${AC1}:$H$${ACN})`]] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ATOT, R_ATOT+1, 1, 2), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  [2,3,4].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_ATOT, R_ATOT+1, c, c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'PERCENT', pattern: '0.0%' },
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ATOT, R_ATOT+1, 7, 8), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  blankRow(R_BLANK2);

  // ── SECTION 2: By Account ─────────────────────────────────────────
  sectionHdr(R_BAHDR, '  Allocation by Account');
  colHdr(R_BACOLS, [
    'Account ID','Account Name','Owner / Household','Current Value',
    'Portfolio %','# Holdings','','','','','',
  ], 6);

  for (let i = 0; i < ACCOUNT_IDS.length; i++) {
    const r0 = R_BAD0 + i;
    const r  = r0 + 1;
    const acctId = ACCOUNT_IDS[i];
    const bg = i % 2 === 0 ? C.panel : C.altRow;

    vals.push({ range: `${S}!A${r}`, values: [[acctId]] });
    vals.push({ range: `${S}!B${r}`, values: [[`=IFERROR(VLOOKUP(A${r},${ACT}!$A$6:$B$305,2,FALSE),"")`]] });
    vals.push({ range: `${S}!C${r}`, values: [[`=IFERROR(VLOOKUP(A${r},${ACT}!$A$6:$C$305,3,FALSE),"")`]] });
    vals.push({ range: `${S}!D${r}`, values: [[`=IFERROR(SUMPRODUCT((${HOLD}!$C$6:$C$1005=A${r})*(${HOLD}!$O$6:$O$1005)),0)`]] });
    vals.push({ range: `${S}!E${r}`, values: [[`=IFERROR(D${r}/SUM($D$${BA1}:$D$${BAN}),0)`]] });
    vals.push({ range: `${S}!F${r}`, values: [[`=IFERROR(SUMPRODUCT((${HOLD}!$C$6:$C$1005=A${r})*(${HOLD}!$J$6:$J$1005>0)*1),0)`]] });

    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 6), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { fontSize: 9, fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r0, endIndex: r0+1 },
      properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  }

  // Account number formats
  fmt.push({ repeatCell: { range: gridRange(SID, R_BAD0, R_BATOT, 3, 4), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID, R_BAD0, R_BATOT, 4, 5), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  // Formula tint on computed cols
  [1,2,3,4,5].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_BAD0, R_BATOT, c, c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });
  fmt.push({ repeatCell: { range: gridRange(SID, R_BAD0, R_BATOT, 0, 1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Account total row
  totalRowStyle(R_BATOT, 'TOTAL', 6);
  vals.push({ range: `${S}!D${BATOT}`, values: [[`=SUM($D$${BA1}:$D$${BAN})`]] });
  vals.push({ range: `${S}!E${BATOT}`, values: [[`=IFERROR(SUM($E$${BA1}:$E$${BAN}),0)`]] });
  vals.push({ range: `${S}!F${BATOT}`, values: [[`=SUM($F$${BA1}:$F$${BAN})`]] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_BATOT, R_BATOT+1, 3, 4), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID, R_BATOT, R_BATOT+1, 4, 5), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  blankRow(R_BLANK3);

  // ── SECTION 3: By Investor ────────────────────────────────────────
  sectionHdr(R_BIHDR, '  Allocation by Investor / Household');
  colHdr(R_BICOLS, [
    'Investor / Household','Current Value','Portfolio %','# Accounts','# Holdings','','','','','','',
  ], 5);

  for (let i = 0; i < INVESTORS.length; i++) {
    const r0 = R_BID0 + i;
    const r  = r0 + 1;
    const inv = INVESTORS[i];
    const bg = i % 2 === 0 ? C.panel : C.altRow;

    vals.push({ range: `${S}!A${r}`, values: [[inv]] });
    vals.push({ range: `${S}!B${r}`, values: [[`=IFERROR(SUMPRODUCT((${HOLD}!$B$6:$B$1005=A${r})*(${HOLD}!$O$6:$O$1005)),0)`]] });
    vals.push({ range: `${S}!C${r}`, values: [[`=IFERROR(B${r}/SUM($B$${BI1}:$B$${BIN}),0)`]] });
    vals.push({ range: `${S}!D${r}`, values: [[`=IFERROR(COUNTIF(${ACT}!$C$6:$C$305,A${r}),0)`]] });
    vals.push({ range: `${S}!E${r}`, values: [[`=IFERROR(SUMPRODUCT((${HOLD}!$B$6:$B$1005=A${r})*(${HOLD}!$J$6:$J$1005>0)*1),0)`]] });

    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 5), cell: { userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { fontSize: 9, fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r0, endIndex: r0+1 },
      properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  }

  // Investor number formats
  fmt.push({ repeatCell: { range: gridRange(SID, R_BID0, R_BITOT, 1, 2), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID, R_BID0, R_BITOT, 2, 3), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  [1,2,3,4].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_BID0, R_BITOT, c, c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });
  fmt.push({ repeatCell: { range: gridRange(SID, R_BID0, R_BITOT, 0, 1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Investor total row
  totalRowStyle(R_BITOT, 'TOTAL', 5);
  vals.push({ range: `${S}!B${BITOT}`, values: [[`=SUM($B$${BI1}:$B$${BIN})`]] });
  vals.push({ range: `${S}!C${BITOT}`, values: [[`=IFERROR(SUM($C$${BI1}:$C$${BIN}),0)`]] });
  vals.push({ range: `${S}!D${BITOT}`, values: [[`=SUM($D$${BI1}:$D$${BIN})`]] });
  vals.push({ range: `${S}!E${BITOT}`, values: [[`=SUM($E$${BI1}:$E$${BIN})`]] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_BITOT, R_BITOT+1, 1, 2), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID, R_BITOT, R_BITOT+1, 2, 3), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  blankRow(R_BLANK4);

  // ── Charts ────────────────────────────────────────────────────────
  // Chart 1: Donut — Portfolio by Asset Class (anchored top-right)
  fmt.push({ addChart: { chart: {
    spec: {
      title: 'Portfolio by Asset Class',
      titleTextFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        domain: {
          sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_AD0, endRowIndex: R_ATOT, startColumnIndex: 0, endColumnIndex: 1 }] },
        },
        series: {
          sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_AD0, endRowIndex: R_ATOT, startColumnIndex: 1, endColumnIndex: 2 }] },
        },
        threeDimensional: false,
        pieHole: 0.4,
      },
      backgroundColor: hex(C.bg),
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_ACHDR, columnIndex: 12 },
      widthPixels: 380, heightPixels: 280,
    }},
  }}});

  // Chart 2: Horizontal bar — Current % vs Target %
  fmt.push({ addChart: { chart: {
    spec: {
      title: 'Current vs. Target Allocation',
      titleTextFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'BOTTOM_LEGEND',
        domains: [{ domain: {
          sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_ACOLS, endRowIndex: R_ATOT, startColumnIndex: 0, endColumnIndex: 1 }] },
        }}],
        series: [
          {
            series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_ACOLS, endRowIndex: R_ATOT, startColumnIndex: 2, endColumnIndex: 3 }] }},
            targetAxis: 'LEFT_AXIS',
          },
          {
            series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_ACOLS, endRowIndex: R_ATOT, startColumnIndex: 3, endColumnIndex: 4 }] }},
            targetAxis: 'LEFT_AXIS',
          },
        ],
        headerCount: 1,
      },
      backgroundColor: hex(C.bg),
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_BAHDR, columnIndex: 12 },
      widthPixels: 380, heightPixels: 330,
    }},
  }}});

  // Chart 3: Pie — Allocation by Account
  fmt.push({ addChart: { chart: {
    spec: {
      title: 'Portfolio by Account',
      titleTextFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        domain: {
          sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_BAD0, endRowIndex: R_BATOT, startColumnIndex: 1, endColumnIndex: 2 }] },
        },
        series: {
          sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_BAD0, endRowIndex: R_BATOT, startColumnIndex: 3, endColumnIndex: 4 }] },
        },
        threeDimensional: false,
        pieHole: 0.3,
      },
      backgroundColor: hex(C.bg),
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_ACHDR, columnIndex: 19 },
      widthPixels: 380, heightPixels: 280,
    }},
  }}});

  // Chart 4: Column — Allocation by Investor
  fmt.push({ addChart: { chart: {
    spec: {
      title: 'Portfolio by Investor',
      titleTextFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        domains: [{ domain: {
          sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_BICOLS, endRowIndex: R_BITOT, startColumnIndex: 0, endColumnIndex: 1 }] },
        }}],
        series: [{
          series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_BICOLS, endRowIndex: R_BITOT, startColumnIndex: 1, endColumnIndex: 2 }] }},
          targetAxis: 'LEFT_AXIS',
        }],
        headerCount: 1,
      },
      backgroundColor: hex(C.bg),
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_BIHDR, columnIndex: 19 },
      widthPixels: 380, heightPixels: 280,
    }},
  }}});

  // ── Flush ─────────────────────────────────────────────────────────
  await batchUpdate(id, fmt, '08-allocation');
  await valuesBatchUpdate(id, vals, '08-allocation');
  console.log('✓ Portfolio Allocation tab complete');
})();
