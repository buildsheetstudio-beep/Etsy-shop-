'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID   = sheetMap['Contributions & Savings'];
const S     = "'Contributions & Savings'";
const SETUP = "'Portfolio Setup'";
const ACT   = "'Account Tracker'";
const REF   = "'Reference Data'";

// ── Row layout (0-indexed) ────────────────────────────────────────────
const R_TITLE   = 0;   // rows 0-1 merged
const R_NOTE    = 2;
const R_CARDHDR = 3;
const R_CARD    = 4;
const R_BLANK1  = 5;
const R_CTRL    = 6;   // year filter (B7 = $B$7)
const R_BLANK2  = 7;
const R_LOGHDR  = 8;   // Contribution Log header
const R_LOGCOLS = 9;   // col headers
const R_DATA0   = 10;  // data start (row 11, 1-indexed)

const CTRL_CELL = '$B$7';  // year filter (0-indexed row 6, col B = 1-indexed B7)
const DATA_R1   = R_DATA0 + 1; // = 11 (1-indexed first data row)

// ── Build 104 records (4/month × 26 months: Jan 2024 – Feb 2026) ─────
const TXDATA = [];
for (let m = 0; m < 26; m++) {
  const year  = m < 12 ? 2024 : (m < 24 ? 2025 : 2026);
  const month = (m % 12) + 1;
  const date  = `${month}/15/${year}`;
  const done  = year < 2026 || (year === 2026 && month <= 2);
  const status = done ? 'Completed' : 'Planned';

  // [date, accountId, contribType, amount, empMatch, status, notes]
  TXDATA.push([date, 'ACT-003', 'Recurring', 1800 + (m % 3 === 0 ? 150 : 0), 600, status, '']);
  TXDATA.push([date, 'ACT-006', 'Recurring', 1200 + (m % 4 === 0 ? 100 : 0), 400, status, '']);
  TXDATA.push([date, 'ACT-007', 'Recurring', 800,  0, status, '']);
  TXDATA.push([date, m % 2 === 0 ? 'ACT-002' : 'ACT-005', 'Recurring', 500, 0, status, '']);
}
const N = TXDATA.length;  // 104

// Summary section rows start after data
const R_ANN_HDR  = R_DATA0 + N + 1;  // Annual Summary header
const R_ANN_COLS = R_ANN_HDR + 1;
const R_ANN_D0   = R_ANN_HDR + 2;    // 3 years
const R_ANN_TOT  = R_ANN_D0 + 3;
const R_BLANK3   = R_ANN_TOT + 1;
const R_ACC_HDR  = R_BLANK3 + 1;     // By Account header
const R_ACC_COLS = R_ACC_HDR + 1;
const R_ACC_D0   = R_ACC_COLS + 1;   // 13 accounts
const R_ACC_TOT  = R_ACC_D0 + 13;
const R_BLANK4   = R_ACC_TOT + 1;
const R_MON_HDR  = R_BLANK4 + 1;     // Monthly Totals header (chart data)
const R_MON_COLS = R_MON_HDR + 1;
const R_MON_D0   = R_MON_COLS + 1;   // 26 months
const R_MON_END  = R_MON_D0 + 25;    // last month (0-indexed)

const YEARS = [2024, 2025, 2026];
const ACCOUNT_IDS = [
  'ACT-001','ACT-002','ACT-003','ACT-004','ACT-005','ACT-006',
  'ACT-007','ACT-008','ACT-009','ACT-010','ACT-011','ACT-012','ACT-013',
];
const MONTH_LABELS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function colLetter(c) { return String.fromCharCode(65 + c); }

(async () => {
  const vals = [];
  const fmt  = [];

  // ── Background ────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID,0,1600,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // ── Column widths ─────────────────────────────────────────────────
  [
    [0,90],[1,88],[2,55],[3,48],[4,80],[5,165],[6,115],[7,115],[8,100],[9,100],[10,105],[11,100],[12,165],
    [13,16],[14,100],[15,100],[16,100],[17,100],
  ].forEach(([ci,px]) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // ── Title banner ──────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_TITLE, R_TITLE+2, 0, 13), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A1`, values: [['CONTRIBUTIONS & SAVINGS\nTrack all contributions by account, investor, and type with employer match']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_TITLE, R_TITLE+2, 0, 13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_TITLE, endIndex: R_TITLE+2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' }});
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' }});

  // ── Note ─────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_NOTE, R_NOTE+1, 0, 13), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A3`, values: [['Enter each contribution separately. Employer matches are tracked in their own column. Use the Year filter to update summary cards.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_NOTE, R_NOTE+1, 0, 13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.info),
    textFormat: { fontSize: 8, italic: true, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_NOTE, endIndex: R_NOTE+1 },
    properties: { pixelSize: 20 }, fields: 'pixelSize' }});

  // ── Summary cards ─────────────────────────────────────────────────
  const CARDS = [
    { label: 'Your Contributions (Year)',  val: `=IFERROR(SUMPRODUCT(($C$${DATA_R1}:$C$500=${CTRL_CELL})*($I$${DATA_R1}:$I$500)),0)`, type: 'CURRENCY', pat: '$#,##0' },
    { label: 'Employer Match (Year)',      val: `=IFERROR(SUMPRODUCT(($C$${DATA_R1}:$C$500=${CTRL_CELL})*($J$${DATA_R1}:$J$500)),0)`, type: 'CURRENCY', pat: '$#,##0' },
    { label: 'Total w/ Match (Year)',      val: `=IFERROR(SUMPRODUCT(($C$${DATA_R1}:$C$500=${CTRL_CELL})*($K$${DATA_R1}:$K$500)),0)`, type: 'CURRENCY', pat: '$#,##0' },
    { label: '% of Annual Goal',          val: `=IFERROR(SUMPRODUCT(($C$${DATA_R1}:$C$500=${CTRL_CELL})*($I$${DATA_R1}:$I$500))/${SETUP}!$B$13,0)`, type: 'PERCENT', pat: '0.0%' },
  ];
  const cardCols = [[0,3],[3,6],[6,9],[9,13]];
  CARDS.forEach((card, i) => {
    const [c0, c1] = cardCols[i];
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARDHDR, R_CARDHDR+1, c0, c1), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARD, R_CARD+1, c0, c1), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${colLetter(c0)}${R_CARDHDR+1}`, values: [[card.label]] });
    vals.push({ range: `${S}!${colLetter(c0)}${R_CARD+1}`, values: [[card.val]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARDHDR, R_CARDHDR+1, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrB), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARD, R_CARD+1, c0, c1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.highlight), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: card.type, pattern: card.pat },
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARDHDR, endIndex: R_CARD+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ── Year filter ───────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_BLANK1, endIndex: R_BLANK1+1 },
    properties: { pixelSize: 8 }, fields: 'pixelSize' }});
  vals.push({ range: `${S}!A7`, values: [['Filter Year:']] });
  vals.push({ range: `${S}!B7`, values: [['2025']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_CTRL, R_CTRL+1, 0, 1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID, R_CTRL, R_CTRL+1, 1, 2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE', borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ setDataValidation: { range: gridRange(SID, R_CTRL, R_CTRL+1, 1, 2),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      { userEnteredValue: '2024' }, { userEnteredValue: '2025' }, { userEnteredValue: '2026' },
    ]}, showCustomUi: true, strict: false } }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CTRL, endIndex: R_CTRL+1 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_BLANK2, endIndex: R_BLANK2+1 },
    properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ── Contribution Log section header ───────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_LOGHDR, R_LOGHDR+1, 0, 13), mergeType: 'MERGE_ALL' }});
  vals.push({ range: `${S}!A9`, values: [['  Contribution Log  —  Add new contributions below row 10. Col A auto-generates IDs.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_LOGHDR, R_LOGHDR+1, 0, 13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrA), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_LOGHDR, endIndex: R_LOGHDR+1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // ── Contribution Log column headers ───────────────────────────────
  const LOG_HEADERS = [
    'Contrib. ID','Date','Year','Mo.','Account ID','Account Name','Owner',
    'Type','Amount ($)','Emp. Match ($)','Total ($)','Status','Notes',
  ];
  vals.push({ range: `${S}!A${R_LOGCOLS+1}`, values: [LOG_HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_LOGCOLS, R_LOGCOLS+1, 0, 13), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrC), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_LOGCOLS, endIndex: R_LOGCOLS+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Freeze log header rows + col A-D
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: R_LOGCOLS+1, frozenColumnCount: 4 } }, fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' }});

  // ── Data values (push by column for efficiency) ───────────────────
  const fmtColB  = [];  // dates
  const fmtColE  = [];  // account IDs
  const fmtColH  = [];  // types
  const fmtColI  = [];  // amounts
  const fmtColJ  = [];  // employer match
  const fmtColL  = [];  // status
  const fmtColM  = [];  // notes
  const fmtColA  = [];  // ID formula
  const fmtColC  = [];  // year formula
  const fmtColD  = [];  // month formula
  const fmtColF  = [];  // account name formula
  const fmtColG  = [];  // owner formula
  const fmtColK  = [];  // total formula

  TXDATA.forEach((row, i) => {
    const r = DATA_R1 + i;
    fmtColB.push([row[0]]);
    fmtColE.push([row[1]]);
    fmtColH.push([row[2]]);
    fmtColI.push([row[3]]);
    fmtColJ.push([row[4]]);
    fmtColL.push([row[5]]);
    fmtColM.push([row[6]]);
    fmtColA.push([`=IF(B${r}="","","CTB-"&TEXT(ROW()-${DATA_R1-1},"0000"))`]);
    fmtColC.push([`=IFERROR(YEAR(B${r}),"")`]);
    fmtColD.push([`=IFERROR(MONTH(B${r}),"")`]);
    fmtColF.push([`=IFERROR(VLOOKUP(E${r},${ACT}!$A$6:$B$305,2,FALSE),"")`]);
    fmtColG.push([`=IFERROR(VLOOKUP(E${r},${ACT}!$A$6:$C$305,3,FALSE),"")`]);
    fmtColK.push([`=IFERROR(I${r}+J${r},0)`]);
  });

  vals.push({ range: `${S}!B${DATA_R1}`, values: fmtColB });
  vals.push({ range: `${S}!E${DATA_R1}`, values: fmtColE });
  vals.push({ range: `${S}!H${DATA_R1}`, values: fmtColH });
  vals.push({ range: `${S}!I${DATA_R1}`, values: fmtColI });
  vals.push({ range: `${S}!J${DATA_R1}`, values: fmtColJ });
  vals.push({ range: `${S}!L${DATA_R1}`, values: fmtColL });
  vals.push({ range: `${S}!M${DATA_R1}`, values: fmtColM });
  vals.push({ range: `${S}!A${DATA_R1}`, values: fmtColA });
  vals.push({ range: `${S}!C${DATA_R1}`, values: fmtColC });
  vals.push({ range: `${S}!D${DATA_R1}`, values: fmtColD });
  vals.push({ range: `${S}!F${DATA_R1}`, values: fmtColF });
  vals.push({ range: `${S}!G${DATA_R1}`, values: fmtColG });
  vals.push({ range: `${S}!K${DATA_R1}`, values: fmtColK });

  // Row styling for data rows
  for (let i = 0; i < N; i++) {
    const r0 = R_DATA0 + i;
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 13), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r0, endIndex: r0+1 },
      properties: { pixelSize: 18 }, fields: 'pixelSize' }});
  }

  // Formula tint on computed cols (A,C,D,F,G,K)
  [0,2,3,5,6,10].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_DATA0, R_DATA0+N, c, c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });
  // Input tint on entry cols (B,E,H,I,J,L,M)
  [1,4,7,8,9,11,12].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_DATA0, R_DATA0+N, c, c+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });

  // Number formats for data rows
  fmt.push({ repeatCell: { range: gridRange(SID, R_DATA0, R_DATA0+N, 1, 2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'MMM D, YYYY' },
  }}, fields: 'userEnteredFormat.numberFormat' }});
  [8,9,10].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_DATA0, R_DATA0+N, c, c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Year + month cols: center
  [2,3].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_DATA0, R_DATA0+N, c, c+1), cell: { userEnteredFormat: {
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Dropdowns for data columns E, H, L
  // Account ID dropdown
  fmt.push({ setDataValidation: { range: gridRange(SID, R_DATA0, R_DATA0+N, 4, 5),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${ACT}!$A$6:$A$305` }] }, showCustomUi: true, strict: false } }});
  // Contribution Type dropdown
  fmt.push({ setDataValidation: { range: gridRange(SID, R_DATA0, R_DATA0+N, 7, 8),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$138:$A$143` }] }, showCustomUi: true, strict: false } }});
  // Status dropdown
  fmt.push({ setDataValidation: { range: gridRange(SID, R_DATA0, R_DATA0+N, 11, 12),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$192:$A$196` }] }, showCustomUi: true, strict: false } }});

  // Conditional: Status color
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_DATA0, R_DATA0+N, 11, 12)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Completed' }] },
      format: { backgroundColor: hex(C.success) },
    },
  }, index: 0 }});
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_DATA0, R_DATA0+N, 11, 12)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Skipped' }] },
      format: { backgroundColor: hex(C.attention) },
    },
  }, index: 0 }});
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_DATA0, R_DATA0+N, 11, 12)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Planned' }] },
      format: { backgroundColor: hex(C.info) },
    },
  }, index: 0 }});

  // ── Annual Summary ────────────────────────────────────────────────
  const DATA_LAST = DATA_R1 + N - 1;  // last data row, 1-indexed
  const DATA_RNG  = `$${DATA_LAST + 100}`;  // generous range ceiling

  function sectionHdr(row, label, nc) {
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, nc), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${row+1}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, 0, nc), cell: { userEnteredFormat: {
      backgroundColor: hex(C.hdrA), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: row, endIndex: row+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  }

  sectionHdr(R_ANN_HDR, '  Annual Contribution Summary', 7);
  vals.push({ range: `${S}!A${R_ANN_COLS+1}`, values: [['Year','Your Contributions','Employer Match','Total (Incl. Match)','Annual Goal','% of Goal','']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ANN_COLS, R_ANN_COLS+1, 0, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrC), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  YEARS.forEach((yr, i) => {
    const r0 = R_ANN_D0 + i;
    const r  = r0 + 1;
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    const DR  = `$C$${DATA_R1}:$C$500`;
    const IRP = `$I$${DATA_R1}:$I$500`;
    const JRP = `$J$${DATA_R1}:$J$500`;
    const KRP = `$K$${DATA_R1}:$K$500`;

    vals.push({ range: `${S}!A${r}`, values: [[yr]] });
    vals.push({ range: `${S}!B${r}`, values: [[`=IFERROR(SUMPRODUCT((${DR}=${yr})*(${IRP})),0)`]] });
    vals.push({ range: `${S}!C${r}`, values: [[`=IFERROR(SUMPRODUCT((${DR}=${yr})*(${JRP})),0)`]] });
    vals.push({ range: `${S}!D${r}`, values: [[`=IFERROR(SUMPRODUCT((${DR}=${yr})*(${KRP})),0)`]] });
    vals.push({ range: `${S}!E${r}`, values: [[`=${SETUP}!$B$13`]] });
    vals.push({ range: `${S}!F${r}`, values: [[`=IFERROR(B${r}/E${r},0)`]] });

    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 7), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r0, endIndex: r0+1 },
      properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  });
  // Annual total row
  const r_ann_tot = R_ANN_TOT + 1;
  vals.push({ range: `${S}!A${r_ann_tot}`, values: [['TOTAL']] });
  vals.push({ range: `${S}!B${r_ann_tot}`, values: [[`=SUM(B${R_ANN_D0+1}:B${R_ANN_D0+3})`]] });
  vals.push({ range: `${S}!C${r_ann_tot}`, values: [[`=SUM(C${R_ANN_D0+1}:C${R_ANN_D0+3})`]] });
  vals.push({ range: `${S}!D${r_ann_tot}`, values: [[`=SUM(D${R_ANN_D0+1}:D${R_ANN_D0+3})`]] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ANN_TOT, R_ANN_TOT+1, 0, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_ANN_TOT, endIndex: R_ANN_TOT+1 },
    properties: { pixelSize: 22 }, fields: 'pixelSize' }});
  // Number formats for annual section
  [1,2,3,4].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_ANN_D0, R_ANN_TOT+1, c, c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0' },
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ANN_D0, R_ANN_TOT, 5, 6), cell: { userEnteredFormat: {
    numberFormat: { type: 'PERCENT', pattern: '0.0%' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // ── By Account ────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_BLANK3, endIndex: R_BLANK3+1 },
    properties: { pixelSize: 10 }, fields: 'pixelSize' }});
  sectionHdr(R_ACC_HDR, '  Total Contributions by Account (All Years)', 6);
  vals.push({ range: `${S}!A${R_ACC_COLS+1}`, values: [['Account ID','Account Name','Owner','Total Contributions','Employer Match','Total (Incl. Match)']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ACC_COLS, R_ACC_COLS+1, 0, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrC), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_ACC_COLS, endIndex: R_ACC_COLS+1 },
    properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  ACCOUNT_IDS.forEach((acct, i) => {
    const r0 = R_ACC_D0 + i;
    const r  = r0 + 1;
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    const ER  = `$E$${DATA_R1}:$E$500`;

    vals.push({ range: `${S}!A${r}`, values: [[acct]] });
    vals.push({ range: `${S}!B${r}`, values: [[`=IFERROR(VLOOKUP(A${r},${ACT}!$A$6:$B$305,2,FALSE),"")`]] });
    vals.push({ range: `${S}!C${r}`, values: [[`=IFERROR(VLOOKUP(A${r},${ACT}!$A$6:$C$305,3,FALSE),"")`]] });
    vals.push({ range: `${S}!D${r}`, values: [[`=IFERROR(SUMPRODUCT((${ER}=A${r})*($I$${DATA_R1}:$I$500)),0)`]] });
    vals.push({ range: `${S}!E${r}`, values: [[`=IFERROR(SUMPRODUCT((${ER}=A${r})*($J$${DATA_R1}:$J$500)),0)`]] });
    vals.push({ range: `${S}!F${r}`, values: [[`=IFERROR(D${r}+E${r},0)`]] });

    fmt.push({ repeatCell: { range: gridRange(SID, r0, r0+1, 0, 6), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r0, endIndex: r0+1 },
      properties: { pixelSize: 20 }, fields: 'pixelSize' }});
  });
  [3,4,5].forEach(c => {
    fmt.push({ repeatCell: { range: gridRange(SID, R_ACC_D0, R_ACC_TOT+1, c, c+1), cell: { userEnteredFormat: {
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0' },
    }}, fields: 'userEnteredFormat.numberFormat' }});
  });
  // Account total
  const r_acc_tot = R_ACC_TOT + 1;
  vals.push({ range: `${S}!A${r_acc_tot}`, values: [['TOTAL']] });
  vals.push({ range: `${S}!D${r_acc_tot}`, values: [[`=SUM(D${R_ACC_D0+1}:D${R_ACC_D0+13})`]] });
  vals.push({ range: `${S}!E${r_acc_tot}`, values: [[`=SUM(E${R_ACC_D0+1}:E${R_ACC_D0+13})`]] });
  vals.push({ range: `${S}!F${r_acc_tot}`, values: [[`=SUM(F${R_ACC_D0+1}:F${R_ACC_D0+13})`]] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ACC_TOT, R_ACC_TOT+1, 0, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_ACC_TOT, endIndex: R_ACC_TOT+1 },
    properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ── Monthly totals (chart data) ───────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_BLANK4, endIndex: R_BLANK4+1 },
    properties: { pixelSize: 10 }, fields: 'pixelSize' }});
  sectionHdr(R_MON_HDR, '  Monthly Contribution Totals (chart data)', 5);
  vals.push({ range: `${S}!A${R_MON_COLS+1}`, values: [['Year','Month','Period','Contributions','Emp. Match']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_MON_COLS, R_MON_COLS+1, 0, 5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.hdrC), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  const monthDataFormulas = { A: [], B: [], C: [], D: [], E: [] };
  for (let m = 0; m < 26; m++) {
    const yr  = 2024 + Math.floor(m / 12);
    const mo  = (m % 12) + 1;
    const lbl = `${MONTH_LABELS[mo-1]} ${yr}`;
    const r   = R_MON_D0 + m + 1;  // 1-indexed
    const DR  = `$C$${DATA_R1}:$C$500`;
    const MR  = `$D$${DATA_R1}:$D$500`;
    monthDataFormulas.A.push([yr]);
    monthDataFormulas.B.push([mo]);
    monthDataFormulas.C.push([lbl]);
    monthDataFormulas.D.push([`=IFERROR(SUMPRODUCT((${DR}=${yr})*(${MR}=${mo})*($I$${DATA_R1}:$I$500)),0)`]);
    monthDataFormulas.E.push([`=IFERROR(SUMPRODUCT((${DR}=${yr})*(${MR}=${mo})*($J$${DATA_R1}:$J$500)),0)`]);
  }
  vals.push({ range: `${S}!A${R_MON_D0+1}`, values: monthDataFormulas.A });
  vals.push({ range: `${S}!B${R_MON_D0+1}`, values: monthDataFormulas.B });
  vals.push({ range: `${S}!C${R_MON_D0+1}`, values: monthDataFormulas.C });
  vals.push({ range: `${S}!D${R_MON_D0+1}`, values: monthDataFormulas.D });
  vals.push({ range: `${S}!E${R_MON_D0+1}`, values: monthDataFormulas.E });

  fmt.push({ repeatCell: { range: gridRange(SID, R_MON_D0, R_MON_END+1, 3, 5), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // ── Charts ────────────────────────────────────────────────────────
  // Chart 1: Column — Monthly contributions trend
  fmt.push({ addChart: { chart: {
    spec: {
      title: 'Monthly Contributions (Jan 2024 – Feb 2026)',
      titleTextFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'BOTTOM_LEGEND',
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_MON_COLS, endRowIndex: R_MON_END+1, startColumnIndex: 2, endColumnIndex: 3 }] }}}}],
        series: [
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_MON_COLS, endRowIndex: R_MON_END+1, startColumnIndex: 3, endColumnIndex: 4 }] }}, targetAxis: 'LEFT_AXIS' },
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_MON_COLS, endRowIndex: R_MON_END+1, startColumnIndex: 4, endColumnIndex: 5 }] }}, targetAxis: 'LEFT_AXIS' },
        ],
        headerCount: 1,
      },
      backgroundColor: hex(C.bg),
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_LOGHDR, columnIndex: 14 },
      widthPixels: 480, heightPixels: 280,
    }},
  }}});

  // Chart 2: Column — Annual comparison
  fmt.push({ addChart: { chart: {
    spec: {
      title: 'Annual Contributions by Year',
      titleTextFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'BOTTOM_LEGEND',
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_ANN_COLS, endRowIndex: R_ANN_TOT, startColumnIndex: 0, endColumnIndex: 1 }] }}}}],
        series: [
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_ANN_COLS, endRowIndex: R_ANN_TOT, startColumnIndex: 1, endColumnIndex: 2 }] }}, targetAxis: 'LEFT_AXIS' },
          { series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_ANN_COLS, endRowIndex: R_ANN_TOT, startColumnIndex: 2, endColumnIndex: 3 }] }}, targetAxis: 'LEFT_AXIS' },
        ],
        headerCount: 1,
      },
      backgroundColor: hex(C.bg),
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_ANN_HDR, columnIndex: 14 },
      widthPixels: 380, heightPixels: 260,
    }},
  }}});

  // Chart 3: Pie — Contributions by Account
  fmt.push({ addChart: { chart: {
    spec: {
      title: 'Contributions by Account',
      titleTextFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        domain: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_ACC_D0, endRowIndex: R_ACC_TOT, startColumnIndex: 1, endColumnIndex: 2 }] }},
        series: { sourceRange: { sources: [{ sheetId: SID, startRowIndex: R_ACC_D0, endRowIndex: R_ACC_TOT, startColumnIndex: 3, endColumnIndex: 4 }] }},
        threeDimensional: false, pieHole: 0.4,
      },
      backgroundColor: hex(C.bg),
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_ACC_HDR, columnIndex: 14 },
      widthPixels: 380, heightPixels: 280,
    }},
  }}});

  // ── Flush ─────────────────────────────────────────────────────────
  await batchUpdate(id, fmt, '10-contributions');
  await valuesBatchUpdate(id, vals, '10-contributions');
  console.log('✓ Contributions & Savings tab complete');
})();
