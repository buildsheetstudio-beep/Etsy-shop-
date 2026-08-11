'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Net Worth Tracker'];
const S   = "'Net Worth Tracker'";
const ACC = "'Account Tracker'";
const HLD = "'Holdings'";

// ─── Row constants (0-indexed) ────────────────────────────────────────────────
const R_TITLE   = 0;
const R_NOTE    = 2;
const R_CARDHDR = 3;
const R_CARD    = 4;

const R_AHDR    = 6;
const R_ACOLS   = 7;
const R_AD0     = 8;
const NA        = 7;
const R_ATOT    = R_AD0 + NA;          // 15

const R_LHDR    = R_ATOT + 2;          // 17
const R_LCOLS   = R_LHDR + 1;          // 18
const R_LD0     = R_LCOLS + 1;         // 19
const NL        = 7;
const R_LTOT    = R_LD0 + NL;          // 26
const R_NW      = R_LTOT + 1;          // 27

const R_MHDR    = R_NW + 2;            // 29
const R_MNOTE   = R_MHDR + 1;          // 30
const R_MCOLS   = R_MNOTE + 1;         // 31
const R_MD0     = R_MCOLS + 1;         // 32
const NM        = 60;
const R_MD_END  = R_MD0 + NM - 1;      // 91

// 1-indexed convenience
const AT1  = R_ATOT + 1;   // 16
const LT1  = R_LTOT + 1;   // 27
const NW1  = R_NW + 1;     // 28
const MD1  = R_MD0 + 1;    // 33
const MDN  = R_MD_END + 1; // 92

// ─── Auto-calc formulas (Holdings → Account Tracker Tax Treatment col F) ──────
// Holdings: col C = Account ID, col B = Owner, col O = Current Value
// Account Tracker: col A = Account ID, col F = Tax Treatment; data starts row 6
const vlk = `IFERROR(VLOOKUP(${HLD}!$C$6:$C$1005,${ACC}!$A$6:$F$305,6,FALSE),"")`;

function investFml(owner) {
  const of = owner ? `(${HLD}!$B$6:$B$1005="${owner}")*` : '';
  return `=IFERROR(SUMPRODUCT(${of}(${vlk}="Taxable")*(${HLD}!$O$6:$O$1005)),0)`;
}

function retireFml(owner) {
  const of = owner ? `(${HLD}!$B$6:$B$1005="${owner}")*` : '';
  return `=IFERROR(SUMPRODUCT(${of}((${vlk}="Tax-Deferred")+(${vlk}="Tax-Free / Roth"))*(${HLD}!$O$6:$O$1005)),0)`;
}

const OWNERS = ['Daniel Walsh', 'Emily Walsh', 'Joint Household'];

// Asset categories and their sample data [daniel, emily, joint] (null = auto-calc)
const ASSET_CATS = [
  { name: 'Investment Accounts',  data: null },
  { name: 'Retirement Accounts',  data: null },
  { name: 'Cash & Savings',       data: [28500, 22000, 45000] },
  { name: 'Real Estate',          data: [0, 0, 485000] },
  { name: 'Business Ownership',   data: [0, 0, 0] },
  { name: 'Vehicles',             data: [18500, 22000, 0] },
  { name: 'Other Assets',         data: [5000, 3500, 0] },
];

const LIAB_CATS = [
  { name: 'Mortgage',          data: [0, 0, 312000] },
  { name: 'Student Loans',     data: [0, 18500, 0] },
  { name: 'Auto Loans',        data: [0, 14200, 0] },
  { name: 'Credit Cards',      data: [2800, 1400, 0] },
  { name: 'Personal Loans',    data: [0, 0, 0] },
  { name: 'Business Debt',     data: [0, 0, 0] },
  { name: 'Other Liabilities', data: [0, 0, 0] },
];

// Monthly log: 60 months Jan 2021–Dec 2025
function buildMonthlyLog() {
  const rows = [];
  let assets = 820000;
  let liabs  = 345000;

  // Deterministic pseudo-random via simple LCG
  let seed = 17;
  const rand = () => {
    seed = (seed * 1664525 + 1013904223) & 0x7fffffff;
    return seed / 0x7fffffff;
  };

  for (let m = 0; m < NM; m++) {
    const year  = 2021 + Math.floor(m / 12);
    const month = (m % 12) + 1;
    rows.push({ date: `${month}/1/${year}`, assets: Math.round(assets), liabs: Math.round(liabs) });

    // 2022 was a down year (months 12-23)
    const isDown = (m >= 12 && m <= 23);
    const ret    = isDown
      ? -0.014 + (rand() - 0.5) * 0.018
      :  0.008 + (rand() - 0.5) * 0.012;
    const contrib = 2800 + rand() * 400;

    assets = assets * (1 + ret) + contrib;
    liabs  = liabs  * 0.9983; // ~0.17% paydown/month
    if (m >= 36 && m <= 47) liabs -= 250; // extra paydown 2024
  }
  return rows;
}

(async () => {
  const fmt = [];
  const vals = [];

  // ── Background ───────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 500, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column widths ────────────────────────────────────────────────────────────
  const COL_W = [180, 120, 120, 120, 120, 80, 180, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20];
  COL_W.forEach((px, ci) => fmt.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
    properties: { pixelSize: px }, fields: 'pixelSize' } }));

  // ── Title banner (rows 0-1) ──────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 2, 0, 13), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['NET WORTH TRACKER\nAssets • Liabilities • Monthly History • Growth Trends']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 2, 0, 13),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 14,
        foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' } });

  // ── Note row ─────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_NOTE, R_NOTE + 1, 0, 13), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_NOTE + 1}`, values: [['Investment & Retirement values are auto-calculated from the Holdings tab. Enter Cash, Real Estate, Vehicle, and Liability amounts manually.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_NOTE, R_NOTE + 1, 0, 13),
    cell: { userEnteredFormat: { backgroundColor: hex(C.info),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.text) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', wrapStrategy: 'CLIP' } },
    fields: 'userEnteredFormat' } });

  // ── Summary cards (rows 3-4) ─────────────────────────────────────────────────
  const cards = [
    { label: 'NET WORTH',          val: `=$E$${NW1}`,                          fmt: '$#,##0',   color: C.primary,   tcol: C.primaryText },
    { label: 'TOTAL ASSETS',       val: `=$E$${AT1}`,                          fmt: '$#,##0',   color: C.hdrC,      tcol: C.primaryText },
    { label: 'TOTAL LIABILITIES',  val: `=$E$${LT1}`,                          fmt: '$#,##0',   color: C.secondary, tcol: C.primaryText },
    { label: 'ASSET/DEBT RATIO',   val: `=IFERROR($E$${AT1}/$E$${LT1},"—")`,  fmt: '0.00"x"',  color: C.highlight, tcol: C.text },
  ];

  const cardCols = [[0,3],[3,6],[6,9],[9,12]];
  cards.forEach(({ label, val, fmt: nf, color, tcol }, i) => {
    const [c1, c2] = cardCols[i];
    // Header
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARDHDR, R_CARDHDR + 1, c1, c2), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colLetter(c1)}${R_CARDHDR + 1}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARDHDR, R_CARDHDR + 1, c1, c2),
      cell: { userEnteredFormat: { backgroundColor: hex(color),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(tcol), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
      fields: 'userEnteredFormat' } });
    // Value
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARD, R_CARD + 1, c1, c2), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${colLetter(c1)}${R_CARD + 1}`, values: [[val]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARD, R_CARD + 1, c1, c2),
      cell: { userEnteredFormat: { backgroundColor: hex(color),
        textFormat: { bold: true, fontSize: 18, foregroundColor: hex(tcol), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'TOP',
        numberFormat: { type: 'NUMBER', pattern: nf } } },
      fields: 'userEnteredFormat' } });
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARDHDR, endIndex: R_CARDHDR + 1 },
    properties: { pixelSize: 20 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARD, endIndex: R_CARD + 1 },
    properties: { pixelSize: 42 }, fields: 'pixelSize' } });

  // ── ASSETS section header ─────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_AHDR, R_AHDR + 1, 0, 7), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_AHDR + 1}`, values: [['CURRENT ASSETS']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_AHDR, R_AHDR + 1, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_AHDR, endIndex: R_AHDR + 1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Asset column headers
  const ASSET_HDRS = ['Asset Category', 'Daniel Walsh', 'Emily Walsh', 'Joint Household', 'Total', '% of Assets', 'Notes'];
  vals.push({ range: `${S}!A${R_ACOLS + 1}`, values: [ASSET_HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ACOLS, R_ACOLS + 1, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });

  // Asset data rows
  const assetRows = [];
  ASSET_CATS.forEach((cat, i) => {
    const r1 = R_AD0 + i + 1; // 1-indexed
    let bFml, cFml, dFml;

    if (cat.name === 'Investment Accounts') {
      bFml = investFml('Daniel Walsh');
      cFml = investFml('Emily Walsh');
      dFml = investFml('Joint Household');
    } else if (cat.name === 'Retirement Accounts') {
      bFml = retireFml('Daniel Walsh');
      cFml = retireFml('Emily Walsh');
      dFml = retireFml('Joint Household');
    } else {
      bFml = cat.data[0];
      cFml = cat.data[1];
      dFml = cat.data[2];
    }

    const eFml = `=IFERROR(B${r1}+C${r1}+D${r1},0)`;
    const fFml = `=IFERROR(E${r1}/$E$${AT1},0)`;
    assetRows.push([cat.name, bFml, cFml, dFml, eFml, fFml, '']);
  });
  vals.push({ range: `${S}!A${R_AD0 + 1}`, values: assetRows });

  // Asset formatting
  fmt.push({ repeatCell: { range: gridRange(SID, R_AD0, R_ATOT, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  // Alternate rows
  for (let i = 0; i < NA; i += 2) {
    fmt.push({ repeatCell: { range: gridRange(SID, R_AD0 + i, R_AD0 + i + 1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor' } });
  }
  // Number formats (cols B-E currency, col F %)
  fmt.push({ repeatCell: { range: gridRange(SID, R_AD0, R_ATOT, 1, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$#,##0' } } },
    fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_AD0, R_ATOT, 5, 6),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat' } });

  // Total Assets row
  vals.push({ range: `${S}!A${AT1}`, values: [['Total Assets',
    `=IFERROR(SUM(B${R_AD0 + 1}:B${AT1 - 1}),0)`,
    `=IFERROR(SUM(C${R_AD0 + 1}:C${AT1 - 1}),0)`,
    `=IFERROR(SUM(D${R_AD0 + 1}:D${AT1 - 1}),0)`,
    `=IFERROR(SUM(B${AT1}:D${AT1}),0)`,
    '',
    '']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_ATOT, R_ATOT + 1, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrC),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      numberFormat: { type: 'NUMBER', pattern: '$#,##0' } } },
    fields: 'userEnteredFormat' } });

  // ── LIABILITIES section ───────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_LHDR, R_LHDR + 1, 0, 7), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_LHDR + 1}`, values: [['CURRENT LIABILITIES']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_LHDR, R_LHDR + 1, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.secondary),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_LHDR, endIndex: R_LHDR + 1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Liability column headers
  const LIAB_HDRS = ['Liability Category', 'Daniel Walsh', 'Emily Walsh', 'Joint Household', 'Total', '% of Liab.', 'Notes'];
  vals.push({ range: `${S}!A${R_LCOLS + 1}`, values: [LIAB_HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_LCOLS, R_LCOLS + 1, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.secondaryDk),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });

  // Liability data rows
  const liabRows = LIAB_CATS.map((cat, i) => {
    const r1 = R_LD0 + i + 1;
    const eFml = `=IFERROR(B${r1}+C${r1}+D${r1},0)`;
    const fFml = `=IFERROR(E${r1}/$E$${LT1},0)`;
    return [cat.name, cat.data[0], cat.data[1], cat.data[2], eFml, fFml, ''];
  });
  vals.push({ range: `${S}!A${R_LD0 + 1}`, values: liabRows });

  // Liability formatting
  fmt.push({ repeatCell: { range: gridRange(SID, R_LD0, R_LTOT, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' },
      verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  for (let i = 0; i < NL; i += 2) {
    fmt.push({ repeatCell: { range: gridRange(SID, R_LD0 + i, R_LD0 + i + 1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor' } });
  }
  fmt.push({ repeatCell: { range: gridRange(SID, R_LD0, R_LTOT, 1, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$#,##0' } } },
    fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_LD0, R_LTOT, 5, 6),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat' } });

  // Total Liabilities row
  vals.push({ range: `${S}!A${LT1}`, values: [['Total Liabilities',
    `=IFERROR(SUM(B${R_LD0 + 1}:B${LT1 - 1}),0)`,
    `=IFERROR(SUM(C${R_LD0 + 1}:C${LT1 - 1}),0)`,
    `=IFERROR(SUM(D${R_LD0 + 1}:D${LT1 - 1}),0)`,
    `=IFERROR(SUM(B${LT1}:D${LT1}),0)`,
    '', '']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_LTOT, R_LTOT + 1, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.attention),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      numberFormat: { type: 'NUMBER', pattern: '$#,##0' } } },
    fields: 'userEnteredFormat' } });

  // Net Worth row
  vals.push({ range: `${S}!A${NW1}`, values: [['NET WORTH',
    `=IFERROR(B${AT1}-B${LT1},0)`,
    `=IFERROR(C${AT1}-C${LT1},0)`,
    `=IFERROR(D${AT1}-D${LT1},0)`,
    `=IFERROR(E${AT1}-E${LT1},0)`,
    '', '']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_NW, R_NW + 1, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      numberFormat: { type: 'NUMBER', pattern: '$#,##0' } } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_NW, endIndex: R_NW + 1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── MONTHLY LOG section ───────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_MHDR, R_MHDR + 1, 0, 9), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_MHDR + 1}`, values: [['MONTHLY NET WORTH LOG']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_MHDR, R_MHDR + 1, 0, 9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_MHDR, endIndex: R_MHDR + 1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Monthly log note
  fmt.push({ mergeCells: { range: gridRange(SID, R_MNOTE, R_MNOTE + 1, 0, 9), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_MNOTE + 1}`, values: [['Enter Total Assets and Total Liabilities each month for an ongoing net worth trend. Net Worth, MoM, and YoY change are calculated automatically.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_MNOTE, R_MNOTE + 1, 0, 9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formula),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.secText) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });

  // Monthly log column headers
  const LOG_HDRS = ['Month', 'Total Assets', 'Total Liabilities', 'Net Worth', 'MoM Change', 'MoM %', 'YoY Change', 'YoY %', 'Notes'];
  vals.push({ range: `${S}!A${R_MCOLS + 1}`, values: [LOG_HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_MCOLS, R_MCOLS + 1, 0, 9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });

  // Monthly log column widths (override for A-I)
  const LOG_COL_W = [110, 120, 120, 120, 110, 80, 110, 80, 180];
  LOG_COL_W.forEach((px, ci) => fmt.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
    properties: { pixelSize: px }, fields: 'pixelSize' } }));

  // Monthly log data
  const logData = buildMonthlyLog();
  const logValues = logData.map(({ date, assets, liabs }, i) => {
    const r1  = MD1 + i;
    const dFml = `=IFERROR(B${r1}-C${r1},0)`;
    const eMomFml = i === 0 ? '' : `=IFERROR(D${r1}-D${r1 - 1},0)`;
    const fMomFml = i === 0 ? '' : `=IFERROR((D${r1}-D${r1 - 1})/D${r1 - 1},0)`;
    const gYoYFml = i < 12  ? '' : `=IFERROR(D${r1}-D${r1 - 12},0)`;
    const hYoYFml = i < 12  ? '' : `=IFERROR((D${r1}-D${r1 - 12})/D${r1 - 12},0)`;
    return [date, assets, liabs, dFml, eMomFml, fMomFml, gYoYFml, hYoYFml, ''];
  });
  vals.push({ range: `${S}!A${MD1}`, values: logValues });

  // Monthly log formatting
  fmt.push({ repeatCell: { range: gridRange(SID, R_MD0, R_MD_END + 1, 0, 9),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  for (let i = 0; i < NM; i += 2) {
    fmt.push({ repeatCell: { range: gridRange(SID, R_MD0 + i, R_MD0 + i + 1, 0, 9),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor' } });
  }
  // Date format col A
  fmt.push({ repeatCell: { range: gridRange(SID, R_MD0, R_MD_END + 1, 0, 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm yyyy' } } },
    fields: 'userEnteredFormat.numberFormat' } });
  // Currency cols B, C, D, E, G
  [1, 2, 3, 4, 6].forEach(ci => fmt.push({ repeatCell: { range: gridRange(SID, R_MD0, R_MD_END + 1, ci, ci + 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$#,##0' } } },
    fields: 'userEnteredFormat.numberFormat' } }));
  // Percent cols F, H
  [5, 7].forEach(ci => fmt.push({ repeatCell: { range: gridRange(SID, R_MD0, R_MD_END + 1, ci, ci + 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' } } },
    fields: 'userEnteredFormat.numberFormat' } }));

  // Conditional format: NW positive → success, negative → attention
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_MD0, R_MD_END + 1, 3, 4)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
      format: { backgroundColor: hex(C.success) } }
  }, index: 0 } });
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_MD0, R_MD_END + 1, 3, 4)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
      format: { backgroundColor: hex(C.attention) } }
  }, index: 1 } });
  // MoM change: green positive, red negative
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_MD0, R_MD_END + 1, 4, 5)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
      format: { textFormat: { foregroundColor: hex(C.success) } } }
  }, index: 2 } });
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_MD0, R_MD_END + 1, 4, 5)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
      format: { textFormat: { foregroundColor: hex(C.attention) } } }
  }, index: 3 } });

  // ── Freeze rows ───────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID,
    gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  // ─── Apply all formatting + values ───────────────────────────────────────────
  await batchUpdate(id, fmt, 'net-worth-fmt');
  await valuesBatchUpdate(id, vals, 'net-worth-vals');

  // ─── Charts ──────────────────────────────────────────────────────────────────
  const charts = [];

  // 1. Net Worth over time — LINE chart (monthly log col D)
  charts.push({ addChart: { chart: {
    spec: { title: 'Net Worth Over Time', basicChart: {
      chartType: 'LINE',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Month' },
        { position: 'LEFT_AXIS', title: 'Net Worth ($)' },
      ],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_MCOLS, R_MD_END + 1, 0, 1)] } } }],
      series: [
        { series: { sourceRange: { sources: [gridRange(SID, R_MCOLS, R_MD_END + 1, 3, 4)] } },
          targetAxis: 'LEFT_AXIS', color: hex(C.primary), lineStyle: { width: 2 } },
      ],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_AHDR, columnIndex: 8 },
      widthPixels: 480, heightPixels: 250,
    } },
  } } });

  // 2. Assets vs Liabilities over time — LINE (multi-series)
  charts.push({ addChart: { chart: {
    spec: { title: 'Assets vs Liabilities', basicChart: {
      chartType: 'LINE',
      legendPosition: 'BOTTOM_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Month' },
        { position: 'LEFT_AXIS', title: 'Amount ($)' },
      ],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_MCOLS, R_MD_END + 1, 0, 1)] } } }],
      series: [
        { series: { sourceRange: { sources: [gridRange(SID, R_MCOLS, R_MD_END + 1, 1, 2)] } },
          targetAxis: 'LEFT_AXIS', color: hex(C.success) },
        { series: { sourceRange: { sources: [gridRange(SID, R_MCOLS, R_MD_END + 1, 2, 3)] } },
          targetAxis: 'LEFT_AXIS', color: hex(C.attention) },
      ],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_AHDR + 10, columnIndex: 8 },
      widthPixels: 480, heightPixels: 250,
    } },
  } } });

  // 3. Asset Breakdown — PIE chart (category totals col E)
  charts.push({ addChart: { chart: {
    spec: { title: 'Asset Breakdown', pieChart: {
      legendPosition: 'RIGHT_LEGEND',
      pieHole: 0.4,
      domain: { sourceRange: { sources: [gridRange(SID, R_ACOLS, R_ATOT, 0, 1)] } },
      series: { sourceRange: { sources: [gridRange(SID, R_ACOLS, R_ATOT, 4, 5)] } },
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_LHDR, columnIndex: 8 },
      widthPixels: 380, heightPixels: 230,
    } },
  } } });

  // 4. Liability Breakdown — PIE chart
  charts.push({ addChart: { chart: {
    spec: { title: 'Liability Breakdown', pieChart: {
      legendPosition: 'RIGHT_LEGEND',
      pieHole: 0.4,
      domain: { sourceRange: { sources: [gridRange(SID, R_LCOLS, R_LTOT, 0, 1)] } },
      series: { sourceRange: { sources: [gridRange(SID, R_LCOLS, R_LTOT, 4, 5)] } },
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_NW + 3, columnIndex: 8 },
      widthPixels: 380, heightPixels: 230,
    } },
  } } });

  // 5. YoY Net Worth Growth — COLUMN chart (col G = YoY Change)
  charts.push({ addChart: { chart: {
    spec: { title: 'Year-over-Year Net Worth Growth', basicChart: {
      chartType: 'COLUMN',
      legendPosition: 'NO_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Month' },
        { position: 'LEFT_AXIS', title: 'YoY Change ($)' },
      ],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_MCOLS, R_MD_END + 1, 0, 1)] } } }],
      series: [
        { series: { sourceRange: { sources: [gridRange(SID, R_MCOLS, R_MD_END + 1, 6, 7)] } },
          targetAxis: 'LEFT_AXIS', color: hex(C.info) },
      ],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_MHDR, columnIndex: 9 },
      widthPixels: 480, heightPixels: 250,
    } },
  } } });

  await batchUpdate(id, charts, 'net-worth-charts');
  console.log('✓ Net Worth Tracker complete');
})();

function colLetter(i) {
  return String.fromCharCode(65 + i);
}
