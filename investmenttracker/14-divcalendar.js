'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Dividend Calendar'];
const S   = "'Dividend Calendar'";
const DIV = "'Dividend Income'";

// ─── Row constants (0-indexed) ────────────────────────────────────────────────
const R_TITLE   = 0;
const R_CTRLHDR = 2;
const R_CTRL    = 3;    // year dropdown at B4 (1-indexed)
const R_GHDR    = 4;
const R_GCOLS   = 5;
const R_GD0     = 6;
const NS        = 6;    // 6 dividend-paying tickers
const R_GTOT    = R_GD0 + NS;  // = 12

// 1-indexed
const GD1  = R_GD0 + 1;   // = 7
const GDN  = R_GTOT;      // = 12
const GTOT1 = R_GTOT + 1; // = 13

// ─── Calendar securities ──────────────────────────────────────────────────────
// [ticker, name, freq, shares, divPerShareAnnualized, costBasis, qualified]
const SECS = [
  ['BND',  'Vanguard Total Bond Market ETF',  'Monthly',   300, 0.32*12, 73.80,  false],
  ['VTI',  'Vanguard Total Stock Market ETF', 'Quarterly', 300, 0.88*4,  200.00, true ],
  ['SCHD', 'Schwab US Dividend Equity ETF',   'Quarterly', 200, 0.26*4,  27.80,  true ],
  ['VNQ',  'Vanguard Real Estate ETF',        'Quarterly', 120, 0.72*4,  86.00,  false],
  ['AMT',  'American Tower Corp.',            'Quarterly',  50, 1.62*4,  215.00, false],
  ['SPG',  'Simon Property Group Inc.',       'Quarterly',  30, 2.05*4,  155.00, false],
];

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

// Year filter lives at B4 (1-indexed) = row 3 (0-indexed), col 1
const YEAR_CELL = '$B$4';

// Formula: net dividends for a ticker in a given month, using the year filter
function monthFml(ticker_ref, month) {
  const divRange = `${DIV}!$B$6:$B$2005`;
  return `=IFERROR(SUMPRODUCT(` +
    `(YEAR(${divRange})=VALUE(${YEAR_CELL}))*` +
    `(MONTH(${divRange})=${month})*` +
    `(${DIV}!$C$6:$C$2005=${ticker_ref})*` +
    `(${DIV}!$S$6:$S$2005="Received")*` +
    `(${DIV}!$O$6:$O$2005)` +
    `),0)`;
}

(async () => {
  const fmt  = [];
  const vals = [];

  // ── Background ────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 500, 0, 20),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const COL_W = [155, 60, 60, 70, 65, 65, 70, 65, 65, 70, 65, 65, 70, 80, 100, 75, 75];
  // A=ticker/security, B-M=12 months, N=annual, O=freq, P=YOC%, Q=shares
  COL_W.forEach((px, ci) => fmt.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
    properties: { pixelSize: px }, fields: 'pixelSize' } }));

  // ── Title banner ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 2, 0, 17), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['DIVIDEND CALENDAR\nSchedule • Monthly estimates • Payment timing by security']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 2, 0, 17),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' } });

  // ── Year control ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_CTRLHDR, R_CTRLHDR + 1, 0, 3), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_CTRLHDR + 1}`, values: [['DIVIDEND SCHEDULE']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_CTRLHDR, R_CTRLHDR + 1, 0, 3),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CTRLHDR, endIndex: R_CTRLHDR + 1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Year label + dropdown
  vals.push({ range: `${S}!A${R_CTRL + 1}`, values: [['Select Year:', '2025']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_CTRL, R_CTRL + 1, 0, 1),
    cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_CTRL, R_CTRL + 1, 1, 2),
    cell: { userEnteredFormat: { backgroundColor: hex(C.input),
      textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.primary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      numberFormat: { type: 'NUMBER', pattern: '0' } } },
    fields: 'userEnteredFormat' } });
  fmt.push({ setDataValidation: { range: gridRange(SID, R_CTRL, R_CTRL + 1, 1, 2),
    rule: { condition: { type: 'ONE_OF_LIST',
      values: [{ userEnteredValue: '2024' },{ userEnteredValue: '2025' },{ userEnteredValue: '2026' }] },
      showCustomUi: true, strict: false } } });

  // ── Calendar section header ───────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_GHDR, R_GHDR + 1, 0, 17), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_GHDR + 1}`, values: [[`=IFERROR("Monthly Dividend Schedule — "&TEXT(${YEAR_CELL},"0"),"Monthly Dividend Schedule")`]] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_GHDR, R_GHDR + 1, 0, 17),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_GHDR, endIndex: R_GHDR + 1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // ── Column headers ────────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A${R_GCOLS + 1}`,
    values: [['Security', ...MONTHS, 'Annual Total', 'Frequency', 'Yield on Cost', 'Shares']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_GCOLS, R_GCOLS + 1, 0, 17),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  // Month cols slightly lighter
  fmt.push({ repeatCell: { range: gridRange(SID, R_GCOLS, R_GCOLS + 1, 1, 13),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrC) } },
    fields: 'userEnteredFormat.backgroundColor' } });

  // ── Grid data rows ────────────────────────────────────────────────────────────
  const gridRows = SECS.map(([ticker, , freq, shares, annualDPS, cost, qual], i) => {
    const r1  = GD1 + i;
    const yoc = annualDPS / cost;
    const monthCells = MONTHS.map((_, mi) => monthFml(`A${r1}`, mi + 1));
    const annTotal  = `=IFERROR(SUM(B${r1}:M${r1}),0)`;
    return [ticker, ...monthCells, annTotal, freq, yoc, shares];
  });

  // Push tickers separately (values) and formulas separately
  const tickerAndOtherCols = SECS.map(([ticker, , freq, shares, annualDPS, cost], i) => {
    const r1  = GD1 + i;
    const yoc = annualDPS / cost;
    return [ticker, '', '', '', '', '', '', '', '', '', '', '', '', '', freq, yoc, shares];
  });
  const formulaCols = SECS.map(([ticker, , freq, shares], i) => {
    const r1  = GD1 + i;
    const mf  = MONTHS.map((_, mi) => monthFml(`A${r1}`, mi + 1));
    return ['', ...mf, `=IFERROR(SUM(B${r1}:M${r1}),0)`, '', '', ''];
  });

  // Combine into full row data
  const fullRows = SECS.map(([ticker, , freq, shares, annualDPS, cost], i) => {
    const r1  = GD1 + i;
    const yoc = annualDPS / cost;
    const mf  = MONTHS.map((_, mi) => monthFml(`A${r1}`, mi + 1));
    return [ticker, ...mf, `=IFERROR(SUM(B${r1}:M${r1}),0)`, freq, yoc, shares];
  });
  vals.push({ range: `${S}!A${GD1}`, values: fullRows });

  // ── Total row ─────────────────────────────────────────────────────────────────
  const totMonths = MONTHS.map((_, mi) => `=IFERROR(SUM(${cl(mi + 1)}${GD1}:${cl(mi + 1)}${GDN}),0)`);
  vals.push({ range: `${S}!A${GTOT1}`, values: [['MONTHLY TOTAL', ...totMonths,
    `=IFERROR(SUM(N${GD1}:N${GDN}),0)`, '', '', '']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_GTOT, R_GTOT + 1, 0, 17),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      numberFormat: { type: 'NUMBER', pattern: '$#,##0.00' } } },
    fields: 'userEnteredFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_GTOT, R_GTOT + 1, 0, 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '@' } } },
    fields: 'userEnteredFormat.numberFormat' } });

  // ── Data row formatting ───────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, R_GD0, R_GTOT, 0, 17),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  for (let i = 0; i < NS; i += 2) {
    fmt.push({ repeatCell: { range: gridRange(SID, R_GD0 + i, R_GD0 + i + 1, 0, 17),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor' } });
  }
  // Currency: month cols B-M + annual total N
  fmt.push({ repeatCell: { range: gridRange(SID, R_GD0, R_GTOT, 1, 14),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$#,##0.00' },
      horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat' } });
  // YOC % col P (index 15)
  fmt.push({ repeatCell: { range: gridRange(SID, R_GD0, R_GTOT, 15, 16),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });
  // Bold ticker col
  fmt.push({ repeatCell: { range: gridRange(SID, R_GD0, R_GTOT, 0, 1),
    cell: { userEnteredFormat: { textFormat: { bold: true } } },
    fields: 'userEnteredFormat.textFormat' } });

  // Highlight months with payments > 0 (conditional format: green background)
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_GD0, R_GTOT, 1, 13)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.01' }] },
      format: { backgroundColor: hex(C.success) } }
  }, index: 0 } });

  // ── Second section: Payment Timing Schedule ───────────────────────────────────
  const R_SCHED_HDR  = R_GTOT + 2;
  const R_SCHED_COLS = R_SCHED_HDR + 1;
  const R_SCHED_DATA = R_SCHED_COLS + 1;

  fmt.push({ mergeCells: { range: gridRange(SID, R_SCHED_HDR, R_SCHED_HDR + 1, 0, 8), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_SCHED_HDR + 1}`, values: [['PAYMENT SCHEDULE']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_SCHED_HDR, R_SCHED_HDR + 1, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_SCHED_HDR, endIndex: R_SCHED_HDR + 1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  vals.push({ range: `${S}!A${R_SCHED_COLS + 1}`,
    values: [['Ticker','Security Name','Frequency','Q1 Months','Q2 Months','Q3 Months','Q4 Months','Pay Day']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_SCHED_COLS, R_SCHED_COLS + 1, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });

  const SCHEDULE_DATA = [
    ['BND',  'Vanguard Total Bond Market ETF',  'Monthly',   'Jan / Feb / Mar','Apr / May / Jun','Jul / Aug / Sep','Oct / Nov / Dec','~15th'],
    ['VTI',  'Vanguard Total Stock Market ETF', 'Quarterly', 'Mar',            'Jun',            'Sep',            'Dec',            '~20th'],
    ['SCHD', 'Schwab US Dividend Equity ETF',   'Quarterly', 'Mar',            'Jun',            'Sep',            'Dec',            '~20th'],
    ['VNQ',  'Vanguard Real Estate ETF',        'Quarterly', 'Mar',            'Jun',            'Sep',            'Dec',            '~20th'],
    ['AMT',  'American Tower Corp.',            'Quarterly', 'Feb',            'May',            'Aug',            'Nov',            '~20th'],
    ['SPG',  'Simon Property Group Inc.',       'Quarterly', 'Jan',            'Apr',            'Jul',            'Oct',            '~20th'],
  ];
  vals.push({ range: `${S}!A${R_SCHED_DATA + 1}`, values: SCHEDULE_DATA });
  fmt.push({ repeatCell: { range: gridRange(SID, R_SCHED_DATA, R_SCHED_DATA + NS, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  for (let i = 0; i < NS; i += 2) {
    fmt.push({ repeatCell: { range: gridRange(SID, R_SCHED_DATA + i, R_SCHED_DATA + i + 1, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor' } });
  }
  fmt.push({ repeatCell: { range: gridRange(SID, R_SCHED_DATA, R_SCHED_DATA + NS, 0, 1),
    cell: { userEnteredFormat: { textFormat: { bold: true } } },
    fields: 'userEnteredFormat.textFormat' } });

  // ── Freeze ────────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID,
    gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'divcal-fmt');
  await valuesBatchUpdate(id, vals, 'divcal-vals');

  // ── Charts ────────────────────────────────────────────────────────────────────
  const charts = [];

  // 1. Monthly totals COLUMN chart
  charts.push({ addChart: { chart: {
    spec: { title: 'Monthly Dividend Totals', basicChart: {
      chartType: 'COLUMN', legendPosition: 'NO_LEGEND',
      axis: [{ position: 'BOTTOM_AXIS', title: 'Month' },
             { position: 'LEFT_AXIS',   title: 'Dividends ($)' }],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_GCOLS, R_GCOLS + 1, 1, 13)] } } }],
      series: [{ series: { sourceRange: { sources: [gridRange(SID, R_GTOT, R_GTOT + 1, 1, 13)] } },
        targetAxis: 'LEFT_AXIS', color: hex(C.success) }],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_TITLE, columnIndex: 17 },
      widthPixels: 480, heightPixels: 260,
    } },
  } } });

  // 2. Annual dividends by security — PIE
  charts.push({ addChart: { chart: {
    spec: { title: 'Annual Income by Security', pieChart: {
      legendPosition: 'RIGHT_LEGEND',
      pieHole: 0.4,
      domain: { sourceRange: { sources: [gridRange(SID, R_GCOLS, R_GTOT, 0, 1)] } },
      series: { sourceRange: { sources: [gridRange(SID, R_GCOLS, R_GTOT, 13, 14)] } },
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_GHDR + 8, columnIndex: 17 },
      widthPixels: 380, heightPixels: 240,
    } },
  } } });

  // 3. Yield on Cost by security — BAR
  charts.push({ addChart: { chart: {
    spec: { title: 'Yield on Cost by Security', basicChart: {
      chartType: 'BAR', legendPosition: 'NO_LEGEND',
      axis: [{ position: 'LEFT_AXIS',   title: 'Security' },
             { position: 'BOTTOM_AXIS', title: 'Yield on Cost %' }],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_GCOLS, R_GTOT, 0, 1)] } } }],
      series: [{ series: { sourceRange: { sources: [gridRange(SID, R_GCOLS, R_GTOT, 15, 16)] } },
        targetAxis: 'BOTTOM_AXIS', color: hex(C.hdrC) }],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_SCHED_HDR, columnIndex: 8 },
      widthPixels: 380, heightPixels: 240,
    } },
  } } });

  await batchUpdate(id, charts, 'divcal-charts');
  console.log('✓ Dividend Calendar complete');
})();

function cl(i) { return String.fromCharCode(65 + i); }
