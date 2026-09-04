'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Watchlist'];
const S   = "'Watchlist'";
const PRC = "'Price Updates'";
const REF = "'Reference Data'";

// ─── Row constants (0-indexed) ────────────────────────────────────────────────
const R_TITLE   = 0;
const R_NOTE    = 2;
const R_CARDHDR = 3;
const R_CARD    = 4;
const R_WHDR    = 6;
const R_WCOLS   = 7;
const R_WD0     = 8;
const NW        = 25;
const R_WD_END  = R_WD0 + NW - 1;  // 32

// 1-indexed
const WD1  = R_WD0 + 1;   // 9
const WDN  = R_WD_END + 1; // 33

// ─── Watchlist entries ────────────────────────────────────────────────────────
// [ticker, status, risk, targetPrice, targetAlloc%, positionSize$, priority, dateAdded, targetDate, notes]
const ENTRIES = [
  ['AVGO', 'Researching',     'High',        195.00, '',  '',       'High',  '1/15/2026','6/30/2026', 'Semiconductor / AI exposure'],
  ['MSCI', 'Monitoring',      'Moderate',    580.00, '',  '',       'Medium','9/1/2025', '12/31/2026','Index business moat — high quality'],
  ['O',    'Monitoring',      'Low',          50.00, 0.5, 5000,     'Medium','3/1/2025', '6/30/2026', 'Monthly dividend REIT'],
  ['MAIN', 'Ready for Review','Moderate',     45.00, 0.5, 5000,     'High',  '11/1/2025','3/31/2026', 'Monthly div, BDC sector'],
  ['JEPI', 'Researching',     'Moderate',     56.00, '',  '',       'Medium','12/1/2025','6/30/2026', 'Covered call income strategy'],
  ['VYMI', 'Monitoring',      'Low',          68.00, '',  '',       'Low',   '6/1/2025', '12/31/2026','International dividend diversification'],
  ['DVY',  'Ready for Review','Low',         118.00, 1.0, 10000,    'Medium','10/1/2025','3/31/2026', 'High dividend value ETF'],
  ['ABBV', 'Monitoring',      'Moderate',    150.00, '',  '',       'Medium','2/1/2025', '9/30/2026', 'Pharma dividend grower, ~4% yield'],
  ['QYLD', 'Researching',     'High',         16.50, '',  '',       'Low',   '1/1/2026', '6/30/2026', 'High-yield covered call ETF'],
  ['BX',   'Researching',     'High',        135.00, '',  '',       'Medium','1/15/2026','12/31/2026','Private equity exposure'],
  ['KKR',  'Researching',     'High',        115.00, '',  '',       'Low',   '2/1/2026', '12/31/2026','Alternative asset management'],
  ['MCD',  'Waiting for Price','Low',        270.00, '',  '',       'Low',   '8/1/2025', '12/31/2026','Dividend aristocrat — wait for pullback'],
  ['LOW',  'Waiting for Price','Low',        220.00, '',  '',       'Low',   '7/1/2025', '6/30/2027', 'Home improvement retailer'],
  ['COST', 'Monitoring',      'Low',         850.00, '',  '',       'Low',   '11/1/2024','12/31/2027','Consumer staples compounder'],
  ['AMZN', 'Researching',     'Moderate',    195.00, '',  '',       'Medium','1/1/2026', '6/30/2026', 'Cloud + retail long-term growth'],
  ['META', 'Monitoring',      'Moderate',    550.00, '',  '',       'Medium','10/1/2025','12/31/2026','AI-driven revenue growth'],
  ['AVLQ', 'Ready for Review','Low',          28.00, 2.0, 20000,    'Medium','12/15/2025','2/28/2026','Value factor ETF tilt'],
  ['DIVO', 'Ready for Review','Low',          40.00, 1.5, 15000,    'High',  '1/15/2026','2/28/2026', 'Active dividend-focused ETF'],
  ['BNDX', 'Monitoring',      'Low',          48.00, '',  '',       'Low',   '6/1/2025', '12/31/2026','International bond diversification'],
  ['TIPS', 'Monitoring',      'Low',         107.00, '',  '',       'Low',   '5/1/2025', '12/31/2026','Inflation-protected treasuries'],
  ['WPC',  'Monitoring',      'Moderate',     55.00, '',  '',       'Medium','9/1/2025', '6/30/2026', 'Net lease REIT diversifier'],
  ['VYM',  'Purchased',       'Low',         115.00, '',  '',       'Medium','3/1/2024', '6/1/2025',  'Position opened — moved to Holdings'],
  ['HASI', 'Researching',     'High',         25.00, '',  '',       'Low',   '1/15/2026','12/31/2026','Sustainable infrastructure finance'],
  ['MLPA', 'Researching',     'High',         36.00, '',  '',       'Low',   '2/1/2026', '12/31/2027','Midstream energy MLP income'],
  ['NUSI', 'Researching',     'High',         18.00, '',  '',       'Low',   '2/1/2026', '6/30/2026', 'Risk-managed income ETF'],
];

const STATUSES   = ['Researching','Monitoring','Waiting for Price','Ready for Review','Purchased','Removed'];
const RISK_LVLS  = ['Low','Moderate','High','Speculative','Not Assessed'];

(async () => {
  const fmt  = [];
  const vals = [];

  // ── Background ────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 500, 0, 22),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const COL_W = [30, 72, 170, 110, 120, 100, 95, 90, 90, 80, 75, 90, 75, 90, 90, 195];
  COL_W.forEach((px, ci) => fmt.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
    properties: { pixelSize: px }, fields: 'pixelSize' } }));

  // ── Title banner ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 2, 0, 16), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['WATCHLIST\nSecurities under consideration • Price targets • Risk assessment']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 2, 0, 16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' } });

  // ── Note ──────────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_NOTE, R_NOTE + 1, 0, 16), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_NOTE + 1}`, values: [['Current Price auto-populates from Price Updates tab for tracked securities. Target Price, Allocation %, and Position Size are manually entered. Upside % compares target vs. current.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_NOTE, R_NOTE + 1, 0, 16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.info),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.text) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });

  // ── Summary cards ─────────────────────────────────────────────────────────────
  const cDefs = [
    { label: 'RESEARCHING',    val: `=IFERROR(COUNTIF($F$${WD1}:$F$${WDN},"Researching"),0)`,      nf:'0', color:C.info,      tcol:C.text  },
    { label: 'READY TO BUY',   val: `=IFERROR(COUNTIF($F$${WD1}:$F$${WDN},"Ready for Review"),0)`, nf:'0', color:C.success,   tcol:C.text  },
    { label: 'MONITORING',     val: `=IFERROR(COUNTIF($F$${WD1}:$F$${WDN},"Monitoring"),0)`,       nf:'0', color:C.highlight, tcol:C.text  },
    { label: 'HIGH PRIORITY',  val: `=IFERROR(COUNTIFS($M$${WD1}:$M$${WDN},"High",$F$${WD1}:$F$${WDN},"<>Purchased",$F$${WD1}:$F$${WDN},"<>Removed"),0)`, nf:'0', color:C.warning, tcol:C.text },
  ];
  const cCols = [[0,4],[4,8],[8,12],[12,16]];
  cDefs.forEach(({ label, val, nf, color, tcol }, i) => {
    const [c1, c2] = cCols[i];
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARDHDR, R_CARDHDR + 1, c1, c2), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${cl(c1)}${R_CARDHDR + 1}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARDHDR, R_CARDHDR + 1, c1, c2),
      cell: { userEnteredFormat: { backgroundColor: hex(color),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(tcol), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
      fields: 'userEnteredFormat' } });
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARD, R_CARD + 1, c1, c2), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${cl(c1)}${R_CARD + 1}`, values: [[val]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARD, R_CARD + 1, c1, c2),
      cell: { userEnteredFormat: { backgroundColor: hex(color),
        textFormat: { bold: true, fontSize: 22, foregroundColor: hex(tcol), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'TOP',
        numberFormat: { type: 'NUMBER', pattern: nf } } },
      fields: 'userEnteredFormat' } });
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARDHDR, endIndex: R_CARDHDR + 1 },
    properties: { pixelSize: 20 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_CARD, endIndex: R_CARD + 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  // ── Watchlist section header ───────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_WHDR, R_WHDR + 1, 0, 16), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_WHDR + 1}`, values: [['SECURITIES UNDER CONSIDERATION']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_WHDR, R_WHDR + 1, 0, 16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_WHDR, endIndex: R_WHDR + 1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // ── Column headers ────────────────────────────────────────────────────────────
  const HDRS = ['#','Ticker','Security Name','Type','Asset Class','Status','Risk Level',
                'Target Price','Current Price','Upside %','Target Alloc%','Position Size $',
                'Priority','Date Added','Target Date','Notes'];
  vals.push({ range: `${S}!A${R_WCOLS + 1}`, values: [HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_WCOLS, R_WCOLS + 1, 0, 16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });

  // ── Watchlist data ────────────────────────────────────────────────────────────
  const wRows = ENTRIES.map(([ticker, status, risk, targetPx, alloc, posSize, priority, dateAdded, targetDate, notes], i) => {
    const r1       = WD1 + i;
    const numFml   = i + 1;
    const nameFml  = `=IFERROR(VLOOKUP(B${r1},${PRC}!$B$6:$C$1005,2,FALSE),"—")`;
    const typeFml  = `=IFERROR(VLOOKUP(B${r1},${PRC}!$B$6:$D$1005,3,FALSE),"—")`;
    const aclsFml  = `=IFERROR(VLOOKUP(B${r1},${PRC}!$B$6:$E$1005,4,FALSE),"—")`;
    const curPxFml = `=IFERROR(VLOOKUP(B${r1},${PRC}!$B$6:$G$1005,6,FALSE),"—")`;
    const upsideFml= `=IFERROR(IF(OR(H${r1}="",I${r1}="","—"=I${r1},"—"=H${r1}),"—",(H${r1}/I${r1})-1),"—")`;
    return [numFml, ticker, nameFml, typeFml, aclsFml, status, risk,
            targetPx, curPxFml, upsideFml, alloc, posSize, priority, dateAdded, targetDate, notes];
  });
  vals.push({ range: `${S}!A${WD1}`, values: wRows });

  // ── Row formatting ────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, R_WD0, R_WD_END + 1, 0, 16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  for (let i = 0; i < NW; i += 2) {
    fmt.push({ repeatCell: { range: gridRange(SID, R_WD0 + i, R_WD0 + i + 1, 0, 16),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor' } });
  }
  // Number formats
  fmt.push({ repeatCell: { range: gridRange(SID, R_WD0, R_WD_END + 1, 7, 9),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$#,##0.00' }, horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_WD0, R_WD_END + 1, 9, 10),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '+0.0%;-0.0%;—' }, horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_WD0, R_WD_END + 1, 10, 11),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.0%' }, horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_WD0, R_WD_END + 1, 11, 12),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '$#,##0' }, horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_WD0, R_WD_END + 1, 13, 15),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });
  // Row # bold/center
  fmt.push({ repeatCell: { range: gridRange(SID, R_WD0, R_WD_END + 1, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER',
      textFormat: { bold: true, foregroundColor: hex(C.secText) } } },
    fields: 'userEnteredFormat' } });
  // Ticker bold
  fmt.push({ repeatCell: { range: gridRange(SID, R_WD0, R_WD_END + 1, 1, 2),
    cell: { userEnteredFormat: { textFormat: { bold: true } } },
    fields: 'userEnteredFormat.textFormat' } });

  // ── Data validation ───────────────────────────────────────────────────────────
  fmt.push({ setDataValidation: { range: gridRange(SID, R_WD0, R_WD_END + 1, 5, 6),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$105:$A$110` }] },
      showCustomUi: true, strict: true } } });
  fmt.push({ setDataValidation: { range: gridRange(SID, R_WD0, R_WD_END + 1, 6, 7),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$113:$A$117` }] },
      showCustomUi: true, strict: true } } });
  fmt.push({ setDataValidation: { range: gridRange(SID, R_WD0, R_WD_END + 1, 12, 13),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$186:$A$189` }] },
      showCustomUi: true, strict: true } } });

  // ── Status conditional formats ────────────────────────────────────────────────
  const statColors = [
    ['Researching',     C.info],
    ['Ready for Review',C.success],
    ['Monitoring',      C.highlight],
    ['Waiting for Price',C.warning],
    ['Purchased',       C.altRow],
    ['Removed',         C.border],
  ];
  statColors.forEach(([text, bg], idx) => {
    fmt.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, R_WD0, R_WD_END + 1, 0, 16)],
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$F${WD1}="${text}"` }] },
        format: { backgroundColor: hex(bg) } }
    }, index: idx } });
  });
  // Upside %: positive (current below target) → green text
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_WD0, R_WD_END + 1, 9, 10)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
      format: { textFormat: { foregroundColor: hex(C.hdrA), bold: true } } }
  }, index: statColors.length } });
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_WD0, R_WD_END + 1, 9, 10)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
      format: { textFormat: { foregroundColor: hex(C.attention), bold: true } } }
  }, index: statColors.length + 1 } });

  // ── Freeze ────────────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID,
    gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'watchlist-fmt');
  await valuesBatchUpdate(id, vals, 'watchlist-vals');

  // ── Summary tables + charts ───────────────────────────────────────────────────
  const sfmt = [];
  const svals = [];

  // Status breakdown (cols R-S, 0-indexed 17-18)
  const R_STAT_HDR  = R_WHDR;
  const R_STAT_COLS = R_WCOLS;
  const R_STAT_DATA = R_WD0;

  svals.push({ range: `${S}!R${R_STAT_HDR + 1}`, values: [['STATUS BREAKDOWN']] });
  sfmt.push({ mergeCells: { range: gridRange(SID, R_STAT_HDR, R_STAT_HDR + 1, 17, 19), mergeType: 'MERGE_ALL' } });
  sfmt.push({ repeatCell: { range: gridRange(SID, R_STAT_HDR, R_STAT_HDR + 1, 17, 19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });

  svals.push({ range: `${S}!R${R_STAT_COLS + 1}`, values: [['Status', 'Count']] });
  sfmt.push({ repeatCell: { range: gridRange(SID, R_STAT_COLS, R_STAT_COLS + 1, 17, 19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });

  const statRows = STATUSES.map(st => [st, `=IFERROR(COUNTIF($F$${WD1}:$F$${WDN},"${st}"),0)`]);
  svals.push({ range: `${S}!R${R_STAT_DATA + 1}`, values: statRows });
  sfmt.push({ repeatCell: { range: gridRange(SID, R_STAT_DATA, R_STAT_DATA + STATUSES.length, 17, 19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' } } },
    fields: 'userEnteredFormat' } });

  // Risk breakdown below
  const R_RISK_HDR  = R_STAT_DATA + STATUSES.length + 1;
  const R_RISK_COLS = R_RISK_HDR + 1;
  const R_RISK_DATA = R_RISK_COLS + 1;

  svals.push({ range: `${S}!R${R_RISK_HDR + 1}`, values: [['RISK BREAKDOWN']] });
  sfmt.push({ mergeCells: { range: gridRange(SID, R_RISK_HDR, R_RISK_HDR + 1, 17, 19), mergeType: 'MERGE_ALL' } });
  sfmt.push({ repeatCell: { range: gridRange(SID, R_RISK_HDR, R_RISK_HDR + 1, 17, 19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrC),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });

  svals.push({ range: `${S}!R${R_RISK_COLS + 1}`, values: [['Risk Level', 'Count']] });
  sfmt.push({ repeatCell: { range: gridRange(SID, R_RISK_COLS, R_RISK_COLS + 1, 17, 19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });

  const riskRows = RISK_LVLS.map(rl => [rl, `=IFERROR(COUNTIF($G$${WD1}:$G$${WDN},"${rl}"),0)`]);
  svals.push({ range: `${S}!R${R_RISK_DATA + 1}`, values: riskRows });
  sfmt.push({ repeatCell: { range: gridRange(SID, R_RISK_DATA, R_RISK_DATA + RISK_LVLS.length, 17, 19),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' } } },
    fields: 'userEnteredFormat' } });

  // Set col widths for R-S
  sfmt.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 17, endIndex: 19 },
    properties: { pixelSize: 130 }, fields: 'pixelSize' } });

  await batchUpdate(id, sfmt, 'watchlist-sfmt');
  await valuesBatchUpdate(id, svals, 'watchlist-svals');

  // Charts
  const charts = [];
  const STAT_RANGE_END = R_STAT_DATA + STATUSES.length;
  const RISK_RANGE_END = R_RISK_DATA + RISK_LVLS.length;

  // 1. Status breakdown — horizontal BAR
  charts.push({ addChart: { chart: {
    spec: { title: 'Watchlist by Status', basicChart: {
      chartType: 'BAR', legendPosition: 'NO_LEGEND',
      axis: [{ position: 'LEFT_AXIS', title: 'Status' },
             { position: 'BOTTOM_AXIS', title: 'Count' }],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_STAT_COLS, STAT_RANGE_END, 17, 18)] } } }],
      series: [{ series: { sourceRange: { sources: [gridRange(SID, R_STAT_COLS, STAT_RANGE_END, 18, 19)] } },
        targetAxis: 'BOTTOM_AXIS', color: hex(C.primary) }],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_CARDHDR, columnIndex: 16 },
      widthPixels: 320, heightPixels: 200,
    } },
  } } });

  // 2. Risk level breakdown — PIE
  charts.push({ addChart: { chart: {
    spec: { title: 'Watchlist by Risk Level', pieChart: {
      legendPosition: 'RIGHT_LEGEND', pieHole: 0.4,
      domain: { sourceRange: { sources: [gridRange(SID, R_RISK_COLS, RISK_RANGE_END, 17, 18)] } },
      series: { sourceRange: { sources: [gridRange(SID, R_RISK_COLS, RISK_RANGE_END, 18, 19)] } },
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_WHDR, columnIndex: 16 },
      widthPixels: 320, heightPixels: 220,
    } },
  } } });

  await batchUpdate(id, charts, 'watchlist-charts');
  console.log('✓ Watchlist complete');
})();

function cl(i) { return String.fromCharCode(65 + i); }
