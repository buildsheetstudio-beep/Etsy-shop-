'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Goals & Milestones'];
const S   = "'Goals & Milestones'";
const REF = "'Reference Data'";

// ─── Row constants (0-indexed) ────────────────────────────────────────────────
const R_TITLE   = 0;
const R_NOTE    = 2;
const R_CARDHDR = 3;
const R_CARD    = 4;
const R_GHDR    = 6;
const R_GCOLS   = 7;
const R_GD0     = 8;
const NG        = 20;
const R_GD_END  = R_GD0 + NG - 1;  // 27

// 1-indexed
const GD1 = R_GD0 + 1;   // 9
const GDN = R_GD_END + 1; // 28

// ─── Sample goals ─────────────────────────────────────────────────────────────
// [Name, Type, Owner, Status, Priority, Target, Current, StartDate, TargetDate, Notes]
const GOALS = [
  ['Reach $1M Portfolio Value',     'Portfolio Value',        'Joint Household', 'In Progress','High',   1000000, `=IFERROR(SUM('Holdings'!$O$6:$O$1005),0)`, '1/1/2024',  '12/31/2026','Primary 2026 milestone'],
  ['Max Daniel 401(k) 2026',        'Annual Contribution',   'Daniel Walsh',    'In Progress','High',    23000,   12600,                                          '1/1/2026',  '12/31/2026','$600/pay period'],
  ['Max Daniel Roth IRA 2026',      'Annual Contribution',   'Daniel Walsh',    'In Progress','High',     7000,    3500,                                           '1/1/2026',  '12/31/2026','$583/month'],
  ['Max Emily Roth IRA 2026',       'Annual Contribution',   'Emily Walsh',     'In Progress','High',     7000,    3500,                                           '1/1/2026',  '12/31/2026','$583/month'],
  ['Max Emily 403(b) 2026',         'Annual Contribution',   'Emily Walsh',     'In Progress','High',    23000,    9800,                                           '1/1/2026',  '12/31/2026','University plan'],
  ['$36K Household Contributions',  'Annual Contribution',   'Joint Household', 'In Progress','High',    36000,   16200,                                           '1/1/2026',  '12/31/2026','Combined all accounts'],
  ['$50K Emergency Fund',           'Emergency Savings',     'Joint Household', 'Achieved',   'High',    50000,   50000,                                           '1/1/2022',  '3/31/2024', 'Fully funded — ACHIEVED'],
  ['Reach $1M Net Worth',           'Net Worth',             'Joint Household', 'In Progress','High',  1000000,   `=IFERROR('Net Worth Tracker'!$E$28,0)`,        '1/1/2024',  '12/31/2028',''],
  ['$25K Annual Dividend Income',   'Dividend Income',       'Joint Household', 'In Progress','Medium',  25000,    8400,                                           '1/1/2025',  '12/31/2030',''],
  ["Pay Off Emily's Student Loans", 'Net Worth',             'Emily Walsh',     'In Progress','Medium',  18500,    4300,                                           '9/1/2024',  '6/30/2028', '$4,300 paid so far'],
  ["Child's 529 College Fund",      'Education Funding',     'Joint Household', 'In Progress','Medium', 200000,   28500,                                           '6/1/2020',  '9/1/2040',  'Target $200K by 2040'],
  ['Reach $2M Portfolio (2050)',    'Portfolio Value',        'Joint Household', 'In Progress','Medium',2000000,   `=IFERROR(SUM('Holdings'!$O$6:$O$1005),0)`,    '1/1/2025',  '12/31/2050','Long-term milestone'],
  ['FI Number: $2.5M Portfolio',   'Financial Independence', 'Joint Household', 'In Progress','Medium',2500000,   `=IFERROR(SUM('Holdings'!$O$6:$O$1005),0)`,    '1/1/2025',  '12/31/2055','25x annual spend'],
  ['15% International Allocation',  'Asset Allocation',      'Joint Household', 'In Progress','Low',     15,       8.5,                                            '1/1/2026',  '12/31/2027','% target — currently underweight'],
  ['5% Crypto Allocation',          'Asset Allocation',      'Joint Household', 'Not Started','Low',      5,       1.8,                                            '1/1/2027',  '12/31/2028','% target — research phase'],
  ['Mortgage Extra Paydown',        'Home Purchase',          'Joint Household', 'In Progress','Low',    73000,   33000,                                            '6/1/2021',  '12/31/2035','Extra $250/month principal'],
  ['Investment Property D/P',       'Home Purchase',          'Joint Household', 'Not Started','Low',   100000,       0,                                           '1/1/2027',  '12/31/2029','$100K down payment fund'],
  ['Daniel HSA Fully Invested',     'Other',                  'Daniel Walsh',    'Achieved',  'Medium',      1,       1,                                            '1/1/2025',  '12/31/2025','100% invested in VTI — DONE'],
  ['Review All Beneficiaries',      'Other',                  'Joint Household', 'Achieved',  'High',        1,       1,                                            '1/1/2025',  '12/31/2025','All accounts reviewed — DONE'],
  ['Capture Full Employer Matches', 'Annual Contribution',    'Joint Household', 'In Progress','High',       1,    0.95,                                            '1/1/2025',  '12/31/2026','Daniel 4%, Emily 5%'],
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Background ────────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 500, 0, 16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const COL_W = [30, 220, 150, 120, 100, 80, 120, 120, 85, 95, 95, 200];
  COL_W.forEach((px, ci) => fmt.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
    properties: { pixelSize: px }, fields: 'pixelSize' } }));

  // ── Title banner ──────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 2, 0, 12), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A1`, values: [['GOALS & MILESTONES\nFinancial targets • Progress tracking • Priority management']] });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 2, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 46 }, fields: 'pixelSize' } });

  // ── Note row ──────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_NOTE, R_NOTE + 1, 0, 12), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_NOTE + 1}`, values: [['Goals 1–6 reference live Holdings values. Asset Allocation targets use percentage points (e.g. 15 = 15%). Boolean goals (rows 18–20) use 0/1 scale. Add your own goals in rows below 28.']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_NOTE, R_NOTE + 1, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.info),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.text) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });

  // ── Summary cards ─────────────────────────────────────────────────────────────
  const cardDefs = [
    { label: 'GOALS ACHIEVED',      val: `=IFERROR(COUNTIF($E$${GD1}:$E$${GDN},"Achieved"),0)`,                                  fmt: '0',    color: C.success,   tcol: C.text },
    { label: 'IN PROGRESS',         val: `=IFERROR(COUNTIF($E$${GD1}:$E$${GDN},"In Progress"),0)`,                               fmt: '0',    color: C.info,      tcol: C.text },
    { label: 'HIGH PRIORITY OPEN',  val: `=IFERROR(COUNTIFS($F$${GD1}:$F$${GDN},"High",$E$${GD1}:$E$${GDN},"<>Achieved"),0)`, fmt: '0',    color: C.warning,   tcol: C.text },
    { label: 'AVG PROGRESS (ACTIVE)',val: `=IFERROR(AVERAGEIF($E$${GD1}:$E$${GDN},"In Progress",$I$${GD1}:$I$${GDN}),0)`,      fmt: '0%',   color: C.highlight, tcol: C.text },
  ];
  const cardCols = [[0,3],[3,6],[6,9],[9,12]];
  cardDefs.forEach(({ label, val, fmt: nf, color, tcol }, i) => {
    const [c1, c2] = cardCols[i];
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARDHDR, R_CARDHDR + 1, c1, c2), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${col(c1)}${R_CARDHDR + 1}`, values: [[label]] });
    fmt.push({ repeatCell: { range: gridRange(SID, R_CARDHDR, R_CARDHDR + 1, c1, c2),
      cell: { userEnteredFormat: { backgroundColor: hex(color),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(tcol), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
      fields: 'userEnteredFormat' } });
    fmt.push({ mergeCells: { range: gridRange(SID, R_CARD, R_CARD + 1, c1, c2), mergeType: 'MERGE_ALL' } });
    vals.push({ range: `${S}!${col(c1)}${R_CARD + 1}`, values: [[val]] });
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

  // ── Goals section header ──────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, R_GHDR, R_GHDR + 1, 0, 12), mergeType: 'MERGE_ALL' } });
  vals.push({ range: `${S}!A${R_GHDR + 1}`, values: [['FINANCIAL GOALS']] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_GHDR, R_GHDR + 1, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrA),
      textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: R_GHDR, endIndex: R_GHDR + 1 },
    properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // ── Column headers ────────────────────────────────────────────────────────────
  const HDRS = ['#','Goal Name','Goal Type','Owner','Status','Priority',
                'Target Value','Current Value','Progress %','Start Date','Target Date','Notes'];
  vals.push({ range: `${S}!A${R_GCOLS + 1}`, values: [HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID, R_GCOLS, R_GCOLS + 1, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.hdrB),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primaryText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });

  // ── Goal data rows ────────────────────────────────────────────────────────────
  const goalRows = GOALS.map(([name, type, owner, status, priority, target, current, start, end, notes], i) => {
    const r1 = GD1 + i;
    const progFml = `=IFERROR(H${r1}/G${r1},0)`;
    return [i + 1, name, type, owner, status, priority, target, current, progFml, start, end, notes];
  });
  vals.push({ range: `${S}!A${GD1}`, values: goalRows });

  // ── Data row formatting ───────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, R_GD0, R_GD_END + 1, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.panel),
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat' } });
  for (let i = 0; i < NG; i += 2) {
    fmt.push({ repeatCell: { range: gridRange(SID, R_GD0 + i, R_GD0 + i + 1, 0, 12),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor' } });
  }
  // Number format: col # (center), Target+Current ($#,##0), Progress % (0%), dates (m/d/yyyy)
  fmt.push({ repeatCell: { range: gridRange(SID, R_GD0, R_GD_END + 1, 0, 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER',
      textFormat: { bold: true, foregroundColor: hex(C.secText) } } },
    fields: 'userEnteredFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_GD0, R_GD_END + 1, 6, 8),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0.##' },
      horizontalAlignment: 'RIGHT' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_GD0, R_GD_END + 1, 8, 9),
    cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0%' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, R_GD0, R_GD_END + 1, 9, 11),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' },
      horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat' } });

  // ── Data validation ───────────────────────────────────────────────────────────
  const dv = (range, refRange) => ({ setDataValidation: { range,
    rule: { condition: { type: 'ONE_OF_RANGE',
      values: [{ userEnteredValue: `=${refRange}` }] },
      showCustomUi: true, strict: true } } });

  fmt.push(dv(gridRange(SID, R_GD0, R_GD_END + 1, 2, 3), `${REF}!$A$84:$A$95`));  // Goal Types
  fmt.push(dv(gridRange(SID, R_GD0, R_GD_END + 1, 3, 4), `${REF}!$A$4:$A$6`));    // Owners
  fmt.push(dv(gridRange(SID, R_GD0, R_GD_END + 1, 4, 5), `${REF}!$A$98:$A$102`)); // Goal Statuses
  fmt.push(dv(gridRange(SID, R_GD0, R_GD_END + 1, 5, 6), `${REF}!$A$186:$A$189`));// Priority

  // ── Conditional formats ───────────────────────────────────────────────────────
  const statusColors = [
    ['Achieved',    C.success],
    ['In Progress', C.info],
    ['Delayed',     C.warning],
    ['Reassess',    C.attention],
    ['Not Started', C.altRow],
  ];
  statusColors.forEach(([text, bg], idx) => {
    fmt.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(SID, R_GD0, R_GD_END + 1, 0, 12)],
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA',
          values: [{ userEnteredValue: `=$E${GD1}="${text}"` }] },
        format: { backgroundColor: hex(bg) } }
    }, index: idx } });
  });

  // Progress % green when >= 100%
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_GD0, R_GD_END + 1, 8, 9)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1' }] },
      format: { textFormat: { bold: true, foregroundColor: hex(C.hdrA) } } }
  }, index: statusColors.length } });

  // Priority HIGH = bold goal name
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(SID, R_GD0, R_GD_END + 1, 1, 2)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA',
        values: [{ userEnteredValue: `=$F${GD1}="High"` }] },
      format: { textFormat: { bold: true } } }
  }, index: statusColors.length + 1 } });

  // ── Freeze rows ───────────────────────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID,
    gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'goals-fmt');
  await valuesBatchUpdate(id, vals, 'goals-vals');

  // ── Chart: Goal Status Breakdown (pie) ───────────────────────────────────────
  // Build a summary mini-table for the chart: col N = status label, col O = count
  const STATUSES = ['Not Started','In Progress','Achieved','Delayed','Reassess'];
  const summaryRange = `${S}!N${R_GHDR + 1}`;
  const summaryData = STATUSES.map(s => [s, `=IFERROR(COUNTIF($E$${GD1}:$E$${GDN},"${s}"),0)`]);
  await valuesBatchUpdate(id, [{ range: summaryRange, values: summaryData }], 'goals-summary');

  const pieStart = R_GHDR;  // 0-indexed
  const pieEnd   = pieStart + STATUSES.length; // exclusive

  const charts = [];
  charts.push({ addChart: { chart: {
    spec: { title: 'Goals by Status', pieChart: {
      legendPosition: 'RIGHT_LEGEND',
      pieHole: 0.4,
      domain: { sourceRange: { sources: [gridRange(SID, pieStart, pieEnd, 13, 14)] } },
      series: { sourceRange: { sources: [gridRange(SID, pieStart, pieEnd, 14, 15)] } },
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_CARDHDR, columnIndex: 12 },
      widthPixels: 340, heightPixels: 200,
    } },
  } } });

  // Progress bars chart: goal names vs progress %
  charts.push({ addChart: { chart: {
    spec: { title: 'Goal Progress', basicChart: {
      chartType: 'BAR',
      legendPosition: 'NO_LEGEND',
      axis: [
        { position: 'BOTTOM_AXIS', title: 'Progress %' },
        { position: 'LEFT_AXIS',   title: '' },
      ],
      domains: [{ domain: { sourceRange: { sources: [gridRange(SID, R_GCOLS, R_GD_END + 1, 1, 2)] } } }],
      series: [{ series: { sourceRange: { sources: [gridRange(SID, R_GCOLS, R_GD_END + 1, 8, 9)] } },
        targetAxis: 'BOTTOM_AXIS' }],
      headerCount: 1,
    } },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: R_GHDR, columnIndex: 12 },
      widthPixels: 380, heightPixels: 400,
    } },
  } } });

  await batchUpdate(id, charts, 'goals-charts');
  console.log('✓ Goals & Milestones complete');
})();

function col(i) { return String.fromCharCode(65 + i); }
