'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Progress & Performance'];
const S = "'Progress & Performance'";
const LOG = "'Daily Workout Log'";

// 12 months of sample body weight data (Oct 2025 – Sep 2026)
const BODY_METRICS = [
  ['2025-10-01', 185, 18.5, 'Starting weight'],
  ['2025-11-01', 183, 18.0, 'Down 2 lbs'],
  ['2025-12-01', 182, 17.8, ''],
  ['2026-01-01', 180, 17.2, 'New Year — feeling strong'],
  ['2026-02-01', 179, 17.0, ''],
  ['2026-03-01', 178, 16.5, 'Hitting PRs'],
  ['2026-04-01', 177, 16.2, ''],
  ['2026-05-01', 176, 15.8, ''],
  ['2026-06-01', 175, 15.5, 'Down 10 lbs total'],
  ['2026-07-01', 176, 15.6, 'Slight rebound — normal'],
  ['2026-08-01', 175, 15.3, ''],
  ['2026-09-01', 174, 15.0, '1-year progress: -11 lbs, -3.5% BF'],
];

// Top exercises to track performance for
const TRACK_EXS = [
  ['EX-0001','Bench Press'],
  ['EX-0007','Deadlift'],
  ['EX-0036','Back Squat'],
  ['EX-0014','Overhead Press'],
  ['EX-0008','Barbell Row'],
  ['EX-0020','Barbell Curl'],
  ['EX-0042','Romanian Deadlift'],
  ['EX-0046','Hip Thrust'],
  ['EX-0058','Outdoor Run'],
  ['EX-0076','Kettlebell Swing'],
];

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title (no frozenColumnCount — spans all 20 cols)
  vals.push({ range: `${S}!A1`, values: [['PROGRESS & PERFORMANCE ANALYTICS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Track your body metrics and monitor key exercise performance trends over time.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // ── SECTION 1: Body Metrics ────────────────────────────────────────────────
  vals.push({ range: `${S}!A3`, values: [['BODY METRICS TRACKER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,0,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Row 4: Body metrics headers
  vals.push({ range: `${S}!A4`, values: [['Date','Body Weight (lb)','Body Fat %','Notes','','Date','Body Weight (lb)','Body Fat %','Notes','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Rows 5-16: 12 months of body metrics (split across 2 groups of 6 for compactness)
  BODY_METRICS.forEach(([date, weight, bf, note], i) => {
    const r = 5 + i;
    // Format varies: left 6 rows in cols A-D, rows 7-12 could go in cols F-I, but simpler to just stack all 12 rows A-D
    vals.push({ range: `${S}!A${r}:D${r}`, values: [[date, weight, bf, note]] });
    const bg = i % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,4), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
      numberFormat: { type: 'DATE', pattern: 'mmm yyyy' },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,3), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
      numberFormat: { type: 'NUMBER', pattern: '0.0' },
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  });

  // Row 17: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 16, endIndex: 17 }, properties: { pixelSize: 16 }, fields: 'pixelSize' }});

  // ── SECTION 2: Exercise Performance Tracker ────────────────────────────────
  vals.push({ range: `${S}!A18`, values: [['KEY EXERCISE PERFORMANCE (Auto-Updated from Daily Log)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,17,18,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,17,18,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 18 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Row 19: Exercise performance headers
  vals.push({ range: `${S}!A19`, values: [['Exercise ID','Exercise Name','Total Sets','Best Weight (lb)','Best Reps','Best Duration (min)','Best Distance','Best PR Value','Sessions']] });
  fmt.push({ repeatCell: { range: gridRange(SID,18,19,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 18, endIndex: 19 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Rows 20-29: 10 tracked exercises
  TRACK_EXS.forEach(([exId, exName], i) => {
    const r = 20 + i;
    const ri = r - 1;
    vals.push({ range: `${S}!A${r}:B${r}`, values: [[exId, exName]] });
    // C: Total Sets
    vals.push({ range: `${S}!C${r}`, values: [[`=SUMPRODUCT((${LOG}!$K$8:$K$5000=A${r})*1)`]] });
    // D: Best Weight
    vals.push({ range: `${S}!D${r}`, values: [[`=IFERROR(MAXIFS(${LOG}!$Q$8:$Q$5000,${LOG}!$K$8:$K$5000,A${r}),"")`]] });
    // E: Best Reps
    vals.push({ range: `${S}!E${r}`, values: [[`=IFERROR(MAXIFS(${LOG}!$R$8:$R$5000,${LOG}!$K$8:$K$5000,A${r}),"")`]] });
    // F: Best Duration
    vals.push({ range: `${S}!F${r}`, values: [[`=IFERROR(MAXIFS(${LOG}!$S$8:$S$5000,${LOG}!$K$8:$K$5000,A${r}),"")`]] });
    // G: Best Distance
    vals.push({ range: `${S}!G${r}`, values: [[`=IFERROR(MAXIFS(${LOG}!$T$8:$T$5000,${LOG}!$K$8:$K$5000,A${r}),"")`]] });
    // H: Best PR Value
    vals.push({ range: `${S}!H${r}`, values: [[`=IFERROR(MAXIFS(${LOG}!$AC$8:$AC$5000,${LOG}!$K$8:$K$5000,A${r}),"")`]] });
    // I: Sessions (distinct dates for this exercise)
    vals.push({ range: `${S}!I${r}`, values: [[`=SUMPRODUCT((${LOG}!$K$8:$K$5000=A${r})*IFERROR(1/COUNTIF(${LOG}!$C$8:$C$5000,${LOG}!$C$8:$C$5000),0))`]] });

    const bg = i % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,ri,ri+1,0,9), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 30: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 29, endIndex: 30 }, properties: { pixelSize: 16 }, fields: 'pixelSize' }});

  // ── SECTION 3: Monthly Performance Summary ─────────────────────────────────
  vals.push({ range: `${S}!A31`, values: [['MONTHLY PERFORMANCE BREAKDOWN (All Years)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,30,31,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,30,31,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.accent), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 30, endIndex: 31 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  vals.push({ range: `${S}!A32`, values: [['Period','Workout Days','Total Sets','Est. Volume (lb)','Avg RPE']] });
  fmt.push({ repeatCell: { range: gridRange(SID,31,32,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 31, endIndex: 32 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // Q4 2025, Q1-Q3 2026 summary rows
  const QUARTERS = [
    ['Q4 2025 (Oct–Dec)', 2025, [10,11,12]],
    ['Q1 2026 (Jan–Mar)', 2026, [1,2,3]],
    ['Q2 2026 (Apr–Jun)', 2026, [4,5,6]],
    ['Q3 2026 (Jul–Sep)', 2026, [7,8,9]],
  ];
  QUARTERS.forEach(([label, yr, months], qi) => {
    const r = 33 + qi;
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    const moFilter = months.map(m => `(MONTH(${LOG}!$C$8:$C$5000)=${m})`).join('+');
    vals.push({ range: `${S}!B${r}`, values: [[`=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=${yr})*((${moFilter})>0)*(LEN(${LOG}!$B$8:$B$5000)>0)*IFERROR(1/COUNTIF(${LOG}!$C$8:$C$5000,${LOG}!$C$8:$C$5000),0))`]] });
    vals.push({ range: `${S}!C${r}`, values: [[`=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=${yr})*((${moFilter})>0)*(LEN(${LOG}!$B$8:$B$5000)>0)*1)`]] });
    vals.push({ range: `${S}!D${r}`, values: [[`=SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=${yr})*((${moFilter})>0)*IFERROR(${LOG}!$AB$8:$AB$5000*1,0))`]] });
    vals.push({ range: `${S}!E${r}`, values: [[`=IFERROR(SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=${yr})*((${moFilter})>0)*IFERROR(${LOG}!$W$8:$W$5000*1,0))/SUMPRODUCT((YEAR(${LOG}!$C$8:$C$5000)=${yr})*((${moFilter})>0)*(LEN(${LOG}!$W$8:$W$5000)>0)*1),"")`]] });
    const bg = qi % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,5), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 10, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});
  });

  // Column widths
  const WIDTHS = [90,170,80,110,100,110,90,90,80,120];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1:2
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 2 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '08-progress values');
  await batchUpdate(id, fmt, '08-progress format');
  console.log('✅  Progress & Performance done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
