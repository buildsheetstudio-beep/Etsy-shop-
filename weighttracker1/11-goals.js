'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Goals & Milestones'];
const S = "'Goals & Milestones'";

// 9 goals: mix of active, achieved, paused
// [name, type, status, start_date, target_date, start_val, target_val, current_val, notes]
const GOALS = [
  ['Walk 10,000 steps daily for 30 days', 'Habit',            'Achieved', '2024-11-01','2024-11-30', '',   '',   '',   'Completed challenge'],
  ['Log meals consistently for 90 days',  'Habit',            'Active',   '2026-06-01','2026-08-29', '',   '',   '',   '90-day streak goal'],
  ['Complete 100 workout sessions',        'Fitness',          'Active',   '2025-10-01','2026-09-30', 0,    100,  75,   'Tracking in Workout tab'],
  ['Waist measurement goal',              'Body Measurements', 'Active',   '2025-01-01','2026-12-31', 34.5, 31.0, 32.8, 'Monthly measurement'],
  ['Drink 8+ glasses water daily',        'Habit',            'Active',   '2026-01-01','2026-12-31', '',   '',   '',   'Habit HB-002 supports this'],
  ['30 consecutive days of yoga',         'Fitness',          'Paused',   '2026-04-01','2026-05-01', 0,    30,   12,   'Paused due to travel'],
  ['Read 12 books this year',             'Habit',            'Active',   '2026-01-01','2026-12-31', 0,    12,   7,    ''],
  ['Consistent sleep schedule',           'Habit',            'Active',   '2026-06-01','2026-09-30', '',   '',   '',   'In bed by 10:30pm'],
  ['Reduce sodium intake',                'Nutrition',        'Active',   '2026-07-01','2026-12-31', 2400, 1800, 2100, 'mg per day average'],
];

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5000,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['GOALS & MILESTONES']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,11), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Track your personal goals. Progress % is calculated automatically for measurable goals. Non-measurable goals show status only.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,11), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Stats
  const STATS = [
    ['Goals Tracked', `=COUNTA(B8:B500)`],
    ['Active Goals',  `=COUNTIF(D8:D500,"Active")`],
    ['Achieved',      `=COUNTIF(D8:D500,"Achieved")`],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,0,4), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,4,9), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!E${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,4), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,4,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 6: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Row 7: headers
  vals.push({ range: `${S}!A7`, values: [['Goal ID','Goal Name','Goal Type','Status','Start Date','Target Date','Start Value','Target Value','Current Value','Progress %','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // A: Goal ID formula
  fmt.push({ repeatCell: { range: gridRange(SID,7,500,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B8="","","GL-"&TEXT(ROW()-7,"000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // J: Progress % formula (neutral direction logic: handles both higher-is-better and lower-is-better)
  // If G8 > H8 (target < start = lower-is-better): progress = (G-I)/(G-H)
  // If H8 > G8 (target > start = higher-is-better): progress = (I-G)/(H-G)
  // If start/target are empty: show ""
  fmt.push({ repeatCell: { range: gridRange(SID,7,500,9,10), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(IF(OR(G8="",H8="",I8=""),"",IF(H8>G8,(I8-G8)/(H8-G8),(G8-I8)/(G8-H8))),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER', numberFormat: { type: 'PERCENT', pattern: '0%' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push goal data (B=name, C=type, D=status, E=start_date, F=target_date, G=start_val, H=target_val, I=current_val, J=formula, K=notes)
  vals.push({ range: `${S}!B8`, values: GOALS.map(g => [g[0], g[1], g[2], g[3], g[4], g[5], g[6], g[7], '', g[8]]) });

  // Row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,500,0,11),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date formats
  fmt.push({ repeatCell: { range: gridRange(SID,7,500,4,6), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [75, 260, 130, 90, 100, 100, 90, 90, 90, 80, 220];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '11-goals values');
  await batchUpdate(id, fmt, '11-goals format');
  console.log(`✅  Goals & Milestones done — ${GOALS.length} goals.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
