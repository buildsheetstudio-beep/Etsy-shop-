'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Weekly Check-In'];
const S = "'Weekly Check-In'";

// 12 weekly check-ins: Jun 23 - Sep 8, 2026
// [week_start, weight, energy(1-5), sleep_qual(1-5), stress(1-5), workouts, mood, wins, focus_next]
const CHECKINS = [
  ['2026-06-23', 178.4, 4, 4, 2, 3, 'Motivated and consistent',      'Logged every meal this week',          'Keep up water intake'],
  ['2026-06-30', 177.9, 3, 3, 3, 2, 'Steady, a bit tired',           'Completed 2 workouts despite busy week','Prioritize sleep'],
  ['2026-07-07', 178.8, 4, 4, 2, 3, 'Great energy this week',        'Hit water goal 6/7 days',              'Add evening walk'],
  ['2026-07-14', 177.5, 5, 5, 2, 4, 'Best week in a while!',         'Yoga 3x + strength 1x',               'Keep this routine'],
  ['2026-07-21', 178.1, 3, 3, 4, 2, 'High-stress work week',         'Maintained logging despite stress',    'Stress management this week'],
  ['2026-07-28', 177.2, 4, 4, 3, 3, 'Feeling more grounded',         'Meditation streak: 7 days',            'Continue meditation habit'],
  ['2026-08-04', 177.8, 4, 3, 2, 3, 'Good overall',                  '75 workouts logged overall',           'Mix in more variety'],
  ['2026-08-11', 176.9, 4, 4, 2, 4, 'Strong week physically',        'Ran first 5K in months',              'Build up cardio base'],
  ['2026-08-18', 177.4, 3, 4, 3, 3, 'Consistent week',               'Prepped meals Sunday – saved time',   'Meal prep next Sunday too'],
  ['2026-08-25', 176.5, 5, 5, 1, 4, 'Feeling fantastic',             'Habit streak for all 9 active habits','Maintain the momentum'],
  ['2026-09-01', 176.8, 4, 4, 2, 3, 'Good end-of-summer energy',     'Tried two new healthy recipes',       'Focus on sleep consistency'],
  ['2026-09-08', 176.2, 4, 4, 2, 3, 'Into fall routine now',         'Best month of workouts logged',       'Set fall season goals'],
];

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5000,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['WEEKLY CHECK-IN']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Reflect weekly. Rate Energy, Sleep Quality, and Stress 1–5. This is a judgment-free check-in for your own awareness.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,10), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Stats
  const STATS = [
    ['Check-Ins Logged',  `=COUNTA(B8:B500)`],
    ['Avg Energy (1-5)',  `=IFERROR(ROUND(AVERAGE(C8:C500),1),"")`],
    ['Avg Sleep Qual(1-5)',`=IFERROR(ROUND(AVERAGE(D8:D500),1),"")`],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,0,4), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,4,8), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!E${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,4), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,4,8), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 6: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Row 7: headers
  vals.push({ range: `${S}!A7`, values: [['Check-In ID','Week Start','Weight (lbs)','Energy (1-5)','Sleep Quality (1-5)','Stress (1-5)','Workouts This Week','Mood / How I Feel','Wins This Week','Focus for Next Week']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,10), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // A: Check-In ID formula
  fmt.push({ repeatCell: { range: gridRange(SID,7,500,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B8="","","CI-"&TEXT(ROW()-7,"000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push check-in data (B=week_start, C=weight, D=energy, E=sleep, F=stress, G=workouts, H=mood, I=wins, J=focus)
  vals.push({ range: `${S}!B8`, values: CHECKINS.map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8]]) });

  // Row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,500,0,10),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format col B
  fmt.push({ repeatCell: { range: gridRange(SID,7,500,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [80, 100, 90, 90, 110, 90, 110, 200, 220, 220];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '12-checkin values');
  await batchUpdate(id, fmt, '12-checkin format');
  console.log(`✅  Weekly Check-In done — ${CHECKINS.length} entries.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
