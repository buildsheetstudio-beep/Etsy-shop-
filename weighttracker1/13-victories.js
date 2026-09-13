'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Non-Scale Victories'];
const S = "'Non-Scale Victories'";

// 26 non-scale victories
// [date, category, description, how_it_felt, notes]
const VICTORIES = [
  ['2024-11-15','Energy & Vitality',   'Walked up 4 flights of stairs without getting winded',    'Surprised and proud',          ''],
  ['2024-12-03','Sleep Quality',       'First full week of 7+ hours of sleep',                    'Refreshed and clear-headed',   ''],
  ['2024-12-18','Physical Strength',   'Did my first unassisted push-up in years',                'Empowered',                    'Sets of 5!'],
  ['2025-01-07','Confidence',          'Wore a jacket I hadn\'t been able to button for 2 years', 'Genuinely happy',              ''],
  ['2025-01-20','Mood',                'Went 2 weeks without a low-mood afternoon slump',         'Stable and positive',          ''],
  ['2025-02-05','Endurance',           'Hiked 5 miles without feeling exhausted the next day',    'Strong',                       ''],
  ['2025-02-22','Health Markers',      'Doctor mentioned blood pressure is in healthy range',     'Reassured',                    ''],
  ['2025-03-10','Clothing Fit',        'Old jeans fit comfortably again',                         'Nostalgic and pleased',        ''],
  ['2025-04-01','Sleep Quality',       'Slept through the night 5 nights in a row',               'Finally rested',               ''],
  ['2025-04-18','Energy & Vitality',   'Played with my kids for 45 min without needing to sit',  'Present and joyful',           ''],
  ['2025-05-02','Physical Strength',   'Can do 3 sets of 10 push-ups consistently',               'Progress is real!',            ''],
  ['2025-05-20','Social',              'Said yes to a spontaneous hiking invite (wouldn\'t before)','Glad I went',                 ''],
  ['2025-06-05','Flexibility',         'Touched my toes for first time in 5 years',               'Shocked and delighted',        ''],
  ['2025-06-25','Confidence',          'Wore a sleeveless top and felt comfortable',               'Easy choice this time',        ''],
  ['2025-07-10','Endurance',           'Ran a 5K without stopping',                               'Personal milestone',           'First 5K!'],
  ['2025-07-28','Health Markers',      'Resting heart rate down to 62 bpm',                       'Quietly proud',                ''],
  ['2025-08-14','Mood',                'Handled a stressful day without emotional eating',        'In control',                   ''],
  ['2025-09-02','Energy & Vitality',   'Morning energy without needing 3 cups of coffee',        'Game changer',                 ''],
  ['2025-10-08','Social',              'Volunteered for a charity 5K walk',                       'Community connection',         ''],
  ['2026-01-15','Physical Strength',   'Can do a proper plank for 90 seconds',                   'Stronger core!',               ''],
  ['2026-02-28','Clothing Fit',        'Bought clothes two sizes smaller than 2 years ago',      'Milestone moment',             ''],
  ['2026-04-10','Flexibility',         'Full yoga practice without modifying any poses',          'Body awareness improved',      ''],
  ['2026-05-18','Energy & Vitality',   'Ran 10K distance for the first time',                    'Beyond what I thought possible','10K!'],
  ['2026-07-04','Confidence',          'Wore a swimsuit at the beach, no hesitation',            'Freedom',                      ''],
  ['2026-08-12','Health Markers',      'Annual checkup: all markers in healthy range',            'Peace of mind',                ''],
  ['2026-09-01','Endurance',           '100th workout session logged',                            'Consistent over a full year!', 'Milestone: 100 sessions'],
];

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5000,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['NON-SCALE VICTORIES']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,6), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Celebrate every win that isn\'t about a number on the scale. These victories are just as real — and often more meaningful.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,6), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Stats
  const STATS = [
    ['Victories Recorded',   `=COUNTA(B8:B500)`],
    ['Most Common Category', `=IFERROR(INDEX(C8:C500,MATCH(MAX(COUNTIF(C8:C500,C8:C500)),COUNTIF(C8:C500,C8:C500),0)),"")`],
    ['Wins This Year',       `=COUNTIFS(B8:B500,">="&DATE(YEAR(TODAY()),1,1),B8:B500,"<="&TODAY())`],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,0,2), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,2,6), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!C${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,2,6), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 6: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Row 7: headers
  vals.push({ range: `${S}!A7`, values: [['Victory ID','Date','Category','Victory Description','How It Made Me Feel','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // A: Victory ID formula
  fmt.push({ repeatCell: { range: gridRange(SID,7,500,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B8="","","NSV-"&TEXT(ROW()-7,"000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push victory data (B=date, C=category, D=description, E=how_it_felt, F=notes)
  vals.push({ range: `${S}!B8`, values: VICTORIES });

  // Row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,500,0,6),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format col B
  fmt.push({ repeatCell: { range: gridRange(SID,7,500,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Row height for content rows (allow text wrapping)
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 500 }, properties: { pixelSize: 50 }, fields: 'pixelSize' }});

  // Wrap text for description/feeling cols
  fmt.push({ repeatCell: { range: gridRange(SID,7,500,3,6), cell: { userEnteredFormat: {
    wrapStrategy: 'WRAP', verticalAlignment: 'TOP',
  }}, fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' }});

  // Column widths
  const WIDTHS = [80, 100, 140, 320, 220, 180];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '13-victories values');
  await batchUpdate(id, fmt, '13-victories format');
  console.log(`✅  Non-Scale Victories done — ${VICTORIES.length} victories.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
