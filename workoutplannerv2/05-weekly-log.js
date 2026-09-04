'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MG_COLORS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WDL = sheetMap['📝 Weekly Detail Log'];

// Columns: A=Week#, B=Day, C=Exercise Name, D=Muscle Group, E=Sets, F=Reps, G=Weight, H=RPE (1-10), I=Notes
const HEADERS = ['Week #', 'Day', 'Exercise Name', 'Muscle Group', 'Sets', 'Reps', 'Weight (kg/lbs)', 'RPE (1-10)', 'Notes'];

const ROWS = [
  // Week 1
  [1, 'Monday',    'Bench Press',            'Chest',     4, 8,  80,  8,  'Last set to failure'],
  [1, 'Monday',    'Incline DB Press',       'Chest',     3, 10, 30,  7,  ''],
  [1, 'Monday',    'Cable Fly',              'Chest',     3, 12, 15,  7,  'Squeeze at top'],
  [1, 'Tuesday',   'Deadlift',               'Back',      4, 5,  120, 9,  'Keep core braced'],
  [1, 'Tuesday',   'Pull-ups',               'Back',      4, 8,  0,   8,  'Bodyweight only'],
  [1, 'Tuesday',   'Seated Cable Row',       'Back',      3, 12, 60,  7,  ''],
  [1, 'Wednesday', 'Treadmill Run',          'Cardio',    1, '', '',  6,  '5km @ 5:45/km'],
  [1, 'Thursday',  'Barbell Squat',          'Legs',      4, 8,  90,  8,  'Pause at bottom'],
  [1, 'Thursday',  'Romanian Deadlift',      'Legs',      3, 10, 70,  7,  ''],
  [1, 'Thursday',  'Leg Press',              'Legs',      3, 12, 150, 7,  ''],
  [1, 'Friday',    'Overhead Press',         'Shoulders', 4, 8,  55,  8,  ''],
  [1, 'Friday',    'Lateral Raise',          'Shoulders', 3, 15, 12,  7,  'Slow eccentric'],
  [1, 'Friday',    'Face Pull',              'Shoulders', 3, 15, 25,  6,  'Warm down'],
  // Week 2
  [2, 'Monday',    'Incline Bench Press',    'Chest',     4, 8,  72,  8,  ''],
  [2, 'Monday',    'Chest Dip',              'Chest',     3, 10, 0,   8,  'Bodyweight + lean forward'],
  [2, 'Tuesday',   'Weighted Pull-ups',      'Back',      4, 6,  10,  9,  '+10kg belt'],
  [2, 'Tuesday',   'Barbell Row',            'Back',      4, 8,  80,  8,  'Elbows tight'],
  [2, 'Wednesday', 'Bike Intervals',         'Cardio',    8, '', '',  9,  '30s on / 90s off × 8'],
  [2, 'Thursday',  'Bulgarian Split Squat',  'Glutes',    3, 10, 20,  9,  'Hard balance'],
  [2, 'Thursday',  'Hip Thrust',             'Glutes',    4, 12, 80,  8,  'New PR weight!'],
  [2, 'Friday',    'Arnold Press',           'Shoulders', 3, 10, 20,  8,  ''],
  [2, 'Friday',    'Bicep Curl',             'Arms',      4, 12, 16,  7,  ''],
  [2, 'Friday',    'Tricep Pushdown',        'Arms',      4, 12, 30,  7,  ''],
  [2, 'Saturday',  'Plank Hold',             'Core',      3, '', '',  6,  '60-second holds'],
  [2, 'Saturday',  'Hollow Body Hold',       'Core',      3, '', '',  7,  '45-second holds'],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'📝 Weekly Detail Log'!A1:I1", values: [HEADERS] },
    { range: "'📝 Weekly Detail Log'!A2:I26", values: ROWS },
  ], 'wdl-values');

  const reqs = [];

  // Base fill
  reqs.push({ repeatCell: { range: gridRange(WDL,0,40,0,9), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.navyDark) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Header
  reqs.push({ repeatCell: { range: gridRange(WDL,0,1,0,9), cell: { userEnteredFormat: { backgroundColor: hex(C.cobalt), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Alternating
  for (let i = 1; i < 26; i += 2) {
    reqs.push({ repeatCell: { range: gridRange(WDL,i,i+1,0,9), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  }

  // Week # col — centered, bold cobalt
  reqs.push({ repeatCell: { range: gridRange(WDL,1,26,0,1), cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.cobalt) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });

  // Day col — centered
  reqs.push({ repeatCell: { range: gridRange(WDL,1,26,1,2), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Muscle group col D (idx 3) — color coded
  const MUSCLE_ROWS = {
    'Chest': [0,1,2,13,14], 'Back': [3,4,5,15,16], 'Cardio': [6,17],
    'Legs': [7,8,9], 'Shoulders': [10,11,12,20], 'Glutes': [18,19],
    'Arms': [21,22], 'Core': [23,24],
  };
  Object.entries(MUSCLE_ROWS).forEach(([mg, rows]) => {
    rows.forEach(r => {
      const ri = r + 1;
      reqs.push({ repeatCell: { range: gridRange(WDL,ri,ri+1,3,4), cell: { userEnteredFormat: { backgroundColor: hex(MG_COLORS[mg]), textFormat: { bold: true } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    });
  });

  // Numeric cols (Sets, Reps, Weight, RPE) — center
  [4,5,6,7].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(WDL,1,26,ci,ci+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });
  });

  // Input bg
  reqs.push({ repeatCell: { range: gridRange(WDL,1,26,0,9), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: WDL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: WDL, dimension: 'ROWS', startIndex: 1, endIndex: 26 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,70],[1,90],[2,210],[3,120],[4,60],[5,60],[6,110],[7,80],[8,200]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: WDL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'wdl-format');
  console.log('Weekly Detail Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
