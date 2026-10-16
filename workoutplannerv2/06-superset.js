'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MG_COLORS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SS = sheetMap['⚡ Superset & Circuit Builder'];

// Columns: A=Circuit Name, B=Exercise Order, C=Exercise Name, D=Reps or Duration, E=Rest Between (sec), F=Total Rounds, G=Muscle Group Focus, H=Notes
const HEADERS = ['Circuit / Superset Name', 'Order', 'Exercise Name', 'Reps or Duration', 'Rest Between (sec)', 'Total Rounds', 'Muscle Group Focus', 'Notes'];

const ROWS = [
  // Upper Body Superset A (5 exercises, 4 rounds)
  ['Upper Body Superset A', 1, 'Push-up',                '15 reps',        0,   4, 'Chest',     'No rest between exercises 1-5'],
  ['Upper Body Superset A', 2, 'Dumbbell Row',            '10 reps each',   0,   4, 'Back',      'Use 20kg DBs'],
  ['Upper Body Superset A', 3, 'Lateral Raise',           '12 reps',        0,   4, 'Shoulders', ''],
  ['Upper Body Superset A', 4, 'Tricep Dip',              '12 reps',        0,   4, 'Arms',      'Use bench or parallel bars'],
  ['Upper Body Superset A', 5, 'Bicep Curl',              '12 reps',        90,  4, 'Arms',      '90s rest after full round'],
  // Leg Burner Circuit (4 exercises, 3 rounds)
  ['Leg Burner Circuit',    1, 'Barbell Squat',           '10 reps',        0,   3, 'Legs',      '60-70% 1RM'],
  ['Leg Burner Circuit',    2, 'Walking Lunge',            '12 steps each',  0,   3, 'Legs',      'Hold DBs optional'],
  ['Leg Burner Circuit',    3, 'Jump Squat',              '15 reps',        0,   3, 'Legs',      'Explosive power'],
  ['Leg Burner Circuit',    4, 'Wall Sit',                '45 seconds',     120, 3, 'Legs',      '2-min rest after full round'],
  // Arms Superset B (4 exercises, 4 rounds)
  ['Arms Superset B',       1, 'Barbell Curl',            '10 reps',        0,   4, 'Arms',      'Strict form, no swinging'],
  ['Arms Superset B',       2, 'Skull Crusher',           '10 reps',        0,   4, 'Arms',      'Use EZ bar'],
  ['Arms Superset B',       3, 'Hammer Curl',             '12 reps each',   0,   4, 'Arms',      ''],
  ['Arms Superset B',       4, 'Overhead Tricep Extension','12 reps',       90,  4, 'Arms',      '90s rest after full round'],
  // Core Destroyer Circuit (4 exercises, 3 rounds)
  ['Core Destroyer Circuit',1, 'Plank Hold',              '60 seconds',     0,   3, 'Core',      'Perfect alignment'],
  ['Core Destroyer Circuit',2, 'Russian Twist',           '20 reps',        0,   3, 'Core',      'With or without weight'],
  ['Core Destroyer Circuit',3, 'Leg Raise',               '15 reps',        0,   3, 'Core',      'Slow and controlled'],
  ['Core Destroyer Circuit',4, 'Mountain Climbers',       '30 seconds',     60,  3, 'Core',      '60s rest after full round'],
  // Glute & Hip Circuit
  ['Glute & Hip Circuit',   1, 'Hip Thrust',              '12 reps',        0,   3, 'Glutes',    'Heavy load, full lockout'],
  ['Glute & Hip Circuit',   2, 'Cable Kickback',          '15 reps each',   0,   3, 'Glutes',    ''],
  ['Glute & Hip Circuit',   3, 'Sumo Squat',              '15 reps',        90,  3, 'Glutes',    '90s rest after full round'],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'⚡ Superset & Circuit Builder'!A1:H1", values: [HEADERS] },
    { range: "'⚡ Superset & Circuit Builder'!A2:H21", values: ROWS },
  ], 'ss-values');

  const reqs = [];

  // Base
  reqs.push({ repeatCell: { range: gridRange(SS,0,30,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.navyDark) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Header
  reqs.push({ repeatCell: { range: gridRange(SS,0,1,0,8), cell: { userEnteredFormat: { backgroundColor: hex('#F57C00'), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Circuit group colors — alternate between slight tints for each circuit
  const circuitColors = ['#FFF3E0','#E8F5E9','#EDE7F6','#E3F2FD','#FCE4EC'];
  const circuitGroups = [
    { name: 'Upper Body Superset A', rows: [0,1,2,3,4],     color: circuitColors[0] },
    { name: 'Leg Burner Circuit',    rows: [5,6,7,8],        color: circuitColors[1] },
    { name: 'Arms Superset B',       rows: [9,10,11,12],     color: circuitColors[2] },
    { name: 'Core Destroyer Circuit',rows: [13,14,15,16],    color: circuitColors[3] },
    { name: 'Glute & Hip Circuit',   rows: [17,18,19],       color: circuitColors[4] },
  ];
  circuitGroups.forEach(({ rows, color }) => {
    rows.forEach(r => {
      const ri = r + 1;
      reqs.push({ repeatCell: { range: gridRange(SS,ri,ri+1,0,8), cell: { userEnteredFormat: { backgroundColor: hex(color) } }, fields: 'userEnteredFormat(backgroundColor)' } });
    });
  });

  // Circuit name — bold, cobalt
  reqs.push({ repeatCell: { range: gridRange(SS,1,21,0,1), cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.cobalt) } } }, fields: 'userEnteredFormat(textFormat)' } });

  // Order — center
  reqs.push({ repeatCell: { range: gridRange(SS,1,21,1,2), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });

  // Muscle group col G (idx 6) — color coded
  const mgRows = {
    'Chest': [0], 'Back': [1], 'Shoulders': [2], 'Arms': [3,4,9,10,11,12], 'Legs': [5,6,7,8], 'Core': [13,14,15,16], 'Glutes': [17,18,19],
  };
  Object.entries(mgRows).forEach(([mg, rows]) => {
    rows.forEach(r => {
      const ri = r + 1;
      reqs.push({ repeatCell: { range: gridRange(SS,ri,ri+1,6,7), cell: { userEnteredFormat: { backgroundColor: hex(MG_COLORS[mg]), textFormat: { bold: true } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    });
  });

  // Numeric cols (Rest, Rounds) — center
  [4,5].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(SS,1,21,ci,ci+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SS, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SS, dimension: 'ROWS', startIndex: 1, endIndex: 21 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,190],[1,60],[2,200],[3,120],[4,120],[5,100],[6,130],[7,200]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SS, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'ss-format');
  console.log('Superset & Circuit Builder complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
