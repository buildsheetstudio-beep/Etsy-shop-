'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MG_COLORS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WL = sheetMap['💪 Workout Log'];

// Columns: A=Date, B=Muscle Group/Focus, C=Workout Type, D=Exercises Summary, E=Duration (min), F=Completed, G=Notes
const HEADERS = ['Date', 'Muscle Group / Focus', 'Workout Type', 'Exercises Summary', 'Duration (min)', 'Completed ✓', 'Notes'];

// 30 rows — June 2025, mixing all types
const ROWS = [
  ['=DATE(2025,6,1)',  'Chest',     'Strength',         'Bench Press 4×8, Incline DB Press 3×10, Cable Fly 3×12',        55, true,  'New weight PR on bench press'],
  ['=DATE(2025,6,2)',  'Back',      'Strength',         'Deadlift 4×5, Pull-ups 4×8, Seated Row 3×12',                   60, true,  'Deadlift felt strong'],
  ['=DATE(2025,6,3)',  'Cardio',    'Cardio',           'Morning run — 5k steady state',                                   30, true,  '5:45/km pace'],
  ['=DATE(2025,6,4)',  'Legs',      'Strength',         'Squat 4×8, Romanian Deadlift 3×10, Leg Press 3×12',             65, true,  'Quads really loaded'],
  ['=DATE(2025,6,5)',  'Shoulders', 'Strength',         'OHP 4×8, Lateral Raise 3×15, Face Pull 3×15',                   50, true,  'Solid shoulder day'],
  ['=DATE(2025,6,6)',  'Core',      'Circuit/Superset', 'Upper Body Superset A — 4 rounds',                               40, true,  'Used circuit from builder tab'],
  ['=DATE(2025,6,7)',  'Rest',      'Rest',             '',                                                                 0, true,  'Active recovery — light walk'],
  ['=DATE(2025,6,8)',  'Arms',      'Strength',         'Bicep Curl 4×12, Tricep Pushdown 4×12, Hammer Curl 3×10',       45, true,  ''],
  ['=DATE(2025,6,9)',  'Cardio',    'HIIT',             '20-min HIIT — Bike intervals 8×30s sprint/90s rest',             25, true,  'Heart rate peaked Z5'],
  ['=DATE(2025,6,10)', 'Back',      'Strength',         'Pull-ups 5×8, Barbell Row 4×8, Face Pull 3×15',                 55, true,  ''],
  ['=DATE(2025,6,11)', 'Legs',      'Strength',         'Leg Burner Circuit — 3 rounds',                                  50, true,  'Legs absolutely destroyed'],
  ['=DATE(2025,6,12)', 'Chest',     'Strength',         'Incline Bench 4×8, Chest Dip 3×10, Pec Deck 3×15',             52, true,  ''],
  ['=DATE(2025,6,13)', 'Cardio',    'Cardio',           'Outdoor cycling — 20km easy Z2 ride',                            55, true,  '22km/h avg'],
  ['=DATE(2025,6,14)', 'Rest',      'Rest',             '',                                                                 0, true,  ''],
  ['=DATE(2025,6,15)', 'Glutes',    'Strength',         'Hip Thrust 4×12, Bulgarian Split Squat 3×10, Cable Kickback 3×15',58, true,'New hip thrust PR'],
  ['=DATE(2025,6,16)', 'Core',      'Strength',         'Core Destroyer Circuit — 3 rounds',                              35, true,  'Core on fire'],
  ['=DATE(2025,6,17)', 'Shoulders', 'Strength',         'OHP 4×6, Arnold Press 3×10, Cable Lateral Raise 3×15',          48, true,  ''],
  ['=DATE(2025,6,18)', 'Cardio',    'Cardio',           'Treadmill run — 8km tempo run',                                   45, true,  '4:58/km pace — new 8k PB!'],
  ['=DATE(2025,6,19)', 'Arms',      'Circuit/Superset', 'Arms Superset B — 4 rounds',                                     42, true,  ''],
  ['=DATE(2025,6,20)', 'Rest',      'Rest',             '',                                                                 0, true,  ''],
  ['=DATE(2025,6,21)', 'Chest',     'Strength',         'Bench Press 5×5 — Heavy Day, Incline DB 3×10',                  58, false, 'Felt off — session cut short'],
  ['=DATE(2025,6,22)', 'Legs',      'HIIT',             'Leg Burner Circuit — 4 rounds (added extra round)',               55, true,  'New PR in circuit time'],
  ['=DATE(2025,6,23)', 'Back',      'Strength',         'Deadlift 4×4, Weighted Pull-ups 4×6, Cable Row 3×12',           62, true,  'Deadlift volume PR'],
  ['=DATE(2025,6,24)', 'Cardio',    'Cardio',           'Row machine — 5000m steady state',                               30, true,  'Split: 2:05/500m'],
  ['=DATE(2025,6,25)', 'Glutes',    'Strength',         'Hip Thrust 4×10, RDL 3×12, Glute Cable Kickback 3×15',          50, true,  ''],
  ['=DATE(2025,6,26)', 'Shoulders', 'Mobility/Stretching','Full upper body mobility — 5-min shoulder circles, band pull-aparts',30,true,'Recovery session'],
  ['=DATE(2025,6,27)', 'Core',      'Circuit/Superset', 'Core Destroyer Circuit — 4 rounds',                              38, true,  'Fastest time yet'],
  ['=DATE(2025,6,28)', 'Rest',      'Rest',             '',                                                                 0, true,  ''],
  ['=DATE(2025,6,29)', 'Arms',      'Strength',         'Bicep Curl 4×12, Skull Crusher 4×10, Cable Curl 3×15',          44, true,  ''],
  ['=DATE(2025,6,30)', 'Cardio',    'Cardio',           'Easy 4km jog — recovery run',                                    25, true,  'Z2 the whole way, HR avg 135'],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'💪 Workout Log'!A1:G1", values: [HEADERS] },
    { range: "'💪 Workout Log'!A2:G31", values: ROWS },
  ], 'wl-values');

  const reqs = [];

  // Base fill
  reqs.push({ repeatCell: { range: gridRange(WL,0,50,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.navyDark) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Header
  reqs.push({ repeatCell: { range: gridRange(WL,0,1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.cobalt), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Alternating rows
  for (let i = 1; i < 31; i += 2) {
    reqs.push({ repeatCell: { range: gridRange(WL,i,i+1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  }

  // Date column — YYYY-MM-DD format
  reqs.push({ repeatCell: { range: gridRange(WL,1,31,0,1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Duration — center
  reqs.push({ repeatCell: { range: gridRange(WL,1,31,4,5), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Completed — center
  reqs.push({ repeatCell: { range: gridRange(WL,1,31,5,6), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Muscle group color-coding column B
  const MUSCLE_ROWS = {
    'Chest': [0,11], 'Back': [1,9,22], 'Cardio': [2,8,12,17,23,29], 'Legs': [3,10,21],
    'Shoulders': [4,16,25], 'Core': [5,15,26], 'Rest': [6,13,19,27], 'Arms': [7,18,28], 'Glutes': [14,24],
  };
  const { MG_COLORS } = require('./lib');
  Object.entries(MUSCLE_ROWS).forEach(([mg, rows]) => {
    rows.forEach(r => {
      const ri = r + 1; // 0-indexed row offset, +1 for header
      reqs.push({ repeatCell: { range: gridRange(WL, ri, ri+1, 1, 2), cell: { userEnteredFormat: { backgroundColor: hex(MG_COLORS[mg]), textFormat: { bold: true } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
    });
  });

  // Input cell bg for editable columns (A, B, C, D, E, G — not F formula)
  reqs.push({ repeatCell: { range: gridRange(WL,1,31,0,5), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(WL,1,31,6,7), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: WL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: WL, dimension: 'ROWS', startIndex: 1, endIndex: 31 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths: Date, MuscleGroup, WorkoutType, Summary, Duration, Completed, Notes
  [[0,110],[1,140],[2,130],[3,320],[4,110],[5,100],[6,200]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: WL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'wl-format');
  console.log('Workout Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
