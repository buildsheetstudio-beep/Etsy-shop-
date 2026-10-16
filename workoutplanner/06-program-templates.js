'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PTL = sheetMap['📚 Program Templates'];

// Columns: A=Program Name, B=Day/Session, C=Muscle Group, D=Exercise, E=Sets × Reps, F=Notes
// NO col A freeze — full-width title A1:F1
// Row 1: title, Row 2: headers, Rows 3-17: 15 template rows

const templates = [
  // Push/Pull/Legs — 3 exercises each
  ['Push Day',   'Day 1 (Mon/Thu)', 'Chest',     'Bench Press',             '4 × 8',   'Focus on full ROM'],
  ['Push Day',   'Day 1 (Mon/Thu)', 'Shoulders', 'Overhead Press',          '3 × 10',  'Seated or standing'],
  ['Push Day',   'Day 1 (Mon/Thu)', 'Arms',      'Tricep Extension',        '3 × 12',  'Cable or dumbbell'],
  ['Pull Day',   'Day 2 (Tue/Fri)', 'Back',      'Deadlift',                '4 × 5',   'Rest 3 min between'],
  ['Pull Day',   'Day 2 (Tue/Fri)', 'Back',      'Barbell Row',             '4 × 8',   'Overhand grip'],
  ['Pull Day',   'Day 2 (Tue/Fri)', 'Arms',      'Dumbbell Curl',           '3 × 12',  'Supinate at top'],
  ['Legs Day',   'Day 3 (Wed/Sat)', 'Legs',      'Squat',                   '4 × 6',   'High bar or low bar'],
  ['Legs Day',   'Day 3 (Wed/Sat)', 'Legs',      'Leg Press',               '4 × 10',  'Full depth'],
  ['Legs Day',   'Day 3 (Wed/Sat)', 'Glutes',    'Hip Thrust',              '3 × 12',  'Hold at top 1sec'],
  // Upper/Lower — 2 exercises each
  ['Upper Body', 'Day 1',           'Chest',     'Incline Dumbbell Press',  '4 × 10',  ''],
  ['Upper Body', 'Day 1',           'Back',      'Lat Pulldown',            '4 × 10',  'Wide grip'],
  ['Lower Body', 'Day 2',           'Legs',      'Romanian Deadlift',       '4 × 8',   '3-sec eccentric'],
  // Full Body — 3 exercises
  ['Full Body',  'Day A',           'Full Body', 'Clean and Press',         '3 × 8',   'Explosive movement'],
  // Bro Split — chest focus
  ['Bro Split',  'Chest Day',       'Chest',     'Bench Press',             '5 × 5',   'Linear progression'],
  // Custom
  ['Custom',     'Core Day',        'Core',      'Plank',                   '3 × 60s', 'Add weight vest'],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'📚 Program Templates'!A1", values: [['📚 Program Templates — Workout Library']] },
    { range: "'📚 Program Templates'!A2", values: [['Program Name','Day / Session','Muscle Group','Exercise','Sets × Reps','Notes']] },
    { range: "'📚 Program Templates'!A3", values: templates },
  ], 'ptl-data');

  const reqs = [];

  // Full-width merge A1:F1 (no col A standalone — this sheet has no col freeze)
  reqs.push({ mergeCells: { range: gridRange(PTL, 0, 1, 0, 6), mergeType: 'MERGE_ALL' } });

  reqs.push({
    repeatCell: {
      range: gridRange(PTL, 0, 1, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.steelGray),
          textFormat: { foregroundColor: hex(C.white), fontSize: 13, bold: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  reqs.push({
    repeatCell: {
      range: gridRange(PTL, 1, 2, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.darkGray),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Group colors by program
  const programColors = {
    'Push Day':     C.cobaltBlue,
    'Pull Day':     C.teal,
    'Legs Day':     C.coral,
    'Upper Body':   C.limeGreen,
    'Lower Body':   C.limeGreen,
    'Full Body':    C.amber,
    'Bro Split':    C.purple,
    'Custom':       C.steelGray,
  };

  for (let r = 0; r < templates.length; r++) {
    const progColor = programColors[templates[r][0]] || C.steelGray;
    reqs.push({
      repeatCell: {
        range: gridRange(PTL, r + 2, r + 3, 0, 6),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 0 ? C.offWhite : C.altRow),
            textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
    // Color-code Program Name col A for each row
    reqs.push({
      repeatCell: {
        range: gridRange(PTL, r + 2, r + 3, 0, 1),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(progColor),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  }

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: PTL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: PTL, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: PTL, dimension: 'ROWS', startIndex: 2, endIndex: 17 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths
  [[0,130],[1,145],[2,110],[3,175],[4,90],[5,200]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: PTL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'ptl-format');
  console.log('Program Templates complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
