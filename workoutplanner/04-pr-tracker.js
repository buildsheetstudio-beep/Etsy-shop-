'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PRT = sheetMap['🏆 PR Tracker'];

// Columns: A=Exercise, B=Current PR (lbs, MAXIFS), C=Date Achieved (MINIFS),
//          D=Last Reps, E=Target (lbs), F=Notes
// Row 1: title, Row 2: headers, Rows 3-14: 12 exercises

const exercises = [
  // [exercise, lastReps, target, notes]
  ['Bench Press',    5, 245, 'Flat barbell'],
  ['Squat',          5, 315, 'Back squat'],
  ['Deadlift',       5, 335, 'Conventional'],
  ['Overhead Press', 6, 155, 'Standing strict'],
  ['Barbell Row',    8, 205, 'Pendlay style'],
  ['Pull-up',        8, 0,   'Bodyweight — track reps'],
  ['Dumbbell Curl',  12,55,  'Per arm'],
  ['Tricep Extension',15,70, 'Cable pushdown'],
  ['Leg Press',      12,450, 'Sled weight'],
  ['Romanian Deadlift',10,225,'3-second eccentric'],
  ['Hip Thrust',     12,245, 'Barbell'],
  ['Lat Pulldown',   12,165, 'Wide grip'],
];

(async () => {
  // ── Headers ───────────────────────────────────────────────────────────────
  await valuesBatchUpdate(id, [
    { range: "'🏆 PR Tracker'!A1", values: [['🏆']] },
    { range: "'🏆 PR Tracker'!B1", values: [['PR Tracker — Personal Records']] },
    { range: "'🏆 PR Tracker'!A2", values: [['Exercise','Current PR (lbs)','Date Achieved','Last Reps','Target (lbs)','Notes']] },
  ], 'prt-headers');

  // ── Data rows with formulas ────────────────────────────────────────────────
  const dataRows = exercises.map(([exercise, lastReps, target, notes], i) => {
    const r = i + 3;
    return [
      exercise,
      `=IFERROR(MAXIFS('📝 Workout Log'!$G$3:$G$32,'📝 Workout Log'!$D$3:$D$32,A${r}),0)`,
      `=IFERROR(MINIFS('📝 Workout Log'!$A$3:$A$32,'📝 Workout Log'!$D$3:$D$32,A${r},'📝 Workout Log'!$G$3:$G$32,B${r}),"")`,
      lastReps,
      target,
      notes,
    ];
  });

  await valuesBatchUpdate(id, [
    { range: "'🏆 PR Tracker'!A3", values: dataRows },
  ], 'prt-data');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  reqs.push({ mergeCells: { range: gridRange(PRT, 0, 1, 1, 6), mergeType: 'MERGE_ALL' } });

  reqs.push({
    repeatCell: {
      range: gridRange(PRT, 0, 1, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.headerBlue),
          textFormat: { foregroundColor: hex(C.amber), fontSize: 16, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  reqs.push({
    repeatCell: {
      range: gridRange(PRT, 0, 1, 1, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.headerBlue),
          textFormat: { foregroundColor: hex(C.amber), fontSize: 13, bold: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  reqs.push({
    repeatCell: {
      range: gridRange(PRT, 1, 2, 0, 6),
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

  for (let r = 2; r < 14; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(PRT, r, r + 1, 0, 6),
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
  }

  // Formula cols B, C = formulaBg
  reqs.push({
    repeatCell: {
      range: gridRange(PRT, 2, 14, 1, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
    },
  });

  // Date format for col C
  reqs.push({
    repeatCell: {
      range: gridRange(PRT, 2, 14, 2, 3),
      cell: {
        userEnteredFormat: {
          numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
        },
      },
      fields: 'userEnteredFormat(numberFormat)',
    },
  });

  // Center align D, E
  reqs.push({
    repeatCell: {
      range: gridRange(PRT, 2, 14, 3, 5),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(horizontalAlignment)',
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: PRT, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: PRT, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: PRT, dimension: 'ROWS', startIndex: 2, endIndex: 14 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths
  [[0,160],[1,110],[2,110],[3,80],[4,100],[5,180]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: PRT, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'prt-format');
  console.log('PR Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
