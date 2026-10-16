'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WKL = sheetMap['📝 Workout Log'];

// Columns: A=Date, B=Day(formula), C=Muscle Group, D=Exercise, E=Workout Type,
//          F=Sets, G=Weight(lbs), H=Reps, I=Duration(min), J=Completed, K=Notes
// Row 1: title (A1=emoji, B1:K1=merged), Row 2: headers, Rows 3-32: data (30 rows)

const rawData = [
  // [date, muscleGroup, exercise, workoutType, sets, weight, reps, duration, completed, notes]
  ['2026-07-01','Chest','Bench Press','Strength',4,225,8,45,true,''],
  ['2026-07-02','Back','Deadlift','Strength',4,315,5,50,true,'New weight PR'],
  ['2026-07-03','Legs','Squat','Strength',4,275,6,55,true,''],
  ['2026-07-04','Cardio','Running','Cardio',1,0,30,30,true,'5K easy run'],
  ['2026-07-05','Mobility','Yoga','Yoga',1,0,60,60,true,'Active recovery'],
  ['2026-07-06','Shoulders','Overhead Press','Strength',4,135,8,40,true,''],
  ['2026-07-07','Arms','Dumbbell Curl','Hypertrophy',3,45,12,35,true,''],
  ['2026-07-08','Chest','Incline Dumbbell Press','Hypertrophy',3,70,12,40,true,''],
  ['2026-07-09','Back','Barbell Row','Strength',4,185,8,45,true,''],
  ['2026-07-10','Legs','Leg Press','Hypertrophy',4,400,12,50,true,''],
  ['2026-07-11','Core','Plank','Endurance',3,0,60,25,true,'60sec holds'],
  ['2026-07-12','Cardio','Jump Rope','HIIT',5,0,3,20,true,'3min rounds'],
  ['2026-07-13','Glutes','Hip Thrust','Hypertrophy',4,225,12,40,true,'New PR!'],
  ['2026-07-14','Back','Lat Pulldown','Hypertrophy',3,150,12,35,true,''],
  ['2026-07-15','Arms','Tricep Extension','Hypertrophy',3,60,15,30,true,''],
  ['2026-07-16','Legs','Romanian Deadlift','Strength',3,200,10,45,true,''],
  ['2026-07-17','Chest','Bench Press','Strength',5,230,5,50,true,'New PR!'],
  ['2026-07-18','Cardio','Cycling','Cardio',1,0,45,45,true,'15mi ride'],
  ['2026-07-19','Full Body','Burpees','HIIT',4,0,15,30,true,''],
  ['2026-07-20','Shoulders','Overhead Press','Strength',4,140,6,40,true,''],
  ['2026-07-21','Back','Pull-up','Strength',4,0,8,35,true,'Bodyweight'],
  ['2026-07-22','Legs','Squat','Strength',4,280,5,55,true,'New PR!'],
  ['2026-07-23','Core','Russian Twist','Hypertrophy',3,25,20,25,true,''],
  ['2026-07-24','Full Body','Clean and Press','Circuit',3,95,8,45,true,''],
  ['2026-07-25','Mobility','Stretching','Stretching',1,0,30,30,true,'Foam roll'],
  ['2026-07-26','Mobility','Yoga','Yoga',1,0,60,60,true,'Rest day yoga'],
  ['2026-07-27','Chest','Bench Press','Hypertrophy',4,210,10,45,false,'Planned'],
  ['2026-07-28','Back','Deadlift','Strength',4,320,5,50,false,'Planned'],
  ['2026-07-29','Legs','Leg Press','Hypertrophy',4,420,10,45,false,'Planned'],
  ['2026-07-30','Shoulders','Overhead Press','Strength',4,145,5,40,false,'Planned'],
];

(async () => {
  // ── Headers ───────────────────────────────────────────────────────────────
  const headerValues = [
    [['📝'],['Workout Log','','','','','','','','','']],
    [['Date','Day','Muscle Group','Exercise','Workout Type','Sets','Weight (lbs)','Reps','Duration (min)','Completed','Notes']],
  ];

  await valuesBatchUpdate(id, [
    { range: "'📝 Workout Log'!A1", values: [['📝']] },
    { range: "'📝 Workout Log'!B1", values: [['Workout Log']] },
    { range: "'📝 Workout Log'!A2", values: [['Date','Day','Muscle Group','Exercise','Workout Type','Sets','Weight (lbs)','Reps','Duration (min)','Completed','Notes']] },
  ], 'wkl-headers');

  // ── Data rows (3-32) ─────────────────────────────────────────────────────
  const dataRows = rawData.map((row, i) => {
    const r = i + 3;
    return [
      row[0],                        // A: Date
      `=TEXT(A${r},"ddd")`,          // B: Day
      row[1],                        // C: Muscle Group
      row[2],                        // D: Exercise
      row[3],                        // E: Workout Type
      row[4],                        // F: Sets
      row[5],                        // G: Weight (lbs)
      row[6],                        // H: Reps
      row[7],                        // I: Duration (min)
      row[8],                        // J: Completed
      row[9],                        // K: Notes
    ];
  });

  await valuesBatchUpdate(id, [
    { range: "'📝 Workout Log'!A3", values: dataRows },
  ], 'wkl-data');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Merge B1:K1 for title
  reqs.push({ mergeCells: { range: gridRange(WKL, 0, 1, 1, 11), mergeType: 'MERGE_ALL' } });

  // Title row A1 (emoji cell)
  reqs.push({
    repeatCell: {
      range: gridRange(WKL, 0, 1, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.headerBlue),
          textFormat: { foregroundColor: hex(C.white), fontSize: 14, bold: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Title row B1:K1
  reqs.push({
    repeatCell: {
      range: gridRange(WKL, 0, 1, 1, 11),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.headerBlue),
          textFormat: { foregroundColor: hex(C.white), fontSize: 13, bold: true },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Header row A2:K2
  reqs.push({
    repeatCell: {
      range: gridRange(WKL, 1, 2, 0, 11),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.cobaltBlue),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Alternate data rows
  for (let r = 2; r < 32; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(WKL, r, r + 1, 0, 11),
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

  // Formula cells (B col = day, light blue bg)
  reqs.push({
    repeatCell: {
      range: gridRange(WKL, 2, 32, 1, 2),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.formulaBg),
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
    },
  });

  // Numeric alignment for F, G, H, I cols
  reqs.push({
    repeatCell: {
      range: gridRange(WKL, 2, 32, 5, 10),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(horizontalAlignment)',
    },
  });

  // Date format for col A
  reqs.push({
    repeatCell: {
      range: gridRange(WKL, 2, 32, 0, 1),
      cell: {
        userEnteredFormat: {
          numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
    },
  });

  // Row heights
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: WKL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 38 }, fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: WKL, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
      properties: { pixelSize: 28 }, fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: WKL, dimension: 'ROWS', startIndex: 2, endIndex: 32 },
      properties: { pixelSize: 22 }, fields: 'pixelSize',
    },
  });

  // Column widths: A=90, B=45, C=110, D=160, E=110, F=45, G=90, H=55, I=90, J=85, K=120
  const colWidths = [90, 45, 110, 160, 110, 45, 90, 55, 90, 85, 120];
  colWidths.forEach((w, ci) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: WKL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
        properties: { pixelSize: w }, fields: 'pixelSize',
      },
    });
  });

  await batchUpdate(id, reqs, 'wkl-format');
  console.log('Workout Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
