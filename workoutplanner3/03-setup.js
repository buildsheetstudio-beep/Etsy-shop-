'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Workout Setup'];
const S = "'Workout Setup'";

// 83 exercises: [Name, Category, Primary Muscle, Secondary Muscle, TrackingMode, Equipment, DefSets, DefReps, DefRestSec, DefDurMin, DefDist, Notes]
const EXERCISES = [
  // Chest (6)
  ['Bench Press',            'Strength',     'Chest',     'Triceps',    'Weight + Reps','Barbell',          3,10,90, '','',''],
  ['Incline Bench Press',    'Strength',     'Chest',     'Shoulders',  'Weight + Reps','Barbell',          3,10,90, '','',''],
  ['Dumbbell Flyes',         'Strength',     'Chest',     '',           'Weight + Reps','Dumbbells',        3,12,60, '','',''],
  ['Push-Ups',               'Bodyweight',   'Chest',     'Triceps',    'Reps Only',    'None / Bodyweight',3,15,45, '','',''],
  ['Cable Crossover',        'Strength',     'Chest',     '',           'Weight + Reps','Cable',            3,12,60, '','',''],
  ['Chest Dip',              'Bodyweight',   'Chest',     'Triceps',    'Reps Only',    'Pull-Up Bar',      3,12,60, '','',''],
  // Back (7)
  ['Deadlift',               'Strength',     'Back',      'Hamstrings', 'Weight + Reps','Barbell',          4,5, 180,'','',''],
  ['Barbell Row',            'Strength',     'Back',      'Biceps',     'Weight + Reps','Barbell',          4,8, 90, '','',''],
  ['Pull-Up',                'Bodyweight',   'Back',      'Biceps',     'Reps Only',    'Pull-Up Bar',      3,8, 90, '','',''],
  ['Lat Pulldown',           'Strength',     'Back',      'Biceps',     'Weight + Reps','Cable',            3,10,75, '','',''],
  ['Seated Cable Row',       'Strength',     'Back',      'Biceps',     'Weight + Reps','Cable',            3,10,75, '','',''],
  ['Dumbbell Row',           'Strength',     'Back',      'Biceps',     'Weight + Reps','Dumbbells',        3,10,60, '','',''],
  ['T-Bar Row',              'Strength',     'Back',      'Biceps',     'Weight + Reps','Barbell',          3,8, 90, '','',''],
  // Shoulders (6)
  ['Overhead Press',         'Strength',     'Shoulders', 'Triceps',    'Weight + Reps','Barbell',          4,8, 120,'','',''],
  ['Lateral Raise',          'Strength',     'Shoulders', '',           'Weight + Reps','Dumbbells',        3,15,60, '','',''],
  ['Front Raise',            'Strength',     'Shoulders', '',           'Weight + Reps','Dumbbells',        3,12,60, '','',''],
  ['Rear Delt Fly',          'Strength',     'Shoulders', 'Back',       'Weight + Reps','Dumbbells',        3,15,60, '','',''],
  ['Arnold Press',           'Strength',     'Shoulders', 'Triceps',    'Weight + Reps','Dumbbells',        3,10,90, '','',''],
  ['Cable Lateral Raise',    'Strength',     'Shoulders', '',           'Weight + Reps','Cable',            3,15,60, '','',''],
  // Biceps (5)
  ['Barbell Curl',           'Strength',     'Biceps',    '',           'Weight + Reps','Barbell',          3,10,60, '','',''],
  ['Dumbbell Curl',          'Strength',     'Biceps',    '',           'Weight + Reps','Dumbbells',        3,12,60, '','',''],
  ['Hammer Curl',            'Strength',     'Biceps',    'Forearms',   'Weight + Reps','Dumbbells',        3,12,60, '','',''],
  ['Preacher Curl',          'Strength',     'Biceps',    '',           'Weight + Reps','Barbell',          3,10,75, '','',''],
  ['Cable Curl',             'Strength',     'Biceps',    '',           'Weight + Reps','Cable',            3,12,60, '','',''],
  // Triceps (5)
  ['Tricep Dip',             'Bodyweight',   'Triceps',   'Chest',      'Reps Only',    'Pull-Up Bar',      3,12,60, '','',''],
  ['Skull Crusher',          'Strength',     'Triceps',   '',           'Weight + Reps','Barbell',          3,10,75, '','',''],
  ['Tricep Pushdown',        'Strength',     'Triceps',   '',           'Weight + Reps','Cable',            3,15,60, '','',''],
  ['Overhead Tricep Extension','Strength',   'Triceps',   '',           'Weight + Reps','Dumbbells',        3,12,60, '','',''],
  ['Close-Grip Bench Press', 'Strength',     'Triceps',   'Chest',      'Weight + Reps','Barbell',          3,8, 90, '','',''],
  // Core (6)
  ['Plank',                  'Bodyweight',   'Core',      'Abs',        'Time',         'None / Bodyweight',3,'',45, 1, '',''],
  ['Crunches',               'Bodyweight',   'Abs',       'Core',       'Reps Only',    'None / Bodyweight',3,20,45, '','',''],
  ['Russian Twist',          'Bodyweight',   'Core',      'Abs',        'Reps Only',    'None / Bodyweight',3,20,45, '','',''],
  ['Dead Bug',               'Bodyweight',   'Core',      '',           'Reps Only',    'None / Bodyweight',3,10,45, '','',''],
  ['Bicycle Crunch',         'Bodyweight',   'Abs',       'Core',       'Reps Only',    'None / Bodyweight',3,20,30, '','',''],
  ['Hanging Leg Raise',      'Bodyweight',   'Abs',       'Core',       'Reps Only',    'Pull-Up Bar',      3,10,60, '','',''],
  // Quads (6)
  ['Back Squat',             'Strength',     'Quads',     'Glutes',     'Weight + Reps','Barbell',          4,8, 180,'','',''],
  ['Leg Press',              'Strength',     'Quads',     'Glutes',     'Weight + Reps','Machine',          3,12,90, '','',''],
  ['Front Squat',            'Strength',     'Quads',     'Core',       'Weight + Reps','Barbell',          3,8, 120,'','',''],
  ['Leg Extension',          'Strength',     'Quads',     '',           'Weight + Reps','Machine',          3,15,60, '','',''],
  ['Bulgarian Split Squat',  'Strength',     'Quads',     'Glutes',     'Weight + Reps','Dumbbells',        3,10,90, '','',''],
  ['Step-Up',                'Strength',     'Quads',     'Glutes',     'Weight + Reps','Dumbbells',        3,12,60, '','',''],
  // Hamstrings (4)
  ['Romanian Deadlift',      'Strength',     'Hamstrings','Glutes',     'Weight + Reps','Barbell',          3,10,90, '','',''],
  ['Leg Curl',               'Strength',     'Hamstrings','',           'Weight + Reps','Machine',          3,12,60, '','',''],
  ['Good Morning',           'Strength',     'Hamstrings','Back',       'Weight + Reps','Barbell',          3,10,90, '','',''],
  ['Nordic Curl',            'Bodyweight',   'Hamstrings','',           'Reps Only',    'None / Bodyweight',3,6, 120,'','',''],
  // Glutes (4)
  ['Hip Thrust',             'Strength',     'Glutes',    'Hamstrings', 'Weight + Reps','Barbell',          3,12,90, '','',''],
  ['Glute Bridge',           'Bodyweight',   'Glutes',    'Core',       'Reps Only',    'None / Bodyweight',3,15,45, '','',''],
  ['Sumo Deadlift',          'Strength',     'Glutes',    'Back',       'Weight + Reps','Barbell',          4,6, 180,'','',''],
  ['Cable Kickback',         'Strength',     'Glutes',    '',           'Weight + Reps','Cable',            3,15,45, '','',''],
  // Calves (3)
  ['Standing Calf Raise',    'Strength',     'Calves',    '',           'Weight + Reps','Machine',          4,15,45, '','',''],
  ['Seated Calf Raise',      'Strength',     'Calves',    '',           'Weight + Reps','Machine',          3,15,45, '','',''],
  ['Single-Leg Calf Raise',  'Bodyweight',   'Calves',    '',           'Reps Only',    'None / Bodyweight',3,15,30, '','',''],
  // Bodyweight (5)
  ['Bodyweight Squat',       'Bodyweight',   'Quads',     'Glutes',     'Reps Only',    'None / Bodyweight',3,20,45, '','',''],
  ['Reverse Lunge',          'Bodyweight',   'Quads',     'Glutes',     'Reps Only',    'None / Bodyweight',3,12,45, '','',''],
  ['Mountain Climber',       'Bodyweight',   'Core',      'Quads',      'Reps Only',    'None / Bodyweight',3,20,30, '','',''],
  ['Burpee',                 'Bodyweight',   'Full Body', '',           'Reps Only',    'None / Bodyweight',3,10,60, '','',''],
  ['Box Jump',               'Bodyweight',   'Quads',     'Glutes',     'Reps Only',    'Bench',            3,8, 90, '','',''],
  // Cardio (10)
  ['Outdoor Run',            'Cardio',       'Full Body', '',           'Distance',     'None / Bodyweight','','','',30,3,'km or miles'],
  ['Treadmill Run',          'Cardio',       'Full Body', '',           'Distance',     'Treadmill',        '','','',30,3,''],
  ['Treadmill Walk',         'Cardio',       'Full Body', '',           'Distance',     'Treadmill',        '','','',30,2,''],
  ['Stationary Bike',        'Cardio',       'Full Body', '',           'Distance',     'Bike',             '','','',30,10,''],
  ['Rowing Machine',         'Cardio',       'Full Body', 'Back',       'Time',         'Rowing Machine',   '','','',20,'',''],
  ['Jump Rope',              'Cardio',       'Full Body', 'Calves',     'Time',         'None / Bodyweight','','','',15,'',''],
  ['Elliptical',             'Cardio',       'Full Body', '',           'Time',         'Cardio Machine',   '','','',30,'',''],
  ['Stair Climber',          'Cardio',       'Legs',      'Glutes',     'Time',         'Cardio Machine',   '','','',20,'',''],
  ['Cycling Outdoors',       'Cardio',       'Full Body', '',           'Distance',     'Bike',             '','','',45,15,''],
  ['Swimming Laps',          'Cardio',       'Full Body', 'Back',       'Distance',     'None / Bodyweight','','','',30,1,'km'],
  // Mobility / Flexibility (8)
  ['Hip Flexor Stretch',     'Mobility',     'Lower Body','',           'Time',         'None / Bodyweight','','','',1, '','Hold 60s each side'],
  ['Hamstring Stretch',      'Flexibility',  'Hamstrings','',           'Time',         'None / Bodyweight','','','',1, '','Hold 45s each side'],
  ['Chest Opener Stretch',   'Flexibility',  'Chest',     'Shoulders',  'Time',         'None / Bodyweight','','','',1, '','Hold 60s'],
  ['Thoracic Rotation',      'Mobility',     'Back',      'Core',       'Reps Only',    'None / Bodyweight',2,10,'', '','',''],
  ['Cat-Cow Stretch',        'Mobility',     'Back',      'Core',       'Reps Only',    'None / Bodyweight',2,10,'', '','',''],
  ['Hip Circles',            'Mobility',     'Lower Body','Core',       'Reps Only',    'None / Bodyweight',2,10,'', '','',''],
  ['Ankle Mobility Drill',   'Mobility',     'Calves',    '',           'Reps Only',    'None / Bodyweight',2,15,'', '','',''],
  ['Shoulder Mobility Circles','Mobility',   'Shoulders', '',           'Reps Only',    'None / Bodyweight',2,15,'', '','',''],
  // Conditioning / HIIT (8)
  ['Kettlebell Swing',       'Conditioning', 'Full Body', 'Glutes',     'Weight + Reps','Kettlebell',       4,15,60, '','',''],
  ['Battle Rope Waves',      'Conditioning', 'Upper Body','',           'Time',         'None / Bodyweight',4,'',45, 0.5,'',''],
  ['Med Ball Slam',          'Conditioning', 'Full Body', 'Core',       'Reps Only',    'None / Bodyweight',3,12,60, '','',''],
  ['Sled Push',              'Conditioning', 'Full Body', '',           'Distance',     'None / Bodyweight','','','',1, 20,''],
  ['Dumbbell Thruster',      'Conditioning', 'Full Body', '',           'Weight + Reps','Dumbbells',        3,10,90, '','',''],
  ['Box Jump Over',          'Conditioning', 'Full Body', '',           'Reps Only',    'Bench',            3,8, 90, '','',''],
  ['Jump Squat',             'Bodyweight',   'Quads',     'Glutes',     'Reps Only',    'None / Bodyweight',3,15,45, '','',''],
  ['Kettlebell Clean & Press','Conditioning','Full Body', 'Shoulders',  'Weight + Reps','Kettlebell',       3,8, 120,'','',''],
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Background ──────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID,0,1100,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Row 1: Title ─────────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['🏋️ WORKOUT SETUP & EXERCISE LIBRARY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // ── Row 2: Subtitle ───────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A2`, values: [['Configure your planner settings below, then add your exercises in the library starting at row 15.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // ── Row 3: Global Setup Header ────────────────────────────────────────────────
  vals.push({ range: `${S}!A3`, values: [['⚙️  GLOBAL SETTINGS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ── Setup fields (rows 4-14) ──────────────────────────────────────────────────
  const SETTINGS = [
    ['Planner Start Date',       '2025-10-01'],
    ['Planner Start Month',      'October'],
    ['Planner Start Year',       2025],
    ['Week Start Day',           'Monday'],
    ['Default Workout Location', 'Gym'],
    ['Preferred Weight Unit',    'lb'],
    ['Preferred Distance Unit',  'miles'],
    ['Weekly Workout Goal',      4],
    ['Monthly Workout Goal',     16],
    ['Streak Target',            30],
    ['Default Rest Day(s)',      'Saturday, Sunday'],
    ['Selected Reporting Year',  2026],
    ['Selected Reporting Month', 'September'],
  ];
  SETTINGS.forEach(([label, val], i) => {
    const r = 4 + i;
    vals.push({ range: `${S}!A${r}:B${r}`, values: [[label, val]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial' },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // ── Library section header (row 15) ──────────────────────────────────────────
  vals.push({ range: `${S}!A15`, values: [['📚  EXERCISE LIBRARY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,14,15,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,14,15,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 14, endIndex: 15 }, properties: { pixelSize: 28 }, fields: 'pixelSize' }});

  // ── Column headers (row 16, index 15) ────────────────────────────────────────
  const HDRS = ['Exercise ID','Exercise Name','Category','Primary Muscle','Secondary Muscle','Tracking Mode','Equipment','Default Sets','Default Reps','Default Rest (s)','Default Duration (min)','Default Distance','Active?','Notes'];
  vals.push({ range: `${S}!A16`, values: [HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID,15,16,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 15, endIndex: 16 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // ── Exercise ID formula rows 17-1016 (data rows 17-1016 = index 16-1015)
  // Exercise library data starts at row 17 (spreadsheet row) = index 16
  // The formula row in Exercise Library refers to row 17 onwards
  // But the spec says Exercise Library rows 15:1014 for data
  // Let me re-check: "Create rows 15:1014" for library
  // Row 15 in the spec = our index 14 (which is the section header)
  // Actually the spec says data starts at row 15 of the Workout Setup
  // But we have: row 1=title, rows 2=subtitle, row 3=settings header, rows 4-14=settings, row 15=library header, row 16=column headers, rows 17+=exercise data
  // So data starts at row 17 (index 16), Exercise ID = EX-TEXT(ROW()-16,"0000")
  fmt.push({ repeatCell: { range: gridRange(SID,16,1016,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B17="","","EX-"&TEXT(ROW()-16,"0000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // ── Push exercise data ────────────────────────────────────────────────────────
  EXERCISES.forEach((ex, i) => {
    const row = 17 + i;
    vals.push({ range: `${S}!B${row}:N${row}`, values: [[...ex, true]] }); // B-M = exercise fields, N = Active? (TRUE)
  });

  // ── Alternating row bg for library area ───────────────────────────────────────
  for (let i = 0; i < 200; i++) {
    const bg = i % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,16+i,17+i,1,14), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }

  // ── Column widths ─────────────────────────────────────────────────────────────
  const WIDTHS = [90, 200, 110, 120, 120, 120, 130, 80, 80, 90, 100, 90, 60, 200];
  WIDTHS.forEach((w, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }});
  });

  // ── Freeze rows 1:16 only (no column freeze — title spans all 16 cols) ────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 16 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '03-setup values');
  await batchUpdate(id, fmt, '03-setup format');
  console.log(`✅  Workout Setup done — ${EXERCISES.length} exercises loaded.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
