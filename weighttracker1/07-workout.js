'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Workout Tracker'];
const S = "'Workout Tracker'";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function fmtDate(d) { return d.toISOString().split('T')[0]; }

// ~100 workouts: 2-3 per week over ~40 weeks (Oct 2025 - Aug 2026)
const WORKOUT_TYPES = [
  { type: 'Strength',        name: 'Upper Body Push',    intensity: 'Moderate', dur: 50 },
  { type: 'Strength',        name: 'Lower Body',         intensity: 'Hard',     dur: 55 },
  { type: 'Strength',        name: 'Upper Body Pull',    intensity: 'Moderate', dur: 50 },
  { type: 'Cardio',          name: 'Brisk Walk',         intensity: 'Light',    dur: 45 },
  { type: 'Cardio',          name: 'Cycling',            intensity: 'Moderate', dur: 40 },
  { type: 'Yoga/Flexibility','name': 'Morning Yoga',     intensity: 'Very Light',dur: 30 },
  { type: 'Cardio',          name: 'Jog/Run',            intensity: 'Hard',     dur: 35 },
  { type: 'Strength',        name: 'Full Body Circuit',  intensity: 'Hard',     dur: 45 },
];
const FEELINGS = ['Felt great!', 'Good energy', 'Solid session', 'A bit tired but finished strong', 'Low energy today', 'Felt strong', 'Challenging but worth it', 'Needed this'];

const START = new Date('2025-10-06');
const workoutRows = [];
let wIdx = 0;
for (let week = 0; week < 40; week++) {
  // Mon and Thu each week
  [0, 3].forEach((offset, di) => {
    if (workoutRows.length < 104) {
      const date = fmtDate(addDays(START, week * 7 + offset));
      const w = WORKOUT_TYPES[wIdx % WORKOUT_TYPES.length];
      const feeling = FEELINGS[(week + di) % FEELINGS.length];
      workoutRows.push([date, w.type, w.name, w.dur, w.intensity, feeling, '']);
      wIdx++;
    }
  });
  // Extra Saturday session every other week
  if (week % 2 === 0 && workoutRows.length < 104) {
    const date = fmtDate(addDays(START, week * 7 + 5));
    const w = WORKOUT_TYPES[(wIdx + 2) % WORKOUT_TYPES.length];
    workoutRows.push([date, w.type, w.name, w.dur, 'Light', 'Easy day', '']);
    wIdx++;
  }
}

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,10100,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['WORKOUT TRACKER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Log workouts by type, duration, and intensity. Use the shaded Log ID column for reference in other tabs.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Stats
  const STATS = [
    ['Sessions Logged',      `=COUNTA(B8:B5000)`],
    ['Total Duration (min)', `=IFERROR(SUM(E8:E5000),"")`],
    ['Unique Workout Types', `=IFERROR(SUMPRODUCT(1/COUNTIF(C8:C5000,C8:C5000)*(C8:C5000<>"")),"")` ],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,0,3), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,3,8), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!D${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,3,8), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 6: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Row 7: headers
  vals.push({ range: `${S}!A7`, values: [['Log ID','Date','Workout Type','Workout Name','Duration (min)','Intensity','How I Felt','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // A: Log ID formula
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B8="","","WO-"&TEXT(ROW()-7,"000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push workout data (B=date, C=type, D=name, E=dur, F=intensity, G=feeling, H=notes)
  vals.push({ range: `${S}!B8`, values: workoutRows });

  // Row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,10100,0,8),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format col B
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [80, 100, 120, 150, 90, 100, 180, 200];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '07-workout values');
  await batchUpdate(id, fmt, '07-workout format');
  console.log(`✅  Workout Tracker done — ${workoutRows.length} sessions.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
