'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Daily Workout Log'];
const S = "'Daily Workout Log'";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function fmtDate(d) { return d.toISOString().split('T')[0]; }

// 104 session dates: Mon + Thu each week, 52 weeks (Oct 6 2025 – Sep 28 2026)
const START = new Date('2025-10-06');
const SESSIONS = [];
for (let w = 0; w < 52; w++) {
  SESSIONS.push({ date: fmtDate(addDays(START, w * 7)),     wtype: w % 5 });
  SESSIONS.push({ date: fmtDate(addDays(START, w * 7 + 3)), wtype: (w * 2 + 1) % 5 });
}

// 5 workout types (rotating)
const WORKOUTS = [
  { type: 'Strength', name: 'Push Day',       status: 'Completed', intensity: 'Moderate',  rpe: 7, loc: 'Gym',
    exercises: [['EX-0001',135,10,'',''],['EX-0002',155, 8,'',''],['EX-0014',95,8,'',''],['EX-0027',50,15,'','']] },
  { type: 'Strength', name: 'Pull Day',       status: 'Completed', intensity: 'Hard',       rpe: 8, loc: 'Gym',
    exercises: [['EX-0007',225, 5,'',''],['EX-0008',135, 8,'',''],['EX-0009','',8,'',''],['EX-0020',65,10,'','']] },
  { type: 'Strength', name: 'Leg Day',        status: 'Completed', intensity: 'Hard',       rpe: 8, loc: 'Gym',
    exercises: [['EX-0036',185, 8,'',''],['EX-0042',155,10,'',''],['EX-0037',270,12,'',''],['EX-0046',135,12,'','']] },
  { type: 'Cardio',   name: 'Cardio Session', status: 'Completed', intensity: 'Moderate',   rpe: 6, loc: 'Outdoors',
    exercises: [['EX-0058','','',30,3.1],['EX-0063','','',15,''],['EX-0076',35,15,'',''],['EX-0030','','',1,'']] },
  { type: 'HIIT',     name: 'HIIT Circuit',   status: 'Completed', intensity: 'Very Hard',  rpe: 9, loc: 'Gym',
    exercises: [['EX-0056','',10,'',''],['EX-0055','',20,'',''],['EX-0076',35,15,'',''],['EX-0080',40,10,'','']] },
];

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,10100,0,30), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title (merged A:AD = cols 0:30, no frozenColumnCount — title spans all 30 cols)
  vals.push({ range: `${S}!A1`, values: [['DAILY WORKOUT LOG']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,30), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,30), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['One row per exercise set. Enter Exercise ID from Workout Setup tab. Shaded columns auto-fill via formula.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,30), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,30), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Quick stats
  const STATS = [
    ['Total Sessions Logged', `=SUMPRODUCT((LEN(B8:B5000)>0)*(P8:P5000=1)*IFERROR(1/COUNTIF(B8:B5000,B8:B5000),0))`],
    ['Total Sets Logged',     `=SUMPRODUCT((LEN(B8:B5000)>0)*1)`],
    ['Total Est. Volume (lb)',`=SUMPRODUCT(IFERROR(AB8:AB5000*1,0))`],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,0,5), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,5,9), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!F${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,5), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,5,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 6: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Row 7: Column headers
  vals.push({ range: `${S}!A7`, values: [['Log ID','Session ID','Date','Year','Month','Day','Status','Rest Day?','Workout Type','Workout Name','Exercise ID','Exercise Name','Primary Muscle','Secondary Muscle','Tracking Mode','Set #','Weight (lb)','Reps','Duration (min)','Distance','Speed / Pace','Rest (s)','RPE','Intensity','Location','Planned?','Completed?','Est. Volume','PR Value','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,30), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // Formula columns (rows 8+ = index 7+)
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(C8="","","LOG-"&TEXT(ROW()-7,"00000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,3,4), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(YEAR(C8),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,4,5), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(TEXT(C8,"mmmm"),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,5,6), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(TEXT(C8,"dddd"),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // L: Exercise Name (VLOOKUP from Workout Setup col B = col index 2)
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,11,12), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(VLOOKUP(K8,'Workout Setup'!$A$17:$N$1016,2,FALSE),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // M: Primary Muscle (col D = index 4)
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,12,13), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(VLOOKUP(K8,'Workout Setup'!$A$17:$N$1016,4,FALSE),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // N: Secondary Muscle (col E = index 5)
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,13,14), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(VLOOKUP(K8,'Workout Setup'!$A$17:$N$1016,5,FALSE),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // O: Tracking Mode (col F = index 6)
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,14,15), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(VLOOKUP(K8,'Workout Setup'!$A$17:$N$1016,6,FALSE),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // AB: Estimated Volume (Weight × Reps, only for Weight + Reps exercises)
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,27,28), cell: {
    userEnteredValue: { formulaValue: `=IF(O8="Weight + Reps",IFERROR(Q8*R8,""),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, numberFormat: { type: 'NUMBER', pattern: '#,##0' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // AC: PR Candidate Value (best measurable output for this tracking mode)
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,28,29), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(IF(O8="Weight + Reps",Q8*R8,IF(O8="Reps Only",R8,IF(O8="Time",S8,IF(O8="Distance",T8,"")))),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' }, numberFormat: { type: 'NUMBER', pattern: '#,##0.##' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Build sample data rows
  const dataRows = [];
  SESSIONS.forEach(({ date, wtype }, si) => {
    const workout = WORKOUTS[wtype];
    const sessionId = `SESSION-${date.replace(/-/g,'')}-01`;
    const delta = Math.floor(si / 20) * 5; // +5 lb every 20 sessions
    workout.exercises.forEach(([exId, baseW, reps, dur, dist]) => {
      const w = baseW !== '' ? baseW + delta : '';
      // 29 values: B(0) through AD(28)
      dataRows.push([
        sessionId, date, '', '', '',              // B C D(formula) E(formula) F(formula)
        workout.status, false,                     // G H(checkbox)
        workout.type, workout.name, exId,          // I J K
        '', '', '', '',                            // L M N O (formula cols)
        1, w, reps, dur, dist, '',               // P Q R S T U
        90, workout.rpe, workout.intensity, workout.loc, // V W X Y
        false, true,                               // Z(planned) AA(completed)
        '', '', '',                                // AB(formula) AC(formula) AD(notes)
      ]);
    });
  });

  // Push data chunks (B8:AD = 29 cols starting at B)
  for (let i = 0; i < dataRows.length; i += 500) {
    vals.push({ range: `${S}!B${8 + i}`, values: dataRows.slice(i, i + 500) });
  }

  // Alternating row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,10100,0,30),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format for column C
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,2,3), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [80,160,90,50,70,80,100,60,110,140,80,150,110,120,120,50,75,50,90,70,80,60,50,90,90,65,75,80,70,160];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1:7 only (no frozenColumnCount — title spans all 30 cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '04-log values');
  await batchUpdate(id, fmt, '04-log format');
  console.log(`✅  Daily Workout Log done — ${dataRows.length} sample rows.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
