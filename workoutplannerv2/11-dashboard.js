'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];
const WL  = sheetMap['💪 Workout Log'];
const STK = sheetMap['🔥 Streak & Consistency Tracker'];

const WL_NAME  = "'💪 Workout Log'";
const STK_NAME = "'🔥 Streak & Consistency Tracker'";

(async () => {
  // ── Values ────────────────────────────────────────────────────────────────
  await valuesBatchUpdate(id, [
    // Title row 1
    { range: "'📊 Dashboard'!A1", values: [['🏋️ ULTIMATE WORKOUT PLANNER']] },

    // Selectors row 3: A3=Month label, B3=Month input, C3=Max HR, D3=Year label, E3=nothing, F3=Year
    { range: "'📊 Dashboard'!A3", values: [['Month (1-12):']] },
    { range: "'📊 Dashboard'!B3", values: [[6]] },
    { range: "'📊 Dashboard'!C3", values: [[185]] },  // Max Heart Rate — referenced by Cardio zone formula
    { range: "'📊 Dashboard'!D3", values: [['← Max HR (bpm)']] },
    { range: "'📊 Dashboard'!E3", values: [['Year:']] },
    { range: "'📊 Dashboard'!F3", values: [[2025]] },

    // Stat card labels row 5
    { range: "'📊 Dashboard'!B5", values: [['📅 Planned Days', '✅ Completed Days', '😴 Rest Days', '📈 Progress %']] },
    // Stat card values row 7
    { range: "'📊 Dashboard'!B7", values: [[
      `=IFERROR(COUNTIFS(${WL_NAME}!A:A,">="&DATE(F3,B3,1),${WL_NAME}!A:A,"<="&EOMONTH(DATE(F3,B3,1),0),${WL_NAME}!B:B,"<>Rest"),0)`,
      `=IFERROR(COUNTIFS(${WL_NAME}!A:A,">="&DATE(F3,B3,1),${WL_NAME}!A:A,"<="&EOMONTH(DATE(F3,B3,1),0),${WL_NAME}!F:F,TRUE),0)`,
      `=IFERROR(COUNTIFS(${WL_NAME}!A:A,">="&DATE(F3,B3,1),${WL_NAME}!A:A,"<="&EOMONTH(DATE(F3,B3,1),0),${WL_NAME}!B:B,"Rest"),0)`,
      `=IFERROR(IF(B7=0,"",C7/B7),"")`,
    ]] },

    // Stat card labels row 9
    { range: "'📊 Dashboard'!B9", values: [['🔥 Current Streak', '🏆 Longest Streak', '💪 Total Workouts', '⚡ Avg Duration (min)']] },
    // Stat card values row 11
    { range: "'📊 Dashboard'!B11", values: [[
      `=IFERROR(INDEX(${STK_NAME}!C:C,COUNTA(${STK_NAME}!C:C)),"")`,
      `=IFERROR(MAX(${STK_NAME}!C:C),"")`,
      `=IFERROR(COUNTIF(${WL_NAME}!F:F,TRUE),0)`,
      `=IFERROR(AVERAGEIF(${WL_NAME}!F:F,TRUE,${WL_NAME}!E:E),"")`,
    ]] },

    // Physician note row 13
    { range: "'📊 Dashboard'!A13", values: [['⚕️ This tracker is for personal organization only and isn\'t medical advice — check with a physician before starting a new exercise program, especially if you have an existing health condition.']] },

    // ── Chart helper data (below row 20) ─────────────────────────────────────
    // Progress donut data (rows 22-23)
    { range: "'📊 Dashboard'!A22", values: [['Status', 'Count']] },
    { range: "'📊 Dashboard'!A23", values: [['Completed', `=IFERROR(COUNTIFS(${WL_NAME}!A:A,">="&DATE(F3,B3,1),${WL_NAME}!A:A,"<="&EOMONTH(DATE(F3,B3,1),0),${WL_NAME}!F:F,TRUE),0)`]] },
    { range: "'📊 Dashboard'!A24", values: [['Not Yet Done', `=IFERROR(MAX(0,B7-B7+B7-C7),0)`]] },

    // Muscle group split data (rows 26-35)
    { range: "'📊 Dashboard'!A26", values: [['Muscle Group', 'Workouts']] },
    ...[['Chest'],['Back'],['Shoulders'],['Arms'],['Legs'],['Glutes'],['Core'],['Cardio'],['Rest']].map((mg, i) => ({
      range: `'📊 Dashboard'!A${27+i}`,
      values: [[mg[0], `=IFERROR(COUNTIF(${WL_NAME}!B:B,"${mg[0]}"),0)`]],
    })),

    // Daily progress data (rows 37-67, days 1-30 of selected month)
    { range: "'📊 Dashboard'!A37", values: [['Day', 'Completed']] },
    ...Array.from({length: 30}, (_,i) => ({
      range: `'📊 Dashboard'!A${38+i}`,
      values: [[i+1, `=IFERROR(COUNTIFS(${WL_NAME}!A:A,DATE(F3,B3,${i+1}),${WL_NAME}!F:F,TRUE),0)`]],
    })),
  ], 'dsh-values');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Base fill
  reqs.push({ repeatCell: { range: gridRange(DSH,0,70,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.navyDark) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Merges
  reqs.push({ mergeCells: { range: gridRange(DSH,0,1,0,8), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(DSH,12,13,0,8), mergeType: 'MERGE_ALL' } });

  // KPI label rows (4 & 8) — merges per pair
  for (const ri of [4, 8]) {
    reqs.push({ mergeCells: { range: gridRange(DSH,ri,ri+1,1,3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH,ri,ri+1,3,5), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH,ri,ri+1,5,7), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH,ri,ri+1,7,8), mergeType: 'MERGE_ALL' } });
  }
  for (const ri of [6, 10]) {
    reqs.push({ mergeCells: { range: gridRange(DSH,ri,ri+1,1,3), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH,ri,ri+1,3,5), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH,ri,ri+1,5,7), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(DSH,ri,ri+1,7,8), mergeType: 'MERGE_ALL' } });
  }

  // Title row
  reqs.push({ repeatCell: { range: gridRange(DSH,0,1,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.cobalt), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 22 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Selector row (idx 2)
  reqs.push({ repeatCell: { range: gridRange(DSH,2,3,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { foregroundColor: hex(C.navyDark) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  // Input cells B3=month (idx 2,1), C3=MaxHR (2,2), F3=year (2,5)
  reqs.push({ repeatCell: { range: gridRange(DSH,2,3,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.cobalt) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,2,3,2,3), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.brick) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(DSH,2,3,5,6), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.cobalt) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // KPI label rows (idx 4 & 8) — lime green
  for (const ri of [4, 8]) {
    reqs.push({ repeatCell: { range: gridRange(DSH,ri,ri+1,1,8), cell: { userEnteredFormat: { backgroundColor: hex(C.limeGreen), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });
  }

  // KPI value rows (idx 6 & 10) — navy dark bg, white text
  for (const ri of [6, 10]) {
    reqs.push({ repeatCell: { range: gridRange(DSH,ri,ri+1,1,8), cell: { userEnteredFormat: { backgroundColor: hex(C.navyDark), textFormat: { foregroundColor: hex(C.limeGreen), bold: true, fontSize: 28 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  }

  // Progress % — format as %
  reqs.push({ repeatCell: { range: gridRange(DSH,6,7,7,8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0%' } } }, fields: 'userEnteredFormat(numberFormat)' } });
  // Avg Duration — 0 decimal
  reqs.push({ repeatCell: { range: gridRange(DSH,10,11,7,8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0' } } }, fields: 'userEnteredFormat(numberFormat)' } });

  // Physician note (idx 12)
  reqs.push({ repeatCell: { range: gridRange(DSH,12,13,0,8), cell: { userEnteredFormat: { backgroundColor: hex('#F5F5F0'), textFormat: { foregroundColor: hex(C.cobalt), fontSize: 9, italic: true }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 64 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 64 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 64 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Column widths
  [[0,160],[1,140],[2,140],[3,140],[4,140],[5,140],[6,140],[7,140]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'dsh-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
