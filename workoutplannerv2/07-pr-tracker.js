'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PR = sheetMap['🏆 Personal Records (PR) Tracker'];

// Columns: A=Exercise Name, B=Category, C=PR Type, D=Value, E=Unit, F=Date Achieved, G=Current Best (MAXIFS formula), H=Notes
const HEADERS = ['Exercise Name', 'Category', 'PR Type', 'Value', 'Unit', 'Date Achieved', 'Current Best', 'Notes'];

// 12+ PRs — bench press and running have multiple historical entries
const ROWS = [
  ['Bench Press', 'Strength', '1-Rep Max',       80,  'kg', '=DATE(2025,3,15)', '', 'First time hitting 80kg!'],
  ['Bench Press', 'Strength', '1-Rep Max',       85,  'kg', '=DATE(2025,5,1)',  '', 'Slow progress but consistent'],
  ['Bench Press', 'Strength', '1-Rep Max',       90,  'kg', '=DATE(2025,6,1)',  '', 'New PR — felt amazing'],
  ['Squat',       'Strength', '1-Rep Max',       110, 'kg', '=DATE(2025,4,10)', '', 'Paused squat 1RM'],
  ['Squat',       'Strength', '1-Rep Max',       120, 'kg', '=DATE(2025,6,4)',  '', 'Belt + knee sleeves'],
  ['Deadlift',    'Strength', '1-Rep Max',       140, 'kg', '=DATE(2025,5,20)', '', ''],
  ['Pull-ups',    'Strength', 'Max Reps',         12, 'reps','=DATE(2025,4,5)',  '', 'Strict form, no kipping'],
  ['Pull-ups',    'Strength', 'Max Reps',         15, 'reps','=DATE(2025,6,10)', '', 'New max reps PR!'],
  ['Hip Thrust',  'Strength', '1-Rep Max',       100, 'kg', '=DATE(2025,6,15)', '', 'First time hitting 100kg'],
  ['5km Run',     'Cardio',   'Best Time',       '28:45', 'mm:ss', '=DATE(2025,4,20)', '', '5:45/km average'],
  ['5km Run',     'Cardio',   'Best Time',       '27:30', 'mm:ss', '=DATE(2025,6,18)', '', '5:30/km — new PB!'],
  ['10km Run',    'Cardio',   'Best Time',       '58:20', 'mm:ss', '=DATE(2025,5,10)', '', '5:50/km average pace'],
  ['Rowing 2000m','Cardio',   'Best Time',       '7:42',  'mm:ss', '=DATE(2025,6,24)', '', '2:18/500m split'],
];

(async () => {
  // Build rows with MAXIFS formula in col G
  const dataRows = ROWS.map((row, i) => {
    const r = i + 2;
    // MAXIFS only works on numeric values; for time PRs use 0 as placeholder
    const isNumeric = typeof row[3] === 'number';
    const maxifsFormula = isNumeric
      ? `=IFERROR(MAXIFS($D$2:$D$200,$A$2:$A$200,A${r},$C$2:$C$200,C${r}),"")`
      : `=IFERROR(MINIFS($D$2:$D$200,$A$2:$A$200,A${r},$C$2:$C$200,C${r}),"")`;
    return [row[0], row[1], row[2], row[3], row[4], row[5], maxifsFormula, row[7]];
  });

  await valuesBatchUpdate(id, [
    { range: "'🏆 Personal Records (PR) Tracker'!A1:H1", values: [HEADERS] },
    { range: "'🏆 Personal Records (PR) Tracker'!A2:H14", values: dataRows },
  ], 'pr-values');

  const reqs = [];

  // Base
  reqs.push({ repeatCell: { range: gridRange(PR,0,30,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.navyDark) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Header
  reqs.push({ repeatCell: { range: gridRange(PR,0,1,0,8), cell: { userEnteredFormat: { backgroundColor: hex('#C9A227'), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Alternating
  for (let i = 1; i < 14; i += 2) {
    reqs.push({ repeatCell: { range: gridRange(PR,i,i+1,0,8), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  }

  // Exercise name — bold
  reqs.push({ repeatCell: { range: gridRange(PR,1,14,0,1), cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.navyDark) } } }, fields: 'userEnteredFormat(textFormat)' } });

  // Date col F — date format
  reqs.push({ repeatCell: { range: gridRange(PR,1,14,5,6), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Current Best (col G) — formula/read-only tint, bold gold
  reqs.push({ repeatCell: { range: gridRange(PR,1,14,6,7), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true, foregroundColor: hex('#C9A227') }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // Category, PR Type, Value, Unit — centered
  [1,2,3,4].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(PR,1,14,ci,ci+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });
  });

  // Input bg for entry cols
  reqs.push({ repeatCell: { range: gridRange(PR,1,14,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  reqs.push({ repeatCell: { range: gridRange(PR,1,14,7,8), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(backgroundColor)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: PR, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: PR, dimension: 'ROWS', startIndex: 1, endIndex: 14 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,150],[1,100],[2,120],[3,80],[4,80],[5,120],[6,110],[7,200]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: PR, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'pr-format');
  console.log('PR Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
