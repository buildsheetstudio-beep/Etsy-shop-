'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const STK = sheetMap['🔥 Streak & Consistency Tracker'];

// Columns: A=Date, B=Valid Day (formula), C=Running Streak (formula)
const HEADERS = ['Date', 'Valid Day', 'Running Streak'];

const WL = "'💪 Workout Log'";

// 30 rows matching Workout Log dates (June 1-30, 2025)
// Valid Day = OR(Workout Log F=TRUE, Workout Log B="Rest")
// Row 21 (Jun 21) = Workout Log row 21 → Completed=FALSE, MG=Chest → NOT valid (not rest, not completed)
// Row 14 (Jun 14) = Rest day → valid

(async () => {
  // Build date + formula rows
  const valueData = [];
  const formulaRows = [];

  for (let day = 1; day <= 30; day++) {
    const wlRow = day + 1; // Workout Log row for day (row 2 = day 1)
    const stkRow = day + 1; // Streak row for day

    const dateFormula = `=DATE(2025,6,${day})`;
    // Valid Day: TRUE if completed OR rest
    const validFormula = `=IFERROR(OR(${WL}!F${wlRow}=TRUE,${WL}!B${wlRow}="Rest"),FALSE)`;
    // Running Streak: simple counter
    const streakFormula = day === 1
      ? `=IF(B2,1,0)`
      : `=IF(B${stkRow},C${stkRow-1}+1,0)`;

    formulaRows.push([dateFormula, validFormula, streakFormula]);
  }

  await valuesBatchUpdate(id, [
    { range: "'🔥 Streak & Consistency Tracker'!A1:C1", values: [HEADERS] },
    { range: "'🔥 Streak & Consistency Tracker'!A2:C31", values: formulaRows },
  ], 'stk-values');

  // Milestone badge labels (rows 33-35)
  await valuesBatchUpdate(id, [
    { range: "'🔥 Streak & Consistency Tracker'!A33:C35", values: [
      ['🔥 7-Day Streak',   '=MAX(C:C)>=7',   '=IF(MAX(C:C)>=7,"🏆 ACHIEVED!","Not yet — keep going!")'],
      ['🏅 30-Day Streak',  '=MAX(C:C)>=30',  '=IF(MAX(C:C)>=30,"🏆 ACHIEVED!","Not yet — you\'ve got this!")'],
      ['💯 100 Workouts',   `=COUNTIF(${WL}!F:F,TRUE)>=100`, `=IF(COUNTIF(${WL}!F:F,TRUE)>=100,"🏆 ACHIEVED!","Not yet — " & COUNTIF(${WL}!F:F,TRUE) & "/100 logged")`],
    ]} ,
  ], 'stk-badges');

  const reqs = [];

  // Base
  reqs.push({ repeatCell: { range: gridRange(STK,0,40,0,3), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite), textFormat: { foregroundColor: hex(C.navyDark) } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // Header
  reqs.push({ repeatCell: { range: gridRange(STK,0,1,0,3), cell: { userEnteredFormat: { backgroundColor: hex('#E65100'), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Alternating rows
  for (let i = 1; i < 31; i += 2) {
    reqs.push({ repeatCell: { range: gridRange(STK,i,i+1,0,3), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  }

  // Date col — YYYY-MM-DD
  reqs.push({ repeatCell: { range: gridRange(STK,1,31,0,1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Valid Day col (B) — formula tint, centered
  reqs.push({ repeatCell: { range: gridRange(STK,1,31,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)' } });

  // Running Streak col (C) — formula tint, centered, bold
  reqs.push({ repeatCell: { range: gridRange(STK,1,31,2,3), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });

  // Milestone badge section (rows 32-35, idx 31-34)
  // Divider row
  reqs.push({ repeatCell: { range: gridRange(STK,31,32,0,3), cell: { userEnteredFormat: { backgroundColor: hex(C.cobalt) } }, fields: 'userEnteredFormat(backgroundColor)' } });
  // Badge rows
  reqs.push({ repeatCell: { range: gridRange(STK,32,35,0,3), cell: { userEnteredFormat: { backgroundColor: hex('#FFF8DC'), textFormat: { bold: true, fontSize: 12 }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  // Badge label col A
  reqs.push({ repeatCell: { range: gridRange(STK,32,35,0,1), cell: { userEnteredFormat: { textFormat: { foregroundColor: hex('#E65100'), bold: true, fontSize: 13 } } }, fields: 'userEnteredFormat(textFormat)' } });
  // Badge TRUE/FALSE col B
  reqs.push({ repeatCell: { range: gridRange(STK,32,35,1,2), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });
  // Badge status col C
  reqs.push({ repeatCell: { range: gridRange(STK,32,35,2,3), cell: { userEnteredFormat: { textFormat: { foregroundColor: hex(C.sage) }, horizontalAlignment: 'LEFT' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: STK, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: STK, dimension: 'ROWS', startIndex: 1, endIndex: 31 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: STK, dimension: 'ROWS', startIndex: 31, endIndex: 32 }, properties: { pixelSize: 10 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: STK, dimension: 'ROWS', startIndex: 32, endIndex: 35 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Column widths
  [[0,120],[1,100],[2,250]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: STK, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'stk-format');
  console.log('Streak & Consistency Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
