'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MG_COLORS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CAL = sheetMap['📅 Monthly Calendar View'];

// Day-of-week headers in row 3 (after selectors in rows 1-2)
const DOW = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

// For a 7×6 grid starting at row 4 (idx 3), cols A-G
// June 2025: starts on Sunday (index 0) — day 1 is col 0
// Calendar layout for June 2025:
// Week 1: [1,2,3,4,5,6,7]
// Week 2: [8,9,10,11,12,13,14]
// Week 3: [15,16,17,18,19,20,21]
// Week 4: [22,23,24,25,26,27,28]
// Week 5: [29,30,_,_,_,_,_]

const WL_NAME = "'💪 Workout Log'";

function cellFormula(dayNum) {
  if (!dayNum) return '';
  return `=IFERROR(TEXTJOIN(", ",TRUE,FILTER(${WL_NAME}!B:B,${WL_NAME}!A:A=DATE($B$1,$B$2,${dayNum}))),"")`;
}

(async () => {
  // Month/Year selectors + day-of-week headers + day numbers + formulas
  const data = [];

  // Row 1: Month/Year selector labels + values
  data.push({ range: "'📅 Monthly Calendar View'!A1:B1", values: [['Year', 2025]] });
  data.push({ range: "'📅 Monthly Calendar View'!A2:B2", values: [['Month', 6]] });

  // Row 3: DOW headers
  data.push({ range: "'📅 Monthly Calendar View'!A3:G3", values: [DOW] });

  // Day number rows (rows 4-9, idx 3-8) + formula rows (rows interleaved)
  // Each calendar "row" = 2 sheet rows: top row = day number, bottom row = formula
  // But for simplicity let's use 1 row per day with day# in the top part and formula below
  // Actually let's use the simpler single-row approach: day# label + formula in same cell via helper
  // We'll do: row 4 = week 1 day numbers, row 5 = week 1 formulas, etc.

  const weeks = [
    [1,2,3,4,5,6,7],
    [8,9,10,11,12,13,14],
    [15,16,17,18,19,20,21],
    [22,23,24,25,26,27,28],
    [29,30,null,null,null,null,null],
  ];

  weeks.forEach((week, wi) => {
    const labelRow = 4 + wi * 2;       // rows 4,6,8,10,12 (1-indexed)
    const formulaRow = 5 + wi * 2;     // rows 5,7,9,11,13

    const labels = week.map(d => d || '');
    const formulas = week.map(d => cellFormula(d));

    data.push({ range: `'📅 Monthly Calendar View'!A${labelRow}:G${labelRow}`, values: [labels] });
    data.push({ range: `'📅 Monthly Calendar View'!A${formulaRow}:G${formulaRow}`, values: [formulas] });
  });

  await valuesBatchUpdate(id, data, 'cal-values');

  const reqs = [];

  // Base fill
  reqs.push({ repeatCell: { range: gridRange(CAL,0,20,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.offWhite) } }, fields: 'userEnteredFormat(backgroundColor)' } });

  // Selector area (rows 1-2)
  reqs.push({ repeatCell: { range: gridRange(CAL,0,2,0,2), cell: { userEnteredFormat: { backgroundColor: hex(C.cobalt), textFormat: { foregroundColor: hex(C.white), bold: true }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)' } });
  // selector input cells
  reqs.push({ repeatCell: { range: gridRange(CAL,0,1,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { foregroundColor: hex(C.navyDark), bold: true, fontSize: 13 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(CAL,1,2,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { foregroundColor: hex(C.navyDark), bold: true, fontSize: 13 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });

  // DOW header row
  reqs.push({ repeatCell: { range: gridRange(CAL,2,3,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.cobalt), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });

  // Day number rows (idx 3,5,7,9,11)
  [3,5,7,9,11].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(CAL,ri,ri+1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.cobalt), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  });

  // Formula rows (idx 4,6,8,10,12) — workout content
  [4,6,8,10,12].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(CAL,ri,ri+1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { foregroundColor: hex(C.navyDark), fontSize: 9, italic: true }, wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' } });
  });

  // Row heights: selectors=28, DOW header=28, day# rows=22, formula rows=54 each
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 0, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  [3,5,7,9,11].forEach(ri => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });
  });
  [4,6,8,10,12].forEach(ri => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  });

  // Column widths — equal 7 cols
  for (let ci = 0; ci < 7; ci++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: 148 }, fields: 'pixelSize' } });
  }

  await batchUpdate(id, reqs, 'cal-format');
  console.log('Monthly Calendar View complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
