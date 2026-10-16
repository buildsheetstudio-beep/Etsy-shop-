'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CAL = sheetMap['📅 Monthly Calendar'];

// Layout:
//   Row 1: A1="Year:" B1=2026(input) C1="Month:" D1=7(input) — linked to Dashboard in 10-dashboard.js
//   Row 2: day-of-week headers Sun–Sat (A2:G2)
//   Rows 3-8: 6×7 calendar grid (A3:G8) using LET+FILTER+TEXTJOIN formula
//   I1: helper = first Sunday of calendar grid

(async () => {
  // ── Static labels and inputs ──────────────────────────────────────────────
  await valuesBatchUpdate(id, [
    { range: "'📅 Monthly Calendar'!A1", values: [['Year:',2026,'Month:',7]] },
    { range: "'📅 Monthly Calendar'!A2", values: [['Sun','Mon','Tue','Wed','Thu','Fri','Sat']] },
    // Helper: first Sunday of the calendar grid
    { range: "'📅 Monthly Calendar'!I1", values: [['=DATE($B$1,$D$1,1)-WEEKDAY(DATE($B$1,$D$1,1),1)+1']] },
  ], 'cal-labels');

  // ── Calendar grid A3:G8 — all 42 cells share the same formula ──────────────
  // ROW()-3 = row offset (0–5), COLUMN()-1 = col offset (0–6)
  const calFormula = `=LET(dt,$I$1+(ROW()-3)*7+(COLUMN()-1),IF(MONTH(dt)=$D$1,DAY(dt)&CHAR(10)&IFERROR(TEXTJOIN(", ",TRUE,FILTER('📝 Workout Log'!$D$3:$D$32,'📝 Workout Log'!$A$3:$A$32=dt)),""),""))`;
  const calGrid = Array.from({ length: 6 }, () => Array(7).fill(calFormula));

  await valuesBatchUpdate(id, [
    { range: "'📅 Monthly Calendar'!A3", values: calGrid },
  ], 'cal-grid');

  // ── Formatting ────────────────────────────────────────────────────────────
  const reqs = [];

  // Row 1 labels/inputs
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 0, 1, 0, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.cobaltBlue),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Input cells B1, D1 — lighter blue bg to show they're editable
  [1, 3].forEach(ci => {
    reqs.push({
      repeatCell: {
        range: gridRange(CAL, 0, 1, ci, ci + 1),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(C.lightBlue),
            textFormat: { foregroundColor: hex(C.headerBlue), bold: true, fontSize: 12 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  });

  // Day-of-week header row A2:G2
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 1, 2, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.cobaltBlue),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Calendar grid A3:G8
  reqs.push({
    repeatCell: {
      range: gridRange(CAL, 2, 8, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.offWhite),
          textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 },
          horizontalAlignment: 'LEFT',
          verticalAlignment: 'TOP',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  // Borders on calendar grid
  const borderStyle = { style: 'SOLID', color: { red: 0.75, green: 0.75, blue: 0.75 } };
  reqs.push({
    updateBorders: {
      range: gridRange(CAL, 2, 8, 0, 7),
      top: borderStyle, bottom: borderStyle, left: borderStyle, right: borderStyle,
      innerHorizontal: borderStyle, innerVertical: borderStyle,
    },
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'ROWS', startIndex: 2, endIndex: 8 }, properties: { pixelSize: 70 }, fields: 'pixelSize' } });

  // Column widths A-G = 120px each
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: 0, endIndex: 7 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } });
  // Hide helper col I
  reqs.push({ updateDimensionProperties: { range: { sheetId: CAL, dimension: 'COLUMNS', startIndex: 8, endIndex: 9 }, properties: { pixelSize: 0, hiddenByUser: true }, fields: 'pixelSize,hiddenByUser' } });

  await batchUpdate(id, reqs, 'cal-format');
  console.log('Monthly Calendar complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
