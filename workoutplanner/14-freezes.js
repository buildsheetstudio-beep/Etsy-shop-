'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DSH = sheetMap['📊 Dashboard'];
const WKL = sheetMap['📝 Workout Log'];
const PRT = sheetMap['🏆 PR Tracker'];
const BMT = sheetMap['📏 Body Measurements'];
const PTL = sheetMap['📚 Program Templates'];
const CAR = sheetMap['🏃 Cardio Tracker'];
const REC = sheetMap['😴 Rest & Recovery'];
const CAL = sheetMap['📅 Monthly Calendar'];
const REF = sheetMap['📋 Reference Data'];

function freeze(sheetId, rows, cols) {
  return {
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { frozenRowCount: rows, frozenColumnCount: cols } },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
    },
  };
}

function hideSheet(sheetId) {
  return {
    updateSheetProperties: {
      properties: { sheetId, hidden: true },
      fields: 'hidden',
    },
  };
}

(async () => {
  const reqs = [
    // Dashboard: freeze rows 1-3 (title + subtitle + input row)
    freeze(DSH, 3, 0),
    // Workout Log: freeze rows 1-2 (title + headers) + col A (date always visible)
    freeze(WKL, 2, 1),
    // PR Tracker: freeze rows 1-2 (title + headers) + col A (exercise name always visible)
    freeze(PRT, 2, 1),
    // Body Measurements: freeze rows 1-2 + col A (date always visible)
    freeze(BMT, 2, 1),
    // Program Templates: freeze rows 1-2 only (no col A freeze — full-width title)
    freeze(PTL, 2, 0),
    // Cardio Tracker: freeze rows 1-2 + col A (date always visible)
    freeze(CAR, 2, 1),
    // Rest & Recovery: freeze rows 1-2 + col A (date always visible)
    freeze(REC, 2, 1),
    // Monthly Calendar: freeze rows 1-2 (inputs + day headers)
    freeze(CAL, 2, 0),
    // Reference Data: freeze header row then hide
    freeze(REF, 1, 0),
    hideSheet(REF),
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
