'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const DSH = sheetMap['📊 Dashboard'];
const WL  = sheetMap['💪 Workout Log'];
const CAL = sheetMap['📅 Monthly Calendar View'];
const WDL = sheetMap['📝 Weekly Detail Log'];
const SS  = sheetMap['⚡ Superset & Circuit Builder'];
const PR  = sheetMap['🏆 Personal Records (PR) Tracker'];
const CRD = sheetMap['❤️ Cardio & Heart Rate Zone Tracker'];
const BDY = sheetMap['📏 Body Measurements & Progress Photos'];
const STK = sheetMap['🔥 Streak & Consistency Tracker'];

function freeze(sheetId, rows, cols) {
  return {
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { frozenRowCount: rows, frozenColumnCount: cols } },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
    },
  };
}

function hideSheet(sheetId) {
  return { updateSheetProperties: { properties: { sheetId, hidden: true }, fields: 'hidden' } };
}

(async () => {
  const reqs = [
    freeze(REF, 1, 0),
    hideSheet(REF),
    freeze(DSH, 3, 0),                // Dashboard: rows 1-3
    freeze(WL, 1, 1),                 // Workout Log: header + col A
    freeze(CAL, 2, 0),                // Monthly Calendar: rows 1-2 (selector + DOW)
    freeze(WDL, 1, 1),                // Weekly Detail Log: header + col A
    freeze(SS, 1, 1),                 // Superset Builder: header + col A
    freeze(PR, 1, 1),                 // PR Tracker: header + col A
    freeze(CRD, 1, 1),               // Cardio Tracker: header + col A
    freeze(BDY, 1, 0),               // Body Measurements: header only
    freeze(STK, 1, 1),               // Streak: header + col A
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
