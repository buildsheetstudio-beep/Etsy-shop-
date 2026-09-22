'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DSH = sheetMap['✨ Dashboard'];
const CAL = sheetMap['📅 Monthly Calendar View'];
const BUD = sheetMap['💰 Budget & Expense Tracker'];
const TSK = sheetMap['✅ Task Manager'];
const HAB = sheetMap['🔁 Habit Tracker'];
const WHL = sheetMap['⚖️ Life Balance Wheel'];
const RFL = sheetMap['📝 Weekly & Monthly Reflection'];
const JRN = sheetMap['📓 Journal & Gratitude Log'];
const REF = sheetMap['📋 Reference Data'];

function freeze(sheetId, rows, cols) {
  return {
    updateSheetProperties: {
      properties: {
        sheetId,
        gridProperties: { frozenRowCount: rows, frozenColumnCount: cols }
      },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'
    }
  };
}

function hideSheet(sheetId) {
  return {
    updateSheetProperties: {
      properties: { sheetId, hidden: true },
      fields: 'hidden'
    }
  };
}

(async () => {
  const reqs = [];

  // Dashboard: freeze rows 1-4 (title, name, blank, section labels)
  reqs.push(freeze(DSH, 4, 0));

  // Monthly Calendar View: freeze rows 1-2 (title + month/year selectors)
  reqs.push(freeze(CAL, 2, 0));

  // Budget & Task & Habit: freeze row 1 + col A
  reqs.push(freeze(BUD, 1, 1));
  reqs.push(freeze(TSK, 1, 1));
  reqs.push(freeze(HAB, 1, 1));

  // Life Balance Wheel, Reflection, Journal: freeze row 1 only
  reqs.push(freeze(WHL, 1, 0));
  reqs.push(freeze(RFL, 1, 0));
  reqs.push(freeze(JRN, 1, 0));

  // Reference Data: freeze row 1, hide sheet
  reqs.push(freeze(REF, 1, 0));
  reqs.push(hideSheet(REF));

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes and hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
