'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const DSH = sheetMap['📊 Dashboard — Annual Overview'];
const DSL = sheetMap['📅 Daily Symptom Log'];
const HMP = sheetMap['🌈 52-Week Severity Heatmap'];
const MED = sheetMap['💊 Medication & Treatment Tracker'];
const SLP = sheetMap['😴 Sleep Tracker'];
const TRG = sheetMap['⚡ Trigger & Stressor Log'];
const THR = sheetMap['🗓️ Therapy & Appointment Log'];
const SP  = sheetMap['🛡️ Safety Plan & Crisis Resources'];

function freeze(sheetId, rows, cols) {
  return {
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { frozenRowCount: rows, frozenColumnCount: cols } },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
    },
  };
}

function hide(sheetId) {
  return {
    updateSheetProperties: {
      properties: { sheetId, hidden: true },
      fields: 'hidden',
    },
  };
}

(async () => {
  const reqs = [
    // Reference: freeze row 1, hide sheet
    freeze(REF, 1, 0),
    hide(REF),

    // Dashboard: freeze 3 rows (title + input + section header)
    freeze(DSH, 3, 0),

    // Daily Symptom Log: freeze row 1 + col A
    freeze(DSL, 1, 1),

    // Heatmap: freeze row 1 only (merged cells in col A prevent column freeze)
    freeze(HMP, 1, 0),

    // Medication & Treatment Tracker: freeze row 1 + col A
    freeze(MED, 1, 1),

    // Sleep Tracker: freeze row 1 + col A
    freeze(SLP, 1, 1),

    // Trigger & Stressor Log: freeze row 1 + col A
    freeze(TRG, 1, 1),

    // Therapy & Appointment Log: freeze row 1 + col A
    freeze(THR, 1, 1),

    // Safety Plan: freeze row 1 only
    freeze(SP, 1, 0),
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
