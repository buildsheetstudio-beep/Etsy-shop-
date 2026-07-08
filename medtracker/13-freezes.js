'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['📊 Dashboard'];
const EXP  = sheetMap['🧾 Expense Log'];
const HSA  = sheetMap['💳 HSA/FSA Tracker'];
const INS  = sheetMap['📋 Insurance Claims'];
const PRV  = sheetMap['🏥 Provider Log'];
const RX   = sheetMap['💊 Prescription Tracker'];
const PYR  = sheetMap['📈 Prior-Year Comparison'];
const REF  = sheetMap['📋 Reference Data'];

function freeze(sheetId, rows, cols) {
  return {
    updateSheetProperties: {
      properties: {
        sheetId,
        gridProperties: { frozenRowCount: rows, frozenColumnCount: cols },
      },
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
    // Dashboard: freeze title + subtitle + inputs (3 rows), no col freeze
    freeze(DASH, 3, 0),
    // Expense Log: freeze title + header (2 rows), freeze col A (col A standalone = safe)
    freeze(EXP, 2, 1),
    // HSA/FSA: freeze title + header (2 rows), no col freeze
    freeze(HSA, 2, 0),
    // Insurance Claims: freeze title + header (2 rows), freeze col A (standalone = safe)
    freeze(INS, 2, 1),
    // Provider Log: freeze title + header (2 rows), no col freeze (full-width merge)
    freeze(PRV, 2, 0),
    // Prescription Tracker: freeze title + header (2 rows), no col freeze (full-width merge)
    freeze(RX, 2, 0),
    // Prior-Year Comparison: freeze title + header (2 rows), freeze col A (standalone = safe)
    freeze(PYR, 2, 1),
    // Reference Data: freeze header row only, no col freeze, then hide
    freeze(REF, 1, 0),
    hideSheet(REF),
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
