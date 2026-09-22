'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DASH = sheetMap['🏆 Goal Dashboard'];
const GL   = sheetMap['🎯 Goals & Milestones'];
const ACT  = sheetMap['📋 Action Steps'];
const WKL  = sheetMap['📅 Weekly Check-In & Reflection'];
const REF  = sheetMap['📋 Reference Data'];

function freeze(sheetId, rows, cols) {
  return { updateSheetProperties: {
    properties: { sheetId, gridProperties: { frozenRowCount: rows, frozenColumnCount: cols } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }};
}

(async () => {
  const reqs = [];

  // Goal Dashboard: freeze 2 rows (title + input/KPI row), 0 cols
  reqs.push(freeze(DASH, 2, 0));

  // Goals & Milestones: freeze 2 rows (title + headers), 1 col (A = goal name)
  // Safe: A0 is standalone, B0:K0 is the merged title — no merge spans col A+B
  reqs.push(freeze(GL, 2, 1));

  // Action Steps: freeze 4 rows (title + summary header + stats + col headers), 0 cols
  // (Row 1 has a full-width merge from col A, which conflicts with col A freeze)
  reqs.push(freeze(ACT, 4, 0));

  // Weekly Check-In: freeze 2 rows (title + headers), 1 col
  reqs.push(freeze(WKL, 2, 1));

  // Reference Data: freeze 2 rows, 0 cols, hidden
  reqs.push(freeze(REF, 2, 0));
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true },
    fields: 'hidden',
  }});

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hidden tab complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
