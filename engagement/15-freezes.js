'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];
const DSH = sheetMap['🎉 Dashboard'];
const GST = sheetMap['👥 Guest List & RSVP'];
const BUD = sheetMap['💰 Budget & Vendor Tracker'];
const TML = sheetMap['⏰ Day-Of Timeline'];
const SEA = sheetMap['🪑 Seating Chart'];
const TST = sheetMap['🥂 Toast & Speech Planner'];
const PHO = sheetMap['📸 Photo & Shot List'];
const DEC = sheetMap['🌸 Decor, Food & Drinks'];
const CHK = sheetMap['✅ Party Checklist'];

function freeze(sheetId, frozenRowCount, frozenColumnCount) {
  return { updateSheetProperties: {
    properties: { sheetId, gridProperties: { frozenRowCount, frozenColumnCount } },
    fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
  }};
}

(async () => {
  const reqs = [
    freeze(REF, 1, 0),   // Reference: just header row
    freeze(DSH, 3, 0),   // Dashboard: title + subtitle + section header
    // Guest List: 3 rows + col A (standalone A in rows 0-1 allows col freeze)
    freeze(GST, 3, 1),
    freeze(BUD, 3, 0),   // Budget: 3 rows, 0 cols (full-width merges in rows 0-1)
    freeze(TML, 3, 0),   // Timeline: 3 rows
    // Seating Chart: 3 rows + col A (standalone A in rows 0-1 allows col freeze)
    freeze(SEA, 3, 1),
    freeze(TST, 3, 0),   // Toast Planner: 3 rows
    freeze(PHO, 3, 0),   // Photo List: 3 rows
    freeze(DEC, 3, 0),   // Decor: 3 rows
    freeze(CHK, 3, 0),   // Checklist: 3 rows
    // Hide Reference Data tab
    { updateSheetProperties: {
      properties: { sheetId: REF, hidden: true },
      fields: 'hidden',
    }},
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
