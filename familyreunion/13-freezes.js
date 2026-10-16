'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['🏡 Dashboard'];
const CKL = sheetMap['✅ Planning Checklist'];
const GST = sheetMap['👨‍👩‍👧‍👦 Guest List & RSVPs'];
const BDG = sheetMap['💰 Budget & Expenses'];
const ITN = sheetMap['🗓️ Event Itinerary & Activities'];
const FD  = sheetMap['🍖 Food, Drinks & Potluck'];
const ACC = sheetMap['🏠 Accommodation & Travel'];
const REF = sheetMap['📋 Reference Data'];

function freeze(sheetId, frozenRowCount, frozenColumnCount) {
  return { updateSheetProperties: {
    properties: { sheetId, gridProperties: { frozenRowCount, frozenColumnCount } },
    fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
  }};
}

(async () => {
  const reqs = [
    freeze(DSH, 2, 0),  // Dashboard: 2 rows, 0 cols
    freeze(CKL, 3, 0),  // Checklist: 3 rows, 0 cols
    freeze(GST, 3, 0),  // Guest List: 3 rows, 0 cols (row 1 has full-width merge, col freeze not possible)
    freeze(BDG, 2, 0),  // Budget: 2 rows, 0 cols
    freeze(ITN, 3, 0),  // Itinerary: 3 rows, 0 cols
    freeze(FD,  2, 0),  // Food: 2 rows, 0 cols
    freeze(ACC, 2, 0),  // Accommodation: 2 rows, 0 cols
    freeze(REF, 1, 0),  // Reference: 1 row, 0 cols
    // Hide Reference Data tab
    { updateSheetProperties: {
      properties: { sheetId: REF, hidden: true },
      fields: 'hidden',
    }},
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
