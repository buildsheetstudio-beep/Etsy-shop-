'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['🎓 Dashboard'];
const TRK = sheetMap['📋 Application Tracker'];
const DL  = sheetMap['⏰ Deadlines & To-Do'];
const RES = sheetMap['🏆 Results & Awards Log'];
const REF = sheetMap['📋 Reference Data'];

function freeze(sheetId, frozenRowCount, frozenColumnCount) {
  return { updateSheetProperties: {
    properties: { sheetId, gridProperties: { frozenRowCount, frozenColumnCount } },
    fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
  }};
}

(async () => {
  const reqs = [
    freeze(DSH, 2, 0),  // Dashboard: 2 rows (title + gold strip), 0 cols
    freeze(TRK, 3, 0),  // Application Tracker: 3 rows, 0 cols (full-width merges at rows 0-1 prevent col freeze)
    freeze(DL,  3, 0),  // Deadlines: 3 rows (title + subheader + note), 0 cols
    freeze(RES, 3, 0),  // Results: 3 rows, 0 cols
    freeze(REF, 1, 0),  // Reference: 1 row, 0 cols
    // Hide Reference tab
    { updateSheetProperties: {
      properties: { sheetId: REF, hidden: true },
      fields: 'hidden',
    }},
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
