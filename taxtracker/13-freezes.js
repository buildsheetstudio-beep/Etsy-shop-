'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];
const DSH = sheetMap['📊 Dashboard'];
const LOG = sheetMap['📝 Deduction Log'];
const MIL = sheetMap['🚗 Mileage Log'];
const HOM = sheetMap['🏠 Home Office'];
const QTR = sheetMap['📅 Quarterly Tax'];
const CHK = sheetMap['✅ Doc Checklist'];
const MYR = sheetMap['📈 Multi-Year'];

function freeze(sheetId, frozenRowCount, frozenColumnCount) {
  return { updateSheetProperties: {
    properties: { sheetId, gridProperties: { frozenRowCount, frozenColumnCount } },
    fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
  }};
}

(async () => {
  const reqs = [
    freeze(REF, 1, 0),  // Reference: header row only
    freeze(DSH, 3, 0),  // Dashboard: title + subtitle + section header
    freeze(LOG, 2, 0),  // Deduction Log: title + header (full-width merge prevents col freeze)
    freeze(MIL, 2, 0),  // Mileage Log: title + header
    freeze(HOM, 1, 0),  // Home Office: just title row
    freeze(QTR, 3, 0),  // Quarterly Tax: title + info + headers
    freeze(CHK, 3, 0),  // Doc Checklist: title + subtitle + headers
    freeze(MYR, 2, 0),  // Multi-Year: title + header
    // Hide Reference Data tab
    { updateSheetProperties: {
      properties: { sheetId: REF, hidden: true },
      fields: 'hidden',
    }},
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');

  // Print spreadsheet URL
  console.log(`\nSpreadsheet URL:\nhttps://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
