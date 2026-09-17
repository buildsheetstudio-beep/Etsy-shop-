'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const DSH = sheetMap['📊 Dashboard'];
const FM  = sheetMap['👨‍👩‍👧‍👦 Family Members'];
const TV  = sheetMap['🌳 Family Tree / Relationship Viewer'];
const PH  = sheetMap['📸 Photos & Notes'];
const SC  = sheetMap['📜 Source & Citation Tracker'];
const DNA = sheetMap['🧬 DNA & Ethnicity Estimate'];
const MIG = sheetMap['🚢 Immigration & Migration Tracker'];
const LR  = sheetMap['📞 Living Relatives Directory'];

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
    freeze(DSH, 3, 0),          // Dashboard: title + family name + KPI header
    freeze(FM, 1, 1),           // Family Members: header + ID col
    freeze(TV, 1, 0),           // Relationship Viewer: ID input row
    freeze(PH, 1, 0),           // Photos & Notes: header only
    freeze(SC, 1, 1),           // Source & Citation: header + fact col
    freeze(DNA, 1, 0),          // DNA: header only
    freeze(MIG, 1, 0),          // Migration: header only
    freeze(LR, 1, 0),           // Living Relatives: header only
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
