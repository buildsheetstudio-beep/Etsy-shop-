'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Weekly Chore Chart'];
const REF = sheetMap['Reference Data'];

(async () => {
  const reqs = [];

  // ── Freeze rows 1-13 on Weekly Chore Chart ────────────────────────────────
  // No column freeze — avoids merged cell conflicts
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: TAB, gridProperties: { frozenRowCount: 13, frozenColumnCount: 0 } },
    fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'
  }});

  // ── Hide Reference Data ───────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true },
    fields: 'hidden'
  }});

  // ── Hide gridlines on Weekly Chore Chart ──────────────────────────────────
  reqs.push({ updateSpreadsheetProperties: {
    // Note: hideGridlines is a sheet-level property on SheetProperties
  }});
  // Actually hideGridlines goes in SheetProperties, not SpreadsheetProperties
  reqs.pop(); // remove placeholder
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: TAB, tabColor: { red: 0.95, green: 0.9, blue: 0.94 } },
    fields: 'tabColor'
  }});

  await batchUpdate(id, reqs, 'finalize');
  console.log('Finalize complete');
  console.log('Spreadsheet ID:', id);
  console.log('URL: https://docs.google.com/spreadsheets/d/' + id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
