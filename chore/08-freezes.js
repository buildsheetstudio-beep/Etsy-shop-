'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Chore & Allowance Tracker'];
const RWD = sheetMap['Reward & Prize Redemption'];
const REF = sheetMap['Reference Data'];

function freeze(sheetId, rows, cols) {
  return {
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { frozenRowCount: rows, frozenColumnCount: cols } },
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

  // Chore & Allowance Tracker: freeze rows 1-21 (Block A + Block B), no col freeze
  // (col freeze conflicts with merged label row A23:F23)
  reqs.push(freeze(TAB, 21, 0));

  // Reward & Prize Redemption: freeze row 1 only
  reqs.push(freeze(RWD, 1, 0));

  // Reference Data: freeze row 1, hide
  reqs.push(freeze(REF, 1, 0));
  reqs.push(hideSheet(REF));

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes and hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
