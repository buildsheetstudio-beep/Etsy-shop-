'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Chore & Allowance Tracker'];
const RWD = sheetMap['Reward & Prize Redemption'];
const REF = sheetMap['Reference Data'];

function freeze(sheetId, rows, cols) {
  return { updateSheetProperties: {
    properties: { sheetId, gridProperties: { frozenRowCount: rows, frozenColumnCount: cols } },
    fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount'
  } };
}

(async () => {
  const reqs = [];

  // ── Remove merges that cross col A boundary before freezing ───────────────
  // Title A1:E1 merge (added in 04-tracker-format.js) spans col A → must unmerge
  reqs.push({ unmergeCells: { range: gridRange(TAB, 0, 1, 0, 5) } });

  // Member header rows A:Q merges (10 rows at 22,29,36,...,85) also span col A
  for (let m = 0; m < 10; m++) {
    const headerRow = (22 + m * 7) - 1; // 0-indexed
    reqs.push({ unmergeCells: { range: gridRange(TAB, headerRow, headerRow + 1, 0, 17) } });
  }

  // ── Chore & Allowance Tracker: freeze rows 1-14 + col A ──────────────────
  reqs.push(freeze(TAB, 14, 1));

  // Reward & Prize Redemption: freeze header row only
  reqs.push(freeze(RWD, 1, 0));

  // Reference Data: freeze header row, then hide the sheet
  reqs.push(freeze(REF, 1, 0));
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true },
    fields: 'hidden'
  } });

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
