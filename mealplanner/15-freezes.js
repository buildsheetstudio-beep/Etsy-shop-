'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const DSH = sheetMap['📊 Dashboard'];
const WMP = sheetMap['📅 Weekly Meal Plan'];
const RCP = sheetMap['📖 Recipe Book'];
const GRL = sheetMap['🛒 Grocery List'];
const FIN = sheetMap['🥗 Food Inventory'];
const MPR = sheetMap['🍳 Meal Prep Schedule'];
const GBD = sheetMap['💰 Grocery Budget'];
const DRT = sheetMap['🚫 Dietary Restrictions'];
const NGO = sheetMap['🎯 Nutrition Goals'];

function freeze(sheetId, rows, cols) {
  return {
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { frozenRowCount: rows, frozenColumnCount: cols } },
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
    // Dashboard: freeze rows 1-3 (title + subtitle + spacer)
    freeze(DSH, 3, 0),
    // Weekly Meal Plan: freeze row 1 (title) + col A (standalone emoji, safe to freeze)
    freeze(WMP, 1, 1),
    // Recipe Book: freeze row 1 (title) + col A (recipe # standalone)
    freeze(RCP, 1, 1),
    // Grocery List: freeze row 1 (title) + col A (ingredient standalone)
    freeze(GRL, 1, 1),
    // Food Inventory: freeze row 1 (title) + col A (item standalone)
    freeze(FIN, 1, 1),
    // Meal Prep Schedule: freeze row 1 only (full-width merged title)
    freeze(MPR, 1, 0),
    // Grocery Budget: freeze row 1 only
    freeze(GBD, 1, 0),
    // Dietary Restrictions: freeze row 1 only
    freeze(DRT, 1, 0),
    // Nutrition Goals: freeze row 1 only
    freeze(NGO, 1, 0),
    // Reference Data: freeze header row, then hide
    freeze(REF, 1, 0),
    hideSheet(REF),
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
