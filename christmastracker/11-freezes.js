'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const DSH = sheetMap['📊 Dashboard'];
const GRT = sheetMap['🎁 Gift Recipient Tracker'];
const WSH = sheetMap['💡 Gift Idea Wishlist'];
const CDL = sheetMap['💌 Holiday Card & Mailing List'];
const RET = sheetMap['🔄 Return & Exchange Tracker'];

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
    // Dashboard: freeze rows 1-3 (title + subtitle + inputs)
    freeze(DSH, 3, 0),
    // Gift Recipient Tracker: freeze row 1 (headers) + col A (recipient name)
    freeze(GRT, 1, 1),
    // Gift Idea Wishlist: freeze row 1 (headers) + col A (recipient)
    freeze(WSH, 1, 1),
    // Holiday Card: freeze row 1 (headers) + col A (name)
    freeze(CDL, 1, 1),
    // Return Tracker: freeze row 1 (headers) + col A (recipient)
    freeze(RET, 1, 1),
    // Reference Data: freeze header row then hide
    freeze(REF, 1, 0),
    hideSheet(REF),
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
