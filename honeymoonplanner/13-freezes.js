'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const OVW = sheetMap['🌅 Honeymoon Overview'];
const ITN = sheetMap['🗓️ Multi-City Day-by-Day Itinerary'];
const BGT = sheetMap['💰 Budget & Expense Tracker'];
const VLT = sheetMap['🔐 Travel Document & Confirmation Vault'];
const WIS = sheetMap['✨ Activity & Excursion Wishlist'];
const CHK = sheetMap['✅ Research & Reservation Checklist'];
const PKG = sheetMap['🧳 Packing List'];

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
    // Overview: freeze rows 1-3 (title + subtitle + inputs)
    freeze(OVW, 3, 0),
    // Itinerary: freeze row 1 (header) + col A (date) + col B (city)
    freeze(ITN, 1, 2),
    // Budget: freeze row 1 + col A (category)
    freeze(BGT, 1, 1),
    // Document Vault: freeze row 1 + col A (type)
    freeze(VLT, 1, 1),
    // Activity Wishlist: freeze row 1 + col A (activity)
    freeze(WIS, 1, 1),
    // Checklist: freeze row 1 + col A (task)
    freeze(CHK, 1, 1),
    // Packing: freeze row 1 + col A (item)
    freeze(PKG, 1, 1),
    // Reference Data: freeze row 1, then hide
    freeze(REF, 1, 0),
    hideSheet(REF),
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
