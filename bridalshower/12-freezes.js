'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🥂 Dashboard'];
const GST  = sheetMap['👰 Guest List & Seating'];
const BUD  = sheetMap['💰 Budget & Expenses'];
const ITN  = sheetMap['🗓️ Party Itinerary & Games'];
const VEN  = sheetMap['💍 Venue, Vendors & Bridesmaids'];
const GFT  = sheetMap['🎁 Gift Tracker & Wishlist'];
const REF  = sheetMap['📋 Reference Data'];

function freeze(sheetId, rows, cols) {
  return { updateSheetProperties: {
    properties: { sheetId, gridProperties: { frozenRowCount: rows, frozenColumnCount: cols } },
    fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
  }};
}

(async () => {
  const reqs = [];

  // Dashboard: 2 rows (title + subtitle), 0 cols
  reqs.push(freeze(DASH, 2, 0));

  // Guest List: 3 rows (title, summary, col headers), 0 cols
  // (seating section has a full-width A:M merge that would conflict with col-A freeze)
  reqs.push(freeze(GST, 3, 0));

  // Budget: 3 rows (title, allocation header, col headers), 0 cols
  reqs.push(freeze(BUD, 3, 0));

  // Itinerary: 2 rows (title, party info), 0 cols
  reqs.push(freeze(ITN, 2, 0));

  // Vendors: 2 rows (title, section header), 0 cols
  reqs.push(freeze(VEN, 2, 0));

  // Gifts: 2 rows (title, section header), 0 cols
  reqs.push(freeze(GFT, 2, 0));

  // Reference: 1 row, 0 cols, then hide tab
  reqs.push(freeze(REF, 1, 0));
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true },
    fields: 'hidden',
  }});

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & Reference hidden — DONE');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
