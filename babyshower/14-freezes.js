'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// MUST run last — freeze rows/cols AFTER all merges to avoid "can't merge frozen and non-frozen" errors

(async () => {
  const reqs = [];

  // Helper to freeze rows and/or columns
  function freeze(sheetId, frozenRowCount, frozenColumnCount) {
    reqs.push({ updateSheetProperties: {
      properties: {
        sheetId,
        gridProperties: { frozenRowCount, frozenColumnCount },
      },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
    }});
  }

  // Dashboard: freeze rows 1-2 (hero + party info), no col freeze
  freeze(sheetMap['🌸 Dashboard'], 2, 0);

  // Checklist: freeze rows 1-3 (title + summary + headers), no col freeze
  // (can't freeze col A — title merge A1:H1 spans it)
  freeze(sheetMap['✅ Party Checklist'], 3, 0);

  // Guest List: freeze rows 1-3, freeze col A
  // (safe: title is B1:M1 so A1 is standalone — no cross-boundary merge)
  freeze(sheetMap['👥 Guest List & RSVPs'], 3, 1);

  // Budget: freeze rows 1-2 (title + alloc header), no col freeze
  freeze(sheetMap['💰 Budget & Expenses'], 2, 0);

  // Itinerary: freeze rows 1-3 (title + summary + headers), no col freeze
  freeze(sheetMap['🗓️ Party Itinerary'], 3, 0);

  // Décor: freeze rows 1-2 (title + décor header), no col freeze
  freeze(sheetMap['🎀 Décor, Food & Drinks'], 2, 0);

  // Games: freeze rows 1-2 (title + headers), no col freeze
  freeze(sheetMap['🎮 Games & Activities'], 2, 0);

  // Gift Tracker: freeze rows 1-3 (title + summary + headers), no col freeze
  freeze(sheetMap['🎁 Gift Tracker'], 3, 0);

  // Reference Data: freeze row 1 (header), no col freeze — and hide the sheet
  freeze(sheetMap['📋 Reference Data'], 1, 0);
  reqs.push({ updateSheetProperties: {
    properties: {
      sheetId: sheetMap['📋 Reference Data'],
      hidden: true,
    },
    fields: 'hidden',
  }});

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hidden sheets complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
