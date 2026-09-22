const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// Full-width row0/1 merges on all tabs prevent any column freeze.
// Itinerary spec wanted frozenColumnCount:3 (A-C) but row0 spans cols 0-12 → rejected.
// All tabs: rows-only freezes.
const freezes = [
  { sheetId: 732094577,  rows: 3, cols: 0 }, // Trip Overview
  { sheetId: 353766069,  rows: 2, cols: 0 }, // Day-by-Day Itinerary (title + header)
  { sheetId: 1601326951, rows: 1, cols: 0 }, // Travel Budget
  { sheetId: 1289386811, rows: 1, cols: 0 }, // Cost Per Person Calculator
  { sheetId: 1614915889, rows: 1, cols: 0 }, // Currency Converter
  { sheetId: 128079104,  rows: 1, cols: 0 }, // Research & Reservations
  { sheetId: 2144247023, rows: 1, cols: 0 }, // Sightseeing & Activities
  { sheetId: 1291844518, rows: 1, cols: 0 }, // Restaurant & Food Diary
  { sheetId: 1488564375, rows: 1, cols: 0 }, // Visa & Entry Requirements
  { sheetId: 1128916394, rows: 2, cols: 0 }, // Packing Checklist (title + progress)
  { sheetId: 2028483192, rows: 1, cols: 0 }, // Reference Data
];

(async () => {
  const reqs = freezes.map(f => ({
    updateSheetProperties: {
      properties: { sheetId: f.sheetId, gridProperties: { frozenRowCount: f.rows, frozenColumnCount: f.cols } },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
    },
  }));

  // Hide Reference Data
  reqs.push({
    updateSheetProperties: { properties: { sheetId: 2028483192, hidden: true }, fields: 'hidden' },
  });

  await batchUpdate(id, reqs, 'freezes-and-hide-reference');
  console.log('Freezes applied, Reference Data hidden');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
