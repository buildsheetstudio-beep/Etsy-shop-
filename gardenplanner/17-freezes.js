'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH  = sheetMap['Dashboard'];
const PLANT = sheetMap['Plant Library'];
const PS    = sheetMap['Planting Schedule'];
const LAYOUT= sheetMap['Garden Layout'];
const COMP  = sheetMap['Companion Planting Guide'];
const WAT   = sheetMap['Watering & Care Schedule'];
const HARV  = sheetMap['Harvest & Yield Tracker'];
const PEST  = sheetMap['Pest & Disease Log'];
const BUD   = sheetMap['Garden Budget'];
const SEED  = sheetMap['Seed Inventory'];
const JRN   = sheetMap['Garden Journal'];
const REF   = sheetMap['Reference Data'];

(async () => {
  const reqs = [];

  // Dashboard — 2 rows (already set but confirm)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DASH, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Plant Library — 2 rows (merged title prevents col freeze)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: PLANT, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Planting Schedule — 3 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: PS, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Garden Layout — 4 rows (title + legend rows + col numbers), 1 col (row numbers)
  // NOTE: Cannot freeze cols here due to merged title row spanning full width
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: LAYOUT, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Companion Guide — 3 rows (title, legend, header)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: COMP, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Watering — 2 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: WAT, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Harvest — 2 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: HARV, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Pest — 2 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: PEST, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Budget — 3 rows (title, input block, header)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: BUD, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Seeds — 2 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SEED, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Journal — 1 row (title)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: JRN, gridProperties: { frozenRowCount: 1, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Reference Data — 1 row + hide
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true },
    fields: 'hidden',
  }});

  await batchUpdate(id, reqs, 'freezes');

  console.log('Freezes applied + Reference Data hidden');
  console.log('\n✅ All-Season Garden Planner COMPLETE!');
  console.log('URL:', JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json')).url);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
