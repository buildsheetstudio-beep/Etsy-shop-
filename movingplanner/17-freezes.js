'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🏠 Dashboard'];
const TT   = sheetMap['✅ Task Tracker'];
const PT   = sheetMap['📦 Packing Tracker'];
const BUD  = sheetMap['💰 Moving Budget'];
const QT   = sheetMap['🚛 Moving Company Quotes'];
const COA  = sheetMap['📋 Change of Address'];
const UT   = sheetMap['⚡ Utilities & Services'];
const NB   = sheetMap['🏘️ Neighbourhood Comparison'];
const SC   = sheetMap['🏫 School & Childcare'];
const TL   = sheetMap['🗓️ Moving Timeline'];
const INV  = sheetMap['📊 Inventory List'];
const REF  = sheetMap['📋 Reference Data'];

(async () => {
  const reqs = [];

  // Dashboard — 2 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DASH, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Task Tracker — 5 rows (title + summary 2 rows + blank + header)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: TT, gridProperties: { frozenRowCount: 5, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Packing Tracker — 2 rows (title + header)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: PT, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Moving Budget — 5 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: BUD, gridProperties: { frozenRowCount: 5, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Quotes — 5 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: QT, gridProperties: { frozenRowCount: 5, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Change of Address — 5 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: COA, gridProperties: { frozenRowCount: 5, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Utilities — 5 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: UT, gridProperties: { frozenRowCount: 5, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Neighbourhood — 3 rows, NO col freeze (full-width title merge conflicts)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: NB, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // School & Childcare — 3 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SC, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Moving Timeline — 2 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: TL, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Inventory — 5 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: INV, gridProperties: { frozenRowCount: 5, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Reference Data — 1 row + hide
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true, gridProperties: { frozenRowCount: 1 } },
    fields: 'hidden,gridProperties.frozenRowCount',
  }});

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes applied + Reference Data hidden');

  const { url } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
  console.log('\n✅ All-in-One Moving Planner COMPLETE!');
  console.log('URL:', url);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
