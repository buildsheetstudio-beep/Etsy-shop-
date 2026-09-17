'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap, url } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🌸 Dashboard'];
const WJ   = sheetMap['📅 Weekly Pregnancy Journal'];
const ST   = sheetMap['🤒 Symptom & Wellbeing Tracker'];
const AP   = sheetMap['🏥 Appointments Log'];
const BB   = sheetMap['💰 Baby Budget'];
const NP   = sheetMap['🛏️ Nursery Planner'];
const HB   = sheetMap['🎒 Hospital Bag Checklist'];
const BN   = sheetMap['👶 Baby Names'];
const FM   = sheetMap['🍱 Freezer Meal Planner'];
const PP   = sheetMap['🌙 Postpartum Recovery Tracker'];
const NC   = sheetMap['📝 Notes & Contacts'];
const REF  = sheetMap['📋 Reference Data'];

(async () => {
  const reqs = [];

  // Dashboard — freeze 2 rows (title + disclaimer)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DASH, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Weekly Journal — freeze 3 rows (title + disclaimer + header), no col freeze (disclaimer merged full-width)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: WJ, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Symptom Tracker — freeze 5 rows (title + disclaimer + summary labels + summary values + col headers)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: ST, gridProperties: { frozenRowCount: 5, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Appointments — freeze 5 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: AP, gridProperties: { frozenRowCount: 5, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Budget — freeze 2 rows (title + summary)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: BB, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Nursery — freeze 3 rows (title + summary labels + summary values)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: NP, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Hospital Bag — freeze 4 rows (title + summary labels + summary values + col headers)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: HB, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Baby Names — freeze 4 rows (title + summary labels + summary values + col headers), freeze col A (#)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: BN, gridProperties: { frozenRowCount: 4, frozenColumnCount: 1 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Freezer Meals — freeze 4 rows
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: FM, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Postpartum — freeze 5 rows (title + disclaimer + summary labels + summary values + col headers), no col freeze (disclaimer merged full-width)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: PP, gridProperties: { frozenRowCount: 5, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Notes & Contacts — freeze 1 row (title only)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: NC, gridProperties: { frozenRowCount: 1, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Reference Data — freeze 1 row, hide tab
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true, gridProperties: { frozenRowCount: 1 } },
    fields: 'hidden,gridProperties.frozenRowCount',
  }});

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes applied + Reference Data hidden');
  console.log('\n✅ All-in-One Pregnancy & Postpartum Planner COMPLETE!');
  console.log('URL:', url);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
