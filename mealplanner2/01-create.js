'use strict';
const { sheets, hex, batchUpdate, C } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: 'Reference Data',         color: C.gray },
  { title: 'Household Setup',        color: C.butter },
  { title: 'Recipe Book',            color: C.peach },
  { title: 'Recipe Ingredients',     color: C.peach },
  { title: 'Pantry Inventory',       color: C.mint },
  { title: 'Weekly Meal Planner',    color: C.powder },
  { title: 'Automated Grocery List', color: C.blush },
  { title: 'Meal Cost Tracker',      color: C.lavender },
  { title: 'Recipe Tags & Categories', color: C.peach },
  { title: 'Meal Planning Dashboard', color: C.butter },
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Weekly Meal Planner + Automated Grocery List' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          title: t.title,
          index: i,
          tabColor: hex(t.color),
        },
      })),
    },
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  for (const s of res.data.sheets) {
    sheetMap[s.properties.title] = s.properties.sheetId;
  }
  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Spreadsheet ID:', id);
  console.log('URL: https://docs.google.com/spreadsheets/d/' + id);

  // Paint bg on all tabs
  const reqs = TABS.map((_, i) => ({
    repeatCell: {
      range: { sheetId: i, startRowIndex: 0, endRowIndex: 1000, startColumnIndex: 0, endColumnIndex: 30 },
      cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  }));
  await batchUpdate(id, reqs, 'bg-paint');
  console.log('✓ Create complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
