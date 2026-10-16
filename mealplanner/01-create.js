'use strict';
const { getAuth, batchUpdate, C, hex } = require('./lib');
const { google } = require('googleapis');
const fs = require('fs');

const TABS = [
  { title: '📋 Reference Data',        color: C.gray },
  { title: '📊 Dashboard',             color: C.goldenMustard },
  { title: '📅 Weekly Meal Plan',      color: C.deepBasil },
  { title: '📖 Recipe Book',           color: C.goldenMustard },
  { title: '🛒 Grocery List',          color: C.deepBasil },
  { title: '🥗 Food Inventory',        color: C.mutedSage },
  { title: '🍳 Meal Prep Schedule',    color: C.goldenMustard },
  { title: '💰 Grocery Budget',        color: C.warmOrange },
  { title: '🚫 Dietary Restrictions',  color: C.deepBerry },
  { title: '🎯 Nutrition Goals',       color: C.deepBasil },
];

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Meal Planner+' },
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

  const reqs = TABS.map((_, i) => ({
    repeatCell: {
      range: { sheetId: i, startRowIndex: 0, endRowIndex: 1000, startColumnIndex: 0, endColumnIndex: 26 },
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  }));
  await batchUpdate(id, reqs, 'bg');
  console.log('Created:', id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
