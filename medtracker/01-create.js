'use strict';
const { google } = require('googleapis');
const { getAuth, C, hex } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: '📋 Reference Data',         color: C.gray },
  { title: '📊 Dashboard',              color: C.deepTeal },
  { title: '🧾 Expense Log',            color: C.warmCoral },
  { title: '💳 HSA/FSA Tracker',        color: C.mutedSage },
  { title: '📋 Insurance Claims',       color: C.amber },
  { title: '🏥 Provider Log',           color: C.deepTeal },
  { title: '💊 Prescription Tracker',  color: C.warmCoral },
  { title: '📈 Prior-Year Comparison', color: C.mutedSage },
];

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: '🏥 Ultimate Medical Expense Tracker' },
      sheets: TABS.map((t, i) => ({
        properties: {
          title: t.title,
          index: i,
          gridProperties: { rowCount: 300, columnCount: 26 },
          tabColorStyle: { rgbColor: hex(t.color) },
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
  console.log('Created:', id);
  console.log('Sheets:', Object.keys(sheetMap));
})().catch(e => { console.error(e.message); process.exit(1); });
