'use strict';
const { google } = require('googleapis');
const { getAuth, C, hex } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: '📋 Reference Data', color: C.gray },
  { title: '📊 Dashboard',      color: C.deepSapphire },
  { title: '📝 Deduction Log',  color: C.warmGold },
  { title: '🚗 Mileage Log',    color: C.mutedSage },
  { title: '🏠 Home Office',    color: C.warmIvory },
  { title: '📅 Quarterly Tax',  color: C.amber },
  { title: '✅ Doc Checklist',  color: C.mutedSage },
  { title: '📈 Multi-Year',     color: C.deepSapphire },
];

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: '💼 Ultimate Tax Deduction Tracker' },
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
