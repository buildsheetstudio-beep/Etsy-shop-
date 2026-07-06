'use strict';
const { getAuth, google } = require('./lib');
const fs = require('fs');
const path = require('path');

const TABS = [
  { title: '🔐 Account Vault',          tabColor: { red: 0.176, green: 0.169, blue: 0.561 } },
  { title: '🛡️ Security Dashboard',     tabColor: { red: 0.102, green: 0.082, blue: 0.376 } },
  { title: '💳 Subscriptions & Billing', tabColor: { red: 0.290, green: 0.565, blue: 0.851 } },
  { title: '📋 Reference Data',          tabColor: { red: 0.400, green: 0.400, blue: 0.624 } },
];

(async () => {
  const auth   = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'All-in-One Password & Account Vault' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          index: i,
          title: t.title,
          tabColorStyle: { rgbColor: t.tabColor },
          gridProperties: { rowCount: 200, columnCount: 26 },
        },
      })),
    },
  });

  const id = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;

  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  const out = { id, url, sheetMap };
  fs.writeFileSync(path.join(__dirname, 'spreadsheet.json'), JSON.stringify(out, null, 2));

  console.log('Created:', url);
  console.log('ID:', id);
  console.log('SheetMap:', JSON.stringify(sheetMap, null, 2));
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
