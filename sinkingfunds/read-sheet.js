'use strict';
const { sheets } = require('./lib');

const SSID = process.argv[2];
const TAB  = process.argv[3] || 'Transactions';
const RANGE = process.argv[4] || `'${TAB}'!A1:Z20`;

(async () => {
  // First get sheet list
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SSID, includeGridData: false });
  console.log('Tabs:', meta.data.sheets.map(s => `${s.properties.title} (id=${s.properties.sheetId})`).join(', '));

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SSID,
    range: RANGE,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const rows = res.data.values || [];
  rows.forEach((row, i) => console.log(`Row ${i+1}: [${row.map(c => JSON.stringify(c)).join(', ')}]`));
})().catch(e => { console.error(e.message || e); process.exit(1); });
