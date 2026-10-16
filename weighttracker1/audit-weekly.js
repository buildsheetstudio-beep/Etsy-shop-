'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
async function getAuth() {
  const s = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const c = s.installed || s.web;
  const o = new google.auth.OAuth2(c.client_id, c.client_secret, c.redirect_uris[0]);
  o.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return o;
}
(async () => {
  const sheets = google.sheets({ version: 'v4', auth: await getAuth() });

  // First get spreadsheet metadata to list all sheets
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID, fields: 'sheets.properties' });
  console.log('=== SHEETS ===');
  meta.data.sheets.forEach(s => console.log(`  id=${s.properties.sheetId} title="${s.properties.title}"`));

  // Find the weekly schedule tab name from the list
  const weeklySheet = meta.data.sheets.find(s => /weekly/i.test(s.properties.title));
  const weeklyName = weeklySheet ? weeklySheet.properties.title : 'Weekly Schedule';
  console.log(`\nUsing weekly tab: "${weeklyName}"`);

  const r = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: [
      `'${weeklyName}'!A1:Z10`,    // headers + first few rows
      `'${weeklyName}'!A1:Z5`,     // top rows (legend area)
    ],
    valueRenderOption: 'FORMULA',
  });
  for (const vr of r.data.valueRanges) {
    console.log(`\n=== ${vr.range} (FORMULA) ===`);
    (vr.values || []).forEach((row, i) => {
      if (row.some(v => v !== '')) console.log(`  Row ${i+1}: ${JSON.stringify(row)}`);
    });
  }

  // Also get displayed values
  const disp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${weeklyName}'!A1:Z30`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log(`\n=== ${weeklyName} DISPLAYED ===`);
  (disp.data.values || []).forEach((row, i) => {
    if (row.some(v => v !== '')) console.log(`  Row ${i+1}: ${JSON.stringify(row)}`);
  });
})().catch(e => { console.error(e.message || e); process.exit(1); });
