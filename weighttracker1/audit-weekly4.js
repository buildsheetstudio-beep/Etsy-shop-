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

  const r = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: [
      "'Master Event Log'!A1:K8",    // headers + first few data rows
      "'Event Details'!A1:K8",       // headers + first few data rows
      "'Weekly Schedule'!A35:P42",   // legend area displayed values
    ],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  for (const vr of r.data.valueRanges) {
    console.log(`\n=== ${vr.range} ===`);
    (vr.values || []).forEach((row, i) => {
      if (row.some(v => v !== '')) console.log(`  Row ${i+1}: ${JSON.stringify(row)}`);
    });
  }
})().catch(e => { console.error(e.message || e); process.exit(1); });
