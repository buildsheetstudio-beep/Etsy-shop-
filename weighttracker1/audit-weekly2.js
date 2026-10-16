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

  // Full schedule rows + reference categories
  const r = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: [
      "'Weekly Schedule'!A6:O50",      // all schedule rows - formulas
      "'Reference Data'!A1:D50",        // categories + colors
      "'Event Details'!A1:J10",         // column layout of source
    ],
    valueRenderOption: 'FORMULA',
  });
  for (const vr of r.data.valueRanges) {
    console.log(`\n=== ${vr.range} ===`);
    (vr.values || []).forEach((row, i) => {
      if (row.some(v => v !== '')) console.log(`  Row ${i+1}: ${JSON.stringify(row)}`);
    });
  }

  // Get CF rules on Weekly Schedule tab
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Weekly Schedule'"],
    includeGridData: false,
    fields: 'sheets(properties,conditionalFormats)',
  });
  const ws = meta.data.sheets.find(s => s.properties.title === 'Weekly Schedule');
  if (ws && ws.conditionalFormats) {
    console.log('\n=== CONDITIONAL FORMATS (Weekly Schedule) ===');
    ws.conditionalFormats.forEach((cf, i) => {
      console.log(`  CF[${i}]: ranges=${JSON.stringify(cf.ranges)}`);
      console.log(`         rule=${JSON.stringify(cf.booleanRule || cf.gradientRule)}`);
    });
  } else {
    console.log('\n(no conditional formats found on Weekly Schedule)');
  }

  // Also check Event Details column layout (displayed values of header row)
  const edHeaders = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Event Details'!A5:J5",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n=== Event Details headers ===');
  console.log(JSON.stringify(edHeaders.data.values));
})().catch(e => { console.error(e.message || e); process.exit(1); });
