'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const SPREADSHEET_ID = '1YSrKj6U3Zb_al3g7j9e-80qFRGpQVwD4JLe4gh1Ah2k';
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
      "'👥 Guest List & RSVP'!R1:Z15",  // RSVP summary full width
      "'📊 Event Analytics'!L3:M15",     // Rating + Satisfaction cols
      "'📋 Events Log'!A9:A16",          // orphan row numbers
    ],
    valueRenderOption: 'FORMULA',
  });
  for (const vr of r.data.valueRanges) {
    console.log(`\n=== ${vr.range} ===`);
    (vr.values || []).forEach((row, i) => console.log(`  Row ${i+1}: ${JSON.stringify(row)}`));
  }
})().catch(e => { console.error(e.message || e); process.exit(1); });
