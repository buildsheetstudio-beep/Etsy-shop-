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
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(title,sheetId),merges)',
  });
  ['Weekly Schedule','Monthly Calendar','Event Dashboard'].forEach(title => {
    const sh = meta.data.sheets.find(s => s.properties.title === title);
    if (!sh) return;
    const merges = (sh.merges || []).filter(m => m.startRowIndex <= 2 && m.endRowIndex >= 3);
    console.log(`\n"${title}" merges touching row 3 (0-based row 2):`);
    if (merges.length === 0) console.log('  (none)');
    merges.forEach(m => console.log(`  rows ${m.startRowIndex}-${m.endRowIndex}, cols ${m.startColumnIndex}-${m.endColumnIndex}`));
    // Also show ALL merges in rows 0-5 for context
    const topMerges = (sh.merges || []).filter(m => m.startRowIndex < 5);
    console.log('  All merges in rows 0-4:');
    topMerges.forEach(m => console.log(`    rows ${m.startRowIndex}:${m.endRowIndex}, cols ${m.startColumnIndex}:${m.endColumnIndex}`));
  });
})().catch(e => { console.error(e.message || e); process.exit(1); });
