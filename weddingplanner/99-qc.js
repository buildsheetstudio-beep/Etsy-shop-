'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const TOKEN_PATH  = path.join(__dirname, 'tokens.json');
const SECRET_PATH = path.join(__dirname, 'client_secret.json');

function getAuth() {
  const { client_secret, client_id, redirect_uris } =
    JSON.parse(fs.readFileSync(SECRET_PATH)).installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));
const ERROR_RE = /#VALUE!|#REF!|#ERROR!|#NAME\?|#DIV\/0!|#N\/A|#NULL!|#NUM!/;

(async () => {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });

  // Get all sheet names
  const meta = await sheets.spreadsheets.get({ spreadsheetId: id, fields: 'sheets.properties' });
  const sheetNames = meta.data.sheets.map(s => s.properties.title).filter(t => t !== 'Reference Data');

  let totalErrors = 0;
  const results = [];

  for (const name of sheetNames) {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `'${name}'`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    const rows = res.data.values || [];
    const sheetErrors = [];
    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (typeof cell === 'string' && ERROR_RE.test(cell)) {
          const col = String.fromCharCode(65 + ci);
          sheetErrors.push(`  ${col}${ri + 1}: ${cell}`);
        }
      });
    });
    if (sheetErrors.length > 0) {
      totalErrors += sheetErrors.length;
      results.push(`  ✗ ${name} (${sheetErrors.length} error${sheetErrors.length > 1 ? 's' : ''}):`);
      sheetErrors.forEach(e => results.push(e));
    } else {
      results.push(`  ✓ ${name}`);
    }
  }

  console.log('\n=== QC SCAN RESULTS ===');
  results.forEach(r => console.log(r));
  console.log(`\nTotal formula errors: ${totalErrors}`);

  if (totalErrors > 0) {
    console.error('\n❌ QC FAILED — fix all formula errors before delivery.');
    process.exit(1);
  } else {
    console.log('\n✅ QC PASSED — zero formula errors across all tabs.');
  }
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
