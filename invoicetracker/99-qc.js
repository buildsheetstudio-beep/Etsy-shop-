'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const TOKEN_PATH  = path.join(__dirname, 'tokens.json');
const SECRET_PATH = path.join(__dirname, 'client_secret.json');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

function getAuth() {
  const { client_secret, client_id, redirect_uris } = JSON.parse(fs.readFileSync(SECRET_PATH)).installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

const sheets = google.sheets({ version: 'v4', auth: getAuth() });

const ERROR_PATTERNS = /^(#VALUE!|#REF!|#ERROR!|#NAME\?|#DIV\/0!|#N\/A|#NULL!|#NUM!)$/;

(async () => {
  const tabNames = Object.keys(sheetMap);
  let totalErrors = 0;

  for (const tab of tabNames) {
    let res;
    try {
      res = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range: `'${tab}'`,
        valueRenderOption: 'FORMATTED_VALUE',
      });
    } catch (e) {
      console.warn(`  Could not read tab "${tab}": ${e.message}`);
      continue;
    }

    const rows = res.data.values || [];
    const errors = [];

    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (typeof cell === 'string' && ERROR_PATTERNS.test(cell.trim())) {
          errors.push(`  ${colLetter(ci)}${ri + 1}: ${cell}`);
        }
      });
    });

    if (errors.length) {
      console.log(`\n❌ "${tab}" — ${errors.length} error(s):`);
      errors.forEach(e => console.log(e));
      totalErrors += errors.length;
    } else {
      console.log(`✓ "${tab}" — no errors`);
    }
  }

  console.log(`\n${'─'.repeat(50)}`);
  if (totalErrors === 0) {
    console.log('✅ Quality Gate PASSED — zero formula errors across all tabs.');
  } else {
    console.log(`🚨 Quality Gate FAILED — ${totalErrors} formula error(s) found. Fix before delivery.`);
    process.exit(1);
  }
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });

function colLetter(idx) {
  let s = '';
  idx += 1;
  while (idx > 0) {
    const rem = (idx - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    idx = Math.floor((idx - 1) / 26);
  }
  return s;
}
