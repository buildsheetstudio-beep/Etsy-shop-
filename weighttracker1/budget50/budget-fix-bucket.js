'use strict';
/**
 * Fix: Bucket formula (col F) was overwritten when writing transaction data.
 * The transactions write to A:J included column F (index 5), blanking all 68 rows.
 * Solution: re-write F10, copyPaste to F10:F1009.
 */
const { google } = require('googleapis');
const fs = require('fs'), path = require('path');
const AUTH_DIR = path.join(__dirname, '..');
const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'budget-config.json')));
const SID = CFG.spreadsheetId;
const SHEET = CFG.sheets['Transactions'];

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const auth = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'tokens.json'))));
  return auth;
}

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // Step 1: Re-write F10 with the Bucket VLOOKUP formula
  const F10 = `=IFERROR(VLOOKUP(E10,'Start Here'!$B$32:$C$54,2,0),"Uncategorized")`;
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [{
        updateCells: {
          range: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 10, startColumnIndex: 5, endColumnIndex: 6 },
          rows: [{ values: [{ userEnteredValue: { formulaValue: F10 } }] }],
          fields: 'userEnteredValue',
        },
      }],
    },
  });
  console.log('   Wrote F10 formula');

  // Step 2: copyPaste F10 → F10:F1009
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [{
        copyPaste: {
          source:      { sheetId: SHEET, startRowIndex: 9, endRowIndex: 10, startColumnIndex: 5, endColumnIndex: 6 },
          destination: { sheetId: SHEET, startRowIndex: 9, endRowIndex: 1009, startColumnIndex: 5, endColumnIndex: 6 },
          pasteType: 'PASTE_FORMULA', pasteOrientation: 'NORMAL',
        },
      }],
    },
  });
  console.log('   CopyPasted F10 → F10:F1009');

  // Step 3: Spot-check a few cells to verify
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SID,
    range: `'Transactions'!E10:S13`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n   Spot-check rows 10-13 (E=Category, F=Bucket, G=Amount, S=Income):');
  (resp.data.values || []).forEach((row, i) => {
    const [cat, bucket, amt, , , , , , , , , , , , inc] = row;
    console.log(`   Row ${10+i}: Category="${cat}" Bucket="${bucket}" Amount="${amt}" Income="${inc||0}"`);
  });

  // Step 4: Check January totals on Dashboard
  const dash = await sheets.spreadsheets.values.get({
    spreadsheetId: SID,
    range: `'Budget Dashboard'!A9:H9`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const dashRow = (dash.data.values || [])[0] || [];
  console.log(`\n   Dashboard stat cards: ${dashRow.filter(v=>v).join(' | ')}`);

  // Step 5: Check Yearly Jan row
  const yr = await sheets.spreadsheets.values.get({
    spreadsheetId: SID,
    range: `'Yearly Insights'!A7:K7`,
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const yrRow = (yr.data.values || [])[0] || [];
  console.log(`   Yearly January: ${yrRow.join(' | ')}`);

  console.log('\n✅ Bucket formula fix complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
