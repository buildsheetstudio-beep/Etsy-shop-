'use strict';
/**
 * Fix Monthly Calendar month/year sync with Payment Schedule.
 *
 * Root cause:
 *   Monthly Calendar had independent year/month inputs (B2=2026, C2=5) that
 *   are separate from Payment Schedule's inputs (C1=month, C2=year).
 *   If a customer changes the month in one tab and forgets the other, the
 *   calendar shows blank bills with no error — all formulas silently return "".
 *
 * Fix:
 *   Replace the static values in Monthly Calendar B2 and C2 with formulas
 *   that reference the Payment Schedule inputs directly:
 *     MC B2 = ='Payment Schedule'!$C$2   (year)
 *     MC C2 = ='Payment Schedule'!$C$1   (month)
 *
 *   Now there is ONE place to change the month: Payment Schedule row 1/2.
 *   The calendar title, date grid, and all bill lookups update automatically.
 */
const { google } = require('googleapis');
const fs   = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1tmklWErTpmbkoP3mVAYds6gbN1BYauwPvRTWfOmWwZI';
const MC_SID         = 1;   // Monthly Calendar sheet ID

async function getSheets() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds  = secret.installed || secret.web;
  const auth   = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return google.sheets({ version: 'v4', auth });
}

async function fixSync(sheets) {
  console.log('\n── Fix: Link Monthly Calendar month/year to Payment Schedule ───────');

  const requests = [
    // MC B2 (row 1, col 1) = year from PS C2
    {
      updateCells: {
        range: { sheetId: MC_SID, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 1, endColumnIndex: 2 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: "='Payment Schedule'!$C$2" } }] }],
        fields: 'userEnteredValue',
      },
    },
    // MC C2 (row 1, col 2) = month from PS C1
    {
      updateCells: {
        range: { sheetId: MC_SID, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 2, endColumnIndex: 3 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: "='Payment Schedule'!$C$1" } }] }],
        fields: 'userEnteredValue',
      },
    },
  ];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody:   { requests },
  });

  console.log('  ✓ MC B2 = \'Payment Schedule\'!$C$2  (year)');
  console.log('  ✓ MC C2 = \'Payment Schedule\'!$C$1  (month)');
  console.log('  Calendar now auto-follows Payment Schedule month/year');
}

async function verify(sheets) {
  console.log('\n── Verification ──────────────────────────────────────────────────────');

  const [mcForm, mcDisp] = await Promise.all([
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID, range: "'Monthly Calendar'!A1:H8", valueRenderOption: 'FORMULA',
    }),
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID, range: "'Monthly Calendar'!A1:H8", valueRenderOption: 'FORMATTED_VALUE',
    }),
  ]);

  const formRows = mcForm.data.values || [];
  const dispRows = mcDisp.data.values || [];

  console.log('  Monthly Calendar rows 1-8 (formulas → display):');
  formRows.forEach((row, r) => {
    const disp = (dispRows[r] || []).map(v => (v||'').padEnd(16)).join(' | ');
    const form = row.map(v => (v||'').padEnd(30)).join(' | ');
    console.log(`    Row ${r+1}: ${form}`);
    console.log(`         → ${disp}`);
  });

  // Check weekly balance now has real values (not flat $3,500)
  const wb = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID, range: "'Monthly Calendar'!B5:G5", valueRenderOption: 'FORMATTED_VALUE',
  });
  const wbRow = (wb.data.values || [])[0] || [];
  const flat = wbRow.every(v => v === wbRow[0]);
  console.log(`\n  Weekly balance B5:G5: ${wbRow.join(' | ')}`);
  console.log(`  ${flat ? '⚠ All values identical (may still be out of sync)' : '✓ Values vary — bills flowing through correctly'}`);
}

(async () => {
  const sheets = await getSheets();
  await fixSync(sheets);
  await verify(sheets);

  console.log('\n✅ Sync fix complete:');
  console.log('   MC B2 and C2 now reference Payment Schedule month/year');
  console.log('   Customer only needs to change month in ONE place (Payment Schedule tab)');
  console.log('   Calendar will always match the Payment Schedule automatically');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
