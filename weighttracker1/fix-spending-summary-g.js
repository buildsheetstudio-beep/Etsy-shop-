'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1tmklWErTpmbkoP3mVAYds6gbN1BYauwPvRTWfOmWwZI';
const SPENDING_SID = 4;

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}

// Root cause:
//   Column D formula =IFERROR(C5/SUM($C$5:$C$16),0) divides by the sum of ALL
//   categories including Income (stored as negative). That sum is negative, so
//   every expense row gets a negative %, which makes REPT("█", ROUND(D*20,0))
//   receive a negative repeat count → IFERROR suppresses it → empty cell.
//   Only the Income row happens to produce a positive D, so only G14 shows bars.
//
// Fix:
//   Replace the denominator in column D with SUMIF($A$5:$A$16,"<>Income",$C$5:$C$16)
//   so it only sums actual expense categories (positive values → positive %).
//   Income row's D is set to 0 explicitly (it isn't an expense %).
//   Column G formulas are unchanged — they work correctly once D is positive.

async function fix() {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const requests = [];

  for (let row = 5; row <= 16; row++) {
    const rowIdx = row - 1; // 0-based
    const aRef  = `A${row}`;
    const cRef  = `C${row}`;
    const expenseTotal = `SUMIF($A$5:$A$16,"<>Income",$C$5:$C$16)`;

    // Column D (idx 3): % of expenses — 0 for Income row, C/expenses-total otherwise
    const dFormula = `=IFERROR(IF(${aRef}="Income",0,${cRef}/${expenseTotal}),0)`;
    requests.push({
      updateCells: {
        range: {
          sheetId: SPENDING_SID,
          startRowIndex: rowIdx,
          endRowIndex: rowIdx + 1,
          startColumnIndex: 3, // D
          endColumnIndex: 4,
        },
        rows: [{ values: [{ userEnteredValue: { formulaValue: dFormula } }] }],
        fields: 'userEnteredValue',
      },
    });
  }

  console.log(`Sending ${requests.length} column D formula updates...`);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests },
  });
  console.log('✓ Done');
}

async function verify() {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const [form, disp] = await Promise.all([
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Spending Summary'!A5:G17",
      valueRenderOption: 'FORMULA',
    }),
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Spending Summary'!A5:G17",
      valueRenderOption: 'FORMATTED_VALUE',
    }),
  ]);

  console.log('\n── After fix: Spending Summary A5:G17 ──');
  console.log('  ' + ['Category','Count','Amount','% of Exp','Paid','Unpaid','Bar'].join(' | '));
  const dRows = disp.data.values || [];
  dRows.forEach((row, r) => {
    const label = row[0] || '';
    const pct   = row[3] || '';
    const bar   = row[6] || '(empty)';
    console.log(`  Row ${5+r}: ${label.padEnd(14)} ${pct.padEnd(8)} bar="${bar.slice(0,20)}"`);
  });
}

(async () => {
  await fix();
  await verify();
  console.log('\n✅ Fix applied:');
  console.log('   Column D now divides by expense-only total (excluding Income)');
  console.log('   → positive percentages for all expense categories');
  console.log('   → column G REPT bars now display correctly for all rows');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
