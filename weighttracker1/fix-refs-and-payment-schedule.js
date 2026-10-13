'use strict';
/**
 * Fix 1 — Extend Refs column D validation range in Bills Setup column K
 *          from Refs!$D$2:$D$6 → Refs!$D$2:$D$20 so customers can add
 *          more Account Types in the Refs tab (D7:D20) and have them
 *          appear in the dropdown. Also extend Categories and Frequency
 *          ranges and apply validation to all Bills Setup rows 2–50.
 *
 * Fix 2 — Replace the 22 hardcoded static rows in Payment Schedule with
 *          formula-driven rows that pull from Bills Setup.
 *          Each row is visible only when:
 *            - Bills Setup N (Active) = TRUE, AND
 *            - The bill's frequency lands it in the selected month/year
 *          Status column (F) is left blank for manual input each month.
 */
const { google } = require('googleapis');
const fs  = require('fs');
const path = require('path');

const SPREADSHEET_ID  = '1tmklWErTpmbkoP3mVAYds6gbN1BYauwPvRTWfOmWwZI';
const BS_SID          = 0;   // Bills Setup
const PS_SID          = 2;   // Payment Schedule

async function getSheets() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds  = secret.installed || secret.web;
  const auth   = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return google.sheets({ version: 'v4', auth });
}

async function send(sheets, requests) {
  const CHUNK = 50;
  for (let i = 0; i < requests.length; i += CHUNK) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: requests.slice(i, i + CHUNK) },
    });
    if (requests.length > CHUNK) process.stdout.write('.');
  }
  if (requests.length > CHUNK) console.log();
}

// ─── Fix 1: Extend validation ranges ─────────────────────────────────────────
async function fixValidation(sheets) {
  console.log('\n── Fix 1: Extend Bills Setup validation ranges ──────────────────────');

  const range = (col) => ({
    sheetId: BS_SID, startRowIndex: 1, endRowIndex: 50,
    startColumnIndex: col, endColumnIndex: col + 1,
  });

  const listRule = (formula) => ({
    condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: formula }] },
    showCustomUi: true,
    strict: false,
  });

  const boolRule = { condition: { type: 'BOOLEAN' }, showCustomUi: true };

  const requests = [
    // C (idx 2)  — Categories: was A2:A13, extended to A2:A20
    { setDataValidation: { range: range(2),  rule: listRule('=Refs!$A$2:$A$20') } },
    // E (idx 4)  — Type: unchanged (only 2 options)
    { setDataValidation: { range: range(4),  rule: listRule('=Refs!$E$2:$E$3')  } },
    // F (idx 5)  — Frequency: was B2:B8, extended to B2:B15
    { setDataValidation: { range: range(5),  rule: listRule('=Refs!$B$2:$B$15') } },
    // J (idx 9)  — Auto-Pay checkbox
    { setDataValidation: { range: range(9),  rule: boolRule                     } },
    // K (idx 10) — Account Type: was D2:D6 → now D2:D20 (main fix)
    { setDataValidation: { range: range(10), rule: listRule('=Refs!$D$2:$D$20') } },
    // M (idx 12) — Color Tag: was F2:F9, extended to F2:F20
    { setDataValidation: { range: range(12), rule: listRule('=Refs!$F$2:$F$20') } },
    // N (idx 13) — Active checkbox
    { setDataValidation: { range: range(13), rule: boolRule                     } },
  ];

  await send(sheets, requests);
  console.log('  ✓ Column K (Account): Refs!$D$2:$D$6 → Refs!$D$2:$D$20');
  console.log('    Customer adds new account types in Refs tab D7–D20');
  console.log('  ✓ All validations now cover Bills Setup rows 2–50');
}

// ─── Fix 2: Formula-driven Payment Schedule ───────────────────────────────────
async function fixPaymentSchedule(sheets) {
  console.log('\n── Fix 2: Replace static Payment Schedule with Bills Setup formulas ──');

  const requests = [];

  // Clear old static rows 6–40
  requests.push({
    repeatCell: {
      range: { sheetId: PS_SID, startRowIndex: 5, endRowIndex: 40, startColumnIndex: 0, endColumnIndex: 10 },
      cell: {},
      fields: 'userEnteredValue',
    },
  });

  // Write formula rows: PS row k maps to Bills Setup row (k – 4)
  // PS rows 6–55  →  BS rows 2–51
  for (let psRow = 6; psRow <= 55; psRow++) {
    const r   = psRow - 4;  // Bills Setup row number

    const N   = `'Bills Setup'!$N$${r}`;   // Active checkbox
    const Ff  = `'Bills Setup'!$F$${r}`;   // Frequency
    const H   = `'Bills Setup'!$H$${r}`;   // Start Date
    const Gd  = `'Bills Setup'!$G$${r}`;   // Due Day of month
    const mo  = `$C$1`;                     // PS selected month
    const yr  = `$C$2`;                     // PS selected year

    // "Due this month" OR clause — handles all 7 frequency types
    const freqOr = [
      `${Ff}="Monthly"`,
      `${Ff}="Weekly"`,
      `${Ff}="Bi-Weekly"`,
      `AND(${Ff}="Quarterly",IFERROR(MOD(DATEDIF(${H},DATE(${yr},${mo},1),"M"),3)=0,FALSE))`,
      `AND(${Ff}="Semi-Annual",IFERROR(MOD(DATEDIF(${H},DATE(${yr},${mo},1),"M"),6)=0,FALSE))`,
      `AND(${Ff}="Annual",${H}<>"",MONTH(${H})=${mo},YEAR(DATE(${yr},${mo},1))>=YEAR(${H}))`,
      `AND(${Ff}="One-Time",${H}<>"",MONTH(${H})=${mo},YEAR(${H})=${yr})`,
    ].join(',');

    const show     = `IFERROR(AND(${N}=TRUE,OR(${freqOr})),FALSE)`;
    const dueDate  = `IF(${Ff}="One-Time",${H},DATE(${yr},${mo},${Gd}))`;

    // 10 columns: A–J  (F = Status, left blank for manual input)
    const cols = [
      `=IF(${show},ROW()-5,"")`,                                        // A: row #
      `=IF(${show},${dueDate},"")`,                                     // B: due date
      `=IF(${show},'Bills Setup'!$B$${r},"")`,                          // C: bill name
      `=IF(${show},'Bills Setup'!$D$${r},"")`,                          // D: amount
      `=IF(${show},'Bills Setup'!$C$${r},"")`,                          // E: category
      ``,                                                                // F: Status (manual)
      `=IF(${show},'Bills Setup'!$K$${r},"")`,                          // G: account
      `=IF(${show},IF('Bills Setup'!$J$${r}=TRUE,"✅",""),"")`,         // H: auto-pay
      `=IF(${show},'Bills Setup'!$F$${r},"")`,                          // I: frequency
      `=IF(${show},'Bills Setup'!$P$${r},"")`,                          // J: notes
    ];

    requests.push({
      updateCells: {
        range: {
          sheetId: PS_SID,
          startRowIndex: psRow - 1, endRowIndex: psRow,
          startColumnIndex: 0,      endColumnIndex: 10,
        },
        rows: [{
          values: cols.map(f =>
            f ? { userEnteredValue: { formulaValue: f } }
              : { userEnteredValue: { stringValue: '' } }
          ),
        }],
        fields: 'userEnteredValue',
      },
    });
  }

  // Date format on column B (due dates)
  requests.push({
    repeatCell: {
      range: { sheetId: PS_SID, startRowIndex: 5, endRowIndex: 56, startColumnIndex: 1, endColumnIndex: 2 },
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/D/YYYY' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Status dropdown on column F (Paid / Unpaid / Pending / Skipped from Refs C)
  requests.push({
    setDataValidation: {
      range: { sheetId: PS_SID, startRowIndex: 5, endRowIndex: 56, startColumnIndex: 5, endColumnIndex: 6 },
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: '=Refs!$C$2:$C$5' }] },
        showCustomUi: true,
        strict: false,
      },
    },
  });

  console.log(`  Sending ${requests.length} requests (50 formula rows + format + validation)...`);
  await send(sheets, requests);
  console.log('  ✓ Payment Schedule rows 6–55 pull from Bills Setup rows 2–51');
  console.log('  ✓ Each row shows only when Active=TRUE AND due in selected month/year');
  console.log('  ✓ Status column (F) left blank — customer selects Paid/Unpaid/etc. each month');
  console.log('  ✓ Due date formatted M/D/YYYY');
  console.log('  ✓ Status dropdown: Paid / Unpaid / Pending / Skipped');
}

// ─── Verify ───────────────────────────────────────────────────────────────────
async function verify(sheets) {
  console.log('\n── Verification ─────────────────────────────────────────────────────');

  const [dvSnap, psSnap] = await Promise.all([
    sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      ranges: ["'Bills Setup'!K2:K3"],
      includeGridData: true,
      fields: 'sheets(data(rowData(values(dataValidation))))',
    }),
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Payment Schedule'!A6:J10",
      valueRenderOption: 'FORMATTED_VALUE',
    }),
  ]);

  const dv = dvSnap.data.sheets[0]?.data[0]?.rowData?.[0]?.values?.[0]?.dataValidation;
  const dvRange = dv?.condition?.values?.[0]?.userEnteredValue || '(none)';
  console.log('  Bills Setup K2 validation range: ' + dvRange);

  console.log('  Payment Schedule rows 6–10 (displayed):');
  console.log('  ' + ['#','Due Date','Bill Name','Amount','Category','Status','Account','Auto-Pay','Freq','Notes'].join(' | '));
  (psSnap.data.values || []).forEach((row, r) => {
    if (row.some(c => c)) console.log('  Row ' + (r + 6) + ': ' + row.join(' | '));
    else console.log('  Row ' + (r + 6) + ': (hidden — inactive or not due this month)');
  });
}

(async () => {
  const sheets = await getSheets();
  await fixValidation(sheets);
  await fixPaymentSchedule(sheets);
  await verify(sheets);

  console.log('\n✅ Both issues fixed:');
  console.log('   1. Account Types: Refs tab D7–D20 now feeds into the Bills Setup dropdown');
  console.log('   2. Payment Schedule: auto-populates from Bills Setup for selected month/year');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
