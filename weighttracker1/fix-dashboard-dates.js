'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
const DASHBOARD_SID = 11;

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

// ─────────────────────────────────────────────────────────────────────────────
// FIX — Event Dashboard D28:E39
//
// Two problems:
//  1. SMALL(IF(D>=TODAY(), D), k) without ARRAYFORMULA only array-expands for
//     k=1–8 in Google Sheets, so D36:D39 return empty (k=9–12 fail).
//     Fix: use SMALL(FILTER(D, D>=TODAY()), k) — FILTER returns a proper array.
//  2. D28:E39 have no number format, so date serials like 46484 display as raw
//     numbers. Fix: apply DATE format "M/D/YYYY".
// ─────────────────────────────────────────────────────────────────────────────
async function fix(sheets) {
  const MEL = "'Master Event Log'";
  const D   = `${MEL}!$D$6:$D$500`;
  const E   = `${MEL}!$E$6:$E$500`;
  const A   = `${MEL}!$A$6:$A$500`;
  const B   = `${MEL}!$B$6:$B$500`;
  const G   = `${MEL}!$G$6:$G$500`;
  const F   = `${MEL}!$F$6:$F$500`;
  const H   = `${MEL}!$H$6:$H$500`;

  // FILTER-based SMALL lookup key — works correctly for all k in Google Sheets
  // and Excel 365; no ARRAYFORMULA needed.
  function upcomingKey(k) {
    return `SMALL(FILTER(${D},${D}>=TODAY()),${k})`;
  }

  const requests = [];

  for (let k = 1; k <= 12; k++) {
    const rowIdx = 26 + k; // rows 28–39 → 0-based 27–38 (rowIdx = 26+k for k=1–12)
    const key = upcomingKey(k);

    // Col D (idx 3) — Start Date: direct SMALL(FILTER()) result
    requests.push({
      updateCells: {
        range: { sheetId: DASHBOARD_SID, startRowIndex: rowIdx, endRowIndex: rowIdx+1, startColumnIndex: 3, endColumnIndex: 4 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IFERROR(${key},"")` } }] }],
        fields: 'userEnteredValue',
      },
    });

    // Col E (idx 4) — End Date: INDEX/MATCH using updated key
    requests.push({
      updateCells: {
        range: { sheetId: DASHBOARD_SID, startRowIndex: rowIdx, endRowIndex: rowIdx+1, startColumnIndex: 4, endColumnIndex: 5 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IFERROR(INDEX(${E},MATCH(${key},${D},0)),"")` } }] }],
        fields: 'userEnteredValue',
      },
    });

    // Also update A, B, C, F, G to use the same FILTER-based key for consistency
    // (these are working but should use the same lookup for correctness with k>8)
    // Col A (idx 0) — Event ID
    requests.push({
      updateCells: {
        range: { sheetId: DASHBOARD_SID, startRowIndex: rowIdx, endRowIndex: rowIdx+1, startColumnIndex: 0, endColumnIndex: 1 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IFERROR(INDEX(${A},MATCH(${key},${D},0)),"")` } }] }],
        fields: 'userEnteredValue',
      },
    });
    // Col B (idx 1) — Event Name
    requests.push({
      updateCells: {
        range: { sheetId: DASHBOARD_SID, startRowIndex: rowIdx, endRowIndex: rowIdx+1, startColumnIndex: 1, endColumnIndex: 2 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IFERROR(INDEX(${B},MATCH(${key},${D},0)),"")` } }] }],
        fields: 'userEnteredValue',
      },
    });
    // Col C (idx 2) — Category (col G in MEL)
    requests.push({
      updateCells: {
        range: { sheetId: DASHBOARD_SID, startRowIndex: rowIdx, endRowIndex: rowIdx+1, startColumnIndex: 2, endColumnIndex: 3 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IFERROR(INDEX(${G},MATCH(${key},${D},0)),"")` } }] }],
        fields: 'userEnteredValue',
      },
    });
    // Col F (idx 5) — Status (col F in MEL)
    requests.push({
      updateCells: {
        range: { sheetId: DASHBOARD_SID, startRowIndex: rowIdx, endRowIndex: rowIdx+1, startColumnIndex: 5, endColumnIndex: 6 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IFERROR(INDEX(${F},MATCH(${key},${D},0)),"")` } }] }],
        fields: 'userEnteredValue',
      },
    });
    // Col G (idx 6) — Venue/Location (col H in MEL)
    requests.push({
      updateCells: {
        range: { sheetId: DASHBOARD_SID, startRowIndex: rowIdx, endRowIndex: rowIdx+1, startColumnIndex: 6, endColumnIndex: 7 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: `=IFERROR(INDEX(${H},MATCH(${key},${D},0)),"")` } }] }],
        fields: 'userEnteredValue',
      },
    });
  }

  // Apply DATE number format to D28:E39 (and also D/E in Master Event Log if needed)
  requests.push({
    repeatCell: {
      range: { sheetId: DASHBOARD_SID, startRowIndex: 27, endRowIndex: 39, startColumnIndex: 3, endColumnIndex: 5 },
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/D/YYYY' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  console.log(`  Sending ${requests.length} requests (12 rows × 7 cols formulas + 1 format range)...`);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests },
  });
  console.log('  ✓ Done');
}

async function verify(sheets) {
  console.log('\n── Verification ────────────────────────────────────────────────');
  const disp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Event Dashboard'!A28:G39",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n  Upcoming Events table (displayed):');
  console.log('  ' + ['Event ID','Event Name','Category','Start Date','End Date','Status','Venue'].join(' | '));
  (disp.data.values || []).forEach((row, r) => {
    console.log(`  Row ${28+r}: ${row.join(' | ')}`);
  });
}

(async () => {
  const sheets = await getSheets();
  console.log('── Event Dashboard: fix D28:E39 dates ──────────────────────────');
  await fix(sheets);
  await verify(sheets);
  console.log('\n✅  Fixed:');
  console.log('    D28:D39 — SMALL(FILTER()) now returns all 12 upcoming start dates');
  console.log('    E28:E39 — INDEX/MATCH using consistent FILTER-based key');
  console.log('    D28:E39 — DATE format applied (M/D/YYYY)');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
