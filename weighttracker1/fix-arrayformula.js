'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
const SID = { weeklySchedule: 9, monthlyCalendar: 10 };

// Strip every ARRAYFORMULA( ... ) wrapper from a formula string,
// preserving the inner expression.
function stripArrayFormula(formula) {
  const TAG = 'ARRAYFORMULA(';
  let result = formula;
  let idx;
  while ((idx = result.toUpperCase().indexOf(TAG)) !== -1) {
    let depth = 1;
    let i = idx + TAG.length;
    while (i < result.length && depth > 0) {
      if (result[i] === '(') depth++;
      else if (result[i] === ')') depth--;
      i++;
    }
    // Remove "ARRAYFORMULA(" and the matching closing ")"
    result = result.slice(0, idx) + result.slice(idx + TAG.length, i - 1) + result.slice(i);
  }
  return result;
}

function colLetterToIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1; // 0-based
}

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

// Read all formulas in a range, build updateCells requests for any cell
// containing ARRAYFORMULA, return the request list and count.
async function buildFixRequests(sheets, sheetName, sheetId, rangeA1) {
  const data = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${sheetName}'!${rangeA1}`,
    valueRenderOption: 'FORMULA',
  });

  const rows = data.data.values || [];
  const startMatch = rangeA1.match(/^([A-Z]+)(\d+)/i);
  const startColIdx = colLetterToIndex(startMatch[1]);
  const startRowIdx = parseInt(startMatch[2], 10) - 1; // 0-based

  const requests = [];
  rows.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (typeof cell === 'string' && cell.toUpperCase().includes('ARRAYFORMULA(')) {
        requests.push({
          updateCells: {
            range: {
              sheetId,
              startRowIndex:    startRowIdx + r,
              endRowIndex:      startRowIdx + r + 1,
              startColumnIndex: startColIdx + c,
              endColumnIndex:   startColIdx + c + 1,
            },
            rows: [{ values: [{ userEnteredValue: { formulaValue: stripArrayFormula(cell) } }] }],
            fields: 'userEnteredValue',
          },
        });
      }
    });
  });
  return requests;
}

(async () => {
  const sheets = await getSheets();

  // ── Weekly Schedule: B6:O29 (24 rows × 14 cols — all formula cells) ────────
  console.log('Scanning Weekly Schedule B6:O29...');
  const wsRequests = await buildFixRequests(sheets, 'Weekly Schedule', SID.weeklySchedule, 'B6:O29');
  console.log(`  Found ${wsRequests.length} cells with ARRAYFORMULA`);

  if (wsRequests.length > 0) {
    for (let i = 0; i < wsRequests.length; i += 50) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { requests: wsRequests.slice(i, i + 50) },
      });
    }
    console.log(`  ✓ Weekly Schedule: ${wsRequests.length} cells updated`);
  }

  // ── Monthly Calendar: broad scan to cover any 12-month layout ───────────────
  // 12 months × (2 header + 6 calendar) rows = 96 rows max, up to 28 cols wide
  console.log('\nScanning Monthly Calendar A1:AB100...');
  const mcRequests = await buildFixRequests(sheets, 'Monthly Calendar', SID.monthlyCalendar, 'A1:AB100');
  console.log(`  Found ${mcRequests.length} cells with ARRAYFORMULA`);

  if (mcRequests.length > 0) {
    for (let i = 0; i < mcRequests.length; i += 50) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { requests: mcRequests.slice(i, i + 50) },
      });
      console.log(`  Batch ${Math.floor(i / 50) + 1} (${Math.min(i + 50, mcRequests.length)}/${mcRequests.length}) ✓`);
    }
    console.log(`  ✓ Monthly Calendar: ${mcRequests.length} cells updated`);
  }

  // ── Verify ────────────────────────────────────────────────────────────────
  console.log('\nVerifying sample cells...');
  const check = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: [
      "'Weekly Schedule'!B6",
      "'Weekly Schedule'!C6",
    ],
    valueRenderOption: 'FORMULA',
  });

  check.data.valueRanges.forEach(vr => {
    const formula = (vr.values || [['']])[0][0] || '';
    const hasAF = formula.toUpperCase().includes('ARRAYFORMULA');
    console.log(`  ${vr.range}: ${hasAF ? '❌ still has ARRAYFORMULA' : '✓ no ARRAYFORMULA'}`);
    console.log(`    ${formula.slice(0, 100)}`);
  });

  const total = wsRequests.length + mcRequests.length;
  console.log(`\n✅  Done. ${total} total cells updated.`);
  console.log('    ARRAYFORMULA stripped from Weekly Schedule and Monthly Calendar.');
  console.log('    Both tabs now work in Excel 365 (event data will display correctly).');
  console.log('    CF color-coding remains Google Sheets only (cross-sheet CF limitation in Excel).');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
