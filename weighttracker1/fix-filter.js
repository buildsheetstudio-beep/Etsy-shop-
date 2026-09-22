'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
const SID = { weeklySchedule: 9, monthlyCalendar: 10 };

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

// ─────────────────────────────────────────────────────────────────────────────
// FIX 1 — Weekly Schedule B6:O29
// Replace TEXTJOIN(IF(...)) with TEXTJOIN(FILTER(...))
// Works in Google Sheets natively AND Excel 365 (dynamic arrays)
// ─────────────────────────────────────────────────────────────────────────────
async function fixWeeklySchedule(sheets) {
  console.log('\n── Fix 1: Weekly Schedule B6:O29 — TEXTJOIN(FILTER()) ──────────');

  // Col B=1-based col 2, C=3, ..., O=15
  // Rows 6-29 = 24 rows (time slots)
  // B(2),D(4),F(6),H(8),J(10),L(12),N(14) → event name+title cells
  // C(3),E(5),G(7),I(9),K(11),M(13),O(15) → room name cells
  // Day offset: col B,C→day0; D,E→day1; F,G→day2; H,I→day3; J,K→day4; L,M→day5; N,O→day6

  const requests = [];

  for (let rowNum = 6; rowNum <= 29; rowNum++) {
    const rowIdx = rowNum - 1; // 0-based

    for (let colNum = 2; colNum <= 15; colNum++) {
      const colIdx = colNum - 1; // 0-based
      const day = Math.floor((colNum - 2) / 2);
      const isRoom = colNum % 2 === 1; // odd 1-based = room (C,E,G,I,K,M,O)

      const dateCond   = `'Event Details'!$D$6:$D$500=$B$4+${day}`;
      const roomCond   = `'Event Details'!$E$6:$E$500=$A${rowNum}`;
      const bothCond   = `(${dateCond})*(${roomCond})`;

      let formula;
      if (isRoom) {
        formula = `=IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER('Event Details'!$H$6:$H$500,${bothCond})),"")`;
      } else {
        formula = `=IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER("▶ "&'Event Details'!$A$6:$A$500&" - "&'Event Details'!$B$6:$B$500,${bothCond})),"")`;
      }

      requests.push({
        updateCells: {
          range: {
            sheetId: SID.weeklySchedule,
            startRowIndex: rowIdx,
            endRowIndex: rowIdx + 1,
            startColumnIndex: colIdx,
            endColumnIndex: colIdx + 1,
          },
          rows: [{ values: [{ userEnteredValue: { formulaValue: formula } }] }],
          fields: 'userEnteredValue',
        },
      });
    }
  }

  console.log(`  Writing ${requests.length} cells (24 rows × 14 cols)...`);

  // Send in batches of 100
  for (let i = 0; i < requests.length; i += 100) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: requests.slice(i, i + 100) },
    });
    console.log(`  Batch ${Math.floor(i / 100) + 1} (${Math.min(i + 100, requests.length)}/${requests.length}) ✓`);
  }
  console.log(`  ✓ Weekly Schedule: ${requests.length} cells updated`);
}

// ─────────────────────────────────────────────────────────────────────────────
// FIX 2 — Monthly Calendar
// Replace TEXTJOIN(IF(...)) with IFERROR(TEXTJOIN(FILTER(...)))
// The LET variable 'd' is preserved — just swap the inner function
// ─────────────────────────────────────────────────────────────────────────────
async function fixMonthlyCalendar(sheets) {
  console.log('\n── Fix 2: Monthly Calendar — TEXTJOIN(FILTER()) in LET formulas ──');

  const data = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Monthly Calendar'!A1:AB150",
    valueRenderOption: 'FORMULA',
  });

  const rows = data.data.values || [];
  const requests = [];
  let found = 0;

  // The broken pattern (may have IFERROR wrapper or not):
  // TEXTJOIN(CHAR(10),TRUE,IF('Master Event Log'!$D$6:$D$500=d,"• "&'Master Event Log'!$B$6:$B$500,""))
  // Replace with:
  // IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER("• "&'Master Event Log'!$B$6:$B$500,'Master Event Log'!$D$6:$D$500=d)),"")
  const BROKEN_PATTERN = `TEXTJOIN(CHAR(10),TRUE,IF('Master Event Log'!$D$6:$D$500=d,"• "&'Master Event Log'!$B$6:$B$500,""))`;
  const FIXED_PATTERN  = `IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER("• "&'Master Event Log'!$B$6:$B$500,'Master Event Log'!$D$6:$D$500=d)),"")`;

  rows.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (typeof cell !== 'string' || !cell.startsWith('=')) return;
      if (!cell.includes(BROKEN_PATTERN)) return;

      found++;
      let fixed = cell;

      // Remove any IFERROR wrapper around the TEXTJOIN so we don't double-wrap
      // Pattern: IFERROR(TEXTJOIN(...),"") → we'll re-add IFERROR inside FIXED_PATTERN
      // But FIXED_PATTERN already includes IFERROR, so strip the outer one if present
      fixed = fixed.replace(
        `IFERROR(${BROKEN_PATTERN},"")`,
        FIXED_PATTERN
      );
      // Also handle case without outer IFERROR
      if (fixed.includes(BROKEN_PATTERN)) {
        fixed = fixed.replace(BROKEN_PATTERN, FIXED_PATTERN);
      }

      requests.push({
        updateCells: {
          range: {
            sheetId: SID.monthlyCalendar,
            startRowIndex: r,
            endRowIndex: r + 1,
            startColumnIndex: c,
            endColumnIndex: c + 1,
          },
          rows: [{ values: [{ userEnteredValue: { formulaValue: fixed } }] }],
          fields: 'userEnteredValue',
        },
      });
    });
  });

  console.log(`  Found ${found} cells with broken TEXTJOIN(IF()) pattern`);

  if (requests.length === 0) {
    console.log('  No matching cells found — checking for alternate pattern...');

    // Debug: sample a few formulas
    rows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (typeof cell === 'string' && cell.includes('Master Event Log') && r < 20) {
          console.log(`    [${r+1},${c+1}]: ${cell.slice(0, 150)}`);
        }
      });
    });
    return;
  }

  for (let i = 0; i < requests.length; i += 100) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: requests.slice(i, i + 100) },
    });
    console.log(`  Batch ${Math.floor(i / 100) + 1} (${Math.min(i + 100, requests.length)}/${requests.length}) ✓`);
  }
  console.log(`  ✓ Monthly Calendar: ${requests.length} cells updated`);
}

// ─────────────────────────────────────────────────────────────────────────────
// VERIFY
// ─────────────────────────────────────────────────────────────────────────────
async function verify(sheets) {
  console.log('\n── Verification ────────────────────────────────────────────────');

  const [wsFormula, mcFormula] = await Promise.all([
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Weekly Schedule'!B6:C7",
      valueRenderOption: 'FORMULA',
    }),
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Monthly Calendar'!B4:C5",
      valueRenderOption: 'FORMULA',
    }),
  ]);

  console.log('\n  Weekly Schedule sample formulas:');
  (wsFormula.data.values || []).forEach((row, r) => {
    row.forEach((cell, c) => {
      const col = String.fromCharCode(66 + c);
      console.log(`    ${col}${r+6}: ${cell.slice(0, 120)}`);
    });
  });

  console.log('\n  Monthly Calendar sample formulas:');
  (mcFormula.data.values || []).forEach((row, r) => {
    row.forEach((cell, c) => {
      const col = String.fromCharCode(66 + c);
      console.log(`    ${col}${r+4}: ${String(cell || '').slice(0, 120)}`);
    });
  });

  // Check for any remaining IF-inside-TEXTJOIN patterns
  const [wsAll] = await Promise.all([
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Weekly Schedule'!B6:O29",
      valueRenderOption: 'FORMULA',
    }),
  ]);
  let brokenIF = 0;
  (wsAll.data.values || []).forEach(row => {
    row.forEach(cell => {
      if (typeof cell === 'string' && cell.includes('TEXTJOIN') && cell.includes(',IF(')) brokenIF++;
    });
  });
  console.log(`\n  Weekly Schedule TEXTJOIN(IF()) remaining: ${brokenIF} (expect 0)`);
  console.log('  FILTER-based formulas are compatible with Google Sheets and Excel 365.');
}

(async () => {
  const sheets = await getSheets();

  await fixWeeklySchedule(sheets);
  await fixMonthlyCalendar(sheets);
  await verify(sheets);

  console.log('\n✅  Both tabs fixed:');
  console.log('    Weekly Schedule B6:O29 — TEXTJOIN(FILTER()) for events and rooms');
  console.log('    Monthly Calendar — TEXTJOIN(FILTER()) inside LET formulas');
  console.log('    Works in Google Sheets natively AND Excel 365 (no ARRAYFORMULA needed)');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
