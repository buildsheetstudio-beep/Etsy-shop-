'use strict';
/**
 * Fix Monthly Calendar — extend all Payment Schedule range references
 * from $X$6:$X$50 → $X$6:$X$60 so bills in Bills Setup rows 47–51
 * (which land in PS rows 51–55) are visible on the calendar.
 *
 * The previous Payment Schedule fix extended PS formula rows from 22 static
 * rows to 50 formula rows (PS rows 6–55 for BS rows 2–51), but the Monthly
 * Calendar formulas still reference only $B$6:$B$50, missing the last 5 rows.
 *
 * Three reference patterns updated across all calendar formula cells:
 *   'Payment Schedule'!$B$6:$B$50  →  'Payment Schedule'!$B$6:$B$60
 *   'Payment Schedule'!$C$6:$C$50  →  'Payment Schedule'!$C$6:$C$60
 *   'Payment Schedule'!$D$6:$D$50  →  'Payment Schedule'!$D$6:$D$60
 *
 * Cells that do not contain these patterns are left untouched.
 */
const { google } = require('googleapis');
const fs   = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1tmklWErTpmbkoP3mVAYds6gbN1BYauwPvRTWfOmWwZI';
const MC_SID         = 1;  // Monthly Calendar

async function getSheets() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds  = secret.installed || secret.web;
  const auth   = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return google.sheets({ version: 'v4', auth });
}

// Replace all three column patterns in a formula string
function extendRanges(formula) {
  return formula
    .replace(/'Payment Schedule'!\$B\$6:\$B\$50/g, "'Payment Schedule'!$B$6:$B$60")
    .replace(/'Payment Schedule'!\$C\$6:\$C\$50/g, "'Payment Schedule'!$C$6:$C$60")
    .replace(/'Payment Schedule'!\$D\$6:\$D\$50/g, "'Payment Schedule'!$D$6:$D$60");
}

async function fixCalendar(sheets) {
  console.log('\n── Fix: Extend Monthly Calendar Payment Schedule range references ──────');

  // Read all formulas from the calendar area (rows 1–45, cols A–H = 0–7)
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range:         "'Monthly Calendar'!A1:H45",
    valueRenderOption: 'FORMULA',
  });

  const rows    = resp.data.values || [];
  const updates = [];
  let changed   = 0;

  rows.forEach((row, r) => {
    row.forEach((cell, c) => {
      if (typeof cell !== 'string' || !cell.startsWith('=')) return;
      const updated = extendRanges(cell);
      if (updated === cell) return;  // no change needed

      changed++;
      updates.push({
        updateCells: {
          range: {
            sheetId:          MC_SID,
            startRowIndex:    r,
            endRowIndex:      r + 1,
            startColumnIndex: c,
            endColumnIndex:   c + 1,
          },
          rows: [{ values: [{ userEnteredValue: { formulaValue: updated } }] }],
          fields: 'userEnteredValue',
        },
      });
    });
  });

  if (updates.length === 0) {
    console.log('  Nothing to update — all ranges already extended.');
    return;
  }

  console.log(`  ${changed} formula cells to update (rows 1–45, cols A–H)...`);

  // Send in chunks of 50
  const CHUNK = 50;
  for (let i = 0; i < updates.length; i += CHUNK) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody:   { requests: updates.slice(i, i + CHUNK) },
    });
    if (updates.length > CHUNK) process.stdout.write('.');
  }
  if (updates.length > CHUNK) console.log();

  console.log(`  ✓ ${changed} cells updated — PS references now $X$6:$X$60`);
}

async function verify(sheets) {
  console.log('\n── Verification ──────────────────────────────────────────────────────');

  // Check a weekly balance cell and a bill cell
  const resp = await sheets.spreadsheets.values.get({
    spreadsheetId:    SPREADSHEET_ID,
    range:            "'Monthly Calendar'!B5:G5",
    valueRenderOption: 'FORMULA',
  });
  const b5 = (resp.data.values || [])[0] || [];
  console.log('  Row 5 weekly balance formulas (should show :$B$60):');
  b5.forEach((f, i) => {
    const ok = f.includes('$B$60') ? '✓' : '⚠';
    console.log(`    ${ok} Col ${String.fromCharCode(66+i)}: ...${f.slice(Math.max(0, f.length-80))}`);
  });

  const r8 = await sheets.spreadsheets.values.get({
    spreadsheetId:    SPREADSHEET_ID,
    range:            "'Monthly Calendar'!B8",
    valueRenderOption: 'FORMULA',
  });
  const b8 = ((r8.data.values || [])[0] || [])[0] || '';
  const ok8 = (b8.includes('$B$60') && b8.includes('$C$60') && b8.includes('$D$60')) ? '✓' : '⚠';
  console.log(`\n  ${ok8} Row 8 bill cell: ...${b8.slice(Math.max(0, b8.length-100))}`);

  // Confirm no old references remain
  const allResp = await sheets.spreadsheets.values.get({
    spreadsheetId:    SPREADSHEET_ID,
    range:            "'Monthly Calendar'!A1:H45",
    valueRenderOption: 'FORMULA',
  });
  const allRows = allResp.data.values || [];
  let oldRefs = 0;
  allRows.forEach(row => row.forEach(cell => {
    if (typeof cell === 'string' && cell.includes('$B$6:$B$50')) oldRefs++;
  }));
  console.log(`\n  Old $B$6:$B$50 references remaining: ${oldRefs === 0 ? '0 ✓' : oldRefs + ' ⚠ STILL PRESENT'}`);

  // Show displayed values for the first week (should be unchanged)
  const disp = await sheets.spreadsheets.values.get({
    spreadsheetId:    SPREADSHEET_ID,
    range:            "'Monthly Calendar'!B5:H12",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n  Displayed values B5:H12 (weekly balance + week 1):');
  (disp.data.values || []).forEach((row, r) => {
    console.log(`    Row ${r+5}: ${row.map(v => (v||'').padEnd(18)).join(' | ')}`);
  });
}

(async () => {
  const sheets = await getSheets();
  await fixCalendar(sheets);
  await verify(sheets);

  console.log('\n✅ Monthly Calendar fix complete:');
  console.log('   All PS references extended from :$B$50 to :$B$60');
  console.log('   Bills in BS rows 47–51 (PS rows 51–55) now visible on calendar');
  console.log('   Future-proofed: up to 55 bills can appear (well beyond BS 49-row limit)');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
