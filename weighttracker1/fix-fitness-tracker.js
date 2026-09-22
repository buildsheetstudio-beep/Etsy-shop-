'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1bvkKXEI2xpKpyaiOLBBr0LoSv9EwOYlE-DcLfcvwrjk';
const SID = { dashboard: 8, progressPerf: 6 };

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

// Helper: build an updateCells request for a single cell
function cellReq(sheetId, row0, col0, value) {
  const isFormula = typeof value === 'string' && value.startsWith('=');
  return {
    updateCells: {
      range: { sheetId, startRowIndex: row0, endRowIndex: row0+1, startColumnIndex: col0, endColumnIndex: col0+1 },
      rows: [{ values: [{ userEnteredValue: isFormula ? { formulaValue: value } : { stringValue: value } }] }],
      fields: 'userEnteredValue',
    },
  };
}

// Workout Setup row for exercise at position N: row = 16 + N
// Sequential positions 1-5 → rows 17-21
// Column A = Exercise ID (auto-generated), Column B = Exercise Name
function wsIdRef(pos)   { return `'Workout Setup'!$A$${16+pos}`; }
function wsNameRef(pos) { return `'Workout Setup'!B${16+pos}`; }

// ─────────────────────────────────────────────────────────────────────────────
// FIX 1 — Progress & Performance: make quarterly comparison labels + formulas
//          roll forward automatically using YEAR(TODAY())
//
// Rows 33-36 show Q4 (prev year), Q1, Q2, Q3 (current year).
// Labels in col A and SUMPRODUCT formulas in cols B-E all hardcode 2025/2026.
// Strategy: replace hardcoded year literals with YEAR(TODAY()) expressions.
// ─────────────────────────────────────────────────────────────────────────────
async function fixProgressPerf(sheets) {
  console.log('\n── Fix 1: Progress & Performance — rolling quarterly labels ───');

  // Read current formulas for rows 33-36 (0-based: rows 32-35), cols A-E (0-4)
  const data = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Progress & Performance'!A33:E36",
    valueRenderOption: 'FORMULA',
  });
  const rows = data.data.values || [];

  const requests = [];

  rows.forEach((row, r) => {
    const rowNum = 33 + r;         // 1-based sheet row
    const rowIdx = 32 + r;         // 0-based API index

    // Determine which year applies to this row
    // Row 33 = Q4 of PREVIOUS year; rows 34-36 = Q1-Q3 of CURRENT year
    const isPrevYear = rowNum === 33;
    const yearExpr   = isPrevYear ? 'YEAR(TODAY())-1' : 'YEAR(TODAY())';

    // Quarter metadata for label rebuilding
    const qMeta = [
      { q: 4, months: 'Oct–Dec',  year: 'YEAR(TODAY())-1' },
      { q: 1, months: 'Jan–Mar',  year: 'YEAR(TODAY())' },
      { q: 2, months: 'Apr–Jun',  year: 'YEAR(TODAY())' },
      { q: 3, months: 'Jul–Sep',  year: 'YEAR(TODAY())' },
    ][r];

    row.forEach((cell, c) => {
      const colIdx = c; // col A=0, B=1, ..., E=4

      if (colIdx === 0) {
        // Column A: label cell (plain text like "Q4 2025 (Oct–Dec)")
        // Replace with a formula that builds the label dynamically
        const newLabel = `="${'Q' + qMeta.q} "&${qMeta.year}&" (${qMeta.months})"`;
        requests.push(cellReq(SID.progressPerf, rowIdx, 0, newLabel));
        console.log(`  A${rowNum}: "${cell}" → ${newLabel}`);
      } else {
        // Columns B-E: SUMPRODUCT formulas with hardcoded year literal
        if (typeof cell !== 'string' || !cell.startsWith('=')) return;
        // Replace =<year> with =yearExpr in the formula
        // Patterns: )=2025) and )=2026)
        let fixed = cell.replace(/=2025/g, `=${yearExpr}`).replace(/=2026/g, `=${yearExpr}`);
        if (fixed !== cell) {
          requests.push(cellReq(SID.progressPerf, rowIdx, colIdx, fixed));
          console.log(`  ${String.fromCharCode(65+colIdx)}${rowNum}: hardcoded year → ${yearExpr}`);
        }
      }
    });
  });

  if (requests.length === 0) {
    console.log('  No changes needed');
    return;
  }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests },
  });
  console.log(`  ✓ ${requests.length} cells updated (4 labels + 16 formulas)`);
}

// ─────────────────────────────────────────────────────────────────────────────
// FIX 2 — Fitness Dashboard: replace hardcoded exercise IDs and labels with
//          dynamic references to the customer's first 5 exercises in
//          Workout Setup (rows 17-21 = positions 1-5).
//
// Top stat tiles (rows 8-9):  positions 1, 2, 3
// Top PRs table (rows 16-20): positions 1, 2, 3, 4, 5
// All MAXIFS criteria switch from "EX-000N" → Workout Setup!$A$<row>
// All label text cells switch to formulas pulling exercise name from Workout Setup!B<row>
// ─────────────────────────────────────────────────────────────────────────────
async function fixDashboard(sheets) {
  console.log('\n── Fix 2: Fitness Dashboard — dynamic exercise references ─────');

  const DWL = "'Daily Workout Log'";  // shorthand
  const Q   = '$Q$8:$Q$5000';         // Weight (lb) column
  const K   = '$K$8:$K$5000';         // Exercise ID column

  function maxifsFormula(pos) {
    return `=IFERROR(MAXIFS(${DWL}!${Q},${DWL}!${K},${wsIdRef(pos)}),"")`;
  }

  const requests = [
    // ── Top stat tile labels (row 8) — hardcoded exercise names ──────────
    cellReq(SID.dashboard, 7, 0,  `="Best "&IFERROR(${wsNameRef(1)},"—")&" (lb)"`),  // A8
    cellReq(SID.dashboard, 7, 5,  `="Best "&IFERROR(${wsNameRef(2)},"—")&" (lb)"`),  // F8
    cellReq(SID.dashboard, 7, 10, `="Best "&IFERROR(${wsNameRef(3)},"—")&" (lb)"`),  // K8

    // ── Top stat tile values (row 9) — MAXIFS with hardcoded EX-IDs ──────
    cellReq(SID.dashboard, 8, 0,  maxifsFormula(1)),  // A9
    cellReq(SID.dashboard, 8, 5,  maxifsFormula(2)),  // F9
    cellReq(SID.dashboard, 8, 10, maxifsFormula(3)),  // K9

    // ── Top PRs table: exercise name labels (col K, rows 16-20) ──────────
    cellReq(SID.dashboard, 15, 10, `=IFERROR(${wsNameRef(1)},"—")`),  // K16
    cellReq(SID.dashboard, 16, 10, `=IFERROR(${wsNameRef(2)},"—")`),  // K17
    cellReq(SID.dashboard, 17, 10, `=IFERROR(${wsNameRef(3)},"—")`),  // K18
    cellReq(SID.dashboard, 18, 10, `=IFERROR(${wsNameRef(4)},"—")`),  // K19
    cellReq(SID.dashboard, 19, 10, `=IFERROR(${wsNameRef(5)},"—")`),  // K20

    // ── Top PRs table: best value formulas (col P, rows 16-20) ───────────
    cellReq(SID.dashboard, 15, 15, maxifsFormula(1)),  // P16
    cellReq(SID.dashboard, 16, 15, maxifsFormula(2)),  // P17
    cellReq(SID.dashboard, 17, 15, maxifsFormula(3)),  // P18
    cellReq(SID.dashboard, 18, 15, maxifsFormula(4)),  // P19
    cellReq(SID.dashboard, 19, 15, maxifsFormula(5)),  // P20
  ];

  console.log('  Updating:');
  console.log('    A8, F8, K8     — label formulas pulling exercise name from Workout Setup');
  console.log('    A9, F9, K9     — MAXIFS now reference Workout Setup positions 1-3');
  console.log('    K16-K20        — exercise name labels from Workout Setup positions 1-5');
  console.log('    P16-P20        — MAXIFS now reference Workout Setup positions 1-5');

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests },
  });
  console.log(`  ✓ ${requests.length} cells updated`);
}

// ─────────────────────────────────────────────────────────────────────────────
// VERIFY
// ─────────────────────────────────────────────────────────────────────────────
async function verify(sheets) {
  console.log('\n── Verification ────────────────────────────────────────────');

  const [ppDisp, dashFm, dashDisp] = await Promise.all([
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Progress & Performance'!A33:E36",
      valueRenderOption: 'FORMATTED_VALUE',
    }),
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Fitness Dashboard'!A8:P20",
      valueRenderOption: 'FORMULA',
    }),
    sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Fitness Dashboard'!A8:P20",
      valueRenderOption: 'FORMATTED_VALUE',
    }),
  ]);

  console.log('\n  P&P quarterly section (displayed):');
  (ppDisp.data.values || []).forEach((row, r) => {
    console.log(`    Row ${33+r}: ${row.join(' | ').slice(0, 90)}`);
  });

  console.log('\n  Dashboard stat tile labels (row 8, displayed):');
  const r8 = (dashDisp.data.values || [])[0] || [];
  console.log(`    A8=${r8[0] || ''}  F8=${r8[5] || ''}  K8=${r8[10] || ''}`);

  console.log('\n  Dashboard stat tile values (row 9, displayed):');
  const r9 = (dashDisp.data.values || [])[1] || [];
  console.log(`    A9=${r9[0] || '(empty)'}  F9=${r9[5] || '(empty)'}  K9=${r9[10] || '(empty)'}`);

  console.log('\n  Dashboard Top PRs exercise labels (col K, displayed):');
  for (let i = 8; i <= 12; i++) {
    const row = (dashDisp.data.values || [])[i] || [];
    console.log(`    K${i+8}: ${row[10] || '(empty)'}`);
  }

  // Confirm no more hardcoded EX- IDs remain in Dashboard
  const dashFull = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Fitness Dashboard'!A1:T30",
    valueRenderOption: 'FORMULA',
  });
  let hardcodedRemaining = 0;
  (dashFull.data.values || []).forEach(row => {
    row.forEach(cell => {
      if (typeof cell === 'string' && cell.includes('"EX-')) hardcodedRemaining++;
    });
  });
  console.log(`\n  Hardcoded EX- IDs remaining in Dashboard: ${hardcodedRemaining} (expect 0)`);
}

(async () => {
  const sheets = await getSheets();

  await fixProgressPerf(sheets);
  await fixDashboard(sheets);
  await verify(sheets);

  console.log('\n✅  Both fixes applied:');
  console.log('    P&P quarterly labels/formulas now roll forward with YEAR(TODAY())');
  console.log('    Dashboard tiles + PRs table now pull from Workout Setup positions 1-5');
  console.log('    Customers\' first 5 exercises will appear automatically on the Dashboard');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
