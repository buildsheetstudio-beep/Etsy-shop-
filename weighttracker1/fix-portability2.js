'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';

const SID = {
  eventDashboard:   11,
  weeklySchedule:    9,
  monthlyCalendar:  10,
};

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

const AMBER_BG   = { red: 1.0, green: 0.973, blue: 0.831 };
const AMBER_TEXT = { red: 0.35, green: 0.28, blue: 0.0 };

function bannerCell(text) {
  return {
    userEnteredValue: { stringValue: text },
    userEnteredFormat: {
      backgroundColor: AMBER_BG,
      textFormat: { fontSize: 8, italic: true, foregroundColor: AMBER_TEXT },
      verticalAlignment: 'MIDDLE',
    },
  };
}

function amberFill() {
  return { userEnteredFormat: { backgroundColor: AMBER_BG } };
}

(async () => {
  const sheets = await getSheets();

  // ── Fix 3a: Weekly Schedule — row 3 is empty; add visible banner ──────────
  // Row 3 (0-based row 2) has no existing merges.
  // Match merge width of existing header rows: cols 0:14 (A:N).
  console.log('Fix 3a: Weekly Schedule — adding tab-rename banner (row 3)...');
  const WS_NOTE = '⚠️  Do not rename any sheet tabs — the color-coded event highlighting depends on tab names staying exactly as they are.';
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          updateCells: {
            range: { sheetId: SID.weeklySchedule, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 14 },
            rows: [{ values: [bannerCell(WS_NOTE), ...Array(13).fill(amberFill())] }],
            fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
          },
        },
        {
          mergeCells: {
            range: { sheetId: SID.weeklySchedule, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 14 },
            mergeType: 'MERGE_ALL',
          },
        },
      ],
    },
  });
  console.log('  ✓ Weekly Schedule banner written (A3:N3)');

  // ── Fix 3b: Monthly Calendar row 3 — already merged A3:G3 with COUNTA ─────
  // Strategy: unmerge the existing A3:G3, rewrite COUNTA to A3 (merged A3:D3),
  // and put the tab note in E3 (merged E3:G3).
  console.log('\nFix 3b: Monthly Calendar — splitting row 3 for COUNTA + note...');
  const MC_COUNTA = `=IFERROR("Total Events: "&COUNTA('Master Event Log'!B6:B500)&" | Confirmed: "&COUNTIF('Master Event Log'!F6:F500,"Confirmed")&" | Planning: "&COUNTIF('Master Event Log'!F6:F500,"Planning")&" | Completed: "&COUNTIF('Master Event Log'!F6:F500,"Completed"),"")`;
  const MC_NOTE = '⚠️  Do not rename sheet tabs — the calendar color coding depends on tab names staying unchanged.';

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        // Step 1: unmerge the existing A3:G3
        {
          unmergeCells: {
            range: { sheetId: SID.monthlyCalendar, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 7 },
          },
        },
        // Step 2: write COUNTA formula to A3 and note to E3
        {
          updateCells: {
            range: { sheetId: SID.monthlyCalendar, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 7 },
            rows: [{
              values: [
                // A3: COUNTA formula
                {
                  userEnteredValue: { formulaValue: MC_COUNTA },
                  userEnteredFormat: { textFormat: { fontSize: 9, italic: true }, verticalAlignment: 'MIDDLE' },
                },
                // B3, C3, D3: fill for merge
                amberFill(), amberFill(), amberFill(),
                // E3: note
                bannerCell(MC_NOTE),
                // F3, G3: fill for merge
                amberFill(), amberFill(),
              ],
            }],
            fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
          },
        },
        // Step 3: re-merge COUNTA area A3:D3
        {
          mergeCells: {
            range: { sheetId: SID.monthlyCalendar, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 4 },
            mergeType: 'MERGE_ALL',
          },
        },
        // Step 4: merge note area E3:G3
        {
          mergeCells: {
            range: { sheetId: SID.monthlyCalendar, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 4, endColumnIndex: 7 },
            mergeType: 'MERGE_ALL',
          },
        },
      ],
    },
  });
  console.log('  ✓ Monthly Calendar row 3 split: COUNTA in A3:D3, note in E3:G3');

  // ── Fix 4: Event Dashboard — row 3 is the stat-tiles row (has content) ────
  // Strategy: add the Google Sheets Only notice as a cell note (popup annotation)
  // on the A2 subtitle cell so it doesn't disturb any existing content.
  // Cell notes appear as a small triangle indicator and are preserved in copies.
  console.log('\nFix 4: Event Dashboard — adding Google Sheets Only cell note to A2...');
  const GS_NOTE = 'Google Sheets Only ─ This spreadsheet is optimized for Google Sheets.\n\nThe Weekly Schedule and Monthly Calendar tabs use Google-specific functions (ARRAYFORMULA, LET) that are not supported in Microsoft Excel.\n\nAll other tabs (Event Dashboard, Master Event Log, Budget & Expenses, etc.) are fully compatible with Excel 2013+.';
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{
        updateCells: {
          range: { sheetId: SID.eventDashboard, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 1 },
          rows: [{ values: [{ note: GS_NOTE }] }],
          fields: 'note',
        },
      }],
    },
  });
  console.log('  ✓ Cell note added to Event Dashboard A2 (hover to read)');

  // ── Verify ────────────────────────────────────────────────────────────────
  console.log('\nVerifying...');
  const check = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Weekly Schedule'!A3", "'Monthly Calendar'!A3", "'Monthly Calendar'!E3"],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  check.data.valueRanges.forEach(vr => {
    const val = (vr.values || [['']])[0][0] || '';
    console.log(`  ${vr.range}: "${val.slice(0, 70)}"`);
  });

  console.log('\n✅  Fixes 3 & 4 applied:');
  console.log('    Weekly Schedule A3:N3  — visible amber banner warning about tab renaming');
  console.log('    Monthly Calendar A3:D3 — COUNTA stats preserved');
  console.log('    Monthly Calendar E3:G3 — visible amber banner warning about tab renaming');
  console.log('    Event Dashboard A2     — cell note (hover indicator) with GS-only notice');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
