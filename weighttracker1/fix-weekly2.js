'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
const WEEKLY_SHEET_ID = 9;

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

function hexToRgb(hex) {
  return {
    red:   parseInt(hex.slice(1, 3), 16) / 255,
    green: parseInt(hex.slice(3, 5), 16) / 255,
    blue:  parseInt(hex.slice(5, 7), 16) / 255,
  };
}

// Category → hex color (from Reference Data)
// Order matches how they were inserted into CF rules (Business first → CF[15], Other last → CF[0])
const CATEGORY_COLORS = [
  ['Business',   '#C9DCEF'],
  ['Conference', '#D8CDE8'],
  ['Workshop',   '#F2D2B6'],
  ['Networking', '#C5E0DD'],
  ['Community',  '#CBDCC5'],
  ['Nonprofit',  '#E9DBA7'],
  ['School',     '#D5E4F2'],
  ['Fundraiser', '#E8C0B8'],
  ['Social',     '#E9CCD8'],
  ['Birthday',   '#F3D0BB'],
  ['Wedding',    '#DDD0DF'],
  ['Family',     '#EEE2C8'],
  ['Holiday',    '#D0E5D8'],
  ['Sports',     '#C6D9E8'],
  ['Virtual',    '#DDD5ED'],
  ['Other',      '#E3E1DE'],
];

// New CF formula: colors a cell ONLY if it is non-empty AND a matching category event
// exists on the day for that column.
// B6 (no $ anchors) = "this cell" — adjusts to each evaluated cell in the range B6:O29.
// INT((COLUMN()-2)/2) maps column index to day offset (B/C=0, D/E=1, ..., N/O=6).
function cfFormula(category) {
  return `=AND(B6<>"",COUNTIFS(INDIRECT("'Master Event Log'!D6:D500"),$B$4+INT((COLUMN()-2)/2),INDIRECT("'Master Event Log'!G6:G500"),"${category}")>0)`;
}

(async () => {
  const sheets = await getSheets();

  // ── 1. Update all 16 CF rules: add non-empty cell check ───────────────────
  // CF rules were added with index:0 repeatedly, so they sit in reverse order:
  //   CATEGORY_COLORS[0]=Business → CF[15], CATEGORY_COLORS[15]=Other → CF[0]
  console.log('Updating CF rules to only color non-empty cells...');
  const cfRange = {
    sheetId: WEEKLY_SHEET_ID,
    startRowIndex:    5,   // row 6 (0-based)
    endRowIndex:      29,  // row 29 inclusive
    startColumnIndex: 1,   // col B
    endColumnIndex:   15,  // col O inclusive
  };

  const updateCfRequests = CATEGORY_COLORS.map(([cat, hex], i) => ({
    updateConditionalFormatRule: {
      rule: {
        ranges: [cfRange],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: cfFormula(cat) }],
          },
          format: {
            backgroundColor: hexToRgb(hex),
          },
        },
      },
      index: 15 - i,  // Business=15, Conference=14, ..., Other=0
    },
  }));

  // Send in batches of 8
  for (let i = 0; i < updateCfRequests.length; i += 8) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: updateCfRequests.slice(i, i + 8) },
    });
    console.log(`  Batch ${Math.floor(i / 8) + 1} ✓`);
  }
  console.log('  ✓ All 16 CF rules updated');

  // ── 2. Write the category color legend in row 38 (cols A–P) ───────────────
  // One cell per category, 16 categories across 16 columns (A=0 through P=15).
  console.log('\nWriting category color legend (row 38, cols A:P)...');
  const legendCells = CATEGORY_COLORS.map(([cat, hex]) => ({
    userEnteredValue: { stringValue: cat },
    userEnteredFormat: {
      backgroundColor: hexToRgb(hex),
      textFormat: {
        bold: true,
        fontSize: 9,
        foregroundColor: { red: 0.13, green: 0.13, blue: 0.13 },
      },
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'MIDDLE',
    },
  }));

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{
        updateCells: {
          range: {
            sheetId:          WEEKLY_SHEET_ID,
            startRowIndex:    37,  // row 38 (0-based)
            endRowIndex:      38,
            startColumnIndex: 0,   // col A
            endColumnIndex:   16,  // col P inclusive
          },
          rows: [{ values: legendCells }],
          fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
        },
      }],
    },
  });
  console.log('  ✓ Legend written: 16 category swatches in row 38, cols A:P');

  // ── 3. Verify ──────────────────────────────────────────────────────────────
  console.log('\nVerifying legend row...');
  const check = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Weekly Schedule'!A38:P38",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  const row = (check.data.values || [[]])[0];
  console.log(`  Legend cells: ${row.join(' | ')}`);

  console.log('\n✅  Weekly Schedule v2 fixes applied:');
  console.log('    CF rules        — updated all 16 rules to only color NON-EMPTY cells');
  console.log('                      (empty time slots now stay uncolored even if day has an event)');
  console.log('    Legend row 38   — 16 category swatches written across cols A:P with correct colors');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
