'use strict';
// Fix: clear stale hardcoded backgrounds (left from original manual data) and add
// conditional formatting rules so cells turn purple automatically when formula
// text is present — keeping colors in sync with whatever week/month is displayed.
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
const CHUNK = 400;

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}

async function getSheets() {
  const auth = await getAuth();
  return google.sheets({ version: 'v4', auth });
}

(async () => {
  const sheets = await getSheets();

  // ─── Get sheet IDs ────────────────────────────────────────────────────────
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetMap = {};
  for (const s of meta.data.sheets) sheetMap[s.properties.title] = s.properties.sheetId;
  const wsId = sheetMap['Weekly Schedule'];
  const mcId = sheetMap['Monthly Calendar'];
  console.log(`sheetIds — Weekly Schedule: ${wsId}, Monthly Calendar: ${mcId}`);

  // ─── Read time-slot rows (col A = decimal fraction 0–1) ──────────────────
  const wsColAResp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Weekly Schedule'!A1:A80",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const wsColA = (wsColAResp.data.values || []).map(r => r[0]);
  const timeSlotRows = []; // 0-indexed
  for (let i = 0; i < wsColA.length; i++) {
    const v = wsColA[i];
    if (typeof v === 'number' && v > 0 && v < 1) timeSlotRows.push(i);
  }
  console.log(`${timeSlotRows.length} time-slot rows (0-indexed): ${timeSlotRows[0]}–${timeSlotRows[timeSlotRows.length - 1]}`);

  const PURPLE = { red: 216 / 255, green: 205 / 255, blue: 232 / 255 };
  const WHITE  = { red: 1, green: 1, blue: 1 };

  const requests = [];

  // ─── WEEKLY SCHEDULE ──────────────────────────────────────────────────────
  // Per time-slot row: (1) clear old hardcoded bg, (2) add CF rule so any
  // non-empty cell in cols B–O (event + location pairs) goes purple.
  for (const rowIdx of timeSlotRows) {
    requests.push({
      repeatCell: {
        range: { sheetId: wsId, startRowIndex: rowIdx, endRowIndex: rowIdx + 1,
                 startColumnIndex: 1, endColumnIndex: 15 },
        cell: { userEnteredFormat: { backgroundColor: WHITE } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });

    const row1 = rowIdx + 1; // 1-indexed, for formula anchor
    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{ sheetId: wsId, startRowIndex: rowIdx, endRowIndex: rowIdx + 1,
                     startColumnIndex: 1, endColumnIndex: 15 }],
          booleanRule: {
            condition: {
              type: 'CUSTOM_FORMULA',
              values: [{ userEnteredValue: `=LEN(B${row1})>0` }],
            },
            format: { backgroundColor: PURPLE },
          },
        },
        index: 0,
      },
    });
  }
  console.log(`  WS: ${timeSlotRows.length * 2} requests (clear + CF per row)`);

  // ─── MONTHLY CALENDAR ─────────────────────────────────────────────────────
  // monthHeaderRows are 1-indexed. Grid rows = mhRow+2 … mhRow+7 (1-indexed)
  // = mhRow+1 … mhRow+6 (0-indexed).
  // CF condition: cell contains "•" (bullet only appears when there are events).
  const monthHeaderRows = [4, 13, 22, 31, 40, 49, 58, 67, 76, 85, 94, 103];

  for (const mhRow of monthHeaderRows) {
    for (let wk = 0; wk < 6; wk++) {
      const calRowIdx = mhRow + 1 + wk; // 0-indexed grid row
      requests.push({
        repeatCell: {
          range: { sheetId: mcId, startRowIndex: calRowIdx, endRowIndex: calRowIdx + 1,
                   startColumnIndex: 0, endColumnIndex: 7 },
          cell: { userEnteredFormat: { backgroundColor: WHITE } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      });
    }

    const startRowIdx = mhRow + 1; // 0-indexed first grid row
    const endRowIdx   = mhRow + 7; // 0-indexed exclusive upper bound
    const anchorRow1  = mhRow + 2; // 1-indexed first grid row for formula

    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{ sheetId: mcId, startRowIndex: startRowIdx, endRowIndex: endRowIdx,
                     startColumnIndex: 0, endColumnIndex: 7 }],
          booleanRule: {
            condition: {
              type: 'CUSTOM_FORMULA',
              values: [{ userEnteredValue: `=ISNUMBER(SEARCH("•",A${anchorRow1}))` }],
            },
            format: { backgroundColor: PURPLE },
          },
        },
        index: 0,
      },
    });
  }
  console.log(`  MC: ${monthHeaderRows.length * 7} requests (6 clears + 1 CF per month)`);

  // ─── Execute ──────────────────────────────────────────────────────────────
  console.log(`Sending ${requests.length} total requests...`);
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: slice },
    });
    if (requests.length > CHUNK) console.log(`  Chunk ${Math.floor(i / CHUNK) + 1}`);
  }

  console.log('✅  Color alignment fix applied.');
  console.log('    Weekly Schedule  — stale backgrounds cleared; CF colors event/location cells purple when non-empty.');
  console.log('    Monthly Calendar — stale backgrounds cleared; CF colors day cells purple when events (•) are present.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
