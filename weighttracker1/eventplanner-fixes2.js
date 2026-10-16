'use strict';
// Three fixes:
// 1. Date picker on Weekly Schedule B4 — adds DATE data validation so a calendar
//    icon appears when the cell is selected (native Google Sheets date picker).
//    Also replaces the =DATE() formula with a static date so the picker works
//    without clearing the formula bar first.
// 2. Color alignment — clears stale hardcoded backgrounds on both tabs, then
//    adds conditional format rules so colors always follow the formula text.
// 3. Event Setup B17 date — reads the current formula/value and fixes the year
//    or bad reference causing the incorrect date.
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
const CHUNK = 400;
const VCHUNK = 200;

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

// Google Sheets serial → { year, month, day } (handles the 1900 leap-year quirk)
function serialToYMD(serial) {
  const adjusted = serial > 59 ? serial - 1 : serial; // skip the fake 1900-02-29
  const ms = (adjusted - 25569) * 86400000;           // serial 25569 = 1970-01-01
  const d  = new Date(ms);
  return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
}

(async () => {
  const sheets = await getSheets();

  // ─── Sheet metadata ───────────────────────────────────────────────────────
  console.log('Reading spreadsheet metadata...');
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetMap = {};
  for (const s of meta.data.sheets) {
    sheetMap[s.properties.title] = s.properties.sheetId;
  }
  console.log('  Tabs:', Object.keys(sheetMap).join(', '));

  const wsId = sheetMap['Weekly Schedule'];
  const mcId = sheetMap['Monthly Calendar'];
  if (!wsId || !mcId) throw new Error('Could not find Weekly Schedule or Monthly Calendar tabs');

  // Find Event Setup tab (tolerant name match)
  const esName = Object.keys(sheetMap).find(k =>
    /event.?setup/i.test(k) || /setup/i.test(k)
  );
  const esId = esName ? sheetMap[esName] : null;
  console.log(`  Event Setup tab: "${esName}" (id=${esId})`);

  // ─── Read time-slot rows (col A decimal fraction) ─────────────────────────
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
  console.log(`  ${timeSlotRows.length} time-slot rows: ${timeSlotRows[0]}–${timeSlotRows[timeSlotRows.length - 1]}`);

  // ─── Diagnose Event Setup B17 ─────────────────────────────────────────────
  let b17Fix = null;
  if (esName) {
    const [fmlResp, valResp, ctxResp] = await Promise.all([
      sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${esName}'!B17`,
        valueRenderOption: 'FORMULA',
      }),
      sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${esName}'!B17`,
        valueRenderOption: 'UNFORMATTED_VALUE',
      }),
      sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `'${esName}'!A14:C20`,
        valueRenderOption: 'FORMULA',
      }),
    ]);

    const formula17  = ((fmlResp.data.values || [[]])[0] || [])[0] ?? '';
    const rawVal17   = ((valResp.data.values || [[]])[0] || [])[0];
    console.log(`\n  B17 formula:    ${formula17}`);
    console.log(`  B17 raw value:  ${rawVal17}`);
    console.log('  Context A14:C20 (formulas):');
    (ctxResp.data.values || []).forEach((row, i) =>
      console.log(`    Row ${14 + i}: ${JSON.stringify(row)}`));

    // Fix strategy:
    //   a) Formula references year 2026 → replace with 2027
    //   b) Formula references a source that drives a 2026 date → same replacement
    //   c) Static serial in 2026 range (46023–46387) → convert to 2027 equivalent
    if (typeof formula17 === 'string' && formula17.includes('2026')) {
      b17Fix = formula17.replace(/2026/g, '2027');
      console.log(`  → Fix: 2026 → 2027 in formula`);
    } else if (typeof rawVal17 === 'number' && rawVal17 >= 46023 && rawVal17 <= 46387) {
      // 2026 date serial — add 365 to land in 2027 (neither year is a leap year)
      const { month, day } = serialToYMD(rawVal17);
      b17Fix = `=DATE(2027,${month},${day})`;
      console.log(`  → Fix: serial ${rawVal17} → =DATE(2027,${month},${day})`);
    } else {
      console.log('  → B17 value not in 2026 range; logging context above for manual review.');
    }
  }

  // ─── Build requests ───────────────────────────────────────────────────────
  const PURPLE = { red: 216 / 255, green: 205 / 255, blue: 232 / 255 };
  const WHITE  = { red: 1, green: 1, blue: 1 };

  const requests    = [];
  const valueUpdates = [];

  // ── Fix 1: Date picker on Weekly Schedule B4 ─────────────────────────────
  // Replace the =DATE() formula with a static date so the calendar picker works
  // (formula bar shows the formula rather than letting the picker set a new value).
  valueUpdates.push({ range: "'Weekly Schedule'!B4", values: [['4/5/2027']] });

  // DATE_IS_VALID validation → calendar icon appears when cell is selected
  requests.push({
    setDataValidation: {
      range: {
        sheetId: wsId, startRowIndex: 3, endRowIndex: 4,
        startColumnIndex: 1, endColumnIndex: 2,
      },
      rule: {
        condition: { type: 'DATE_IS_VALID' },
        strict: false,
        showCustomUi: true,
      },
    },
  });

  // Format B4 as a readable date
  requests.push({
    repeatCell: {
      range: {
        sheetId: wsId, startRowIndex: 3, endRowIndex: 4,
        startColumnIndex: 1, endColumnIndex: 2,
      },
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'M/d/yyyy' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── Fix 2a: Weekly Schedule — clear stale backgrounds + CF rules ──────────
  for (const rowIdx of timeSlotRows) {
    // Clear hardcoded background for this row's event+location columns (B–O)
    requests.push({
      repeatCell: {
        range: {
          sheetId: wsId, startRowIndex: rowIdx, endRowIndex: rowIdx + 1,
          startColumnIndex: 1, endColumnIndex: 15,
        },
        cell: { userEnteredFormat: { backgroundColor: WHITE } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });

    // CF rule: color purple when any cell in this row is non-empty
    const row1 = rowIdx + 1;
    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{
            sheetId: wsId, startRowIndex: rowIdx, endRowIndex: rowIdx + 1,
            startColumnIndex: 1, endColumnIndex: 15,
          }],
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

  // ── Fix 2b: Monthly Calendar — clear stale backgrounds + CF rules ─────────
  // Grid rows for each month: mhRow+2 … mhRow+7 (1-indexed)
  //                         = mhRow+1 … mhRow+6 (0-indexed)
  const monthHeaderRows = [4, 13, 22, 31, 40, 49, 58, 67, 76, 85, 94, 103];
  for (const mhRow of monthHeaderRows) {
    for (let wk = 0; wk < 6; wk++) {
      const calRowIdx = mhRow + 1 + wk; // 0-indexed
      requests.push({
        repeatCell: {
          range: {
            sheetId: mcId, startRowIndex: calRowIdx, endRowIndex: calRowIdx + 1,
            startColumnIndex: 0, endColumnIndex: 7,
          },
          cell: { userEnteredFormat: { backgroundColor: WHITE } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      });
    }

    // One CF rule covering all 6 grid rows of this month
    const anchorRow1 = mhRow + 2; // 1-indexed first grid row for formula anchor
    requests.push({
      addConditionalFormatRule: {
        rule: {
          ranges: [{
            sheetId: mcId,
            startRowIndex: mhRow + 1, endRowIndex: mhRow + 7,
            startColumnIndex: 0, endColumnIndex: 7,
          }],
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

  // ── Fix 3: Event Setup B17 ────────────────────────────────────────────────
  if (b17Fix && esName) {
    valueUpdates.push({ range: `'${esName}'!B17`, values: [[b17Fix]] });
  }

  // ─── Execute ──────────────────────────────────────────────────────────────
  console.log(`\nSending ${requests.length} batchUpdate requests...`);
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: slice },
    });
    if (requests.length > CHUNK) console.log(`  Chunk ${Math.floor(i / CHUNK) + 1}`);
  }
  console.log('✅  batchUpdate done.');

  console.log(`Sending ${valueUpdates.length} value updates...`);
  for (let i = 0; i < valueUpdates.length; i += VCHUNK) {
    const slice = valueUpdates.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice },
    });
  }
  console.log('✅  valuesBatchUpdate done.');

  console.log('\n✅  All fixes applied:');
  console.log('    1. Weekly Schedule B4 — date picker added (click cell → calendar icon).');
  console.log('    2. Weekly Schedule & Monthly Calendar — stale colors cleared; CF rules');
  console.log('       now auto-color event cells purple whenever formula text is present.');
  console.log(`    3. Event Setup B17 — ${b17Fix ? `updated to: ${b17Fix}` : 'no automatic fix applied (see diagnostic log above)'}.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
