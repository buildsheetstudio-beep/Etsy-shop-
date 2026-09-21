'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
async function getAuth() {
  const s = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const c = s.installed || s.web;
  const o = new google.auth.OAuth2(c.client_id, c.client_secret, c.redirect_uris[0]);
  o.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return o;
}
(async () => {
  const sheets = google.sheets({ version: 'v4', auth: await getAuth() });

  // 1. Get sheet IDs
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets.properties,namedRanges',
  });
  const sheetMap = {};
  meta.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });
  console.log('=== SHEET IDs ===');
  Object.entries(sheetMap).forEach(([t, id]) => console.log(`  ${id}: "${t}"`));

  // 2. Named ranges detail  
  console.log('\n=== NAMED RANGES (full detail) ===');
  (meta.data.namedRanges || []).forEach(nr => {
    const r = nr.range;
    console.log(`  id=${nr.namedRangeId} name="${nr.name}" sheetId=${r.sheetId} rows=${r.startRowIndex+1}:${r.endRowIndex} cols=${r.startColumnIndex+1}:${r.endColumnIndex}`);
  });

  // 3. Find all DV cells with EVT- in their values (event ID dropdowns)
  const dvMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(title,sheetId),data(rowData(values(dataValidation))))',
  });
  console.log('\n=== EVENT-ID DROPDOWN LOCATIONS ===');
  dvMeta.data.sheets.forEach(sheet => {
    const evtCols = new Set();
    (sheet.data || []).forEach(grid => {
      (grid.rowData || []).forEach((row, ri) => {
        (row.values || []).forEach((cell, ci) => {
          const dv = cell.dataValidation;
          if (!dv || dv.condition.type !== 'ONE_OF_LIST') return;
          const vals = (dv.condition.values || []).map(v => v.userEnteredValue || '');
          if (vals.some(v => /^EVT-\d/.test(v))) {
            evtCols.add(ci);
          }
        });
      });
    });
    if (evtCols.size > 0) {
      const cols = [...evtCols].map(ci => String.fromCharCode(65 + ci)).join(', ');
      console.log(`  "${sheet.properties.title}" (sheetId=${sheet.properties.sheetId}): cols ${cols} (indices ${[...evtCols].join(', ')})`);
    }
  });

  // 4. Headers for Guests & Attendees and Event Tasks
  const hdrs = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Guests & Attendees'!A5:Z5", "'Event Tasks'!A5:Z5"],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  hdrs.data.valueRanges.forEach(vr => {
    console.log(`\n=== ${vr.range} headers ===`);
    (vr.values || [[]])[0].forEach((h, i) => {
      if (h) console.log(`  col ${i} (${String.fromCharCode(65+i)}): "${h}"`);
    });
  });

  // 5. Reference Data layout — what's in each column
  const ref = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Reference Data'!A1:Z20",
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n=== REFERENCE DATA — first 3 rows ===');
  (ref.data.values || []).slice(0, 3).forEach((row, i) => {
    if (row.some(v => v)) console.log(`  Row ${i+1}: ${JSON.stringify(row)}`);
  });

  // 6. Weekly Schedule and Monthly Calendar — find header/info rows to add notes
  const notes = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Weekly Schedule'!A1:B5", "'Monthly Calendar'!A1:B5"],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  console.log('\n=== TOP ROWS (for note placement) ===');
  notes.data.valueRanges.forEach(vr => {
    console.log(`\n  ${vr.range}:`);
    (vr.values || []).forEach((row, i) => {
      if (row.some(v => v)) console.log(`    Row ${i+1}: ${JSON.stringify(row)}`);
    });
  });
})().catch(e => { console.error(e.message || e); process.exit(1); });
