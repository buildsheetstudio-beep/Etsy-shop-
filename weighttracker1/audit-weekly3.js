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

  // Check rows 30-45 to understand the legend area
  const r = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: [
      "'Weekly Schedule'!A30:P45",
      "'Weekly Schedule'!A1:P5",
    ],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  for (const vr of r.data.valueRanges) {
    console.log(`\n=== ${vr.range} ===`);
    (vr.values || []).forEach((row, i) => {
      if (row.some(v => v !== '')) console.log(`  Row ${i+1}: ${JSON.stringify(row)}`);
    });
  }

  // Get formula view of rows 30-45
  const f = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Weekly Schedule'!A30:P45",
    valueRenderOption: 'FORMULA',
  });
  console.log('\n=== Weekly Schedule A30:P45 (FORMULA) ===');
  (f.data.values || []).forEach((row, i) => {
    if (row.some(v => v !== '')) console.log(`  Row ${i+1} (sheet row ${i+30}): ${JSON.stringify(row)}`);
  });

  // Get current CF rules
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Weekly Schedule'"],
    includeGridData: false,
    fields: 'sheets(properties,conditionalFormats)',
  });
  const ws = meta.data.sheets.find(s => s.properties.title === 'Weekly Schedule');
  if (ws && ws.conditionalFormats) {
    console.log('\n=== CONDITIONAL FORMATS (Weekly Schedule) ===');
    ws.conditionalFormats.forEach((cf, i) => {
      console.log(`  CF[${i}]: ranges=${JSON.stringify(cf.ranges)}`);
      const rule = cf.booleanRule || cf.gradientRule;
      if (rule && rule.format && rule.format.backgroundColor) {
        const bg = rule.format.backgroundColor;
        console.log(`         color=rgb(${bg.red},${bg.green},${bg.blue})`);
      }
      if (rule && rule.condition) {
        const val = rule.condition.values && rule.condition.values[0];
        console.log(`         formula=${val ? val.userEnteredValue : 'N/A'}`);
      }
    });
  }

  // Get the grid data for row 38 specifically (0-based: row 37)
  const gridMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Weekly Schedule'!A35:P42"],
    includeGridData: true,
    fields: 'sheets(data(rowData(values(userEnteredValue,effectiveValue,formattedValue,userEnteredFormat,effectiveFormat))))',
  });
  console.log('\n=== Grid data rows 35-42 ===');
  const gridSheet = gridMeta.data.sheets[0];
  if (gridSheet && gridSheet.data[0] && gridSheet.data[0].rowData) {
    gridSheet.data[0].rowData.forEach((row, ri) => {
      const sheetRow = ri + 35;
      if (!row.values) return;
      const nonEmpty = row.values.filter(c => c && (c.formattedValue || (c.userEnteredFormat && c.userEnteredFormat.backgroundColor)));
      if (nonEmpty.length > 0) {
        console.log(`  Sheet row ${sheetRow}:`);
        row.values.forEach((cell, ci) => {
          if (cell && (cell.formattedValue || (cell.userEnteredFormat && cell.userEnteredFormat.backgroundColor))) {
            const bg = cell.effectiveFormat && cell.effectiveFormat.backgroundColor;
            console.log(`    col ${ci} (${String.fromCharCode(65+ci)}): val="${cell.formattedValue}" bg=${bg ? JSON.stringify(bg) : 'none'} userBg=${cell.userEnteredFormat && cell.userEnteredFormat.backgroundColor ? JSON.stringify(cell.userEnteredFormat.backgroundColor) : 'none'}`);
          }
        });
      }
    });
  }
})().catch(e => { console.error(e.message || e); process.exit(1); });
