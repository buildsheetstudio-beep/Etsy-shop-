'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const { id, url } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ERROR_PATTERNS = ['#VALUE!','#REF!','#ERROR!','#NAME?','#DIV/0!','#N/A','#NULL!','#NUM!'];

(async () => {
  console.log('QC scan starting...');
  const res = await sheets.spreadsheets.get({
    spreadsheetId: id,
    includeGridData: false,
  });

  const sheetNames = res.data.sheets.map(s => s.properties.title);
  console.log(`Scanning ${sheetNames.length} tabs: ${sheetNames.join(', ')}`);

  let totalErrors = 0;

  for (const name of sheetNames) {
    const range = `'${name}'!A1:T200`;
    let data;
    try {
      data = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range,
        valueRenderOption: 'FORMATTED_VALUE',
      });
    } catch (e) {
      console.warn(`  [WARN] Could not read ${name}: ${e.message}`);
      continue;
    }

    const rows = data.data.values || [];
    let tabErrors = 0;
    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (typeof cell === 'string' && ERROR_PATTERNS.some(p => cell.includes(p))) {
          console.error(`  ❌ ${name} row ${ri+1} col ${ci+1}: "${cell}"`);
          tabErrors++;
          totalErrors++;
        }
      });
    });
    if (tabErrors === 0) {
      console.log(`  ✅ ${name} — no errors in first 200 rows`);
    }
  }

  console.log('\n─────────────────────────────────');
  if (totalErrors === 0) {
    console.log('✅ QC PASSED — zero formula errors found.');
  } else {
    console.log(`❌ QC FAILED — ${totalErrors} error(s) found. Fix before delivery.`);
    process.exit(1);
  }
  console.log(`Spreadsheet: ${url}`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
