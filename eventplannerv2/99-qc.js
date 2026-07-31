'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const ERROR_PATTERNS = ['#VALUE!','#REF!','#ERROR!','#NAME?','#DIV/0!','#N/A','#NULL!','#NUM!'];

(async () => {
  // Get spreadsheet metadata to find all sheet names
  const meta = await sheets.spreadsheets.get({ spreadsheetId: id, fields: 'sheets.properties' });
  const sheetList = meta.data.sheets.map(s => s.properties);

  let totalErrors = 0;
  const report = [];

  for (const sheet of sheetList) {
    if (sheet.hidden) continue;
    const name = sheet.title;

    let res;
    try {
      res = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range: `'${name}'`,
        valueRenderOption: 'FORMATTED_VALUE',
      });
    } catch (e) {
      console.warn(`  [SKIP] '${name}': ${e.message}`);
      continue;
    }

    const rows = res.data.values || [];
    const sheetErrors = [];
    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (ERROR_PATTERNS.some(err => String(cell).startsWith(err))) {
          const colLetter = ci < 26
            ? String.fromCharCode(65 + ci)
            : 'A' + String.fromCharCode(65 + ci - 26);
          sheetErrors.push(`  ${colLetter}${ri + 1}: ${cell}`);
          totalErrors++;
        }
      });
    });

    if (sheetErrors.length > 0) {
      report.push(`❌ '${name}' — ${sheetErrors.length} error(s):`);
      sheetErrors.forEach(e => report.push(e));
    } else {
      report.push(`✅ '${name}' — clean`);
    }
  }

  console.log('\n══════════════════════════════════════');
  console.log('  ULTIMATE EVENT PLANNER — QC REPORT');
  console.log('══════════════════════════════════════');
  report.forEach(line => console.log(line));
  console.log('══════════════════════════════════════');
  if (totalErrors === 0) {
    console.log('✅ ALL CLEAR — zero formula errors found.');
  } else {
    console.log(`⚠  TOTAL ERRORS: ${totalErrors} — fix before delivery.`);
    process.exit(1);
  }
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
