'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ERROR_PATTERNS = ['#VALUE!', '#REF!', '#NAME?', '#DIV/0!', '#N/A', '#NUM!', '#NULL!', '#ERROR!'];

(async () => {
  const tabsToCheck = Object.entries(sheetMap).filter(([name]) => name !== 'Reference Data');
  let totalErrors = 0;
  const report = [];

  for (const [tabName, sheetId] of tabsToCheck) {
    // Read entire tab as unformatted values
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `'${tabName}'`,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const rows = res.data.values || [];
    const tabErrors = [];

    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (cell === null || cell === undefined) return;
        const cellStr = String(cell);
        if (ERROR_PATTERNS.some(e => cellStr === e || cellStr.startsWith(e))) {
          tabErrors.push({ cell: `${String.fromCharCode(65+ci)}${ri+1}`, value: cellStr });
          totalErrors++;
        }
      });
    });

    if (tabErrors.length > 0) {
      report.push({ tab: tabName, errors: tabErrors });
      console.error(`  ✗ ${tabName}: ${tabErrors.length} error(s)`);
      tabErrors.slice(0, 10).forEach(e => console.error(`    ${e.cell}: ${e.value}`));
    } else {
      console.log(`  ✓ ${tabName}: clean`);
    }
  }

  console.log('');
  if (totalErrors === 0) {
    console.log('✓ QC PASSED — 0 formula errors across all tabs');
  } else {
    console.error(`✗ QC FAILED — ${totalErrors} formula error(s) found`);
    process.exit(1);
  }
})().catch(e => { console.error(e.message || e); process.exit(1); });
