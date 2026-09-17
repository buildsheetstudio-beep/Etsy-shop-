'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ERROR_PATTERNS = ['#VALUE!','#REF!','#NAME?','#DIV/0!','#N/A','#NUM!','#NULL!','#ERROR!'];

(async () => {
  let totalErrors = 0;
  const tabNames = Object.keys(sheetMap).filter(n => n !== 'Reference Data');

  for (const tabName of tabNames) {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `'${tabName}'`,
      valueRenderOption: 'UNFORMATTED_VALUE',
    });
    const rows = res.data.values || [];
    const tabErrors = [];
    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        const s = String(cell);
        if (ERROR_PATTERNS.some(e => s === e)) {
          tabErrors.push({ row: ri + 1, col: ci + 1, cell: s });
        }
      });
    });
    if (tabErrors.length > 0) {
      console.error(`  ✗ ${tabName}: ${tabErrors.length} error(s)`);
      tabErrors.slice(0, 10).forEach(e => console.error(`      Row ${e.row}, Col ${e.col}: ${e.cell}`));
      totalErrors += tabErrors.length;
    } else {
      console.log(`  ✓ ${tabName}: 0 errors`);
    }
  }

  if (totalErrors > 0) {
    console.error(`\n✗ QC FAILED — ${totalErrors} formula error(s) found. Fix before delivery.`);
    process.exit(1);
  } else {
    console.log(`\n✓ QC PASSED — 0 formula errors across all tabs`);
  }
})().catch(e => { console.error(e.message || e); process.exit(1); });
