'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const TABS = [
  'Master Recipe Index',
  'Recipe Ingredients',
  'Recipe Instructions',
  'Recipe Details',
  'Categories & Tags',
  'Favorites & Ratings',
  'Search & Filter',
  'Recipe Cost Calculator',
  'Recipe Dashboard',
];

const ERROR_PATTERNS = ['#VALUE!', '#REF!', '#ERROR!', '#NAME?', '#DIV/0!', '#N/A', '#NUM!', '#NULL!'];

(async () => {
  let totalErrors = 0;
  for (const tab of TABS) {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `'${tab}'`,
      valueRenderOption: 'FORMULA',
    });
    const rows = res.data.values || [];

    // Also get rendered values to detect actual errors
    const res2 = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `'${tab}'`,
      valueRenderOption: 'FORMATTED_VALUE',
    });
    const rendered = res2.data.values || [];

    let tabErrors = 0;
    rendered.forEach((row, r) => {
      row.forEach((cell, c) => {
        const cellStr = String(cell);
        if (ERROR_PATTERNS.some(e => cellStr.includes(e))) {
          const col = String.fromCharCode(65 + c);
          console.log(`  ❌ ${tab}!${col}${r+1}: ${cellStr}`);
          tabErrors++;
        }
      });
    });
    if (tabErrors === 0) {
      console.log(`  ✓ ${tab}: no errors`);
    } else {
      totalErrors += tabErrors;
      console.log(`  → ${tabErrors} error(s) in ${tab}`);
    }
  }
  console.log('');
  if (totalErrors === 0) {
    console.log('✅ QC PASS — zero formula errors across all tabs');
  } else {
    console.log(`⚠️  QC: ${totalErrors} total formula errors found — fix before delivery`);
    process.exit(1);
  }
})().catch(e => { console.error(e.message || e); process.exit(1); });
