'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

async function main() {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: id,
    includeGridData: false
  });

  const tabNames = meta.data.sheets.map(s => s.properties.title);
  console.log(`\nQC scan across ${tabNames.length} tabs…\n`);

  const errorPatterns = ['#VALUE!', '#REF!', '#ERROR!', '#NAME?', '#DIV/0!', '#N/A', '#NULL!', '#NUM!'];
  let totalErrors = 0;
  const errorLog = [];

  for (const tab of tabNames) {
    // Fetch all cell values for this tab
    let res;
    try {
      res = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range: `'${tab}'`,
        valueRenderOption: 'FORMATTED_VALUE'
      });
    } catch (e) {
      console.warn(`  [SKIP] Could not read tab: ${tab} — ${e.message}`);
      continue;
    }

    const rows = res.data.values || [];
    const tabErrors = [];

    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        const cellStr = String(cell).trim().toUpperCase();
        for (const pat of errorPatterns) {
          if (cellStr === pat || cellStr.startsWith(pat)) {
            const cellRef = `${colLetter(ci)}${ri + 1}`;
            tabErrors.push({ cellRef, value: cell });
            break;
          }
        }
      });
    });

    if (tabErrors.length > 0) {
      totalErrors += tabErrors.length;
      errorLog.push({ tab, errors: tabErrors });
      console.log(`  ❌  ${tab}: ${tabErrors.length} error(s)`);
      tabErrors.slice(0, 20).forEach(e => {
        console.log(`       ${e.cellRef} = ${e.value}`);
      });
      if (tabErrors.length > 20) {
        console.log(`       … and ${tabErrors.length - 20} more`);
      }
    } else {
      console.log(`  ✓   ${tab}: clean`);
    }
  }

  console.log('\n────────────────────────────────────────');
  if (totalErrors === 0) {
    console.log('✅  ALL TABS CLEAN — no formula errors found.');
  } else {
    console.log(`❌  QUALITY GATE FAILED — ${totalErrors} formula error(s) across ${errorLog.length} tab(s).`);
    console.log('\nSummary of errors:');
    errorLog.forEach(({ tab, errors }) => {
      console.log(`  ${tab}: ${errors.map(e => `${e.cellRef}(${e.value})`).join(', ')}`);
    });
    process.exit(1);
  }
}

function colLetter(index) {
  let s = '';
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

main().catch(e => { console.error(e); process.exit(1); });
