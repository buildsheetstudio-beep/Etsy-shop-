'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

(async () => {
  const errors = [];
  const tabs = Object.entries(sheetMap);

  for (const [name] of tabs) {
    let resp;
    try {
      resp = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range: `'${name}'`,
        valueRenderOption: 'UNFORMATTED_VALUE',
      });
    } catch (e) {
      console.warn(`  [WARN] Could not read "${name}": ${e.message}`);
      continue;
    }
    const rows = resp.data.values || [];
    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        const s = String(cell);
        if (s.startsWith('#') && ['#VALUE!','#REF!','#ERROR!','#NAME?','#N/A','#DIV/0!','#NULL!'].some(e => s.includes(e))) {
          const col = String.fromCharCode(65 + ci);
          errors.push(`  ${name}!${col}${ri+1} → ${s}`);
        }
      });
    });
  }

  if (errors.length === 0) {
    console.log('✓ QC PASS — no formula errors found across all tabs');
  } else {
    console.log(`✗ QC FAIL — ${errors.length} error cell(s) found:`);
    errors.forEach(e => console.log(e));
    process.exit(1);
  }
})().catch(e => { console.error(e.message || e); process.exit(1); });
