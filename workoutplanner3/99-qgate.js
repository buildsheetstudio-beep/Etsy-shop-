'use strict';
const { getSheets } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ERROR_PATTERNS = ['#VALUE!','#REF!','#ERROR!','#DIV/0!','#NAME?','#NUM!','#N/A'];

(async () => {
  const sheets = await getSheets();
  const res = await sheets.spreadsheets.get({
    spreadsheetId: id,
    includeGridData: false,
  });

  const sheetNames = res.data.sheets.map(s => s.properties.title);
  let totalErrors = 0;

  for (const name of sheetNames) {
    // Read cell values for this sheet
    const vRes = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `'${name}'!A1:T200`,
    });

    const rows = vRes.data.values || [];
    const sheetErrors = [];

    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (typeof cell === 'string' && ERROR_PATTERNS.some(e => cell.includes(e))) {
          const col = String.fromCharCode(65 + ci);
          sheetErrors.push(`  ${col}${ri+1} = ${cell}`);
        }
      });
    });

    if (sheetErrors.length > 0) {
      console.log(`\n❌  ${name} (${sheetErrors.length} error(s)):`);
      sheetErrors.slice(0, 20).forEach(e => console.log(e));
      if (sheetErrors.length > 20) console.log(`  ... and ${sheetErrors.length - 20} more`);
      totalErrors += sheetErrors.length;
    } else {
      console.log(`✅  ${name} — no errors`);
    }
  }

  console.log(`\n${ totalErrors === 0 ? '✅  QUALITY GATE PASSED — 0 formula errors' : `❌  QUALITY GATE FAILED — ${totalErrors} total error(s)` }`);
  process.exit(totalErrors === 0 ? 0 : 1);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
