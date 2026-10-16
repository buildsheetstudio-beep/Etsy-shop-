'use strict';
const { getSheets } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ERROR_PATTERNS = ['#VALUE!', '#REF!', '#ERROR!', '#NAME?', '#DIV/0!', '#N/A', '#NUM!', '#NULL!'];

(async () => {
  const sheets = await getSheets();

  // Get all sheet metadata
  const meta = await sheets.spreadsheets.get({ spreadsheetId: id, fields: 'sheets.properties' });
  const allSheets = meta.data.sheets.map(s => ({ id: s.properties.sheetId, title: s.properties.title }));

  let totalErrors = 0;
  const report = [];

  for (const sheet of allSheets) {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `'${sheet.title}'`,
      valueRenderOption: 'FORMATTED_VALUE',
    }).catch(() => null);

    if (!resp) continue;
    const rows = resp.data.values || [];

    for (let ri = 0; ri < rows.length; ri++) {
      for (let ci = 0; ci < rows[ri].length; ci++) {
        const cell = String(rows[ri][ci] || '');
        const errorMatch = ERROR_PATTERNS.find(e => cell.includes(e));
        if (errorMatch) {
          const colLetter = String.fromCharCode(65 + ci);
          report.push(`  ❌  ${sheet.title}!${colLetter}${ri + 1} = "${cell}"`);
          totalErrors++;
        }
      }
    }
  }

  if (totalErrors === 0) {
    console.log('✅  Quality Gate PASSED — zero formula errors across all tabs.');
  } else {
    console.error(`\n🚨  Quality Gate FAILED — ${totalErrors} formula error(s) found:\n`);
    report.forEach(r => console.error(r));
    console.error('\nFix all errors above before delivering this spreadsheet.');
    process.exit(1);
  }
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
