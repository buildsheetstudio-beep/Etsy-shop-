'use strict';
const { getSheets } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ERROR_PATTERNS = ['#VALUE!','#REF!','#ERROR!','#NAME?','#DIV/0!','#N/A','#NUM!','#NULL!'];

(async () => {
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: id });
  const sheetList = meta.data.sheets.map(s => ({
    id: s.properties.sheetId,
    title: s.properties.title,
    hidden: s.properties.hidden,
  }));

  let totalErrors = 0;
  const errorReport = [];

  for (const sheet of sheetList) {
    const range = `'${sheet.title}'`;
    let res;
    try {
      res = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range,
        valueRenderOption: 'FORMATTED_VALUE',
      });
    } catch (e) {
      console.warn(`  Skipped '${sheet.title}': ${e.message}`);
      continue;
    }

    const rows = res.data.values || [];
    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        if (typeof cell === 'string' && ERROR_PATTERNS.some(p => cell.startsWith(p))) {
          const cellAddr = `${String.fromCharCode(65 + ci)}${ri + 1}`;
          errorReport.push({ sheet: sheet.title, cell: cellAddr, value: cell });
          totalErrors++;
        }
      });
    });
  }

  if (totalErrors === 0) {
    console.log('✅  Quality gate passed — no formula errors found.');
  } else {
    console.error(`\n❌  Quality gate FAILED — ${totalErrors} formula error(s) found:\n`);
    errorReport.forEach(e => {
      console.error(`  [${e.sheet}] ${e.cell}: ${e.value}`);
    });
    process.exit(1);
  }
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
