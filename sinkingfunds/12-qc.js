'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ERROR_PATTERNS = ['#VALUE!','#REF!','#NAME?','#DIV/0!','#N/A','#NUM!','#NULL!','#ERROR!'];

(async () => {
  const res = await sheets.spreadsheets.get({
    spreadsheetId: id,
    includeGridData: false,
  });

  const sheetTitles = res.data.sheets.map(s => ({
    title: s.properties.title,
    sheetId: s.properties.sheetId,
  }));

  console.log(`\nScanning ${sheetTitles.length} sheets for formula errors...`);

  let totalErrors = 0;

  for (const { title, sheetId } of sheetTitles) {
    if (title === 'Reference Data') continue; // hidden reference sheet, skip

    const data = await sheets.spreadsheets.values.get({
      spreadsheetId: id,
      range: `'${title}'`,
      valueRenderOption: 'FORMULA',
    });

    const rows = data.data.values || [];
    const sheetErrors = [];

    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        const cellStr = String(cell);
        for (const errPat of ERROR_PATTERNS) {
          if (cellStr === errPat || cellStr.includes(errPat)) {
            const col = String.fromCharCode(65 + ci);
            sheetErrors.push({ cell: `${col}${ri+1}`, value: cellStr });
          }
        }
      });
    });

    if (sheetErrors.length > 0) {
      console.error(`\n❌ ${title}: ${sheetErrors.length} error(s)`);
      sheetErrors.forEach(({ cell, value }) => {
        console.error(`   ${cell}: ${value}`);
      });
      totalErrors += sheetErrors.length;
    } else {
      console.log(`✓ ${title}: no errors`);
    }
  }

  console.log(`\n${'─'.repeat(50)}`);
  if (totalErrors === 0) {
    console.log('✅ QC PASSED — 0 formula errors across all tabs');
  } else {
    console.error(`❌ QC FAILED — ${totalErrors} formula error(s) found. Fix before delivery.`);
    process.exit(1);
  }
})().catch(e => { console.error(e.message || e); process.exit(1); });
