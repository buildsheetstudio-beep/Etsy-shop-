const { sheets } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

(async () => {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: id,
    fields: 'sheets(properties(sheetId,title),merges)',
  });
  meta.data.sheets.forEach(s => {
    const merges = (s.merges || []).filter(m => m.startRowIndex < 5);
    if (!merges.length) return;
    console.log('---', s.properties.title, '(sheetId', s.properties.sheetId + ')');
    merges.forEach(m => console.log('  rows', m.startRowIndex, '-', m.endRowIndex, 'cols', m.startColumnIndex, '-', m.endColumnIndex));
  });
})().catch(e => { console.error(e.errors || e); process.exit(1); });
