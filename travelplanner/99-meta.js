const { sheets } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
(async () => {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: id,
    fields: 'sheets(properties(sheetId,title,hidden,gridProperties,tabColor),conditionalFormats,charts)',
  });
  meta.data.sheets.forEach(s => {
    const p = s.properties;
    console.log(p.sheetId, p.title, 'hidden=' + !!p.hidden, 'frozenRows=' + (p.gridProperties.frozenRowCount || 0), 'frozenCols=' + (p.gridProperties.frozenColumnCount || 0), 'cfRules=' + (s.conditionalFormats || []).length, 'charts=' + (s.charts || []).length);
  });
})().catch(e => { console.error(e.errors || e); process.exit(1); });
