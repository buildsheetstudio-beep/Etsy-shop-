'use strict';
// Expand all sheets to 5200 rows and 30 cols so larger data ranges don't hit grid limits
const { sheets, batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

(async () => {
  const reqs = [];
  Object.values(sheetMap).forEach(sheetId => {
    reqs.push({
      updateSheetProperties: {
        properties: {
          sheetId,
          gridProperties: { rowCount: 5200, columnCount: 30 },
        },
        fields: 'gridProperties(rowCount,columnCount)',
      },
    });
  });
  await batchUpdate(id, reqs, 'expand-grids');
  console.log('✓ All sheets expanded to 5200 rows × 30 cols');
})().catch(e => { console.error(e.message || e); process.exit(1); });
