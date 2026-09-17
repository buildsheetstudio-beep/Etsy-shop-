'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const GRIDS = [
  ['Reference Data',         500,   10],
  ['Master Recipe Index',    2100,  35],
  ['Recipe Ingredients',     12100, 18],
  ['Recipe Instructions',    10100, 11],
  ['Recipe Details',         120,   20],
  ['Categories & Tags',      500,   20],
  ['Favorites & Ratings',    300,   20],
  ['Search & Filter',        300,   20],
  ['Recipe Cost Calculator', 200,   20],
  ['Recipe Dashboard',       200,   20],
];

(async () => {
  for (const [name, rows, cols] of GRIDS) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: id,
      requestBody: { requests: [{
        updateSheetProperties: {
          properties: { sheetId: sheetMap[name], gridProperties: { rowCount: rows, columnCount: cols } },
          fields: 'gridProperties(rowCount,columnCount)',
        }
      }] },
    });
    console.log(`  ✓ ${name}: ${rows} rows × ${cols} cols`);
  }
  console.log('✓ All grids expanded');
})().catch(e => { console.error(e.message || e); process.exit(1); });
