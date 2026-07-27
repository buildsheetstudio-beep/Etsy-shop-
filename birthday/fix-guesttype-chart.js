'use strict';
// Fix: "Confirmed Guests by Type" donut chart has endRowIndex:57 but only 3 rows
// of data exist (header at index 53, Children at 54, Adults at 55). The extra
// empty row at index 56 creates a phantom blank slice. Fix: endRowIndex:56.
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Dashboard'];
const CHART_ID = 144528056;

(async () => {
  const reqs = [{
    updateChartSpec: {
      chartId: CHART_ID,
      spec: {
        title: 'Confirmed Guests by Type',
        pieChart: {
          legendPosition: 'LABELED_LEGEND',
          domain: {
            sourceRange: {
              sources: [{ sheetId: DB, startRowIndex: 53, endRowIndex: 56, startColumnIndex: 9, endColumnIndex: 10 }]
            }
          },
          series: {
            sourceRange: {
              sources: [{ sheetId: DB, startRowIndex: 53, endRowIndex: 56, startColumnIndex: 10, endColumnIndex: 11 }]
            }
          },
          pieHole: 0.45,
        },
      },
    },
  }];
  await batchUpdate(id, reqs, 'fix-guesttype-chart');
  console.log('Fixed: Confirmed Guests by Type chart range trimmed to endRowIndex:56');
  console.log('Spreadsheet:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
