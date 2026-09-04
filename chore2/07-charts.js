'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Chore & Allowance Tracker'];

(async () => {
  const reqs = [];

  // Pie/Donut chart: Weekly Allowance by family member
  // Domain: Chore & Allowance Tracker!I3:I12 (Family Member names)
  // Series: Chore & Allowance Tracker!K3:K12 (Weekly Allowance)
  // Anchored at L1 on the Tracker tab, 400×300px
  const spec = {
    title: 'Weekly Allowance by Member',
    titleTextFormat: { foregroundColor: hex(C.charcoal), bold: true, fontSize: 11 },
    backgroundColor: hex(C.white),
    pieChart: {
      legendPosition: 'RIGHT_LEGEND',
      pieHole: 0.4,
      series: {
        sourceRange: {
          sources: [{ sheetId: TAB, startRowIndex: 2, endRowIndex: 12,
            startColumnIndex: 10, endColumnIndex: 11 }]  // K col (Weekly Allowance)
        }
      },
      domain: {
        sourceRange: {
          sources: [{ sheetId: TAB, startRowIndex: 2, endRowIndex: 12,
            startColumnIndex: 8, endColumnIndex: 9 }]   // I col (Family Member names)
        }
      }
    }
  };

  reqs.push({ addChart: { chart: { spec, position: { overlayPosition: {
    anchorCell: { sheetId: TAB, rowIndex: 0, columnIndex: 11 },  // L1
    widthPixels: 400,
    heightPixels: 300,
    offsetXPixels: 0,
    offsetYPixels: 0,
  } } } } });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
