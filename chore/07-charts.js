'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Chore & Allowance Tracker'];

(async () => {
  const reqs = [];

  // Pie chart: Allowance Distribution by Family Member
  // Domain: A12:A21 (family member names), Series: D12:D21 (Total Earned All Time)
  // Anchored to the right of Block B, at col E row 10 (0-indexed row 9, col 4)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Allowance Distribution by Family Member',
          pieChart: {
            legendPosition: 'LABELED_LEGEND',
            pieHole: 0.4,
            domain: {
              sourceRange: {
                sources: [{ sheetId: TAB, startRowIndex: 11, endRowIndex: 21, startColumnIndex: 0, endColumnIndex: 1 }]
              }
            },
            series: {
              sourceRange: {
                sources: [{ sheetId: TAB, startRowIndex: 11, endRowIndex: 21, startColumnIndex: 3, endColumnIndex: 4 }]
              }
            }
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: TAB, rowIndex: 9, columnIndex: 5 },
            widthPixels: 400,
            heightPixels: 300
          }
        }
      }
    }
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
