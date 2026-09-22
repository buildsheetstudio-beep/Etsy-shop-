'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DB  = sheetMap['🛡️ Security Dashboard'];
const AV  = sheetMap['🔐 Account Vault'];
const SUB = sheetMap['💳 Subscriptions & Billing'];

(async () => {
  const reqs = [];

  // Chart 1: Password Strength Donut — on Security Dashboard
  // Labels: rows 26-29 col A (Strength), Values: col C count
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Password Strength Distribution',
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        domain: { sourceRange: { sources: [{ sheetId: DB, startRowIndex: 25, endRowIndex: 29, startColumnIndex: 0, endColumnIndex: 1 }] } },
        series: { sourceRange: { sources: [{ sheetId: DB, startRowIndex: 25, endRowIndex: 29, startColumnIndex: 2, endColumnIndex: 3 }] } },
        threeDimensional: false,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DB, rowIndex: 5, columnIndex: 4 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 440, heightPixels: 240,
      },
    },
  }}});

  // Chart 2: Account Status Donut — on Security Dashboard
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Account Status Overview',
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        domain: { sourceRange: { sources: [{ sheetId: DB, startRowIndex: 31, endRowIndex: 36, startColumnIndex: 0, endColumnIndex: 1 }] } },
        series: { sourceRange: { sources: [{ sheetId: DB, startRowIndex: 31, endRowIndex: 36, startColumnIndex: 2, endColumnIndex: 3 }] } },
        threeDimensional: false,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: DB, rowIndex: 5, columnIndex: 10 },
        offsetXPixels: 0, offsetYPixels: 0,
        widthPixels: 440, heightPixels: 240,
      },
    },
  }}});

  // Chart 3: Monthly Subscription Spend by Category — column on Subscriptions tab
  // Data from summary block Q10:S21 (categories + monthly cost)
  reqs.push({ addChart: { chart: {
    spec: {
      title: 'Monthly Subscription Spend by Category',
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Monthly Cost ($)' },
          { position: 'LEFT_AXIS', title: 'Category' },
        ],
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: SUB, startRowIndex: 9, endRowIndex: 21, startColumnIndex: 16, endColumnIndex: 17 }] } } }],
        series: [
          {
            series: { sourceRange: { sources: [{ sheetId: SUB, startRowIndex: 9, endRowIndex: 21, startColumnIndex: 17, endColumnIndex: 18 }] } },
            targetAxis: 'BOTTOM_AXIS',
            color: hex(C.deepIndigo),
          },
        ],
        headerCount: 0,
      },
    },
    position: {
      overlayPosition: {
        anchorCell: { sheetId: SUB, rowIndex: 25, columnIndex: 0 },
        offsetXPixels: 0, offsetYPixels: 10,
        widthPixels: 560, heightPixels: 320,
      },
    },
  }}});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
