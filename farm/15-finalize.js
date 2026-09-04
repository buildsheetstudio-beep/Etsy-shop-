'use strict';
const { batchUpdate, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

(async () => {
  const reqs = [];

  // Tab colors
  const tabColors = {
    'Reference Data':          '#9E9E9E',  // hidden/grey — will also be hidden
    'Farm Setup':              '#B86F3C',  // Harvest Copper
    'Income Log':              '#388E3C',  // green
    'Expense Log':             '#C62828',  // red
    'Monthly View':            '#1565C0',  // blue
    'Annual Summary':          '#6A1B9A',  // purple
    'Enterprise Profitability':'#E65100',  // orange
    'Budget vs Actual':        '#00695C',  // teal
    'Profit Goals':            '#F9A825',  // amber
    'Tax Deduction Summary':   '#37474F',  // dark grey
    'Dashboard':               '#315A6B',  // primary Deep Field Blue
  };

  for (const [name, colorHex] of Object.entries(tabColors)) {
    const sheetId = sheetMap[name];
    if (sheetId === undefined) continue;
    reqs.push({
      updateSheetProperties: {
        properties: { sheetId, tabColorStyle: { rgbColor: hex(colorHex) } },
        fields: 'tabColorStyle',
      },
    });
  }

  // Hide Reference Data tab
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: sheetMap['Reference Data'], hidden: true },
      fields: 'hidden',
    },
  });

  // Reorder tabs: Dashboard first, then Farm Setup, then logs, then analytics
  const tabOrder = [
    'Dashboard',
    'Farm Setup',
    'Income Log',
    'Expense Log',
    'Monthly View',
    'Annual Summary',
    'Enterprise Profitability',
    'Budget vs Actual',
    'Profit Goals',
    'Tax Deduction Summary',
    'Reference Data',
  ];
  tabOrder.forEach((name, idx) => {
    const sheetId = sheetMap[name];
    if (sheetId === undefined) return;
    reqs.push({ updateSheetProperties: { properties: { sheetId, index: idx }, fields: 'index' } });
  });

  await batchUpdate(id, reqs, 'finalize');
  console.log('Finalize complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
