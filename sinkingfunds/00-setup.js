'use strict';
const { sheets } = require('./lib');
const fs = require('fs');

const TABS = [
  'Reference Data',
  'Fund Setup & Goals',
  'Contribution Log',
  'Monthly Funding Planner',
  'Savings Forecast',
  'Goals & Milestones',
  'Fund Progress',
  'Annual Summary',
  'Sinking Funds Dashboard',
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Sinking Funds Tracker' },
      sheets: TABS.map((title, i) => ({
        properties: { title, index: i, sheetId: i },
      })),
    },
  });
  const spreadsheetId = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => {
    sheetMap[s.properties.title] = s.properties.sheetId;
  });
  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id: spreadsheetId, sheetMap }, null, 2));
  console.log('✓ Spreadsheet created:', spreadsheetId);
  console.log('  URL: https://docs.google.com/spreadsheets/d/' + spreadsheetId);
  TABS.forEach(t => console.log(`  ${t}: sheetId=${sheetMap[t]}`));
})().catch(e => { console.error(e.message || e); process.exit(1); });
