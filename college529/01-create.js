'use strict';
const { getSheets } = require('./lib');
const fs = require('fs');
const path = require('path');

const TABS = [
  { title: 'Reference Data',          sheetId: 0, rows: 300,  cols: 18, hidden: false },
  { title: 'Beneficiary Setup',        sheetId: 1, rows: 600,  cols: 18, hidden: false },
  { title: 'College Cost Estimator',   sheetId: 2, rows: 200,  cols: 24, hidden: false },
  { title: '529 Accounts',             sheetId: 3, rows: 1100, cols: 26, hidden: false },
  { title: 'Contribution Log',         sheetId: 4, rows: 5100, cols: 14, hidden: false },
  { title: 'Growth Tracker',           sheetId: 5, rows: 2500, cols: 14, hidden: false },
  { title: 'Contribution Planner',     sheetId: 6, rows: 200,  cols: 16, hidden: false },
  { title: 'Goals & Milestones',       sheetId: 7, rows: 1600, cols: 14, hidden: false },
  { title: 'College Savings Dashboard',sheetId: 8, rows: 150,  cols: 18, hidden: false },
];

(async () => {
  const sheets = await getSheets();
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate College Savings & 529 Plan Tracker' },
      sheets: TABS.map(t => ({
        properties: {
          sheetId: t.sheetId,
          title: t.title,
          hidden: t.hidden || false,
          gridProperties: { rowCount: t.rows, columnCount: t.cols },
        }
      })),
    }
  });

  const id  = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  const sheetMap = {};
  TABS.forEach(t => { sheetMap[t.title] = t.sheetId; });

  console.log('Created:', url);
  fs.writeFileSync(
    path.join(__dirname, 'spreadsheet.json'),
    JSON.stringify({ id, url, sheetMap }, null, 2)
  );
  console.log('spreadsheet.json saved.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
