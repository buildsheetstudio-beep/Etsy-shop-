'use strict';
const { getSheets } = require('./lib');
const fs = require('fs');
const path = require('path');

const TABS = [
  { title: 'Reference Data',      sheetId: 0, rows: 200,  cols: 16 },
  { title: 'Master Book Library', sheetId: 1, rows: 1100, cols: 24 },
  { title: 'Book Review & Notes', sheetId: 2, rows: 600,  cols: 20 },
  { title: 'Wishlist',            sheetId: 3, rows: 600,  cols: 20 },
  { title: 'Goals & Challenges',  sheetId: 4, rows: 300,  cols: 16 },
  { title: 'Reading Insights',    sheetId: 5, rows: 200,  cols: 18 },
  { title: 'Search & Filter',     sheetId: 6, rows: 150,  cols: 18 },
  { title: 'Library Gallery',     sheetId: 7, rows: 500,  cols: 14 },
  { title: 'Reading Dashboard',   sheetId: 8, rows: 150,  cols: 20 },
];

(async () => {
  const sheets = await getSheets();
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Book Tracker & Digital Reading Journal' },
      sheets: TABS.map(t => ({
        properties: {
          sheetId: t.sheetId,
          title: t.title,
          hidden: false,
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
