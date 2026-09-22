'use strict';
const { getSheets, batchUpdate } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: 'Reference Data',       sheetId: 0, rows: 200,   cols: 20 },
  { title: 'Workout Setup',        sheetId: 1, rows: 1100,  cols: 16 },
  { title: 'Daily Workout Log',    sheetId: 2, rows: 10100, cols: 30 },
  { title: 'Weekly Planner',       sheetId: 3, rows: 200,   cols: 20 },
  { title: 'Monthly Calendar',     sheetId: 4, rows: 200,   cols: 16 },
  { title: 'Yearly Fitness Planner', sheetId: 5, rows: 200, cols: 20 },
  { title: 'Progress & Performance', sheetId: 6, rows: 200, cols: 20 },
  { title: 'Personal Records',     sheetId: 7, rows: 1100,  cols: 16 },
  { title: 'Fitness Dashboard',    sheetId: 8, rows: 200,   cols: 20 },
];

(async () => {
  const sheets = await getSheets();

  // Create spreadsheet with first tab
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Workout Planner & Fitness Tracker' },
      sheets: [{ properties: { sheetId: TABS[0].sheetId, title: TABS[0].title, gridProperties: { rowCount: TABS[0].rows, columnCount: TABS[0].cols } } }],
    },
  });

  const id = res.data.spreadsheetId;
  const url = `https://docs.google.com/spreadsheets/d/${id}/edit`;
  console.log(`Created: ${url}`);

  // Add remaining tabs
  const addReqs = TABS.slice(1).map(t => ({
    addSheet: { properties: { sheetId: t.sheetId, title: t.title, gridProperties: { rowCount: t.rows, columnCount: t.cols } } },
  }));
  await batchUpdate(id, addReqs, 'add-tabs');

  // Build sheetMap
  const sheetMap = {};
  TABS.forEach(t => { sheetMap[t.title] = t.sheetId; });

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, url, sheetMap }, null, 2));
  console.log('spreadsheet.json saved.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
