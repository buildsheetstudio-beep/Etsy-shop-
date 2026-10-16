'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const path = require('path');

const TABS = [
  { title: 'Reference Data',          sheetId: 0,  rows: 200,  cols: 10 },
  { title: 'School Year Setup',       sheetId: 1,  rows: 200,  cols: 20 },
  { title: 'Student Roster',          sheetId: 2,  rows: 1010, cols: 22 },
  { title: 'Assignments & Assessments', sheetId: 3, rows: 1510, cols: 22 },
  { title: 'Gradebook',               sheetId: 4,  rows: 5010, cols: 20 },
  { title: 'Attendance',              sheetId: 5,  rows: 5010, cols: 16 },
  { title: 'Weekly Planner',          sheetId: 6,  rows: 500,  cols: 18 },
  { title: 'Monthly Planner',         sheetId: 7,  rows: 200,  cols: 16 },
  { title: 'Teacher To-Do',           sheetId: 8,  rows: 1510, cols: 18 },
  { title: 'Printable Weekly Overview', sheetId: 9, rows: 100,  cols: 10 },
  { title: 'Teacher Dashboard',       sheetId: 10, rows: 120,  cols: 18 },
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Teacher Planner' },
      sheets: TABS.map(t => ({
        properties: {
          sheetId: t.sheetId,
          title: t.title,
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
