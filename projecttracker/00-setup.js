'use strict';
const { sheets } = require('./lib');
const fs = require('fs');

const TABS = [
  'Reference Data',
  'Project Setup',
  'Master Task Log',
  'Recurring Tasks',
  'Milestones & Progress',
  'Kanban Board',
  'Gantt Chart',
  'Weekly Schedule',
  'Monthly Calendar',
  'Eisenhower Matrix',
  'Project Dashboard',
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Project Management Tracker' },
      sheets: [{ properties: { title: TABS[0] } }],
    },
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  sheetMap[TABS[0]] = res.data.sheets[0].properties.sheetId;

  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: {
      requests: TABS.slice(1).map(title => ({ addSheet: { properties: { title } } })),
    },
  });

  addRes.data.replies.forEach((reply, i) => {
    sheetMap[TABS[i + 1]] = reply.addSheet.properties.sheetId;
  });

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log(`✓ Spreadsheet created: ${id}`);
  console.log(`  URL: https://docs.google.com/spreadsheets/d/${id}`);
  TABS.forEach(t => console.log(`  ${t}: sheetId ${sheetMap[t]}`));
})().catch(e => { console.error(e.message || e); process.exit(1); });
