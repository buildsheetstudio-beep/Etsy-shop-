'use strict';
const { sheets } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: 'Life Dashboard',         color: { red: 0.608, green: 0.557, blue: 0.769 } }, // #9B8EC4
  { title: 'Annual Goals',           color: { red: 0.722, green: 0.663, blue: 0.851 } }, // #B8A9D9
  { title: 'Quarterly Planner',      color: { red: 0.576, green: 0.773, blue: 0.992 } }, // #93C5FD
  { title: 'Monthly Milestones',     color: { red: 0.957, green: 0.655, blue: 0.725 } }, // #F4A7B9
  { title: 'Habit Tracker',          color: { red: 0.608, green: 0.557, blue: 0.769 } }, // #9B8EC4
  { title: 'Vision & Affirmations',  color: { red: 0.976, green: 0.843, blue: 0.494 } }, // #F9D77E
  { title: 'Reference',              color: { red: 0.847, green: 0.808, blue: 0.941 } }, // #D8CEF0
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Goal Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          title: t.title,
          index: i,
          tabColor: t.color,
          gridProperties: {
            rowCount: i === 4 ? 50 : (i === 5 ? 80 : (i === 1 ? 120 : 80)),
            columnCount: i === 4 ? 45 : 16,
          },
        },
      })),
    },
  });

  const spreadsheetId = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  console.log('Created:', spreadsheetId);
  console.log('URL:', url);
  Object.entries(sheetMap).forEach(([t, id]) => console.log(' ', id, t));

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id: spreadsheetId, url, sheetMap }, null, 2));
  console.log('spreadsheet.json written');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
