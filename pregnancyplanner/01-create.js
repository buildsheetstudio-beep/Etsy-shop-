'use strict';
const { sheets } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: '🌸 Dashboard',                   color: '#6B5B8B', cols: 14, rows: 55 },
  { title: '📅 Weekly Pregnancy Journal',    color: '#8B7DA8', cols: 12, rows: 48 },
  { title: '🤒 Symptom & Wellbeing Tracker', color: '#9B7A8B', cols: 12, rows: 210 },
  { title: '🏥 Appointments Log',            color: '#6B8B7A', cols: 10, rows: 210 },
  { title: '💰 Baby Budget',                 color: '#7A8B6B', cols: 14, rows: 65 },
  { title: '🛏️ Nursery Planner',             color: '#8B9B7A', cols: 10, rows: 55 },
  { title: '🎒 Hospital Bag Checklist',      color: '#6B8B9B', cols: 6,  rows: 75 },
  { title: '👶 Baby Names',                  color: '#9B8B7A', cols: 11, rows: 55 },
  { title: '🍱 Freezer Meal Planner',        color: '#7A6B8B', cols: 11, rows: 35 },
  { title: '🌙 Postpartum Recovery Tracker', color: '#8B6B7A', cols: 11, rows: 25 },
  { title: '📝 Notes & Contacts',            color: '#6B7A8B', cols: 8,  rows: 70 },
  { title: '📋 Reference Data',              color: '#888888', cols: 22, rows: 55 },
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'All-in-One Pregnancy & Postpartum Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          index: i,
          title: t.title,
          tabColor: (() => { const v = t.color.replace('#',''); return { red: parseInt(v.slice(0,2),16)/255, green: parseInt(v.slice(2,4),16)/255, blue: parseInt(v.slice(4,6),16)/255 }; })(),
          gridProperties: { rowCount: t.rows, columnCount: t.cols },
        },
      })),
    },
  });

  const id  = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  const sheetMap = {};
  for (const s of res.data.sheets) sheetMap[s.properties.title] = s.properties.sheetId;

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, url, sheetMap }, null, 2));
  console.log('Created:', url);
  console.log('ID:', id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
