'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const { sheets } = require('./lib');

const TABS = [
  { title: '📋 Dashboard',                          color: '#2E3A35', cols: 10, rows: 55 },
  { title: '✅ Executor Checklist',                 color: '#3D5A4A', cols: 8,  rows: 60 },
  { title: '💼 Assets & Debts',                     color: '#4A5A3A', cols: 11, rows: 60 },
  { title: '👨‍👩‍👧 Beneficiaries',                      color: '#5A4A3A', cols: 14, rows: 30 },
  { title: '🏦 Estate Distribution',                color: '#3A4A5A', cols: 14, rows: 30 },
  { title: '💻 Digital Assets & Online Accounts',   color: '#4A3A5A', cols: 9,  rows: 30 },
  { title: '📞 Beneficiary Communication Log',      color: '#5A3A4A', cols: 11, rows: 30 },
  { title: '🏛️ Professional Services & Fees',        color: '#3A5A5A', cols: 15, rows: 25 },
  { title: '📅 Estate Timeline & Probate Progress', color: '#2E3A35', cols: 9,  rows: 25 },
  { title: '📁 Documents Register',                 color: '#4A5A4A', cols: 8,  rows: 25 },
  { title: '📝 Notes & Contacts',                   color: '#5A5040', cols: 8,  rows: 50 },
  { title: '📋 Reference Data',                     color: '#888888', cols: 25, rows: 55 },
];

(async () => {
  // Create the spreadsheet with all sheets
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Estate Executor Organiser' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          index: i,
          title: t.title,
          tabColor: hexObj(t.color),
          gridProperties: { rowCount: t.rows, columnCount: t.cols },
        },
      })),
    },
  });

  const id = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(path.join(__dirname, 'spreadsheet.json'), JSON.stringify({ id, url, sheetMap }, null, 2));

  console.log('Spreadsheet created:', id);
  console.log('URL:', url);
  console.log('Sheets:', Object.keys(sheetMap).join(', '));
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });

function hexObj(h) {
  const c = h.replace('#', '');
  return { red: parseInt(c.slice(0,2),16)/255, green: parseInt(c.slice(2,4),16)/255, blue: parseInt(c.slice(4,6),16)/255 };
}
