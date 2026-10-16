'use strict';
const { getAuth, hex } = require('./lib');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const TABS = [
  { title: '📋 Reference Data',         color: '#9E9E9E' },
  { title: '🎉 Dashboard',               color: '#5B2A4D' },
  { title: '👥 Guest List & RSVP',       color: '#C08D6B' },
  { title: '💰 Budget & Vendor Tracker', color: '#C9A667' },
  { title: '⏰ Day-Of Timeline',          color: '#8FA888' },
  { title: '🪑 Seating Chart',            color: '#EAD9C4' },
  { title: '🥂 Toast & Speech Planner',  color: '#B5573F' },
  { title: '📸 Photo & Shot List',        color: '#5B2A4D' },
  { title: '🌸 Decor, Food & Drinks',    color: '#C08D6B' },
  { title: '✅ Party Checklist',           color: '#8FA888' },
];

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Engagement Party Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          title: t.title,
          index: i,
          tabColor: hex(t.color),
          gridProperties: { rowCount: 300, columnCount: 26 },
        },
      })),
    },
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(path.join(__dirname, 'spreadsheet.json'), JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created:', id);
  console.log('Sheets:', Object.keys(sheetMap).join(', '));
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
