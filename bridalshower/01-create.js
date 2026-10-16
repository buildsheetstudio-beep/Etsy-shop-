'use strict';
const { getAuth, google } = require('./lib');
const fs = require('fs');
const path = require('path');

function rgb(h) {
  return { red: parseInt(h.slice(1,3),16)/255, green: parseInt(h.slice(3,5),16)/255, blue: parseInt(h.slice(5,7),16)/255 };
}

const TABS = [
  { title: '🥂 Dashboard',                    color: '#8B5E6A' },
  { title: '👰 Guest List & Seating',         color: '#A5707A' },
  { title: '💰 Budget & Expenses',            color: '#7A8B6A' },
  { title: '🗓️ Party Itinerary & Games',      color: '#7A6A8B' },
  { title: '💍 Venue, Vendors & Bridesmaids', color: '#8B7A5E' },
  { title: '🎁 Gift Tracker & Wishlist',      color: '#C9A96E' },
  { title: '📋 Reference Data',               color: '#AAAAAA' },
];

(async () => {
  const auth   = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Bridal Shower Planner' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          index: i,
          title: t.title,
          tabColorStyle: { rgbColor: rgb(t.color) },
          gridProperties: { rowCount: 200, columnCount: 26 },
        },
      })),
    },
  });

  const id  = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(path.join(__dirname,'spreadsheet.json'), JSON.stringify({ id, url, sheetMap }, null, 2));
  console.log('Created:', url);
  console.log('ID:', id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
