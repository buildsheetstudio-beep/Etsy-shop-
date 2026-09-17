'use strict';
const { sheets } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: '🎉 Dashboard',                 color: '#8B4A6B', cols: 16, rows: 55  },
  { title: '📋 Events Log',                color: '#A5607A', cols: 17, rows: 65  },
  { title: '🗓️ Events Calendar',           color: '#7A4A6B', cols: 10, rows: 32  },
  { title: '✅ Task Tracker',              color: '#B8860B', cols: 10, rows: 65  },
  { title: '💰 Event Budget',              color: '#5A7A4A', cols: 18, rows: 55  },
  { title: '👥 Guest List & RSVP',         color: '#6B4A8B', cols: 22, rows: 65  },
  { title: '🏪 Vendor & Supplier Tracker', color: '#8B6A3A', cols: 18, rows: 45  },
  { title: '⏱️ Run-of-Show',              color: '#3A6B8B', cols: 12, rows: 45  },
  { title: '📊 Event Analytics',           color: '#4A6B5A', cols: 14, rows: 35  },
  { title: '📝 Event Notes & Contacts',    color: '#8B4A4A', cols: 8,  rows: 65  },
  { title: '📋 Reference Data',            color: '#888888', cols: 22, rows: 65  },
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'All-in-One Event Planner ✨' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          index: i,
          title: t.title,
          tabColor: (() => { const c = t.color.replace('#',''); return { red: parseInt(c.slice(0,2),16)/255, green: parseInt(c.slice(2,4),16)/255, blue: parseInt(c.slice(4,6),16)/255 }; })(),
          gridProperties: { rowCount: t.rows, columnCount: t.cols },
        },
      })),
    },
  });

  const id  = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, url, sheetMap }, null, 2));
  console.log('Created:', url);
  console.log('ID:', id);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
