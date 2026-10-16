'use strict';
const { sheets, hex, C } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: '🏠 Dashboard',                  color: C.antiqueBrass, cols: 16, rows: 50  },
  { title: '✅ Task Tracker',               color: C.deepForest,   cols: 12, rows: 60  },
  { title: '📦 Packing Tracker',            color: C.mossGreen,    cols: 16, rows: 40  },
  { title: '💰 Moving Budget',              color: C.forestGreen,  cols: 18, rows: 40  },
  { title: '🚛 Moving Company Quotes',      color: C.darkBrass,    cols: 18, rows: 20  },
  { title: '📋 Change of Address',          color: C.sageDark,     cols: 10, rows: 60  },
  { title: '⚡ Utilities & Services',       color: C.tealGreen,    cols: 12, rows: 20  },
  { title: '🏘️ Neighbourhood Comparison',   color: C.slateBlue,    cols: 12, rows: 30  },
  { title: '🏫 School & Childcare',         color: C.warmBrown,    cols: 12, rows: 30  },
  { title: '🗓️ Moving Timeline',            color: C.deepForest,   cols: 10, rows: 20  },
  { title: '📊 Inventory List',             color: C.olive,        cols: 12, rows: 30  },
  { title: '📋 Reference Data',             color: '#888888',      cols: 20, rows: 60  },
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'All-in-One Moving Planner — BuildSheetStudio' },
      sheets: TABS.map((t, i) => ({
        properties: {
          sheetId: i,
          index:   i,
          title:   t.title,
          tabColorStyle: { rgbColor: hex(t.color) },
          gridProperties: { rowCount: t.rows, columnCount: t.cols },
        },
      })),
    },
  });

  const ss = res.data;
  const sheetMap = {};
  ss.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  const out = { id: ss.spreadsheetId, url: ss.spreadsheetUrl, sheetMap };
  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify(out, null, 2));

  console.log('Created:', ss.spreadsheetUrl);
  console.log('ID:', ss.spreadsheetId);
  console.log('Sheets:', Object.keys(sheetMap).join(', '));
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
