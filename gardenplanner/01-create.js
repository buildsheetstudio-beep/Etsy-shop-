'use strict';
const { sheets, hex, C } = require('./lib');
const fs = require('fs');

const TABS = [
  { title: 'Dashboard',               color: C.terracotta,   cols: 16, rows: 50  },
  { title: 'Plant Library',           color: C.olive,        cols: 14, rows: 50  },
  { title: 'Planting Schedule',       color: C.oliveLight,   cols: 16, rows: 30  },
  { title: 'Garden Layout',           color: C.soil,         cols: 35, rows: 40  },
  { title: 'Companion Planting Guide',color: C.sage,         cols: 26, rows: 35  },
  { title: 'Watering & Care Schedule',color: C.terracottaLight, cols: 14, rows: 30 },
  { title: 'Harvest & Yield Tracker', color: C.wheat,        cols: 14, rows: 30  },
  { title: 'Pest & Disease Log',      color: C.warningRed,   cols: 14, rows: 25  },
  { title: 'Garden Budget',           color: C.warmGray,     cols: 18, rows: 30  },
  { title: 'Seed Inventory',          color: C.oliveLight,   cols: 14, rows: 25  },
  { title: 'Garden Journal',          color: C.soil,         cols: 12, rows: 80  },
  { title: 'Reference Data',          color: C.lightGray,    cols: 18, rows: 40  },
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'All-Season Garden Planner — BuildSheetStudio 2025' },
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
