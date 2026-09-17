'use strict';
const { sheets } = require('./lib');
const fs = require('fs');
const path = require('path');

const TABS = [
  { title: 'Reference Data',    sheetId: 0  },
  { title: 'Event Setup',       sheetId: 1  },
  { title: 'Master Event Log',  sheetId: 2  },
  { title: 'Event Details',     sheetId: 3  },
  { title: 'Guests & Attendees',sheetId: 4  },
  { title: 'Vendors & Venues',  sheetId: 5  },
  { title: 'Event Tasks',       sheetId: 6  },
  { title: 'Equipment & Supplies', sheetId: 7 },
  { title: 'Budget & Expenses', sheetId: 8  },
  { title: 'Weekly Schedule',   sheetId: 9  },
  { title: 'Monthly Calendar',  sheetId: 10 },
  { title: 'Event Dashboard',   sheetId: 11 },
];

(async () => {
  // Create spreadsheet
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Event Planner' },
      sheets: TABS.map(t => ({
        properties: {
          sheetId: t.sheetId,
          title: t.title,
          gridProperties: { rowCount: 1500, columnCount: 30 },
        }
      })),
    }
  });

  const id = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  console.log('Created:', url);

  fs.writeFileSync(
    path.join(__dirname, 'spreadsheet.json'),
    JSON.stringify({ id, url }, null, 2)
  );
  console.log('spreadsheet.json saved.');
})().catch(e => { console.error(e.message || e); process.exit(1); });
