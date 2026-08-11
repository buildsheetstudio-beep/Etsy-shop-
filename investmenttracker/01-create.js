'use strict';
const { sheets } = require('./lib');
const fs = require('fs');

const TITLE = 'Ultimate Investment Tracker';

const TABS = [
  { title: 'Reference Data',          rows: 600  },
  { title: 'Portfolio Setup',         rows: 500  },
  { title: 'Account Tracker',         rows: 400  },
  { title: 'Price Updates',           rows: 1100 },
  { title: 'Transaction Log',         rows: 5100 },
  { title: 'Holdings',                rows: 1100 },
  { title: 'Portfolio Allocation',    rows: 500  },
  { title: 'Growth Forecast',         rows: 500  },
  { title: 'Contributions & Savings', rows: 1600 },
  { title: 'Net Worth Tracker',       rows: 500  },
  { title: 'Goals & Milestones',      rows: 600  },
  { title: 'Dividend Income',         rows: 2100 },
  { title: 'Dividend Calendar',       rows: 500  },
  { title: 'Watchlist',               rows: 1100 },
  { title: 'Annual Review',           rows: 500  },
  { title: 'Investment Dashboard',    rows: 500  },
];

(async () => {
  console.log(`Creating spreadsheet: ${TITLE}`);

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: TITLE, locale: 'en_US', timeZone: 'America/New_York' },
      sheets: TABS.map((tab, i) => ({
        properties: {
          sheetId: i,
          title: tab.title,
          index: i,
          gridProperties: { rowCount: tab.rows, columnCount: 26 },
        },
      })),
    },
  });

  const spreadsheetId = res.data.spreadsheetId;
  const spreadsheetUrl = res.data.spreadsheetUrl;

  console.log('Spreadsheet ID:', spreadsheetId);
  console.log('URL:', spreadsheetUrl);

  const sheetMap = {};
  res.data.sheets.forEach(s => {
    sheetMap[s.properties.title] = s.properties.sheetId;
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ updateSheetProperties: {
      properties: { sheetId: 0, hidden: true },
      fields: 'hidden',
    }}]},
  });
  console.log('  Reference Data tab hidden.');

  fs.writeFileSync(
    __dirname + '/spreadsheet.json',
    JSON.stringify({ id: spreadsheetId, url: spreadsheetUrl, sheetMap }, null, 2)
  );

  console.log('✅ Spreadsheet created — spreadsheet.json written.');
  TABS.forEach((t, i) => console.log(`  [${i}] ${t.title} (${t.rows} rows)`));
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
