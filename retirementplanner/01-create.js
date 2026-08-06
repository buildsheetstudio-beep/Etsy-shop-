'use strict';
const { sheets } = require('./lib');
const fs = require('fs');

const TITLE = 'Ultimate Retirement Planner';

const TABS = [
  { title: 'Reference Data',             hidden: true  },
  { title: 'Personal & Household Setup', hidden: false },
  { title: 'Retirement Accounts',        hidden: false },
  { title: 'Retirement Income',          hidden: false },
  { title: 'Retirement Expenses',        hidden: false },
  { title: 'Scenario Planner',           hidden: false },
  { title: 'Growth Forecast',            hidden: false },
  { title: 'Milestone Tracker',          hidden: false },
  { title: 'Withdrawal Strategy',        hidden: false },
  { title: 'Retirement Readiness',       hidden: false },
  { title: 'Retirement Dashboard',       hidden: false },
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
          gridProperties: { rowCount: 1000, columnCount: 26 },
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

  // Hide Reference Data tab (sheetId 0) now that other sheets exist
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
  TABS.forEach((t, i) => console.log(`  [${i}] ${t.title}${t.hidden ? ' (hidden)' : ''}`));
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
