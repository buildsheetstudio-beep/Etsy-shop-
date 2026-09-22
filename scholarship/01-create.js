'use strict';
const { getAuth, google } = require('./lib');
const fs = require('fs');

(async () => {
  const auth   = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  const drive  = google.drive({ version: 'v3', auth });

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Scholarship & Bursary Tracker — BuildSheetStudio' },
      sheets: [
        { properties: { title: '🎓 Dashboard',              sheetId: 0, index: 0 } },
        { properties: { title: '📋 Application Tracker',   sheetId: 1, index: 1 } },
        { properties: { title: '⏰ Deadlines & To-Do',     sheetId: 2, index: 2 } },
        { properties: { title: '🏆 Results & Awards Log',  sheetId: 3, index: 3 } },
        { properties: { title: '📋 Reference Data',        sheetId: 4, index: 4 } },
      ],
    },
  });

  const id = res.data.spreadsheetId;
  console.log('Created:', id);

  // Set background of all sheets to white (default cream replacement)
  const reqs = [];

  // Tab colors
  const tabColors = [
    { sheetId: 0, rgb: { red: 0.102, green: 0.227, blue: 0.545 } }, // #1A3A8B
    { sheetId: 1, rgb: { red: 0.102, green: 0.337, blue: 0.769 } }, // #1A56C4
    { sheetId: 2, rgb: { red: 0.788, green: 0.659, blue: 0.000 } }, // #C9A800
    { sheetId: 3, rgb: { red: 0.102, green: 0.420, blue: 0.227 } }, // #1A6B3A
    { sheetId: 4, rgb: { red: 0.600, green: 0.600, blue: 0.600 } }, // grey
  ];
  tabColors.forEach(({ sheetId, rgb }) => {
    reqs.push({ updateSheetProperties: {
      properties: { sheetId, tabColorStyle: { rgbColor: rgb } },
      fields: 'tabColorStyle',
    }});
  });

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: id, requestBody: { requests: reqs } });

  // Share publicly readable
  await drive.permissions.create({
    fileId: id,
    requestBody: { role: 'reader', type: 'anyone' },
  });

  const sheetMap = {
    '🎓 Dashboard': 0,
    '📋 Application Tracker': 1,
    '⏰ Deadlines & To-Do': 2,
    '🏆 Results & Awards Log': 3,
    '📋 Reference Data': 4,
  };

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({
    id,
    url: `https://docs.google.com/spreadsheets/d/${id}/edit`,
    sheetMap,
  }, null, 2));

  console.log('URL:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
  console.log('Done');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
