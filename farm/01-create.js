'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const { batchUpdate, gridRange, hex, C } = require('./lib');

function getAuth() {
  const s = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
  const { client_id, client_secret, redirect_uris } = s.installed;
  const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
  return auth;
}

(async () => {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  const TABS = [
    { title: 'Reference Data',          sheetId: 0,  index: 0  },
    { title: 'Farm Setup',              sheetId: 1,  index: 1  },
    { title: 'Income Log',              sheetId: 2,  index: 2  },
    { title: 'Expense Log',             sheetId: 3,  index: 3  },
    { title: 'Monthly View',            sheetId: 4,  index: 4  },
    { title: 'Annual Summary',          sheetId: 5,  index: 5  },
    { title: 'Enterprise Profitability',sheetId: 6,  index: 6  },
    { title: 'Budget vs Actual',        sheetId: 7,  index: 7  },
    { title: 'Profit Goals',            sheetId: 8,  index: 8  },
    { title: 'Tax Deduction Summary',   sheetId: 9,  index: 9  },
    { title: 'Dashboard',               sheetId: 10, index: 10 },
  ];

  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Farm Income & Expense Tracker' },
      sheets: TABS.map(t => ({ properties: { title: t.title, sheetId: t.sheetId, index: t.index } }))
    }
  });

  const id = res.data.spreadsheetId;
  const sheetMap = {};
  for (const s of res.data.sheets) sheetMap[s.properties.title] = s.properties.sheetId;

  // Baseline background on all tabs
  const reqs = [];
  for (const sid of Object.values(sheetMap)) {
    reqs.push({ repeatCell: {
      range: gridRange(sid, 0, 1000, 0, 30),
      cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
      fields: 'userEnteredFormat.backgroundColor'
    }});
  }
  await batchUpdate(id, reqs, 'bg');

  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, sheetMap }, null, 2));
  console.log('Created:', id);
  console.log('Sheets:', JSON.stringify(sheetMap, null, 2));
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
