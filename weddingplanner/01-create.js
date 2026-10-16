'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SECRET_PATH = path.join(__dirname, 'client_secret.json');
const TOKEN_PATH  = path.join(__dirname, 'tokens.json');

function getAuth() {
  const { client_secret, client_id, redirect_uris } =
    JSON.parse(fs.readFileSync(SECRET_PATH)).installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
  return oAuth2Client;
}

const sheets = google.sheets({ version: 'v4', auth: getAuth() });

// Tab order: sheetId matches position
const TABS = [
  { title: 'Reference Data',       sheetId: 0,  rowCount: 200,  columnCount: 20  },
  { title: 'Wedding Setup',        sheetId: 1,  rowCount: 200,  columnCount: 10  },
  { title: 'Guest List & RSVP',    sheetId: 2,  rowCount: 600,  columnCount: 39  },
  { title: 'Invitation Tracker',   sheetId: 3,  rowCount: 500,  columnCount: 20  },
  { title: 'Seating Chart',        sheetId: 4,  rowCount: 600,  columnCount: 20  },
  { title: 'Table Planner',        sheetId: 5,  rowCount: 200,  columnCount: 14  },
  { title: 'Vendors & Venue',      sheetId: 6,  rowCount: 350,  columnCount: 24  },
  { title: 'Reception Planner',    sheetId: 7,  rowCount: 200,  columnCount: 16  },
  { title: 'Food, Drinks & Cake',  sheetId: 8,  rowCount: 300,  columnCount: 16  },
  { title: 'Budget & Payments',    sheetId: 9,  rowCount: 500,  columnCount: 16  },
  { title: 'Wedding Checklist',    sheetId: 10, rowCount: 500,  columnCount: 16  },
  { title: 'Dashboard',            sheetId: 11, rowCount: 300,  columnCount: 14  },
];

(async () => {
  // Create spreadsheet with first tab
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Wedding Planner' },
      sheets: [{
        properties: {
          sheetId: TABS[0].sheetId,
          title: TABS[0].title,
          index: 0,
          gridProperties: { rowCount: TABS[0].rowCount, columnCount: TABS[0].columnCount },
        },
      }],
    },
  });

  const id = res.data.spreadsheetId;
  console.log('Created:', id);

  // Add remaining sheets
  const addReqs = TABS.slice(1).map((t, i) => ({
    addSheet: {
      properties: {
        sheetId: t.sheetId,
        title: t.title,
        index: i + 1,
        gridProperties: { rowCount: t.rowCount, columnCount: t.columnCount },
      },
    },
  }));

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: id,
    requestBody: { requests: addReqs },
  });

  const sheetMap = {};
  TABS.forEach(t => { sheetMap[t.title] = t.sheetId; });

  fs.writeFileSync(
    path.join(__dirname, 'spreadsheet.json'),
    JSON.stringify({ id, sheetMap }, null, 2)
  );

  console.log('✓ All 12 tabs created');
  console.log('URL:', `https://docs.google.com/spreadsheets/d/${id}`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
