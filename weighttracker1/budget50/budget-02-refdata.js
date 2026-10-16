'use strict';
/**
 * Step 2: Populate Reference Data tab (sheetId 0) with all lookup lists.
 * This tab is the source for all data validation dropdowns.
 */
const { google } = require('googleapis');
const fs = require('fs'), path = require('path');
const AUTH_DIR = path.join(__dirname, '..');
const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'budget-config.json')));
const SID = CFG.spreadsheetId;
const REF_SID = CFG.sheets['Reference Data'];

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const auth = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(AUTH_DIR, 'tokens.json'))));
  return auth;
}

function sv(v) { return { userEnteredValue: { stringValue: String(v) } }; }

(async () => {
  const auth = await getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // ── Lookup lists ──────────────────────────────────────────────────────────
  // Columns A/B/C/D/E/F/G/H each hold a distinct lookup list
  // Row 1 = header, data from row 2

  const listHeaders = ['Cadence','Buckets','Transaction Types','Payment Methods',
                       'Bill Frequencies','Goal Types','Payment Statuses','Tab Navigation'];
  const lists = {
    Cadence:            ['Weekly','Biweekly','Monthly'],
    Buckets:            ['Needs','Wants','Future','Income','Transfer','Uncategorized'],
    'Transaction Types':['Income','Expense','Savings Transfer','Transfer In','Transfer Out','Refund'],
    'Payment Methods':  ['Bank','Credit Card','Cash','Other'],
    'Bill Frequencies': ['Weekly','Biweekly','Monthly','Quarterly','Yearly'],
    'Goal Types':       ['Emergency Fund','Sinking Fund','Savings','Extra Debt Payment'],
    'Payment Statuses': ['Scheduled','Paid','Late','Overdue'],
    'Tab Navigation':   ['Start Here','Budget Dashboard','Transactions','Bills & Goals','Yearly Insights'],
  };

  // Build rows: header in row 1 (0-indexed: row 0), data below
  const maxRows = Math.max(...Object.values(lists).map(a => a.length));
  const allRows = [];

  // Header row
  allRows.push(listHeaders.map(h => sv(h)));

  // Data rows
  for (let r = 0; r < maxRows; r++) {
    const row = Object.values(lists).map(arr => r < arr.length ? sv(arr[r]) : sv(''));
    allRows.push(row);
  }

  // ── Keyword-to-category mapping (column J/K) ───────────────────────────
  // Used in Transactions CSV staging to auto-suggest category
  const kwHeaders = ['Keyword (contains)','Suggested Category'];
  const kwData = [
    ['rent','Housing'], ['lease','Housing'],
    ['electric','Utilities'], ['gas bill','Utilities'], ['water','Utilities'],
    ['internet','Utilities'], ['cell phone','Utilities'], ['phone bill','Utilities'],
    ['grocery','Groceries'], ['freshmart','Groceries'], ['whole foods','Groceries'],
    ['transit','Transportation'], ['uber','Transportation'], ['lyft','Transportation'],
    ['gas station','Transportation'], ['parking','Transportation'],
    ['insurance','Insurance'], ['safeGuard','Insurance'],
    ['doctor','Healthcare'], ['pharmacy','Healthcare'], ['clinic','Healthcare'],
    ['paycheck','Paycheck'], ['employer','Paycheck'], ['direct deposit','Paycheck'],
    ['restaurant','Dining Out'], ['bistro','Dining Out'], ['cafe','Coffee'],
    ['netflix','Streaming'], ['spotify','Streaming'], ['streamflow','Streaming'],
    ['gym','Hobbies'], ['fitness','Hobbies'],
    ['amazon','Shopping'], ['clothing','Shopping'],
    ['emergency fund','Emergency Fund'], ['savings','Emergency Fund'],
    ['vacation','Vacation Fund'], ['student loan','Extra Debt Payment'],
    ['transfer','Account Transfer'],
  ];

  // Write main lists to columns A-H (indices 0-7)
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [
        {
          updateCells: {
            range: { sheetId: REF_SID, startRowIndex: 0, endRowIndex: allRows.length,
                     startColumnIndex: 0, endColumnIndex: 8 },
            rows: allRows.map(row => ({
              values: row.map(cell => ({
                userEnteredValue: cell.userEnteredValue,
                userEnteredFormat: { textFormat: { bold: row === allRows[0] } },
              })),
            })),
            fields: 'userEnteredValue,userEnteredFormat.textFormat.bold',
          },
        },
      ],
    },
  });

  // Write keyword mapping to columns J-K (indices 9-10)
  const kwRows = [kwHeaders, ...kwData].map(row =>
    row.map(v => ({ userEnteredValue: { stringValue: String(v) } }))
  );
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: `'Reference Data'!J1:K${kwRows.length}`, values: [kwHeaders, ...kwData] },
      ],
    },
  });

  // Header row formatting
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SID,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: { sheetId: REF_SID, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 12 },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.208, green: 0.282, blue: 0.243 }, // deep forest
                textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
              },
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat)',
          },
        },
        // Column A-H width
        ...['A','B','C','D','E','F','G','H'].map((_, i) => ({
          updateDimensionProperties: {
            range: { sheetId: REF_SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
            properties: { pixelSize: 160 },
            fields: 'pixelSize',
          },
        })),
        // Keyword columns J-K
        {
          updateDimensionProperties: {
            range: { sheetId: REF_SID, dimension: 'COLUMNS', startIndex: 9, endIndex: 11 },
            properties: { pixelSize: 200 },
            fields: 'pixelSize',
          },
        },
        // Freeze row 1
        { updateSheetProperties: { properties: { sheetId: REF_SID, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } },
      ],
    },
  });

  console.log('✅ Reference Data populated:');
  Object.entries(lists).forEach(([k, v]) => console.log(`   ${k}: ${v.join(', ')}`));
  console.log(`   Keyword mapping: ${kwData.length} rows`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
