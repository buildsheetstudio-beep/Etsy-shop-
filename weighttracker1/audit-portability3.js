'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
async function getAuth() {
  const s = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const c = s.installed || s.web;
  const o = new google.auth.OAuth2(c.client_id, c.client_secret, c.redirect_uris[0]);
  o.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return o;
}
(async () => {
  const sheets = google.sheets({ version: 'v4', auth: await getAuth() });

  // 1. Check if any formula uses the named range names
  const TABS = ['Event Dashboard','Event Setup','Master Event Log','Event Details',
                'Guests & Attendees','Vendors & Venues','Event Tasks',
                'Equipment & Supplies','Budget & Expenses','Weekly Schedule','Monthly Calendar'];
  const namedRangeNames = ['REF_CONTRACT_STATUS','REF_VENUE_TYPE','REF_PAYMENT_STATUS',
    'REF_RECURRENCE','REF_VENDOR_CATEGORY','REF_ATTENDANCE_STATUS','REF_REVENUE_TYPE',
    'REF_PRIORITY','REF_EXPENSE_CATEGORY','REF_EVENT_CATEGORY','REF_EVENT_TYPE',
    'REF_MEAL_CHOICE','REF_EVENT_STATUS','REF_EQUIPMENT_STATUS','REF_TASK_STATUS',
    'REF_CATEGORY_COLORS','REF_TICKET_TYPE'];

  const namedRangeUsage = {};
  for (const tab of TABS) {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'${tab}'!A1:Z500`,
      valueRenderOption: 'FORMULA',
    });
    (res.data.values || []).forEach((row, ri) => {
      row.forEach((cell) => {
        if (typeof cell !== 'string' || !cell.startsWith('=')) return;
        namedRangeNames.forEach(name => {
          if (cell.includes(name)) {
            if (!namedRangeUsage[name]) namedRangeUsage[name] = [];
            namedRangeUsage[name].push(tab);
          }
        });
      });
    });
  }
  console.log('=== NAMED RANGE USAGE IN FORMULAS ===');
  if (Object.keys(namedRangeUsage).length === 0) {
    console.log('  None of the 17 named ranges are referenced in any formula.');
  } else {
    for (const [name, tabs] of Object.entries(namedRangeUsage)) {
      console.log(`  ${name}: used in ${[...new Set(tabs)].join(', ')}`);
    }
  }

  // 2. Check data validation — are any using range-based (not hardcoded) sources?
  const dvMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(title),data(rowData(values(dataValidation))))',
  });
  console.log('\n=== DATA VALIDATION — RANGE-BASED (non-hardcoded) DROPDOWNS ===');
  let rangeDvCount = 0;
  dvMeta.data.sheets.forEach(sheet => {
    (sheet.data || []).forEach(grid => {
      (grid.rowData || []).forEach((row, ri) => {
        (row.values || []).forEach((cell, ci) => {
          const dv = cell.dataValidation;
          if (!dv) return;
          if (dv.condition.type === 'ONE_OF_RANGE') {
            const colLetter = String.fromCharCode(65 + ci);
            rangeDvCount++;
            const src = dv.condition.values && dv.condition.values[0];
            console.log(`  "${sheet.properties.title}"!${colLetter}${ri+1}: ${src ? src.userEnteredValue : 'no source'}`);
          }
        });
      });
    });
  });
  if (rangeDvCount === 0) console.log('  None found — all dropdowns use hardcoded ONE_OF_LIST values.');

  // 3. Check Monthly Calendar full formula (it's truncated in sample)
  const cal = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Monthly Calendar'!A6",
    valueRenderOption: 'FORMULA',
  });
  console.log('\n=== MONTHLY CALENDAR A6 — FULL FORMULA ===');
  console.log((cal.data.values || [[]])[0][0] || '(empty)');

  // 4. Check Weekly Schedule B4 (week number cell)
  const ws = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Weekly Schedule'!G4:I4",
    valueRenderOption: 'FORMULA',
  });
  console.log('\n=== WEEKLY SCHEDULE — WEEKNUM cells ===');
  (ws.data.values || []).forEach((row, ri) => {
    row.forEach((cell, ci) => {
      if (cell) console.log(`  ${String.fromCharCode(71+ci)}4: ${cell}`);
    });
  });

  // 5. Check Budget & Expenses formulas — are there any complex ones?
  const budget = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Budget & Expenses'!A1:Z50",
    valueRenderOption: 'FORMULA',
  });
  console.log('\n=== BUDGET & EXPENSES — unique formula patterns ===');
  const seen = new Set();
  (budget.data.values || []).forEach((row, ri) => {
    row.forEach((cell, ci) => {
      if (typeof cell !== 'string' || !cell.startsWith('=')) return;
      const key = cell.slice(0, 60);
      if (!seen.has(key)) { seen.add(key); console.log(`  ${String.fromCharCode(65+ci)}${ri+1}: ${cell.slice(0, 100)}`); }
    });
  });
})().catch(e => { console.error(e.message || e); process.exit(1); });
