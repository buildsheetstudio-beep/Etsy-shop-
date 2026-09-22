'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
const VCHUNK = 200;

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}

async function getSheets() {
  const auth = await getAuth();
  return google.sheets({ version: 'v4', auth });
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  const sheets = await getSheets();
  for (let i = 0; i < data.length; i += VCHUNK) {
    const slice = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice },
    });
    if (data.length > VCHUNK) console.log(`  ${label} chunk ${Math.floor(i / VCHUNK) + 1}`);
  }
}

function colLetter(idx) {
  let s = '';
  let n = idx + 1;
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

(async () => {
  const sheets = await getSheets();

  // ─── WEEKLY SCHEDULE ──────────────────────────────────────────────────────
  console.log('Reading Weekly Schedule col A for time slot rows...');
  const wsColAResp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Weekly Schedule'!A1:A80",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const wsColA = (wsColAResp.data.values || []).map(r => r[0]);

  // Time slot rows = col A has a decimal fraction between 0 and 1 (time value)
  const timeSlotRows = [];
  for (let i = 0; i < wsColA.length; i++) {
    const v = wsColA[i];
    if (typeof v === 'number' && v > 0 && v < 1) timeSlotRows.push(i + 1); // 1-indexed
  }
  console.log(`  Found ${timeSlotRows.length} time-slot rows: ${timeSlotRows[0]}–${timeSlotRows[timeSlotRows.length - 1]}`);

  // Day columns (B4 = week start, Monday +0 … Sunday +6)
  // Event data cols: B D F H J L N  |  Location cols: C E G I K M O
  const dayDefs = [
    { eCol: 'B', lCol: 'C', off: 0 }, // Monday
    { eCol: 'D', lCol: 'E', off: 1 }, // Tuesday
    { eCol: 'F', lCol: 'G', off: 2 }, // Wednesday
    { eCol: 'H', lCol: 'I', off: 3 }, // Thursday
    { eCol: 'J', lCol: 'K', off: 4 }, // Friday
    { eCol: 'L', lCol: 'M', off: 5 }, // Saturday
    { eCol: 'N', lCol: 'O', off: 6 }, // Sunday
  ];

  const wsUpdates = [];
  for (const row of timeSlotRows) {
    for (const { eCol, lCol, off } of dayDefs) {
      // Event cell: show "▶ ID - Name" for each matching event
      const evtFml = `=IFERROR(TEXTJOIN(CHAR(10),TRUE,IF(('Event Details'!$D$6:$D$500=$B$4+${off})*('Event Details'!$E$6:$E$500=$A${row}),"▶ "&'Event Details'!$A$6:$A$500&" - "&'Event Details'!$B$6:$B$500,"")),"")`;
      // Location cell: show room for each matching event
      const locFml = `=IFERROR(TEXTJOIN(CHAR(10),TRUE,IF(('Event Details'!$D$6:$D$500=$B$4+${off})*('Event Details'!$E$6:$E$500=$A${row}),'Event Details'!$H$6:$H$500,"")),"")`;
      wsUpdates.push({ range: `'Weekly Schedule'!${eCol}${row}`, values: [[evtFml]] });
      wsUpdates.push({ range: `'Weekly Schedule'!${lCol}${row}`, values: [[locFml]] });
    }
  }

  console.log(`  Writing ${wsUpdates.length} Weekly Schedule formula cells...`);
  await valuesBatchUpdate(SPREADSHEET_ID, wsUpdates, 'weekly-schedule');
  console.log('✅  Weekly Schedule done.\n');

  // ─── MONTHLY CALENDAR ─────────────────────────────────────────────────────
  console.log('Reading Monthly Calendar header row to locate month-start date column...');

  // Confirm which column in the month header row holds the date serial
  const mcHdrResp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Monthly Calendar'!A4:G4",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const hdrRow = ((mcHdrResp.data.values || [[]])[0] || []);
  let msColIdx = 0;
  for (let c = 0; c < hdrRow.length; c++) {
    if (typeof hdrRow[c] === 'number' && hdrRow[c] > 40000) { msColIdx = c; break; }
  }
  const msCol = colLetter(msColIdx);
  console.log(`  Month-start date in column ${msCol} (row 4 raw value: ${hdrRow[msColIdx]})`);

  // Month header rows (1-indexed): 9-row blocks, 12 months
  // Layout per block: row M = month header, M+1 = day names, M+2..M+7 = 6-week grid, M+8 = separator
  const monthHeaderRows = [4, 13, 22, 31, 40, 49, 58, 67, 76, 85, 94, 103];
  // Calendar columns: A=Sun(0), B=Mon(1), C=Tue(2), D=Wed(3), E=Thu(4), F=Fri(5), G=Sat(6)
  const CAL_COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];

  const mcUpdates = [];
  for (const mhRow of monthHeaderRows) {
    const msRef = `$${msCol}$${mhRow}`; // e.g. $A$4 — absolute reference to month start date

    for (let wk = 0; wk < 6; wk++) {
      const calRow = mhRow + 2 + wk; // calendar grid rows M+2 to M+7
      for (let day = 0; day < 7; day++) {
        const col = CAL_COLS[day];
        const rawOff = wk * 7 + day; // position in 6×7 grid
        // date for this cell = monthStart + rawOff - (WEEKDAY(monthStart,1)-1)
        // WEEKDAY with type=1: 1=Sunday ... 7=Saturday, so -1 gives 0-based Sunday offset
        const dateExpr = `${msRef}+${rawOff}-(WEEKDAY(${msRef},1)-1)`;
        // Show day number + events; blank if outside current month
        const fml = `=IF(${msRef}="","",LET(d,${dateExpr},IF(OR(d<${msRef},MONTH(d)<>MONTH(${msRef})),"",TEXT(DAY(d),"0")&IFERROR(CHAR(10)&TEXTJOIN(CHAR(10),TRUE,IF('Master Event Log'!$D$6:$D$500=d,"• "&'Master Event Log'!$B$6:$B$500,"")),""))))`;
        mcUpdates.push({ range: `'Monthly Calendar'!${col}${calRow}`, values: [[fml]] });
      }
    }
  }

  console.log(`  Writing ${mcUpdates.length} Monthly Calendar formula cells...`);
  await valuesBatchUpdate(SPREADSHEET_ID, mcUpdates, 'monthly-calendar');
  console.log('✅  Monthly Calendar done.\n');

  console.log('✅  All auto-populate formulas applied successfully.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
