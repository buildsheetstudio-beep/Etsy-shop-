'use strict';
// Fix: add ARRAYFORMULA inside every TEXTJOIN+IF so the IF evaluates element-wise.
// Without ARRAYFORMULA, IF(range*range, range, "") does NOT filter correctly in
// Google Sheets — it either returns the whole range (location column shows every room)
// or silently collapses to one row (event column misses some days).
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

(async () => {
  const sheets = await getSheets();

  // ─── WEEKLY SCHEDULE: re-read time slot rows ─────────────────────────────
  console.log('Reading Weekly Schedule col A...');
  const wsResp = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'Weekly Schedule'!A1:A80",
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const wsColA = (wsResp.data.values || []).map(r => r[0]);
  const timeSlotRows = [];
  for (let i = 0; i < wsColA.length; i++) {
    const v = wsColA[i];
    if (typeof v === 'number' && v > 0 && v < 1) timeSlotRows.push(i + 1);
  }
  console.log(`  ${timeSlotRows.length} time-slot rows: ${timeSlotRows[0]}–${timeSlotRows[timeSlotRows.length - 1]}`);

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
      // ARRAYFORMULA ensures IF evaluates element-wise across all 495 rows
      const cond = `('Event Details'!$D$6:$D$500=$B$4+${off})*('Event Details'!$E$6:$E$500=$A${row})`;
      const evtFml =
        `=IFERROR(TEXTJOIN(CHAR(10),TRUE,ARRAYFORMULA(IF(${cond},"▶ "&'Event Details'!$A$6:$A$500&" - "&'Event Details'!$B$6:$B$500,""))),"")`;
      const locFml =
        `=IFERROR(TEXTJOIN(CHAR(10),TRUE,ARRAYFORMULA(IF(${cond},'Event Details'!$H$6:$H$500,""))),"")`;
      wsUpdates.push({ range: `'Weekly Schedule'!${eCol}${row}`, values: [[evtFml]] });
      wsUpdates.push({ range: `'Weekly Schedule'!${lCol}${row}`, values: [[locFml]] });
    }
  }

  console.log(`  Writing ${wsUpdates.length} Weekly Schedule formula cells...`);
  await valuesBatchUpdate(SPREADSHEET_ID, wsUpdates, 'ws-arfml');
  console.log('✅  Weekly Schedule done.\n');

  // ─── MONTHLY CALENDAR: rebuild day cells with ARRAYFORMULA ───────────────
  // Month header cells ($A$[row]) now hold DATE(2027,m,1) from the previous fix.
  // Calendar columns: A=Sun(0) … G=Sat(6); grid rows = headerRow+2 … headerRow+7
  const monthHeaderRows = [4, 13, 22, 31, 40, 49, 58, 67, 76, 85, 94, 103];
  const CAL_COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];

  const mcUpdates = [];
  for (const mhRow of monthHeaderRows) {
    const msRef = `$A$${mhRow}`;
    for (let wk = 0; wk < 6; wk++) {
      const calRow = mhRow + 2 + wk;
      for (let day = 0; day < 7; day++) {
        const col = CAL_COLS[day];
        const rawOff = wk * 7 + day;
        const fml =
          `=IF(${msRef}="","",LET(ms,${msRef},d,ms+${rawOff}-(WEEKDAY(ms,1)-1),` +
          `IF(OR(d<ms,MONTH(d)<>MONTH(ms)),"",` +
          `LET(evts,TEXTJOIN(CHAR(10),TRUE,ARRAYFORMULA(IF('Master Event Log'!$D$6:$D$500=d,"• "&'Master Event Log'!$B$6:$B$500,""))),` +
          `TEXT(DAY(d),"0")&IF(evts="","",CHAR(10)&evts)))))`;
        mcUpdates.push({ range: `'Monthly Calendar'!${col}${calRow}`, values: [[fml]] });
      }
    }
  }

  console.log(`  Writing ${mcUpdates.length} Monthly Calendar formula cells...`);
  await valuesBatchUpdate(SPREADSHEET_ID, mcUpdates, 'mc-arfml');
  console.log('✅  Monthly Calendar done.\n');

  console.log('✅  All ARRAYFORMULA-based formulas written. Events should now display correctly.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
