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

(async () => {
  const updates = [];

  // ─── MONTHLY CALENDAR: advance from 2026 → 2027 ──────────────────────────
  // The event log has events in April 2027. The calendar currently shows Jan–Dec
  // 2026 which has no events. Update all 12 month-header cells and the title.

  // Title row
  updates.push({
    range: "'Monthly Calendar'!A1",
    values: [['MONTHLY CALENDAR — JANUARY THROUGH DECEMBER 2027']],
  });

  // 12 month header cells (rows 4, 13, 22, ..., 103 — each holds the 1st of the month)
  // Day cells already reference these via $A$[headerRow], so they update automatically.
  const monthHeaderRows = [4, 13, 22, 31, 40, 49, 58, 67, 76, 85, 94, 103];
  for (let i = 0; i < 12; i++) {
    updates.push({
      range: `'Monthly Calendar'!A${monthHeaderRows[i]}`,
      values: [[`=DATE(2027,${i + 1},1)`]],
    });
  }

  // Rebuild day-cell formulas with a nested LET so events are appended with
  // CHAR(10) only when there are actually events (no trailing newline on empty days).
  // Calendar columns: A=Sun(0) … G=Sat(6); grid rows = headerRow+2 … headerRow+7
  const CAL_COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G'];
  for (const mhRow of monthHeaderRows) {
    const msRef = `$A$${mhRow}`;
    for (let wk = 0; wk < 6; wk++) {
      const calRow = mhRow + 2 + wk;
      for (let day = 0; day < 7; day++) {
        const col = CAL_COLS[day];
        const rawOff = wk * 7 + day;
        // Use nested LET: outer computes the date, inner avoids double TEXTJOIN call
        const fml =
          `=IF(${msRef}="","",LET(ms,${msRef},d,ms+${rawOff}-(WEEKDAY(ms,1)-1),` +
          `IF(OR(d<ms,MONTH(d)<>MONTH(ms)),"",` +
          `LET(evts,TEXTJOIN(CHAR(10),TRUE,IF('Master Event Log'!$D$6:$D$500=d,"• "&'Master Event Log'!$B$6:$B$500,"")),` +
          `TEXT(DAY(d),"0")&IF(evts="","",CHAR(10)&evts)))))`;
        updates.push({ range: `'Monthly Calendar'!${col}${calRow}`, values: [[fml]] });
      }
    }
  }

  console.log(`Writing ${updates.length} Monthly Calendar updates...`);
  await valuesBatchUpdate(SPREADSHEET_ID, updates, 'mc-fix');
  console.log('✅  Monthly Calendar updated to 2027.\n');

  // ─── WEEKLY SCHEDULE: move B4 to Monday of the first event week ───────────
  // Events start on Apr 7, 2027 (Wednesday). Monday of that week = Apr 5, 2027.
  // B4 is user-editable; setting it to this week ensures events are immediately visible.
  const wsUpdates = [
    { range: "'Weekly Schedule'!B4", values: [['=DATE(2027,4,5)']] },
  ];
  await valuesBatchUpdate(SPREADSHEET_ID, wsUpdates, 'ws-fix');
  console.log('✅  Weekly Schedule B4 set to Mon Apr 5 2027 (week of first events).\n');

  console.log('✅  Done. Events should now appear on both tabs.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
