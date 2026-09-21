'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';
const WEEKLY_SHEET_ID = 9;

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

function hexToRgb(hex) {
  return {
    red:   parseInt(hex.slice(1, 3), 16) / 255,
    green: parseInt(hex.slice(3, 5), 16) / 255,
    blue:  parseInt(hex.slice(5, 7), 16) / 255,
  };
}

// Category → hex color (from Reference Data A2:B18)
const CATEGORY_COLORS = {
  Business:    '#C9DCEF',
  Conference:  '#D8CDE8',
  Workshop:    '#F2D2B6',
  Networking:  '#C5E0DD',
  Community:   '#CBDCC5',
  Nonprofit:   '#E9DBA7',
  School:      '#D5E4F2',
  Fundraiser:  '#E8C0B8',
  Social:      '#E9CCD8',
  Birthday:    '#F3D0BB',
  Wedding:     '#DDD0DF',
  Family:      '#EEE2C8',
  Holiday:     '#D0E5D8',
  Sports:      '#C6D9E8',
  Virtual:     '#DDD5ED',
  Other:       '#E3E1DE',
};

// CF formula: checks if Master Event Log has an event of this category on the day
// matching this column. COLUMN()-2 gives 0-based offset; /2 groups event+room pairs.
// INDIRECT needed because Google Sheets rejects cross-sheet refs in CF formulas.
function cfFormula(category) {
  return `=COUNTIFS(INDIRECT("'Master Event Log'!D6:D500"),$B$4+INT((COLUMN()-2)/2),INDIRECT("'Master Event Log'!G6:G500"),"${category}")>0`;
}

// Build one schedule cell formula
// isRoom: true = pull 'Event Details'!H (Room/Location); false = pull event ID+name
function cellFormula(day, row, isRoom) {
  const dateCond = `('Event Details'!$D$6:$D$500=$B$4+${day})*('Event Details'!$E$6:$E$500=$A${row})`;
  if (isRoom) {
    return `=IFERROR(TEXTJOIN(CHAR(10),TRUE,ARRAYFORMULA(IF(${dateCond},'Event Details'!$H$6:$H$500,""))),"")`;
  }
  return `=IFERROR(TEXTJOIN(CHAR(10),TRUE,ARRAYFORMULA(IF(${dateCond},"▶ "&'Event Details'!$A$6:$A$500&" - "&'Event Details'!$B$6:$B$500,""))),"")`;
}

(async () => {
  const sheets = await getSheets();

  // ── 1. Rebuild all schedule cell formulas (rows 6-29, cols B-O) ───────────
  // Column order: B=day0evt, C=day0room, D=day1evt, E=day1room, ... N=day6evt, O=day6room
  // 24 time-slot rows, each 14 columns wide
  console.log('Building schedule formula grid (rows 6-29)...');
  const scheduleValues = [];
  for (let sheetRow = 6; sheetRow <= 29; sheetRow++) {
    const rowData = [];
    for (let col = 0; col < 14; col++) {
      const day    = Math.floor(col / 2);   // 0-6 for B-O pairs
      const isRoom = (col % 2 === 1);       // odd col = room col
      rowData.push(cellFormula(day, sheetRow, isRoom));
    }
    scheduleValues.push(rowData);
  }
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [{ range: "'Weekly Schedule'!B6:O29", values: scheduleValues }],
    },
  });
  console.log('  ✓ Schedule formulas rebuilt (24 rows × 14 cols)');

  // ── 2. Fix conditional formatting ─────────────────────────────────────────
  // Current CF rules all have format:{} (empty) — no color is applied.
  // Fix: delete all 13 existing rules (highest-index first to avoid shift),
  //      then add 16 new rules (all Reference Data categories) with correct colors.
  console.log('\nDeleting existing CF rules...');
  const NUM_EXISTING_CF = 13;
  const deleteRequests = [];
  for (let i = NUM_EXISTING_CF - 1; i >= 0; i--) {
    deleteRequests.push({ deleteConditionalFormatRule: { sheetId: WEEKLY_SHEET_ID, index: i } });
  }
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: deleteRequests },
  });
  console.log(`  ✓ Deleted ${NUM_EXISTING_CF} empty CF rules`);

  console.log('\nAdding CF rules with colors...');
  const cfRange = {
    sheetId: WEEKLY_SHEET_ID,
    startRowIndex:    5,   // row 6 (0-based)
    endRowIndex:      29,  // row 29 inclusive (0-based end = 29)
    startColumnIndex: 1,   // col B (0-based)
    endColumnIndex:   15,  // col O inclusive (0-based end = 15)
  };

  const addCfRequests = Object.entries(CATEGORY_COLORS).map(([cat, hex]) => ({
    addConditionalFormatRule: {
      rule: {
        ranges: [cfRange],
        booleanRule: {
          condition: {
            type: 'CUSTOM_FORMULA',
            values: [{ userEnteredValue: cfFormula(cat) }],
          },
          format: {
            backgroundColor: hexToRgb(hex),
          },
        },
      },
      index: 0,
    },
  }));

  // Send in batches of 10 to stay under API limits
  for (let i = 0; i < addCfRequests.length; i += 10) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: addCfRequests.slice(i, i + 10) },
    });
  }
  console.log(`  ✓ Added ${Object.keys(CATEGORY_COLORS).length} CF rules with correct background colors`);

  // ── 3. Verify a sample of cells ───────────────────────────────────────────
  console.log('\nVerifying sample cells (FORMULA)...');
  const check = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: [
      "'Weekly Schedule'!F6:I6",   // Fri event/room + Sat event/room for row 6
      "'Weekly Schedule'!F9:I9",   // same for 10:00 row
      "'Weekly Schedule'!F12:I12", // same for 11:30 row
    ],
    valueRenderOption: 'FORMULA',
  });
  for (const vr of check.data.valueRanges) {
    console.log(`  ${vr.range}:`);
    (vr.values || []).forEach(row => {
      // Print just the day+type part to keep output readable
      row.forEach((f, ci) => {
        const dayMatch  = f.match(/\$B\$4\+(\d)/);
        const isRoom    = f.includes("'Event Details'!$H$");
        const timeMatch = f.match(/\$A(\d+)/);
        const day  = dayMatch  ? dayMatch[1]  : '?';
        const aRef = timeMatch ? timeMatch[1] : '?';
        const type = isRoom ? 'room' : 'event';
        console.log(`    col ${ci + 6} (${['F','G','H','I'][ci]}): day+${day} ${type} $A${aRef}`);
      });
    });
  }

  console.log('\n✅  Weekly Schedule fixed:');
  console.log('    Schedule B6:O29   — all 24×14 cells rebuilt with correct day offsets and room/event types');
  console.log('    Conditional CF    — deleted 13 color-less rules; added 16 rules with proper hex colors');
  console.log('    Categories covered:', Object.keys(CATEGORY_COLORS).join(', '));
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
