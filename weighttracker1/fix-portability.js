'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1CCo2yKNdKwNiuzK0E4CuNVzVCP27M72h0d4h81IMEaI';

// Sheet IDs
const SID = {
  referenceData:      0,
  eventDashboard:    11,
  eventSetup:         1,
  masterEventLog:     2,
  eventDetails:       3,
  guestsAttendees:    4,
  vendorsVenues:      5,
  eventTasks:         6,
  equipmentSupplies:  7,
  budgetExpenses:     8,
  weeklySchedule:     9,
  monthlyCalendar:   10,
};

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

// ── Named range definitions (all live on Reference Data sheet, sheetId=0) ────
// Coordinates derived from the API audit — rows/cols are 0-based startIndex / exclusive endIndex
const NAMED_RANGES = [
  { id: 1411749473, name: 'REF_CONTRACT_STATUS',   startRow: 1, endRow: 6,  startCol: 10, endCol: 11 },
  { id: 1020276651, name: 'REF_VENUE_TYPE',         startRow: 1, endRow: 11, startCol:  8, endCol:  9 },
  { id:  954013860, name: 'REF_PAYMENT_STATUS',     startRow: 1, endRow: 6,  startCol: 11, endCol: 12 },
  { id: 1192986111, name: 'REF_RECURRENCE',         startRow: 1, endRow: 7,  startCol: 18, endCol: 19 },
  { id:  196518949, name: 'REF_VENDOR_CATEGORY',    startRow: 1, endRow: 14, startCol:  9, endCol: 10 },
  { id:  522003825, name: 'REF_ATTENDANCE_STATUS',  startRow: 1, endRow: 7,  startCol: 12, endCol: 13 },
  { id: 1385834770, name: 'REF_REVENUE_TYPE',       startRow: 1, endRow: 8,  startCol: 17, endCol: 18 },
  { id: 1235061414, name: 'REF_PRIORITY',           startRow: 1, endRow: 5,  startCol:  6, endCol:  7 },
  { id:  837933930, name: 'REF_EXPENSE_CATEGORY',   startRow: 1, endRow: 14, startCol: 16, endCol: 17 },
  { id:  564255484, name: 'REF_EVENT_CATEGORY',     startRow: 1, endRow: 17, startCol:  3, endCol:  4 },
  { id: 1464294060, name: 'REF_EVENT_TYPE',         startRow: 1, endRow: 6,  startCol:  5, endCol:  6 },
  { id:  170450773, name: 'REF_MEAL_CHOICE',        startRow: 1, endRow: 9,  startCol: 14, endCol: 15 },
  { id:  474623245, name: 'REF_EVENT_STATUS',       startRow: 1, endRow: 7,  startCol:  4, endCol:  5 },
  { id: 1226341843, name: 'REF_EQUIPMENT_STATUS',   startRow: 1, endRow: 7,  startCol: 15, endCol: 16 },
  { id: 1678750679, name: 'REF_TASK_STATUS',        startRow: 1, endRow: 6,  startCol:  7, endCol:  8 },
  { id: 1674879102, name: 'REF_CATEGORY_COLORS',    startRow: 2, endRow: 18, startCol:  0, endCol:  2 },
  { id: 1221679462, name: 'REF_TICKET_TYPE',        startRow: 1, endRow: 8,  startCol: 13, endCol: 14 },
];

function noteCell(text) {
  return {
    userEnteredValue: { stringValue: text },
    userEnteredFormat: {
      backgroundColor: { red: 1.0, green: 0.973, blue: 0.831 }, // soft amber
      textFormat: {
        fontSize: 8,
        italic: true,
        foregroundColor: { red: 0.35, green: 0.28, blue: 0.0 },
      },
      verticalAlignment: 'MIDDLE',
    },
  };
}

(async () => {
  const sheets = await getSheets();

  // ══════════════════════════════════════════════════════════════════════════
  // FIX 1 — Event ID dropdowns: hardcoded list → live range from Master Event Log
  // Affects: Guests & Attendees col A (rows 6:500), Event Tasks col A (rows 6:500)
  // ══════════════════════════════════════════════════════════════════════════
  console.log('Fix 1: Updating event ID dropdowns to live range validation...');
  const liveRangeDv = {
    condition: {
      type: 'ONE_OF_RANGE',
      values: [{ userEnteredValue: "='Master Event Log'!$A$6:$A$500" }],
    },
    showCustomUi: true,
    strict: false,
  };

  const dvRequests = [SID.guestsAttendees, SID.eventTasks].map(sheetId => ({
    setDataValidation: {
      range: {
        sheetId,
        startRowIndex:    5,   // row 6 (0-based)
        endRowIndex:    500,
        startColumnIndex: 0,   // col A
        endColumnIndex:   1,
      },
      rule: liveRangeDv,
    },
  }));
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: dvRequests },
  });
  console.log('  ✓ Guests & Attendees and Event Tasks col A now pull from Master Event Log A6:A500');

  // ══════════════════════════════════════════════════════════════════════════
  // FIX 2 — Named ranges: add missing sheetId=0 (Reference Data)
  // All 17 named ranges had sheetId=undefined; the row/col coordinates were correct
  // ══════════════════════════════════════════════════════════════════════════
  console.log('\nFix 2: Repairing named ranges (adding sheetId for Reference Data)...');
  const namedRangeRequests = NAMED_RANGES.map(nr => ({
    updateNamedRange: {
      namedRange: {
        namedRangeId: String(nr.id),
        name: nr.name,
        range: {
          sheetId:          SID.referenceData,
          startRowIndex:    nr.startRow,
          endRowIndex:      nr.endRow,
          startColumnIndex: nr.startCol,
          endColumnIndex:   nr.endCol,
        },
      },
      fields: 'range',
    },
  }));

  // Send in batches of 8
  for (let i = 0; i < namedRangeRequests.length; i += 8) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: namedRangeRequests.slice(i, i + 8) },
    });
  }
  console.log(`  ✓ All ${NAMED_RANGES.length} named ranges now properly anchored to Reference Data (sheetId=0)`);

  // ══════════════════════════════════════════════════════════════════════════
  // FIX 3 — Tab-rename warning: add info banners on Weekly Schedule (row 3)
  //          and Monthly Calendar (row 3 col B onward, leaving A3 formula intact)
  // ══════════════════════════════════════════════════════════════════════════
  console.log('\nFix 3: Adding tab-rename warning banners...');
  const TAB_NOTE = '⚠️  Do not rename any sheet tabs — the color-coded highlighting on this view depends on the tab names staying exactly as they are.';

  // Weekly Schedule row 3: cols A:N (merge-spanning the width of the schedule)
  const wsNoteRequests = [
    {
      updateCells: {
        range: {
          sheetId:          SID.weeklySchedule,
          startRowIndex:    2,  // row 3 (0-based)
          endRowIndex:      3,
          startColumnIndex: 0,  // col A
          endColumnIndex:   15, // col O
        },
        rows: [{
          values: [
            noteCell(TAB_NOTE),
            ...Array(14).fill({ userEnteredFormat: { backgroundColor: { red: 1.0, green: 0.973, blue: 0.831 } } }),
          ],
        }],
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    },
    // Merge the banner across A3:O3
    {
      mergeCells: {
        range: {
          sheetId: SID.weeklySchedule,
          startRowIndex: 2, endRowIndex: 3,
          startColumnIndex: 0, endColumnIndex: 15,
        },
        mergeType: 'MERGE_ALL',
      },
    },
  ];

  // Monthly Calendar row 3: cols B:N (col A has the COUNTA formula — leave it)
  const mcNoteRequests = [
    {
      updateCells: {
        range: {
          sheetId:          SID.monthlyCalendar,
          startRowIndex:    2,  // row 3 (0-based)
          endRowIndex:      3,
          startColumnIndex: 1,  // col B
          endColumnIndex:   8,  // col H
        },
        rows: [{
          values: [
            noteCell(TAB_NOTE),
            ...Array(6).fill({ userEnteredFormat: { backgroundColor: { red: 1.0, green: 0.973, blue: 0.831 } } }),
          ],
        }],
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    },
    {
      mergeCells: {
        range: {
          sheetId: SID.monthlyCalendar,
          startRowIndex: 2, endRowIndex: 3,
          startColumnIndex: 1, endColumnIndex: 8,
        },
        mergeType: 'MERGE_ALL',
      },
    },
  ];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [...wsNoteRequests, ...mcNoteRequests] },
  });
  console.log('  ✓ Warning banner added to Weekly Schedule row 3 (A3:O3)');
  console.log('  ✓ Warning banner added to Monthly Calendar row 3 (B3:H3)');

  // ══════════════════════════════════════════════════════════════════════════
  // FIX 4 — Google Sheets Only notice on Event Dashboard (row 3)
  // ══════════════════════════════════════════════════════════════════════════
  console.log('\nFix 4: Adding Google Sheets compatibility notice to Event Dashboard...');
  const GS_NOTE = '📋  Google Sheets Only — This spreadsheet is optimized for Google Sheets. The Weekly Schedule and Monthly Calendar tabs use Google-specific functions (ARRAYFORMULA, LET) that are not supported in Microsoft Excel.';

  const dashNoteRequests = [
    {
      updateCells: {
        range: {
          sheetId:          SID.eventDashboard,
          startRowIndex:    2,  // row 3 (0-based)
          endRowIndex:      3,
          startColumnIndex: 0,  // col A
          endColumnIndex:   16, // col P
        },
        rows: [{
          values: [
            noteCell(GS_NOTE),
            ...Array(15).fill({ userEnteredFormat: { backgroundColor: { red: 1.0, green: 0.973, blue: 0.831 } } }),
          ],
        }],
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    },
    {
      mergeCells: {
        range: {
          sheetId: SID.eventDashboard,
          startRowIndex: 2, endRowIndex: 3,
          startColumnIndex: 0, endColumnIndex: 16,
        },
        mergeType: 'MERGE_ALL',
      },
    },
  ];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: dashNoteRequests },
  });
  console.log('  ✓ Google Sheets Only notice added to Event Dashboard row 3 (A3:P3)');

  // ══════════════════════════════════════════════════════════════════════════
  // VERIFY
  // ══════════════════════════════════════════════════════════════════════════
  console.log('\nVerifying...');

  // Check named ranges now have sheetId
  const verMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'namedRanges',
  });
  const broken = (verMeta.data.namedRanges || []).filter(nr => nr.range.sheetId === undefined || nr.range.sheetId === null);
  console.log(`  Named ranges still broken: ${broken.length} (expect 0)`);
  console.log(`  Named ranges repaired:     ${(verMeta.data.namedRanges || []).length - broken.length}`);

  // Check DV on Guests & Attendees col A row 6
  const dvCheck = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Guests & Attendees'!A6", "'Event Tasks'!A6"],
    includeGridData: true,
    fields: 'sheets(data(rowData(values(dataValidation))))',
  });
  dvCheck.data.sheets.forEach((sh, i) => {
    const tab = ['Guests & Attendees', 'Event Tasks'][i];
    const dv = sh.data[0].rowData[0].values[0].dataValidation;
    if (dv) {
      console.log(`  "${tab}" A6 DV type: ${dv.condition.type} — ${(dv.condition.values||[])[0]?.userEnteredValue?.slice(0,40) || 'n/a'}`);
    }
  });

  // Check banner cells
  const banners = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ["'Weekly Schedule'!A3", "'Monthly Calendar'!B3", "'Event Dashboard'!A3"],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  banners.data.valueRanges.forEach(vr => {
    const val = (vr.values || [['']])[0][0] || '';
    console.log(`  ${vr.range}: "${val.slice(0, 60)}..."`);
  });

  console.log('\n✅  All 4 portability fixes applied:');
  console.log('    Fix 1 — Event ID dropdowns now live-pull from Master Event Log (auto-grows with new events)');
  console.log('    Fix 2 — All 17 named ranges repaired with correct Reference Data sheet anchor');
  console.log('    Fix 3 — Tab-rename warning banner on Weekly Schedule and Monthly Calendar');
  console.log('    Fix 4 — Google Sheets Only notice on Event Dashboard');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
