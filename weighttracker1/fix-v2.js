'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1YSrKj6U3Zb_al3g7j9e-80qFRGpQVwD4JLe4gh1Ah2k';

const SID = { dashboard: 0, eventsLog: 1, runOfShow: 7 };

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

// confirmed-guests formula for Events Log col I row r
const cgFml = r =>
  `=IFERROR(COUNTIFS('👥 Guest List & RSVP'!D$5:D$65,B${r},'👥 Guest List & RSVP'!F$5:F$65,"Confirmed"),0)`;

(async () => {
  const sheets = await getSheets();

  // ── BATCH 1: formatting + validation ─────────────────────────────────────
  console.log('Batch 1: Formatting & validation fixes...');
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        // Fix Dashboard D4 — remove currency format, use plain number
        {
          repeatCell: {
            range: { sheetId: SID.dashboard, startRowIndex: 3, endRowIndex: 4, startColumnIndex: 3, endColumnIndex: 4 },
            cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0' } } },
            fields: 'userEnteredFormat.numberFormat',
          },
        },
        // Add dropdown validation to Run-of-Show D2 — sourced from Events Log col B
        {
          setDataValidation: {
            range: { sheetId: SID.runOfShow, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 3, endColumnIndex: 4 },
            rule: {
              condition: {
                type: 'ONE_OF_RANGE',
                values: [{ userEnteredValue: "='📋 Events Log'!$B$3:$B$65" }],
              },
              showCustomUi: true,
              strict: false,
            },
          },
        },
      ],
    },
  });
  console.log('  ✓ Dashboard D4 format fixed; Run-of-Show D2 dropdown added');

  // ── BATCH 2: Value updates ────────────────────────────────────────────────
  console.log('\nBatch 2: Value updates...');
  const valueUpdates = [];

  // ── 1. EVENTS LOG: populate 8 events (rows 3-10) ─────────────────────────
  // Cols: A=# B=EventName C=Category D=Type E=StartDate F=EndDate G=SetupDate
  //       H=StrikeDate I=ConfirmedGuests(formula) J=Venue K=VenueCost L=Status
  //       M=ActualSpend N=Revenue O=Priority P=CreatedBy Q=Notes
  const events = [
    [1, 'Tech Summit 2026',            'Conference',     'In-Person', '9/15/2026',  '9/15/2026',  '9/14/2026',  '9/16/2026',  cgFml(3),  'Grand Convention Centre',  8000, 'Confirmed',  28200, 42500, 'High',   'Sarah Johnson', 'Annual technology summit — 350 attendees'],
    [2, 'Charity Gala 2026',           'Fundraiser',     'In-Person', '11/20/2026', '11/20/2026', '11/19/2026', '11/21/2026', cgFml(4),  'Riverside Ballroom',       5000, 'Confirmed',  0,     0,     'High',   'Sarah Johnson', 'Annual fundraiser gala — goal $80,000'],
    [3, 'Product Launch — Nova X',     'Product Launch', 'Hybrid',    '10/10/2026', '10/10/2026', '10/9/2026',  '10/10/2026', cgFml(5),  'Rooftop Events',           2200, 'Planning',   0,     0,     'High',   'Oliver Chen',   'New product launch — hybrid in-person + livestream'],
    [4, 'Team Building Day',           'Team Building',  'In-Person', '8/5/2026',   '8/5/2026',   '8/5/2026',   '8/5/2026',   cgFml(6),  'Company HQ',               0,    'Completed',  4800,  0,     'Medium', 'Mark Davis',    'Q3 team building activities'],
    [5, 'Annual Awards Night',         'Corporate',      'Hybrid',    '12/5/2026',  '12/5/2026',  '12/4/2026',  '12/6/2026',  cgFml(7),  'Hilton Grand Ballroom',    6500, 'Planning',   0,     0,     'High',   'Sarah Johnson', 'Annual staff awards ceremony'],
    [6, 'Marketing Workshop',          'Workshop',       'In-Person', '7/12/2026',  '7/12/2026',  '7/12/2026',  '7/12/2026',  cgFml(8),  'Co-Work Hub CBD',          400,  'Completed',  950,   1520,  'Medium', 'Lisa Chen',     'Q2 marketing strategy workshop'],
    [7, 'Virtual Workshop — AI Trends','Product Launch', 'Hybrid',    '10/25/2026', '10/25/2026', '10/25/2026', '10/25/2026', cgFml(9),  'Online',                   0,    'Planning',   0,     0,     'Medium', 'Priya Sharma',  'Virtual AI trends webinar — 500 cap'],
    [8, 'Networking Drinks — Q4',      'Social',         'In-Person', '11/8/2026',  '11/8/2026',  '11/8/2026',  '11/8/2026',  cgFml(10), 'Riverside Bar & Grill',    0,    'Planning',   0,     0,     'Low',    'Oliver Chen',   'Q4 informal networking evening'],
  ];
  valueUpdates.push({ range: "'📋 Events Log'!A3:Q10", values: events });
  console.log('  ✓ Events Log: 8 events queued');

  // ── 2. DASHBOARD K4: fix Total Budget formula ─────────────────────────────
  // Was: =SUM(E4:E345) — includes row 17 TOTALS formula + date serials from expense log col E
  // Fix: =SUM(E4:E15) — only allocation rows
  valueUpdates.push({
    range: "'🎉 Dashboard'!K4",
    values: [["=SUM('💰 Event Budget'!E4:E15)"]],
  });
  console.log('  ✓ Dashboard K4: Total Budget formula fixed');

  // ── 3. RUN-OF-SHOW: D2 event selector, D3 count, H2 date/venue, H3 run time
  // D2: set to first event (dropdown validation already added above)
  valueUpdates.push({
    range: "'⏱️ Run-of-Show'!D2",
    values: [['Tech Summit 2026']],
  });

  // D3: count segments matching selected event in D2
  valueUpdates.push({
    range: "'⏱️ Run-of-Show'!D3",
    values: [['=IFERROR(COUNTIF(D5:D45,D2),0)']],
  });

  // H2: dynamic date + venue looked up from Events Log
  const h2Formula =
    `="Date: "&TEXT(IFERROR(INDEX('📋 Events Log'!E$3:E$65,MATCH(D2,'📋 Events Log'!B$3:B$65,0)),0),"d mmmm yyyy")` +
    `&" | Venue: "&IFERROR(INDEX('📋 Events Log'!J$3:J$65,MATCH(D2,'📋 Events Log'!B$3:B$65,0)),"")`;
  valueUpdates.push({ range: "'⏱️ Run-of-Show'!H2", values: [[h2Formula]] });

  // H3: total run time for selected event only (was literal string, not formula)
  const h3Formula =
    `="Total Run Time: "&TEXT(IFERROR(SUMPRODUCT((D5:D45=D2)*(C5:C45<>"Setup")*(F5:F45-E5:E45)*24),0),"0.0")&" hours"`;
  valueUpdates.push({ range: "'⏱️ Run-of-Show'!H3", values: [[h3Formula]] });

  // Update D column in data rows 5-18 from "Tech Summit 2025" → "Tech Summit 2026"
  valueUpdates.push({
    range: "'⏱️ Run-of-Show'!D5:D18",
    values: Array(14).fill(['Tech Summit 2026']),
  });
  console.log('  ✓ Run-of-Show D2/D3/H2/H3 fixed; D5:D18 updated to 2026');

  // ── 4. Add Charity Gala 2026 segments in rows 19-26 ──────────────────────
  // So the user can select Charity Gala 2026 from the dropdown and see its count change
  valueUpdates.push({
    range: "'⏱️ Run-of-Show'!A19:J26",
    values: [
      [1, 'Venue Setup & Decoration',    'Setup',        'Charity Gala 2026', 0.625,       0.729166667, "=TEXT((F19-E19)*24,\"0.0\")&\"h\"", 'Main Ballroom',    'Mark Davis',    'Floral arrangements + centrepieces'],
      [2, 'Volunteer & Staff Briefing',  'Setup',        'Charity Gala 2026', 0.729166667, 0.770833333, "=TEXT((F20-E20)*24,\"0.0\")&\"h\"", 'Staff Room',       'Lisa Chen',     'Dress code: black tie'],
      [3, 'Cocktail Hour & Registration','Registration', 'Charity Gala 2026', 0.770833333, 0.833333333, "=TEXT((F21-E21)*24,\"0.0\")&\"h\"", 'Foyer',            'Priya Sharma',  'Champagne + canapes, silent auction opens'],
      [4, 'Welcome Address',             'Opening',      'Charity Gala 2026', 0.833333333, 0.854166667, "=TEXT((F22-E22)*24,\"0.0\")&\"h\"", 'Main Stage',       'Sarah Johnson', 'Introduce guest speaker'],
      [5, 'Three-Course Dinner',         'Meal',         'Charity Gala 2026', 0.854166667, 0.958333333, "=TEXT((F23-E23)*24,\"0.0\")&\"h\"", 'Ballroom',         'Mark Davis',    'Dietary requirements confirmed with kitchen'],
      [6, 'Live Entertainment — Jazz',   'Entertainment','Charity Gala 2026', 0.854166667, 0.979166667, "=TEXT((F24-E24)*24,\"0.0\")&\"h\"", 'Stage',            'Marcus Bell',   ''],
      [7, 'Fund Appeal & Auction Close', 'Fundraising',  'Charity Gala 2026', 0.979166667, 1.0,         "=TEXT((F25-E25)*24,\"0.0\")&\"h\"", 'Main Stage',       'Oliver Chen',   'Live auction + pledge cards'],
      [8, 'Venue Teardown',              'Teardown',     'Charity Gala 2026', 1.0,         1.083333333, "=TEXT((F26-E26)*24,\"0.0\")&\"h\"", 'All Areas',        'Mark Davis',    ''],
    ],
  });
  console.log('  ✓ Run-of-Show: 8 Charity Gala 2026 segments added (rows 19-26)');

  // Apply TIME format to new RoS rows E19:F26 (same fix as before)
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{
        repeatCell: {
          range: {
            sheetId: SID.runOfShow,
            startRowIndex: 18, endRowIndex: 27,
            startColumnIndex: 4, endColumnIndex: 6,
          },
          cell: { userEnteredFormat: { numberFormat: { type: 'TIME', pattern: 'h:mm AM/PM' } } },
          fields: 'userEnteredFormat.numberFormat',
        },
      }],
    },
  });
  console.log('  ✓ Run-of-Show E19:F26 TIME format applied');

  // ── Execute all value updates ─────────────────────────────────────────────
  const VCHUNK = 10;
  for (let i = 0; i < valueUpdates.length; i += VCHUNK) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: valueUpdates.slice(i, i + VCHUNK) },
    });
    console.log(`  Values chunk ${Math.floor(i / VCHUNK) + 1}/${Math.ceil(valueUpdates.length / VCHUNK)} ✓`);
  }

  // ── Verify key cells ──────────────────────────────────────────────────────
  console.log('\nVerifying...');
  const verify = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SPREADSHEET_ID,
    ranges: [
      "'🎉 Dashboard'!K4",
      "'🎉 Dashboard'!A4:E4",
      "'📋 Events Log'!B3:B10",
      "'📋 Events Log'!E3:E10",
      "'📋 Events Log'!L3:L10",
      "'⏱️ Run-of-Show'!D2",
      "'⏱️ Run-of-Show'!D3",
      "'⏱️ Run-of-Show'!H3",
    ],
    valueRenderOption: 'FORMATTED_VALUE',
  });
  for (const vr of verify.data.valueRanges) {
    console.log(`  ${vr.range}: ${JSON.stringify(vr.values)}`);
  }

  console.log('\n✅  All fixes applied:');
  console.log('    Events Log        — 8 events with dates, status, venue, Confirmed Guests formula');
  console.log('    Dashboard K4      — Total Budget = SUM(E4:E15) only (no double-count)');
  console.log('    Dashboard D4      — Planning count number format fixed (no more "$0")');
  console.log('    Run-of-Show D2    — Dropdown added, sourced from Events Log event names');
  console.log('    Run-of-Show D3    — Now COUNTIF(D5:D45,D2) — counts segments for selected event');
  console.log('    Run-of-Show H2    — Now dynamic date + venue from Events Log via INDEX/MATCH');
  console.log('    Run-of-Show H3    — Now actual formula (was literal string); filters by selected event');
  console.log('    Run-of-Show D5:D18 — Updated "Tech Summit 2025" → "Tech Summit 2026"');
  console.log('    Run-of-Show rows 19-26 — Charity Gala 2026 segments added');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
