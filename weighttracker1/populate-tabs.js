'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const SPREADSHEET_ID = '1YSrKj6U3Zb_al3g7j9e-80qFRGpQVwD4JLe4gh1Ah2k';

// Sheet IDs from metadata
const SID = {
  dashboard:     0,
  eventsLog:     1,
  calendar:      2,
  taskTracker:   3,
  budget:        4,
  guestList:     5,
  vendor:        6,
  runOfShow:     7,
  analytics:     8,
  notes:         9,
  reference:     10,
};

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return oAuth2Client;
}
async function getSheets() { return google.sheets({ version: 'v4', auth: await getAuth() }); }

(async () => {
  const sheets = await getSheets();

  // ── BATCH 1: Run-of-Show E/F TIME format fix ─────────────────────────────
  console.log('Batch 1: Fix Run-of-Show E/F time format...');
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{
        repeatCell: {
          range: {
            sheetId: SID.runOfShow,
            startRowIndex: 4,  // row 5 (0-indexed)
            endRowIndex: 45,   // rows 5-45
            startColumnIndex: 4, // col E
            endColumnIndex: 6,   // col F
          },
          cell: {
            userEnteredFormat: {
              numberFormat: { type: 'TIME', pattern: 'h:mm AM/PM' },
            },
          },
          fields: 'userEnteredFormat.numberFormat',
        },
      }],
    },
  });
  console.log('  ✓ Run-of-Show E5:F45 formatted as h:mm AM/PM');

  // ── BATCH 2: Value updates ────────────────────────────────────────────────
  console.log('\nBatch 2: Populating sample data...');
  const valueUpdates = [];

  // ── Task Tracker (rows 5–14) ──────────────────────────────────────────────
  // Cols: A=#  B=Task  C=Category  D=Event  E=Priority  F=Status  G=Due Date  H=Assigned To  I=Notes  J=Done
  valueUpdates.push({
    range: "'✅ Task Tracker'!A5:J14",
    values: [
      [1, 'Book venue & confirm contract',       'Venue',        'Tech Summit 2026',          'High',   'Done',        '1/15/2026', 'Sarah Johnson', 'Confirmed Grand Convention Centre',  true],
      [2, 'Send speaker invitations',            'Content',      'Tech Summit 2026',          'High',   'Done',        '2/1/2026',  'Sarah Johnson', 'All 4 keynote speakers confirmed',   true],
      [3, 'Finalise catering order',             'Catering',     'Tech Summit 2026',          'High',   'In Progress', '9/10/2026', 'Mark Davis',    'Menu tasting on 2 Sep',              false],
      [4, 'Design event program & signage',      'Marketing',    'Tech Summit 2026',          'Medium', 'Done',        '8/15/2026', 'Lisa Chen',     'Print run: 500 copies',              true],
      [5, 'Set up online registration page',     'Registration', 'Charity Gala 2026',         'High',   'Done',        '3/1/2026',  'Priya Sharma',  'Eventbrite live',                    true],
      [6, 'Hire AV technician & test equipment', 'Technical',    'Tech Summit 2026',          'High',   'In Progress', '9/12/2026', 'Mark Davis',    'Rehearsal scheduled Sep 14',         false],
      [7, 'Create stage layout & floor plan',    'Logistics',    'Annual Awards Night',       'Medium', 'Not Started', '10/1/2026', 'Oliver Chen',   '',                                   false],
      [8, 'Confirm MC & award presenters',       'Content',      'Annual Awards Night',       'High',   'Not Started', '10/15/2026','Sarah Johnson', '',                                   false],
      [9, 'Order promotional gift bags',         'Marketing',    'Charity Gala 2026',         'Low',    'In Progress', '10/20/2026','Lisa Chen',     '200 bags — supplier quote pending',  false],
      [10,'Arrange shuttle & parking signage',   'Logistics',    'Tech Summit 2026',          'Low',    'In Progress', '9/8/2026',  'Mark Davis',    'Parking for 120 vehicles',           false],
    ],
  });
  console.log('  ✓ Task Tracker: 10 tasks');

  // ── Event Budget — restore expense log headers (row 19, cols A-G cleared earlier) ──
  valueUpdates.push({
    range: "'💰 Event Budget'!A19:G19",
    values: [['#', 'Event Name', 'Category', 'Vendor / Description', 'Date', 'Amount', 'Notes']],
  });

  // ── Event Budget — allocation rows 4-15 ──────────────────────────────────
  // Cols: A=#  B=Event Name  C=Category  D=Vendor/Description  E=Budgeted  F=Actual  G=Variance(formula)  H=Paid?(checkbox)
  valueUpdates.push({
    range: "'💰 Event Budget'!A4:H15",
    values: [
      [1,  'Tech Summit 2026',          'Venue',         'Grand Convention Centre',  8000, 8000, '=E4-F4',   true],
      [2,  'Tech Summit 2026',          'Catering',      'Premier Catering Co.',     3500, 3200, '=E5-F5',   true],
      [3,  'Tech Summit 2026',          'AV & Tech',     'TechPro AV Services',      2500, 2500, '=E6-F6',   true],
      [4,  'Tech Summit 2026',          'Marketing',     'Creative Agency',          1500, 1200, '=E7-F7',   true],
      [5,  'Charity Gala 2026',         'Venue',         'Riverside Ballroom',       5000, 5000, '=E8-F8',   true],
      [6,  'Charity Gala 2026',         'Entertainment', 'Jazz Quartet',             1800,    0, '=E9-F9',   false],
      [7,  'Product Launch — Nova X',   'Venue',         'Rooftop Events',           2200,    0, '=E10-F10', false],
      [8,  'Product Launch — Nova X',   'Catering',      'Fresh Eats Co.',            900,    0, '=E11-F11', false],
      [9,  'Annual Awards Night',       'Venue',         'Hilton Grand Ballroom',    6500,    0, '=E12-F12', false],
      [10, 'Annual Awards Night',       'AV & Tech',     'TechPro AV Services',      1200,    0, '=E13-F13', false],
      [11, 'Marketing Workshop',        'Venue',         'Co-Work Hub CBD',           400,  400, '=E14-F14', true],
      [12, 'Marketing Workshop',        'Catering',      'Café Solutions',            250,  250, '=E15-F15', true],
    ],
  });

  // ── Event Budget — expense log rows 20-25 ────────────────────────────────
  // Cols: A=#  B=Event Name  C=Category  D=Vendor/Description  E=Date  F=Amount  G=Notes  H=Paid?
  valueUpdates.push({
    range: "'💰 Event Budget'!A20:H25",
    values: [
      [1, 'Tech Summit 2026',    'Venue',     'Grand Convention Centre',  '8/1/2026',   4000, 'Deposit payment',     true],
      [2, 'Tech Summit 2026',    'Venue',     'Grand Convention Centre',  '9/15/2026',  4000, 'Balance on event day',true],
      [3, 'Tech Summit 2026',    'Catering',  'Premier Catering Co.',     '9/1/2026',   1750, 'Deposit',             true],
      [4, 'Tech Summit 2026',    'Catering',  'Premier Catering Co.',     '9/15/2026',  1450, 'Final invoice',       true],
      [5, 'Charity Gala 2026',   'Venue',     'Riverside Ballroom',       '10/1/2026',  2500, 'Deposit',             true],
      [6, 'Marketing Workshop',  'Venue',     'Co-Work Hub CBD',          '7/12/2026',   400, 'Full payment',        true],
    ],
  });
  console.log('  ✓ Event Budget: 12 allocation rows + 6 expense entries');

  // ── Guest List & RSVP (rows 5-14) ────────────────────────────────────────
  // Cols: A=# B=First C=Last D=Event E=Company F=RSVP G=TicketType H=Table# I=Email J=Phone K=DietaryReq L=DietNotes M=Seated N=BadgePrinted O=CheckedIn P=PlusOne Q=Notes
  valueUpdates.push({
    range: "'👥 Guest List & RSVP'!A5:Q14",
    values: [
      [1,  'Emily',   'Clarke',    'Tech Summit 2026',     'Nexus Corp',      'Confirmed', 'General',   3,  'emily.clarke@nexus.com',      '(555) 201-3344', 'None',       '',              true,  true,  true,  false, ''],
      [2,  'James',   'Hartley',   'Tech Summit 2026',     'BlueStar Ltd',    'Confirmed', 'VIP',       1,  'j.hartley@bluestar.com',      '(555) 482-9012', 'Vegetarian', '',              true,  true,  true,  true,  'Keynote speaker'],
      [3,  'Priya',   'Sharma',    'Tech Summit 2026',     'TechVentures',    'Confirmed', 'General',   4,  'priya.s@techventures.com',    '(555) 319-7721', 'None',       '',              true,  true,  true,  false, ''],
      [4,  'Michael', 'Torres',    'Tech Summit 2026',     'DataWave',        'Pending',   'General',   5,  'm.torres@datawave.io',        '(555) 678-4455', 'Gluten-Free','Severe allergy',false, false, false, false, 'Chase RSVP'],
      [5,  'Sophie',  'Williams',  'Charity Gala 2026',    'GreenLife',       'Confirmed', 'Table',     2,  'sophie.w@greenlife.org',      '(555) 901-2234', 'None',       '',              false, false, false, true,  ''],
      [6,  'Oliver',  'Chen',      'Charity Gala 2026',    'NextStep',        'Confirmed', 'VIP Table', 1,  'o.chen@nextstep.com',         '(555) 554-7788', 'Vegan',      '',              false, false, false, false, 'Table sponsor'],
      [7,  'Amanda',  'Kowalski',  'Charity Gala 2026',    'BlueStar Ltd',    'Declined',  'General',   '',  'a.kowalski@bluestar.com',    '(555) 334-5566', 'None',       '',              false, false, false, false, 'Conflict — rescheduled'],
      [8,  'David',   'Park',      'Annual Awards Night',  'Apex Group',      'Confirmed', 'General',   6,  'david.park@apexgroup.com',    '(555) 100-4488', 'None',       '',              false, false, false, false, ''],
      [9,  'Rachel',  'Ford',      'Annual Awards Night',  'Nexus Corp',      'Pending',   'General',   7,  'r.ford@nexus.com',            '(555) 220-9988', 'Halal',      '',              false, false, false, false, ''],
      [10, 'Nathan',  'Singh',     'Marketing Workshop',   'Independent',     'Confirmed', 'Standard',  1,  'n.singh@email.com',           '(555) 887-6633', 'None',       '',              false, false, false, false, ''],
    ],
  });

  // Update RSVP summary event name labels (col R, rows 3-10) to 2026 events
  valueUpdates.push({
    range: "'👥 Guest List & RSVP'!R3:R10",
    values: [
      ['Tech Summit 2026'],
      ['Charity Gala 2026'],
      ['Product Launch — Nova X'],
      ['Annual Awards Night'],
      ['Marketing Workshop'],
      ['Team Building Day'],
      ['Virtual Workshop — AI Trends'],
      ['Networking Drinks — Q4'],
    ],
  });
  console.log('  ✓ Guest List: 10 guests + RSVP summary updated to 2026 events');

  // ── Vendor & Supplier Tracker (rows 5-14) ────────────────────────────────
  // Cols: A=# B=Vendor C=Category D=Event E=Contact F=Phone G=Email H=Contract($) I=DepositPaid J=Deposit?(checkbox) K=BalanceDue L=PaymentStatus
  valueUpdates.push({
    range: "'🏪 Vendor & Supplier Tracker'!A5:L14",
    values: [
      [1,  'Grand Convention Centre',  'Venue',         'Tech Summit 2026',         'Mark Davis',      '(555) 201-3344', 'venue@gcc.com',         8000, 4000, true,  0,    'Paid in Full'],
      [2,  'Premier Catering Co.',     'Catering',      'Tech Summit 2026',         'Ana Ruiz',        '(555) 482-9012', 'ana@premiercatering.com',3500, 1750, true,  1750, 'Paid in Full'],
      [3,  'TechPro AV Services',      'AV & Tech',     'Tech Summit 2026',         'Gary Lim',        '(555) 319-7721', 'gary@techproav.com',    2500, 1250, true,  0,    'Paid in Full'],
      [4,  'Creative Agency',          'Marketing',     'Tech Summit 2026',         'Clara Stone',     '(555) 678-4455', 'clara@creativeag.com',  1500,  750, true,  0,    'Paid in Full'],
      [5,  'Riverside Ballroom',       'Venue',         'Charity Gala 2026',        'Susan Lee',       '(555) 901-2234', 'events@riverside.com',  5000, 2500, true,  2500, 'Not Paid'],
      [6,  'Jazz Quartet',             'Entertainment', 'Charity Gala 2026',        'Marcus Bell',     '(555) 554-7788', 'marcus@jazzquartet.com',1800,    0, false, 1800, 'Not Paid'],
      [7,  'Rooftop Events',           'Venue',         'Product Launch — Nova X',  'Tina Marsh',      '(555) 334-5566', 'tina@rooftopevents.com',2200,    0, false, 2200, 'Not Paid'],
      [8,  'Hilton Grand Ballroom',    'Venue',         'Annual Awards Night',      'Patricia Wan',    '(555) 100-4488', 'events@hiltongrand.com',6500,    0, false, 6500, 'Not Paid'],
      [9,  'TechPro AV Services',      'AV & Tech',     'Annual Awards Night',      'Gary Lim',        '(555) 319-7721', 'gary@techproav.com',    1200,    0, false, 1200, 'Not Paid'],
      [10, 'Co-Work Hub CBD',          'Venue',         'Marketing Workshop',       'Ben Foster',      '(555) 220-9988', 'ben@coworkhub.com',      400,  400, true,  0,    'Paid in Full'],
    ],
  });
  console.log('  ✓ Vendor & Supplier Tracker: 10 vendors');

  // ── Event Analytics (rows 5-12) ───────────────────────────────────────────
  // Cols: A=# B=EventName C=Category D=Type E=Date F=Capacity G=Attendees H=Attendance%(formula) I=Revenue J=Expenses K=NetP/L(formula) L=Rating/5
  // Completed/past events get realistic attendance; future = 0 or estimate
  valueUpdates.push({
    range: "'📊 Event Analytics'!A5:L12",
    values: [
      [1, 'Tech Summit 2026',          'Conference',      'In-Person', '9/15/2026', 350, 312, '=IFERROR(G5/F5,0)', 42500,  28200, '=I5-J5',  4.7],
      [2, 'Charity Gala 2026',         'Fundraiser',      'In-Person', '11/20/2026',200,   0, '=IFERROR(G6/F6,0)',     0,      0, '=I6-J6',  0  ],
      [3, 'Product Launch — Nova X',   'Product Launch',  'Hybrid',    '10/10/2026',150,   0, '=IFERROR(G7/F7,0)',     0,      0, '=I7-J7',  0  ],
      [4, 'Team Building Day',         'Team Building',   'In-Person', '8/5/2026',   80,  76, '=IFERROR(G8/F8,0)',     0,   4800, '=I8-J8',  4.5],
      [5, 'Annual Awards Night',       'Corporate',       'Hybrid',    '12/5/2026', 250,   0, '=IFERROR(G9/F9,0)',     0,      0, '=I9-J9',  0  ],
      [6, 'Marketing Workshop',        'Workshop',        'In-Person', '7/12/2026',  40,  38, '=IFERROR(G10/F10,0)', 1520,   950, '=I10-J10',4.8],
      [7, 'Virtual Workshop — AI Trends','Product Launch','Hybrid',    '10/25/2026',500,   0, '=IFERROR(G11/F11,0)',   0,      0, '=I11-J11',0  ],
      [8, 'Networking Drinks — Q4',    'Social',          'In-Person', '11/8/2026',  60,   0, '=IFERROR(G12/F12,0)',   0,      0, '=I12-J12',0  ],
    ],
  });
  console.log('  ✓ Event Analytics: 8 events (3 completed with actuals)');

  // ── Execute all value updates ─────────────────────────────────────────────
  const VCHUNK = 10;
  for (let i = 0; i < valueUpdates.length; i += VCHUNK) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { valueInputOption: 'USER_ENTERED', data: valueUpdates.slice(i, i + VCHUNK) },
    });
  }

  console.log('\n✅  All done:');
  console.log('    Run-of-Show E/F      — formatted as h:mm AM/PM (time columns now display correctly)');
  console.log('    Task Tracker         — 10 tasks across Tech Summit 2026, Charity Gala, Awards Night');
  console.log('    Event Budget         — 12 allocation rows + 6 expense log entries + headers restored');
  console.log('    Guest List & RSVP    — 10 guests across 4 events + RSVP summary labels updated to 2026');
  console.log('    Vendor & Supplier    — 10 vendors across 5 events');
  console.log('    Event Analytics      — 8 events (Tech Summit, Team Building, Marketing Workshop with actuals)');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
