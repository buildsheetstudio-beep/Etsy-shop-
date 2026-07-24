'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// Update all 12 sample events to 2026 dates so TODAY()-based dashboard filter works.
// Dashboard filter: Start Date >= TODAY() AND <= TODAY()+90  (TODAY = 2026-07-24)
// Next-90-days window: 2026-07-24 → 2026-10-22
//
// Column layout in '📋 Events Log'!A3:Q14:
// A=# B=EventName C=Category D=Type E=StartDate F=EndDate G=SetupDate H=StrikeDate
// I=ConfGuests(formula) J=Venue K=VenueCost L=Status M=ActualSpend(formula) N=Revenue
// O=Priority P=CreatedBy Q=Notes
//
// Strategy:
//   Past (Completed): rows 5,8,10,12  → dates before 2026-07-24
//   Near-future (in next 90 days): rows 3,4,6,7,9,10 → 2026-07-26 to 2026-10-20
//   Far-future: rows 7,11,13,14 → 2026-11 to 2026-12

(async () => {
  const data = [];

  // Rewrite the full events data with 2026 dates
  // Rows 3-14 (1-indexed in sheet), A through Q (excluding formula cols I, M)
  // We push the whole row and let formulas remain untouched by only writing non-formula cols
  // Use individual range pushes per event to avoid overwriting formula cells in I and M

  const events = [
    // [row, #, name, cat, type, startDate, endDate, setupDate, strikeDate, venue, venueCost, status, revenue, priority, createdBy, notes]
    [3,  1, 'Tech Summit 2026',           'Conference',    'In-Person', '2026-09-15', '2026-09-15', '2026-09-14', '2026-09-16', 'Grand Convention Centre', 8500,  'Confirmed', 12500, 'High',   'Sarah Johnson', 'Annual tech conference'],
    [4,  2, 'Charity Gala 2026',          'Fundraiser',    'In-Person', '2026-10-18', '2026-10-18', '2026-10-18', '2026-10-19', 'Riverside Ballroom',      12000, 'Confirmed', 45000, 'High',   'Sarah Johnson', 'Black tie fundraiser'],
    [5,  3, 'Product Launch — Nova X',    'Product Launch','Hybrid',    '2026-05-05', '2026-05-05', '2026-05-05', '2026-05-05', 'Studio 12 + Online',      3200,  'Completed', 0,     'High',   'Mark Davis',    'New product reveal'],
    [6,  4, 'Team Building Day',          'Team Building', 'In-Person', '2026-08-15', '2026-08-15', '2026-08-15', '2026-08-15', 'Outdoor Adventure Park',  2800,  'Confirmed', 0,     'Medium', 'Lisa Chen',     'Q3 team activity'],
    [7,  5, 'Annual Awards Night',        'Corporate',     'In-Person', '2026-11-27', '2026-11-27', '2026-11-27', '2026-11-28', 'The Grand Hotel',         9500,  'Planning',  0,     'High',   'Sarah Johnson', 'Employee recognition'],
    [8,  6, 'Marketing Workshop',         'Workshop',      'In-Person', '2026-03-20', '2026-03-20', '2026-03-20', '2026-03-20', 'Training Centre B',       800,   'Completed', 2400,  'Medium', 'Mark Davis',    'Digital marketing training'],
    [9,  7, 'Virtual Workshop — AI Trends','Workshop',     'Virtual',   '2026-09-03', '2026-09-03', '',           '',           'Online / Zoom',           0,     'Confirmed', 3800,  'Medium', 'Lisa Chen',     'Online professional development'],
    [10, 8, 'Networking Drinks — Q4',     'Social',        'In-Person', '2026-10-09', '2026-10-09', '2026-10-09', '2026-10-09', 'Rooftop Bar & Lounge',   1800,  'Confirmed', 0,     'Low',    'Mark Davis',    'Quarterly networking mixer'],
    [11, 9, 'Holiday Party 2026',         'Holiday Party', 'In-Person', '2026-12-11', '2026-12-11', '2026-12-11', '2026-12-11', 'The Loft Event Space',    5500,  'Planning',  0,     'High',   'Sarah Johnson', 'End of year celebration'],
    [12, 10,'CEO Birthday Celebration',   'Birthday',      'In-Person', '2026-06-25', '2026-06-25', '2026-06-25', '2026-06-25', 'Private Dining Room',     2200,  'Completed', 0,     'Medium', 'Lisa Chen',     'Surprise 50th birthday'],
    [13, 11,'Investor Day 2026',          'Conference',    'Hybrid',    '2026-10-29', '2026-10-29', '2026-10-29', '2026-10-29', 'HQ Boardroom + Online',   1500,  'Planning',  0,     'High',   'Sarah Johnson', 'Annual investor briefing'],
    [14, 12,'6-Year Anniversary Dinner',  'Anniversary',   'In-Person', '2026-11-07', '2026-11-07', '2026-11-07', '2026-11-07', 'Harbour View Restaurant', 4800,  'Confirmed', 0,     'High',   'Mark Davis',    'Company 6th anniversary'],
  ];

  for (const [row, num, name, cat, type, start, end, setup, strike, venue, venueCost, status, revenue, priority, createdBy, notes] of events) {
    // Push A, B, C, D, E, F, G, H (skip I — formula)
    data.push({ range: `'📋 Events Log'!A${row}:H${row}`, values: [[num, name, cat, type, start, end, setup, strike]] });
    // Push J, K, L (skip M — formula), N, O, P, Q
    data.push({ range: `'📋 Events Log'!J${row}:Q${row}`, values: [[venue, venueCost, status, null, revenue, priority, createdBy, notes]] });
  }

  await valuesBatchUpdate(id, data, 'fix-event-dates');
  console.log('Event dates updated to 2026 — dashboard upcoming events will now populate.');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
