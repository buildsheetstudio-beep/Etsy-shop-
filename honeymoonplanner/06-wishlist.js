'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WIS = sheetMap['✨ Activity & Excursion Wishlist'];

// Columns: A=Activity, B=Category (Ref!G), C=City (lookup from Itinerary col B), D=Estimated Cost, E=Priority, F=Booking Status (Ref!D), G=Link/Notes, H=Notes

const HEADERS = ['Activity / Excursion', 'Category', 'City', 'Est. Cost ($)', 'Priority', 'Booking Status', 'Link', 'Notes'];

const ITN_RANGE = "'🗓️ Multi-City Day-by-Day Itinerary'";

const SAMPLE = [
  ['Eiffel Tower Sunset Visit',          'Sightseeing',         'Paris',     35,  'High',   'Confirmed',  'https://toureiffel.paris', 'Book skip-line tickets'],
  ['Louvre Museum',                       'Cultural & Arts',     'Paris',     22,  'High',   'Confirmed',  'https://louvre.fr',         'Arrive early'],
  ['Seine River Dinner Cruise',           'Dining & Food',       'Paris',    120,  'High',   'Researching','https://bateauxparisiens.com','Romantic option'],
  ['Versailles Palace & Gardens',         'Sightseeing',         'Paris',     29,  'Medium', 'Researching','https://chateauversailles.fr','Take RER train'],
  ['Champagne Tasting Experience',        'Dining & Food',       'Paris',     85,  'Medium', 'Pending',    'https://cavedulouvre.com',  'In-cave experience'],
  ['Montmartre Art Walk',                 'Cultural & Arts',     'Paris',      0,  'Low',    'Not Needed', '',                          'Self-guided, free'],
  ['Santorini Catamaran Cruise',          'Beach & Water',       'Santorini', 220, 'High',   'Confirmed',  'https://santorinicruises.com','Couples BBQ included'],
  ['Akrotiri Archaeological Ruins',       'Cultural & Arts',     'Santorini',  12, 'Medium', 'Researching','https://akrotiri.gr',        'Bring water'],
  ['Oia Sunset Wine Tasting',             'Dining & Food',       'Santorini',  45, 'High',   'Pending',    'https://santowines.gr',     'Book table'],
  ['Couples Spa at Katikies',             'Spa & Wellness',      'Santorini', 190, 'High',   'Confirmed',  'https://katikies.com/spa',   '90-min couples massage'],
  ['Traditional Greek Cooking Class',     'Cultural & Arts',     'Santorini',  95, 'Medium', 'Researching','',                          'Small group, 3 hrs'],
  ['Fira to Oia Clifftop Hike',           'Adventure & Outdoors','Santorini',   0, 'Medium', 'Not Needed', '',                          '10km, bring sunscreen'],
  ['Red Beach Visit',                     'Beach & Water',       'Santorini',   0, 'Low',    'Not Needed', '',                          'Unique volcanic beach'],
  ['Delos Island Archaeology Day Trip',   'Cultural & Arts',     'Mykonos',    35, 'High',   'Researching','https://delosisland.gr',    'Guided tour available'],
  ['Paradise Beach Sunbathing',           'Beach & Water',       'Mykonos',     0, 'Medium', 'Not Needed', '',                          'Sunbeds ~€12/day'],
  ['Mykonos Windmills & Old Town Walk',   'Sightseeing',         'Mykonos',     0, 'Low',    'Not Needed', '',                          'Free, great at sunset'],
  ['Scorpios Sunset Beach Club',          'Nightlife',           'Mykonos',    80, 'Medium', 'Pending',    'https://scorpios.com',       'Trendy, no reservation'],
  ['Nobu Mykonos Dinner',                 'Dining & Food',       'Mykonos',   280, 'High',   'Pending',    'https://noburestaurants.com','Reserve well in advance'],
  ['Jewellery Shopping — Mykonos Town',   'Shopping',            'Mykonos',   300, 'Low',    'Not Needed', '',                          'Budget for souvenirs'],
  ['Hot Air Balloon over Santorini',      'Adventure & Outdoors','Santorini', 360, 'Low',    'Researching','',                          'Sunrise or sunset options'],
];

(async () => {
  // City in col C will be user-entered sample data (no cross-sheet formula for simplicity since itinerary col B is repeated)
  const dataRows = SAMPLE.map(row => [row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7]]);

  await valuesBatchUpdate(id, [
    { range: "'✨ Activity & Excursion Wishlist'!A1:H1", values: [HEADERS] },
    { range: "'✨ Activity & Excursion Wishlist'!A2:H21", values: dataRows },
  ], 'wis-values');

  const reqs = [];

  // Base fill
  reqs.push({
    repeatCell: {
      range: gridRange(WIS, 0, 40, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmIvory), textFormat: { foregroundColor: hex(C.warmCharcoal) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Header
  reqs.push({
    repeatCell: {
      range: gridRange(WIS, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mutedSage),
          textFormat: { foregroundColor: hex(C.warmIvory), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Alternating rows
  for (let i = 1; i < 21; i += 2) {
    reqs.push({
      repeatCell: {
        range: gridRange(WIS, i, i+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
        fields: 'userEnteredFormat(backgroundColor)',
      },
    });
  }

  // Activity name — bold
  reqs.push({ repeatCell: { range: gridRange(WIS, 1, 21, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat(textFormat)' } });

  // City — teal
  reqs.push({ repeatCell: { range: gridRange(WIS, 1, 21, 2, 3), cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.slateTeal) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });

  // Cost — currency
  reqs.push({ repeatCell: { range: gridRange(WIS, 1, 21, 3, 4), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Priority & Status — center
  reqs.push({ repeatCell: { range: gridRange(WIS, 1, 21, 4, 6), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: WIS, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: WIS, dimension: 'ROWS', startIndex: 1, endIndex: 21 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,250],[1,140],[2,100],[3,110],[4,90],[5,120],[6,180],[7,200]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: WIS, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'wis-format');
  console.log('Activity Wishlist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
