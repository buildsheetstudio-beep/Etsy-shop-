'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const ITN = sheetMap['🗓️ Multi-City Day-by-Day Itinerary'];

// Columns: A=Date, B=City, C=Time of Day, D=Activity/Plan, E=Location/Address, F=Cost, G=Booked, H=Notes

const HEADERS = ['Date', 'City', 'Time of Day', 'Activity / Plan', 'Location / Address', 'Estimated Cost ($)', 'Booked ✓', 'Notes'];

// Sample data: Paris → Santorini → Mykonos
const SAMPLE = [
  // Paris
  ['=DATE(2025,6,15)', 'Paris', 'Morning',   'Arrive CDG — Check in to hotel',          'Hotel Le Marais, Paris',          250, false, 'Flight AF1234'],
  ['=DATE(2025,6,15)', 'Paris', 'Afternoon', 'Eiffel Tower visit',                       'Champ de Mars, Paris',              35, true,  'Book tickets online'],
  ['=DATE(2025,6,15)', 'Paris', 'Evening',   'Romantic dinner in Montmartre',            'Le Moulin de la Galette',          120, false, 'Reservation needed'],
  ['=DATE(2025,6,16)', 'Paris', 'Morning',   'Louvre Museum',                            'Rue de Rivoli, Paris',              22, true,  'Skip-the-line tickets'],
  ['=DATE(2025,6,16)', 'Paris', 'Afternoon', 'Seine River Cruise',                       'Port de la Bourdonnais',            18, false, '1-hr cruise'],
  ['=DATE(2025,6,16)', 'Paris', 'Evening',   'Champagne tasting experience',             'Cave du Louvre',                    85, false, ''],
  ['=DATE(2025,6,17)', 'Paris', 'Morning',   'Versailles Palace day trip',               'Palace of Versailles',              29, false, 'Train from Paris'],
  ['=DATE(2025,6,17)', 'Paris', 'Evening',   'Le Jules Verne dinner at Eiffel Tower',    'Eiffel Tower, 2nd Floor',          320, false, 'Reserve 3 months ahead'],
  ['=DATE(2025,6,18)', 'Paris', 'Morning',   'Sacré-Cœur & Montmartre walk',            'Montmartre, Paris',                  0, false, 'Free'],
  ['=DATE(2025,6,18)', 'Paris', 'Afternoon', 'Depart for Santorini (flight)',            'CDG → JTR (Santorini Airport)',    420, true,  'Flight BA789'],
  // Santorini
  ['=DATE(2025,6,18)', 'Santorini', 'Evening',   'Check in — Oia cliffside suite',         'Katikies Hotel, Oia',              400, true,  'Infinity pool view'],
  ['=DATE(2025,6,19)', 'Santorini', 'Morning',   'Sunrise in Oia',                         'Oia Castle viewpoint',               0, false, 'Set alarm for 5:30am'],
  ['=DATE(2025,6,19)', 'Santorini', 'Afternoon', 'Red Beach & Akrotiri ruins',              'Akrotiri, Santorini',               12, false, 'Rent ATV'],
  ['=DATE(2025,6,19)', 'Santorini', 'Evening',   'Caldera dinner — Amoudi Bay seafood',    'Amoudi Bay taverna',               180, false, 'Watch sunset from table'],
  ['=DATE(2025,6,20)', 'Santorini', 'Morning',   'Private catamaran sailing cruise',        'Old Port, Fira',                   220, false, 'Couple\'s package'],
  ['=DATE(2025,6,20)', 'Santorini', 'Afternoon', 'Wine tasting at Santo Wines',             'Santo Wines Winery, Pyrgos',        45, false, ''],
  ['=DATE(2025,6,20)', 'Santorini', 'Evening',   'Spa treatment at hotel',                  'Katikies Spa',                     190, false, 'Couples massage'],
  ['=DATE(2025,6,21)', 'Santorini', 'Morning',   'Fira to Oia hike',                        'Fira, Santorini',                    0, false, '10km — bring water'],
  ['=DATE(2025,6,21)', 'Santorini', 'Afternoon', 'Perissa Black Sand Beach',                'Perissa Beach, Santorini',           0, false, 'Sunbeds ~12/day'],
  ['=DATE(2025,6,21)', 'Santorini', 'Evening',   'Traditional Greek cooking class',         'Fira cooking school',               95, false, ''],
  ['=DATE(2025,6,22)', 'Santorini', 'Morning',   'Ferry to Mykonos',                        'Santorini Port → Mykonos',          55, false, 'Fast ferry 2.5 hrs'],
  ['=DATE(2025,6,22)', 'Santorini', 'Afternoon', 'Check in Mykonos — Explore Old Town',     'Mykonos Town (Hora)',               280, true,  'Boutique hotel'],
  // Mykonos
  ['=DATE(2025,6,22)', 'Mykonos', 'Evening',   'Dinner at Little Venice waterfront',      'Little Venice, Mykonos',           140, false, 'Sunset views'],
  ['=DATE(2025,6,23)', 'Mykonos', 'Morning',   'Paradise Beach morning',                  'Paradise Beach, Mykonos',            0, false, 'Sunbeds available'],
  ['=DATE(2025,6,23)', 'Afternoon', 'Afternoon', 'Windmills of Mykonos & shopping',       'Kato Mili Windmills',                0, false, 'Jewellery stores nearby'],
  ['=DATE(2025,6,23)', 'Mykonos', 'Evening',   'Sunset cocktails at Scorpios',            'Scorpios Beach Club',               80, false, 'No res required'],
  ['=DATE(2025,6,24)', 'Mykonos', 'Morning',   'Delos island day trip',                   'Delos Archaeological Site',         35, false, 'Ferry from Old Port'],
  ['=DATE(2025,6,24)', 'Mykonos', 'Evening',   'Final romantic dinner at Nobu Mykonos',   'Nobu Mykonos Restaurant',          280, false, 'Reserve in advance'],
  ['=DATE(2025,6,25)', 'Mykonos', 'Morning',   'Check out — beach morning before flight', 'Ornos Beach',                        0, false, ''],
  ['=DATE(2025,6,25)', 'Mykonos', 'Afternoon', 'Depart — MYK → home',                    'Mykonos International Airport',    380, true,  'Flight SK456'],
];

(async () => {
  // Headers
  await valuesBatchUpdate(id, [
    { range: "'🗓️ Multi-City Day-by-Day Itinerary'!A1:H1", values: [HEADERS] },
    { range: "'🗓️ Multi-City Day-by-Day Itinerary'!A2:H31", values: SAMPLE },
  ], 'itn-values');

  const reqs = [];

  // Base fill — warm ivory
  reqs.push({
    repeatCell: {
      range: gridRange(ITN, 0, 50, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmIvory), textFormat: { foregroundColor: hex(C.warmCharcoal) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Header row
  reqs.push({
    repeatCell: {
      range: gridRange(ITN, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.slateTeal),
          textFormat: { foregroundColor: hex(C.warmIvory), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Alternating rows
  for (let i = 1; i < 31; i += 2) {
    reqs.push({
      repeatCell: {
        range: gridRange(ITN, i, i+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
        fields: 'userEnteredFormat(backgroundColor)',
      },
    });
  }

  // Date column (A) — date format
  reqs.push({ repeatCell: { range: gridRange(ITN, 1, 31, 0, 1), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // City column (B) — bold teal
  reqs.push({ repeatCell: { range: gridRange(ITN, 1, 31, 1, 2), cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.slateTeal) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });

  // Time of Day column (C) — centered
  reqs.push({ repeatCell: { range: gridRange(ITN, 1, 31, 2, 3), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Cost column (F) — currency format
  reqs.push({ repeatCell: { range: gridRange(ITN, 1, 31, 5, 6), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Booked column (G) — centered
  reqs.push({ repeatCell: { range: gridRange(ITN, 1, 31, 6, 7), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Row height
  reqs.push({ updateDimensionProperties: { range: { sheetId: ITN, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: ITN, dimension: 'ROWS', startIndex: 1, endIndex: 31 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,110],[1,100],[2,100],[3,260],[4,220],[5,130],[6,80],[7,180]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: ITN, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Border around data
  reqs.push({
    updateBorders: {
      range: gridRange(ITN, 0, 31, 0, 8),
      innerHorizontal: { style: 'SOLID', color: hex(C.altRow) },
      innerVertical:   { style: 'SOLID', color: hex(C.altRow) },
      bottom: { style: 'SOLID_MEDIUM', color: hex(C.slateTeal) },
    },
  });

  await batchUpdate(id, reqs, 'itn-format');
  console.log('Itinerary complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
