'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CHK = sheetMap['✅ Research & Reservation Checklist'];

// Columns: A=Task, B=Category (Ref!B), C=Due Date, D=Assigned To (Ref!E), E=Done (checkbox), F=Notes

const HEADERS = ['Task', 'Category', 'Due Date', 'Assigned To', 'Done ✓', 'Notes'];

const SAMPLE = [
  ['Research and compare flight options',             'Flights & Transport',    '=DATE(2025,3,1)',  'Both',      false, 'Use Google Flights & Skyscanner'],
  ['Book outbound flights (x2)',                       'Flights & Transport',    '=DATE(2025,3,15)', 'Partner 1', true,  'Booked AF1234'],
  ['Book connecting flights (Paris → Santorini)',      'Flights & Transport',    '=DATE(2025,3,15)', 'Partner 1', true,  'Booked BA789'],
  ['Book return flights (Mykonos → home)',             'Flights & Transport',    '=DATE(2025,4,1)',  'Partner 1', false, 'Still shopping'],
  ['Book Paris hotel — Hotel Le Marais',               'Accommodation',          '=DATE(2025,3,20)', 'Partner 2', true,  'Booked, confirmation emailed'],
  ['Book Santorini hotel — Katikies Oia',              'Accommodation',          '=DATE(2025,3,20)', 'Partner 2', true,  'Cliffside suite reserved'],
  ['Book Mykonos boutique hotel',                      'Accommodation',          '=DATE(2025,4,5)',  'Partner 2', true,  'Confirmed'],
  ['Arrange Santorini port transfer',                  'Flights & Transport',    '=DATE(2025,5,1)',  'Partner 2', false, 'Hotel may include'],
  ['Research and book catamaran cruise',               'Activities & Excursions','=DATE(2025,4,1)',  'Both',      true,  'Booked through Santorini Cruises'],
  ['Book Le Jules Verne dinner (Eiffel Tower)',        'Activities & Excursions','=DATE(2025,3,1)',  'Partner 1', true,  'Must book 3 months ahead!'],
  ['Book Nobu Mykonos dinner',                         'Activities & Excursions','=DATE(2025,5,1)',  'Partner 1', false, 'Trying to reserve'],
  ['Book couples spa at Katikies',                     'Activities & Excursions','=DATE(2025,4,15)', 'Partner 2', true,  'Confirmed 90-min couples massage'],
  ['Purchase travel insurance (both)',                 'Health & Safety',        '=DATE(2025,4,1)',  'Both',      true,  'World Nomads purchased'],
  ['Check passport expiry dates (6+ months validity)', 'Documents & Visas',     '=DATE(2025,3,1)',  'Both',      true,  'Both valid through 2026+'],
  ['Check visa requirements for all destinations',     'Documents & Visas',     '=DATE(2025,3,1)',  'Both',      true,  'No visa required (EU/UK)'],
  ['Notify bank/credit cards of travel dates',         'Budget & Finance',       '=DATE(2025,6,1)',  'Both',      false, 'Avoid card blocks abroad'],
  ['Set up international phone plan',                  'General',                '=DATE(2025,5,15)', 'Both',      false, 'Check carrier options'],
  ['Download offline maps (Paris, Santorini, Mykonos)','General',               '=DATE(2025,6,10)', 'Both',      false, 'Google Maps offline'],
  ['Research currency exchange rates',                 'Budget & Finance',       '=DATE(2025,5,1)',  'Both',      false, 'Compare rates at home vs abroad'],
  ['Pack prescription medications + copies',           'Health & Safety',        '=DATE(2025,6,13)', 'Both',      false, 'Check airline liquid rules'],
  ['Create emergency contact list',                    'Health & Safety',        '=DATE(2025,6,10)', 'Partner 1', false, 'Embassy, insurance, hotel numbers'],
  ['Confirm all reservations 48h before',              'General',                '=DATE(2025,6,13)', 'Both',      false, 'Email confirmations'],
  ['Pack & weigh bags (check airline limits)',          'Packing',                '=DATE(2025,6,14)', 'Both',      false, 'Most: 23kg checked + 8kg cabin'],
  ['Download all e-tickets and QR codes offline',      'Documents & Visas',     '=DATE(2025,6,14)', 'Both',      false, 'Screenshots in photos app'],
  ['Arrange home care (post, plants, pets etc.)',      'General',                '=DATE(2025,6,10)', 'Both',      false, ''],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'✅ Research & Reservation Checklist'!A1:F1", values: [HEADERS] },
    { range: "'✅ Research & Reservation Checklist'!A2:F26", values: SAMPLE },
  ], 'chk-values');

  const reqs = [];

  // Base fill
  reqs.push({
    repeatCell: {
      range: gridRange(CHK, 0, 40, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmIvory), textFormat: { foregroundColor: hex(C.warmCharcoal) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Header
  reqs.push({
    repeatCell: {
      range: gridRange(CHK, 0, 1, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepTeal),
          textFormat: { foregroundColor: hex(C.warmIvory), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Alternating rows
  for (let i = 1; i < 26; i += 2) {
    reqs.push({
      repeatCell: {
        range: gridRange(CHK, i, i+1, 0, 6),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
        fields: 'userEnteredFormat(backgroundColor)',
      },
    });
  }

  // Task — bold
  reqs.push({ repeatCell: { range: gridRange(CHK, 1, 26, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat(textFormat)' } });

  // Category — teal
  reqs.push({ repeatCell: { range: gridRange(CHK, 1, 26, 1, 2), cell: { userEnteredFormat: { textFormat: { foregroundColor: hex(C.deepTeal) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });

  // Due Date — date format
  reqs.push({ repeatCell: { range: gridRange(CHK, 1, 26, 2, 3), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Assigned To — center
  reqs.push({ repeatCell: { range: gridRange(CHK, 1, 26, 3, 4), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Done checkbox — center
  reqs.push({ repeatCell: { range: gridRange(CHK, 1, 26, 4, 5), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: CHK, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: CHK, dimension: 'ROWS', startIndex: 1, endIndex: 26 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,310],[1,170],[2,110],[3,100],[4,80],[5,220]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: CHK, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'chk-format');
  console.log('Checklist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
