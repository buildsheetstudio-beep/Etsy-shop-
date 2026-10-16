'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const VLT = sheetMap['🔐 Travel Document & Confirmation Vault'];

// Columns: A=Type, B=Provider/Airline, C=Confirmation #, D=Date/Time, E=Status, F=Details, G=Document Link (formula), H=URL (raw, hidden-ish), I=Notes

const HEADERS = ['Document Type', 'Provider / Airline', 'Confirmation #', 'Date & Time', 'Status', 'Details', 'View Document', 'URL (paste here)', 'Notes'];

const SAMPLE = [
  ['Flight',            'Air France',        'AF1234-XYZ',  '=DATE(2025,6,15)', 'Confirmed', 'CDG → Paris depart 10:30am', '', 'https://airfrance.com/manage', 'Print boarding pass 24h before'],
  ['Flight',            'British Airways',   'BA789-ABC',   '=DATE(2025,6,18)', 'Confirmed', 'CDG → JTR depart 2:15pm',   '', 'https://ba.com/manage',        'Online check-in opens 24h prior'],
  ['Flight',            'Aegean Airlines',   'SK456-DEF',   '=DATE(2025,6,25)', 'Confirmed', 'MYK → LHR depart 4:00pm',   '', 'https://aegeanair.com',        ''],
  ['Hotel',             'Hotel Le Marais',   'LM-2025-9812','=DATE(2025,6,15)', 'Confirmed', '3 nights, Suite Romantique', '', 'https://lemarais.com',         'Late checkout requested'],
  ['Hotel',             'Katikies Hotel',    'KAT-7734',    '=DATE(2025,6,18)', 'Confirmed', '4 nights, Infinity Suite',   '', 'https://katikies.com',         'Transfer from port included'],
  ['Hotel',             'Mykonos Boutique',  'MYK-B3391',   '=DATE(2025,6,22)', 'Confirmed', '3 nights, Old Town Suite',   '', 'https://mykboutique.com',      ''],
  ['Tour/Activity',     'Catamaran Santorini','CAT-2025-55', '=DATE(2025,6,20)', 'Confirmed', 'Private couples cruise 10am', '', 'https://santorinicruises.com', 'Includes BBQ & snorkeling'],
  ['Tour/Activity',     'Delos Island Tours','DELOS-882',   '=DATE(2025,6,24)', 'Researching','Morning departure 9am',       '', '',                             'Must book in advance'],
  ['Restaurant',        'Le Jules Verne',    'JV-0615-A',   '=DATE(2025,6,17)', 'Confirmed', '8pm — Eiffel Tower 2nd Floor','', 'https://lejulesverne.com',    'Smart casual dress code'],
  ['Restaurant',        'Nobu Mykonos',      'NOBU-624-M',  '=DATE(2025,6,24)', 'Pending',   '7:30pm — Beachfront',         '', 'https://noburestaurants.com', 'Reservation not confirmed yet'],
  ['Travel Insurance',  'World Nomads',      'WN-2025-7712','=DATE(2025,6,1)',  'Confirmed', 'Couples plan — Jun 15-25',    '', 'https://worldnomads.com',     'Emergency: +1-800-555-WN24'],
  ['Visa',              'France Embassy',    'N/A',         '',                 'Not Needed','EU/UK citizens — no visa',    '', '',                             'Check requirements if non-EU'],
  ['Passport',          'Partner 1',         'P1-PP-9988',  '=DATE(2026,5,30)', 'Confirmed', 'Expires May 2026 — valid',    '', '',                             'Renew if travel > 6 months'],
  ['Passport',          'Partner 2',         'P2-PP-4412',  '=DATE(2027,2,14)', 'Confirmed', 'Expires Feb 2027 — valid',    '', '',                             ''],
  ['Car Rental',        'N/A',               'N/A',         '',                 'Not Needed','Using ferries & taxis',        '', '',                             ''],
];

(async () => {
  // Build data with HYPERLINK formula in col G
  const dataRows = SAMPLE.map((row, i) => {
    const r = i + 2;
    return [
      row[0], row[1], row[2], row[3], row[4], row[5],
      `=IFERROR(IF(H${r}="","",HYPERLINK(H${r},"View Document")),"")`,
      row[7], row[8],
    ];
  });

  await valuesBatchUpdate(id, [
    { range: "'🔐 Travel Document & Confirmation Vault'!A1:I1", values: [HEADERS] },
    { range: "'🔐 Travel Document & Confirmation Vault'!A2:I16", values: dataRows },
  ], 'vlt-values');

  const reqs = [];

  // Base fill
  reqs.push({
    repeatCell: {
      range: gridRange(VLT, 0, 30, 0, 9),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmIvory), textFormat: { foregroundColor: hex(C.warmCharcoal) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Header
  reqs.push({
    repeatCell: {
      range: gridRange(VLT, 0, 1, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.rustRed),
          textFormat: { foregroundColor: hex(C.warmIvory), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Alternating rows
  for (let i = 1; i < 16; i += 2) {
    reqs.push({
      repeatCell: {
        range: gridRange(VLT, i, i+1, 0, 9),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
        fields: 'userEnteredFormat(backgroundColor)',
      },
    });
  }

  // Type col A — bold rust
  reqs.push({ repeatCell: { range: gridRange(VLT, 1, 16, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.rustRed) } } }, fields: 'userEnteredFormat(textFormat)' } });

  // Date col D — date format
  reqs.push({ repeatCell: { range: gridRange(VLT, 1, 16, 3, 4), cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'MMM d, yyyy' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Status col E — center
  reqs.push({ repeatCell: { range: gridRange(VLT, 1, 16, 4, 5), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // View Document col G — teal, centered
  reqs.push({ repeatCell: { range: gridRange(VLT, 1, 16, 6, 7), cell: { userEnteredFormat: { textFormat: { foregroundColor: hex(C.slateTeal), bold: true }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });

  // URL col H — small gray text (data entry column)
  reqs.push({ repeatCell: { range: gridRange(VLT, 1, 16, 7, 8), cell: { userEnteredFormat: { textFormat: { foregroundColor: hex(C.mediumGray), fontSize: 9 }, backgroundColor: hex(C.inputBg) } }, fields: 'userEnteredFormat(textFormat,backgroundColor)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: VLT, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: VLT, dimension: 'ROWS', startIndex: 1, endIndex: 16 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,130],[1,160],[2,140],[3,110],[4,110],[5,240],[6,110],[7,200],[8,180]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: VLT, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'vlt-format');
  console.log('Document Vault complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
