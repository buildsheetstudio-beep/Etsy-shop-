'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BGT = sheetMap['💰 Budget & Expense Tracker'];

// Columns: A=Category, B=Description, C=Estimated ($), D=Actual ($), E=Diff, F=Currency, G=Payment Method, H=Notes

const HEADERS = ['Category', 'Description', 'Estimated ($)', 'Actual ($)', 'Diff ($)', 'Currency', 'Payment Method', 'Notes'];

const SAMPLE = [
  ['Flights',           'Paris Round-trip Flights (x2)',           2400, 2380, '', 'USD', 'Credit Card', 'Booked AF & BA'],
  ['Flights',           'Paris → Santorini Flight (x2)',            840,  820, '', 'USD', 'Credit Card', ''],
  ['Flights',           'Santorini → Mykonos Ferry',                110,  110, '', 'EUR', 'Credit Card', 'Fast ferry'],
  ['Flights',           'Mykonos → Home Flight (x2)',               760,    0, '', 'USD', 'Credit Card', 'Not yet booked'],
  ['Accommodation',     'Hotel Le Marais, Paris (3 nights)',       1050, 1050, '', 'EUR', 'Credit Card', 'Breakfast included'],
  ['Accommodation',     'Katikies Hotel, Santorini (4 nights)',    2400, 2280, '', 'EUR', 'Credit Card', 'Cliff suite'],
  ['Accommodation',     'Boutique Hotel, Mykonos (3 nights)',       840,  840, '', 'EUR', 'Credit Card', 'Old Town location'],
  ['Food & Dining',     'Paris — Restaurants & Cafés',              600,  570, '', 'EUR', 'Cash/Card',   ''],
  ['Food & Dining',     'Santorini — Caldera Dining',               450,  480, '', 'EUR', 'Cash/Card',   'Over budget'],
  ['Food & Dining',     'Mykonos — Restaurants',                    400,    0, '', 'EUR', 'Cash/Card',   'Budget estimate'],
  ['Activities & Tours','Eiffel Tower + Louvre Tickets',             95,   95, '', 'EUR', 'Credit Card', ''],
  ['Activities & Tours','Paris — Seine Cruise + Champagne',         103,  100, '', 'EUR', 'Cash',        ''],
  ['Activities & Tours','Versailles Day Trip',                       29,   29, '', 'EUR', 'Credit Card', ''],
  ['Activities & Tours','Santorini — Catamaran Cruise',             440,  440, '', 'EUR', 'Credit Card', 'Couples package'],
  ['Activities & Tours','Santorini — Wine Tasting',                  90,   90, '', 'EUR', 'Cash',        ''],
  ['Activities & Tours','Santorini — Cooking Class',                190,    0, '', 'EUR', 'Cash',        ''],
  ['Activities & Tours','Mykonos — Delos Island Trip',               70,   70, '', 'EUR', 'Cash',        ''],
  ['Health & Beauty',   'Couples Spa — Santorini',                  380,  380, '', 'EUR', 'Credit Card', 'Prebooked'],
  ['Transport',         'Paris Metro & Taxis',                      120,  115, '', 'EUR', 'Cash/Card',   ''],
  ['Transport',         'Santorini ATV Rental (2 days)',             80,   80, '', 'EUR', 'Cash',        ''],
  ['Shopping',          'Paris Boutiques & Souvenirs',              500,  620, '', 'EUR', 'Cash/Card',   'Over budget'],
  ['Shopping',          'Santorini & Mykonos Shopping',             400,    0, '', 'EUR', 'Cash/Card',   'Estimate'],
  ['Miscellaneous',     'Travel Insurance (x2)',                    320,  320, '', 'USD', 'Credit Card', 'World Nomads'],
  ['Miscellaneous',     'Currency Exchange / ATM Fees',              60,   45, '', 'USD', 'Various',     ''],
  ['Miscellaneous',     'Tips & Gratuities',                        200,  180, '', 'EUR', 'Cash',        ''],
];

(async () => {
  const dataRows = SAMPLE.map((row, i) => {
    const r = i + 2;
    return [
      row[0], row[1], row[2], row[3],
      `=IFERROR(C${r}-D${r},"")`,
      row[5], row[6], row[7],
    ];
  });

  await valuesBatchUpdate(id, [
    { range: "'💰 Budget & Expense Tracker'!A1:H1", values: [HEADERS] },
    { range: "'💰 Budget & Expense Tracker'!A2:H26", values: dataRows },
  ], 'bgt-values');

  const reqs = [];

  // Base fill
  reqs.push({
    repeatCell: {
      range: gridRange(BGT, 0, 50, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmIvory), textFormat: { foregroundColor: hex(C.warmCharcoal) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Header
  reqs.push({
    repeatCell: {
      range: gridRange(BGT, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.amber),
          textFormat: { foregroundColor: hex(C.warmCharcoal), bold: true, fontSize: 11 },
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
        range: gridRange(BGT, i, i+1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
        fields: 'userEnteredFormat(backgroundColor)',
      },
    });
  }

  // Category — bold amber
  reqs.push({ repeatCell: { range: gridRange(BGT, 1, 26, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true, foregroundColor: hex(C.amber) } } }, fields: 'userEnteredFormat(textFormat)' } });

  // Estimated, Actual, Diff columns — currency format
  [2, 3, 4].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(BGT, 1, 26, ci, ci+1), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' }, horizontalAlignment: 'RIGHT' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });
  });

  // Currency & Payment Method — center
  reqs.push({ repeatCell: { range: gridRange(BGT, 1, 26, 5, 7), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: BGT, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: BGT, dimension: 'ROWS', startIndex: 1, endIndex: 26 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,140],[1,240],[2,110],[3,110],[4,100],[5,90],[6,120],[7,180]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: BGT, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'bgt-format');
  console.log('Budget complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
