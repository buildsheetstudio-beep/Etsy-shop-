'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PKG = sheetMap['🧳 Packing List'];

// Columns: A=Item, B=Category (Ref!F), C=Partner (Ref!E), D=Packed (checkbox), E=Notes

const HEADERS = ['Item', 'Category', 'For', 'Packed ✓', 'Notes'];

const SAMPLE = [
  // Documents & Money
  ['Passport (original)',                       'Documents & Money',  'Both',      false, 'Check expiry — 6+ months required'],
  ['Printed flight & hotel confirmations',      'Documents & Money',  'Both',      false, 'Backup of digital tickets'],
  ['Travel insurance documents',                'Documents & Money',  'Both',      false, 'World Nomads policy number'],
  ['Credit cards + backup card',               'Documents & Money',  'Both',      false, 'Notify banks before travel'],
  ['Emergency cash (USD + EUR)',               'Documents & Money',  'Both',      false, '~$200 USD + €200 EUR'],
  ['Driving licence (if renting ATV)',         'Documents & Money',  'Both',      false, 'Santorini ATV hire'],
  // Clothing
  ['Lightweight dresses / sundresses (x5)',    'Clothing',           'Partner 1', false, 'Smart-casual for dinners'],
  ['Linen trousers (x2)',                      'Clothing',           'Partner 1', false, ''],
  ['Swimsuit (x2)',                            'Clothing',           'Partner 1', false, 'Quick-dry'],
  ['Beach cover-up (x2)',                      'Clothing',           'Partner 1', false, ''],
  ['Smart casual outfit for fine dining',      'Clothing',           'Partner 1', false, 'Eiffel Tower dinner'],
  ['Comfortable walking sandals',              'Clothing',           'Partner 1', false, 'Fira-Oia hike ready'],
  ['Heeled sandals for evenings',              'Clothing',           'Partner 1', false, ''],
  ['Lightweight shirts (x4)',                  'Clothing',           'Partner 2', false, 'Linen preferred for heat'],
  ['Smart chinos or dress trousers',           'Clothing',           'Partner 2', false, 'Nobu & Jules Verne dinners'],
  ['Swim shorts (x2)',                         'Clothing',           'Partner 2', false, ''],
  ['Casual shorts (x3)',                       'Clothing',           'Partner 2', false, ''],
  ['Smart shoes for evenings',                 'Clothing',           'Partner 2', false, ''],
  ['Comfortable walking shoes',                'Clothing',           'Partner 2', false, 'Louvre + Versailles = lots of walking'],
  ['Light cardigan / jumper (both)',           'Clothing',           'Both',      false, 'Evenings can be cool'],
  // Toiletries & Beauty
  ['Sunscreen SPF50+ (x2 — Greek islands!)',  'Toiletries & Beauty','Both',      false, 'Reef-safe for snorkelling'],
  ['After-sun lotion',                        'Toiletries & Beauty','Both',      false, ''],
  ['Insect repellent',                        'Toiletries & Beauty','Both',      false, 'Mediterranean midges'],
  ['Shampoo & conditioner (travel size)',     'Toiletries & Beauty','Both',      false, 'Or use hotel'],
  ['Moisturiser & lip balm',                 'Toiletries & Beauty','Both',      false, 'Dry heat'],
  ['Razor & shaving supplies',               'Toiletries & Beauty','Both',      false, ''],
  ['Perfume / cologne (small bottle)',        'Toiletries & Beauty','Both',      false, 'Decant into travel bottle'],
  // Electronics & Gadgets
  ['Smartphones + chargers',                  'Electronics & Gadgets','Both',    false, 'Roaming enabled'],
  ['Universal travel adapter (EU plug)',      'Electronics & Gadgets','Both',    false, 'Type C/E for France & Greece'],
  ['Camera + extra memory cards',             'Electronics & Gadgets','Both',    false, 'Capture memories!'],
  // Beach & Pool
  ['Beach bag',                               'Beach & Pool',       'Both',      false, ''],
  ['Snorkel set (lightweight)',               'Beach & Pool',       'Both',      false, 'For catamaran cruise'],
  ['Waterproof phone pouch',                  'Beach & Pool',       'Both',      false, 'Beach & boat days'],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'🧳 Packing List'!A1:E1", values: [HEADERS] },
    { range: "'🧳 Packing List'!A2:E34", values: SAMPLE },
  ], 'pkg-values');

  const reqs = [];

  // Base fill
  reqs.push({
    repeatCell: {
      range: gridRange(PKG, 0, 50, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warmIvory), textFormat: { foregroundColor: hex(C.warmCharcoal) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Header
  reqs.push({
    repeatCell: {
      range: gridRange(PKG, 0, 1, 0, 5),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.warmCharcoal),
          textFormat: { foregroundColor: hex(C.warmIvory), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Alternating rows
  for (let i = 1; i < 34; i += 2) {
    reqs.push({
      repeatCell: {
        range: gridRange(PKG, i, i+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
        fields: 'userEnteredFormat(backgroundColor)',
      },
    });
  }

  // Item col A — bold
  reqs.push({ repeatCell: { range: gridRange(PKG, 1, 34, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat(textFormat)' } });

  // Category col B — teal
  reqs.push({ repeatCell: { range: gridRange(PKG, 1, 34, 1, 2), cell: { userEnteredFormat: { textFormat: { foregroundColor: hex(C.slateTeal) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(textFormat,horizontalAlignment)' } });

  // For col C — center
  reqs.push({ repeatCell: { range: gridRange(PKG, 1, 34, 2, 3), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Packed col D — center
  reqs.push({ repeatCell: { range: gridRange(PKG, 1, 34, 3, 4), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: PKG, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: PKG, dimension: 'ROWS', startIndex: 1, endIndex: 34 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Column widths
  [[0,310],[1,160],[2,90],[3,80],[4,220]].forEach(([ci, w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: PKG, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'pkg-format');
  console.log('Packing List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
