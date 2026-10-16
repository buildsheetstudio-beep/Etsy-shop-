'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WSH = sheetMap['💡 Gift Idea Wishlist'];

// A Recipient  B Idea  C Price Estimate ($)  D Priority  E Link  F Notes  G Selected (checkbox)

const headers = ['Recipient','Gift Idea / Brainstorm','Est. Price ($)','Priority','Link / Source','Notes','Selected?'];

// 20 rows — several recipients with 2-3 competing ideas to show brainstorming
const rows = [
  ['Sarah (wife)',     'Cashmere sweater set (cream)',       145, 'High',   'nordstrom.com',        'She circled this in October',        true],
  ['Sarah (wife)',     'Silk robe + matching slippers',       90, 'Medium', 'amazon.com',           'Good backup option',                 false],
  ['Sarah (wife)',     'Jewelry organizer + trays',           55, 'Low',    'amazon.com',           'Nice but impersonal',                false],
  ['Dad',              'Electric drill set (cordless)',        78, 'High',   'homedepot.com',        'His old one seized up',              true],
  ['Dad',              'Work-boot insoles 3-pack',            28, 'Medium', 'amazon.com',           'Practical stocking stuffer',         false],
  ['Mom',              'Cookbook (Italian) + cooking class',  65, 'High',   'amazon.com / local',   'Two-part gift — both sourced',       true],
  ['Mom',              'Herb garden kit (indoor)',             35, 'Medium', 'amazon.com',           'Backup if class is full',            false],
  ['Jake (son)',       'LEGO Technic Race set',               60, 'High',   'target.com',           'He circled it in the catalog',       true],
  ['Jake (son)',       'RC monster truck',                    50, 'High',   'target.com',           'Second choice — both are equal',     false],
  ['Emma (daughter)',  'Art supply mega kit',                 55, 'High',   'michaels.com',         'She asked for this specifically',    true],
  ['Emma (daughter)',  'Craft jewelry-making kit',            30, 'Medium', 'amazon.com',           'Good stocking stuffer addition',     false],
  ['Grandma Rose',     'Photo book + silver frame',           50, 'High',   'shutterfly.com',       'Family photos from this year',       true],
  ['Grandma Rose',     'Heating neck wrap (lavender)',        35, 'Medium', 'amazon.com',           'She mentioned her neck hurts',       false],
  ['Grandpa Joe',      'Fishing lure & tackle set',           40, 'High',   'basspro.com',          'He fishes every Sat',                true],
  ['Sister Lily',      'Spa day gift card (Spa Luxe)',        70, 'High',   'spaluxe.com',          'Her favourite treat',                true],
  ['Sister Lily',      'Cashmere socks set',                  25, 'Medium', 'nordstrom.com',        'Nice stocking add-on',               false],
  ['Aunt Carol',       'Wine & cheese basket (premium)',       62, 'Medium', 'harryanddavid.com',    'Upgraded from $55 tier',             true],
  ['Cousin Rachel',    'Yoga mat + foam blocks',              40, 'High',   'amazon.com',           'She started yoga last spring',       true],
  ['BFF Tina',         'Artisan candle set (cedarwood)',      32, 'Medium', 'etsy.com',             'From her favourite Etsy seller',     true],
  ['Boss Karen',       'Succulent + ceramic planter',         27, 'Medium', 'traderjoes.com',       'Safe & appropriate work gift',       true],
];

(async () => {
  await valuesBatchUpdate(id, [
    { range: "'💡 Gift Idea Wishlist'!A1", values: [headers] },
    { range: "'💡 Gift Idea Wishlist'!A2", values: rows },
  ], 'wishlist-values');

  const reqs = [];
  const n = rows.length;

  // Header
  reqs.push({
    repeatCell: {
      range: gridRange(WSH, 0, 1, 0, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  for (let r = 1; r <= n; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(WSH, r, r + 1, 0, 7),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 1 ? C.snowWhite : C.altRow),
            textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  reqs.push({ repeatCell: { range: gridRange(WSH, 1, n + 1, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat(textFormat)' } });
  reqs.push({ repeatCell: { range: gridRange(WSH, 1, n + 1, 2, 3), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(WSH, 1, n + 1, 3, 4), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(horizontalAlignment)' } });
  reqs.push({ repeatCell: { range: gridRange(WSH, 1, n + 1, 6, 7), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(horizontalAlignment,verticalAlignment)' } });

  [140, 210, 90, 80, 160, 190, 80].forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: WSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: WSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'wishlist-format');
  console.log('Gift Idea Wishlist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
