'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GRT = sheetMap['🎁 Gift Recipient Tracker'];

// A Recipient  B Relationship  C Gift Idea  D Store/Link  E Budget  F Actual
// G Diff(=E-F) H Bought  I Wrapped  J Sent/Delivered  K Delivery Method  L Address  M Notes

const headers = ['Recipient','Relationship','Gift Idea','Store / Link','Budget ($)','Actual ($)','Diff ($)','Bought?','Wrapped?','Sent / Delivered?','Delivery Method','Address','Notes'];

// 25 rows spanning all 5 relationship types; mix of bought/wrapped/sent; some over-budget
const rows = [
  // Family (8)
  ['Sarah (wife)',       'Family', 'Cashmere sweater set',        'Nordstrom',      150,  142,  '',true, true, true,  'In Person','',       'She mentioned this in October'],
  ['Dad',                'Family', 'Electric drill set',          'Home Depot',      80,  74.99,'',true, true, true,  'In Person','',       ''],
  ['Mom',                'Family', 'Cookbook + cooking class',    'Amazon',          60,  67.50,'',true, false,false, 'Mail',     '',       'Over budget — class voucher added'],
  ['Jake (son)',         'Family', 'LEGO Technic Race set',       'Target',          65,  62.99,'',true, true, false, 'In Person','',       'Hiding in hall closet'],
  ['Emma (daughter)',    'Family', 'Art supply mega kit',         'Michaels',        55,  58.00,'',true, true, false, 'In Person','',       'Over budget by $3 — worth it'],
  ['Grandma Rose',       'Family', 'Photo book + silver frame',   'Shutterfly',      45,  49.99,'',true, false,false, 'Mail',     '4512 Maple Ln, Portland OR 97201','Ships Dec 15'],
  ['Grandpa Joe',        'Family', 'Fishing lure & tackle set',   'Bass Pro',        40,  38.50,'',true, true, true,  'In Person','',       ''],
  ['Brother Mike',       'Family', 'Wireless noise-cancel earbuds','Best Buy',        80,  79.99,'',true, false,false, 'Mail',     '550 Pine St, Gresham OR 97030', ''],
  // Extended Family (5)
  ['Sister Lily',        'Extended Family','Spa day gift card',   'Spa Luxe',        70,  70,   '',true, false,false, 'In Person','',       'Her favourite place'],
  ['Aunt Carol',         'Extended Family','Wine & cheese basket','Harry & David',   55,  62,   '',true, false,false, 'Mail',     '221 Birchwood Dr, Salem OR 97301','Over budget — upgraded tier'],
  ['Uncle Dave',         'Extended Family','Golf gloves + balls', 'Golf Galaxy',     45,  43,   '',true, true, true,  'In Person','',       ''],
  ['Cousin Rachel',      'Extended Family','Yoga mat + blocks',   'Amazon',          40,  37.49,'',true, false,false, 'Mail',     '89 Elm St Apt 3B, Seattle WA 98101',''],
  ['Cousin Tom',         'Extended Family','Video game (latest)', 'GameStop',        60,  59.99,'',true, false,false, 'Mail',     '89 Elm St Apt 3B, Seattle WA 98101','Same building as Rachel'],
  // Friends (4)
  ['BFF Tina',           'Friends','Artisan candle gift set',     'Etsy',            35,  32,   '',true, true, true,  'In Person','',       'From her fav Etsy shop'],
  ['Neighbor Sue',       'Friends','Holiday cookie tin',          'Homemade',        20,  18.50,'',true, false,false, 'In Person','',       'Baking Dec 22'],
  ['College friend Jamie','Friends','Leather card wallet',        'Amazon',          30,  28.99,'',true, true, true,  'Mail',     '15 Park Blvd, Eugene OR 97401','Shipped Nov 30'],
  ['Dog-sitter Pam',     'Friends','Gift basket (spa)',           'TJ Maxx',         30,  28,   '',true, true, true,  'In Person','',       ''],
  // Coworkers (4)
  ['Boss Karen',         'Coworkers','Succulent planter set',     'Trader Joe\'s',   30,  27,   '',true, true, true,  'In Person','',       'Left on her desk'],
  ['Team gift (office)', 'Coworkers','Coffee shop cards ×5',      'Starbucks',       50,  50,   '',true, false,false, 'In Person','',       'Giving at holiday party Dec 20'],
  ['Work friend Dana',   'Coworkers','Notebook + pen set',        'Papier',          25,  26.50,'',true, false,false, 'In Person','',       'Over budget slightly'],
  ['Office raffle prize','Coworkers','Candle set',                'TJ Maxx',         15,  14.25,'',true, true, true,  'In Person','',       ''],
  // Other (4)
  ['Mail carrier',       'Other',  'Visa gift card',              'Visa',            15,  15,   '',true, false,false, 'In Person','',       'Leave in mailbox Dec 23'],
  ['Hairdresser',        'Other',  'Cash tip + holiday card',     'N/A',             40,  40,   '',true, false,false, 'In Person','',       ''],
  ['Teacher Ms. Parks',  'Other',  'Classroom supply box',        'Staples',         25,  24.50,'',true, true, false, 'In Person','',       'Deliver with Emma'],
  ['Secret Santa draw',  'Other',  'Strategy board game',        'Amazon',           20,  22,   '',false,false,false, 'Mail',     '',       'Over budget — better game option'],
];

(async () => {
  const valueRows = rows.map((row, i) => {
    const r = i + 2;
    return [
      row[0], row[1], row[2], row[3],
      row[4], row[5],
      `=IFERROR(E${r}-F${r},"")`,
      row[7], row[8], row[9],
      row[10], row[11], row[12],
    ];
  });

  await valuesBatchUpdate(id, [
    { range: "'🎁 Gift Recipient Tracker'!A1", values: [headers] },
    { range: "'🎁 Gift Recipient Tracker'!A2", values: valueRows },
  ], 'grt-values');

  const reqs = [];
  const n = rows.length;

  // Header
  reqs.push({
    repeatCell: {
      range: gridRange(GRT, 0, 1, 0, 13),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepCranberry),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  // Alternating rows
  for (let r = 1; r <= n; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(GRT, r, r + 1, 0, 13),
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

  // Col A bold
  reqs.push({ repeatCell: { range: gridRange(GRT, 1, n + 1, 0, 1), cell: { userEnteredFormat: { textFormat: { bold: true } } }, fields: 'userEnteredFormat(textFormat)' } });

  // Col E, F — currency
  reqs.push({ repeatCell: { range: gridRange(GRT, 1, n + 1, 4, 6), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(numberFormat,horizontalAlignment)' } });

  // Col G — diff formula bg + currency
  reqs.push({
    repeatCell: {
      range: gridRange(GRT, 1, n + 1, 6, 7),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.formulaBg),
          textFormat: { italic: true },
          horizontalAlignment: 'CENTER',
          numberFormat: { type: 'NUMBER', pattern: '"$"#,##0.00;[Red]"-$"#,##0.00' },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,numberFormat)',
    },
  });

  // Cols H, I, J — center checkboxes
  reqs.push({ repeatCell: { range: gridRange(GRT, 1, n + 1, 7, 10), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(horizontalAlignment,verticalAlignment)' } });

  // Column widths
  [150, 120, 190, 130, 90, 90, 90, 70, 70, 90, 120, 180, 170].forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: GRT, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: GRT, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'grt-format');
  console.log('Gift Recipient Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
