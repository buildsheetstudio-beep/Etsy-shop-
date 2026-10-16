'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const FIN = sheetMap['🥗 Food Inventory'];

// A=Item B=Category C=Qty D=PurchaseDate E=ExpirationDate F=DaysUntilExp(formula) G=Status(formula)
// Today = 2026-07-13; Use Soon = exp in 0-3 days; Expired = exp < today; Fresh = exp > 3 days
const items = [
  ['Spinach','Produce','1 bag','2026-07-05','2026-07-14'],
  ['Chicken Breast','Meat','2 lbs','2026-07-10','2026-07-15'],
  ['Greek Yogurt','Dairy','32 oz','2026-07-08','2026-07-16'],
  ['Eggs','Dairy','12 count','2026-07-06','2026-07-20'],
  ['Milk','Dairy','1 gallon','2026-07-10','2026-07-18'],
  ['Avocado','Produce','3','2026-07-11','2026-07-14'],
  ['Almond Milk','Dairy','1 carton','2026-07-05','2026-07-25'],
  ['Tomatoes','Produce','4','2026-07-08','2026-07-18'],
  ['Garlic','Produce','1 head','2026-07-01','2026-08-01'],
  ['Rice','Grains','5 lbs','2026-07-01','2027-07-01'],
  ['Olive Oil','Pantry','1 bottle','2026-06-01','2027-06-01'],
  ['Soy Sauce','Pantry','1 bottle','2026-05-01','2027-05-01'],
  ['Broccoli','Produce','1 head','2026-07-09','2026-07-15'],
  ['Salmon','Seafood','1 lb','2026-07-11','2026-07-14'],
  ['Pasta','Grains','1 lb','2026-07-01','2027-07-01'],
  ['Bell Peppers','Produce','3','2026-07-07','2026-07-12'],
  ['Bananas','Produce','5','2026-07-10','2026-07-14'],
  ['Canned Tomatoes','Pantry','2 cans','2026-01-01','2028-01-01'],
  ['Quinoa','Grains','2 lbs','2026-07-01','2027-01-01'],
  ['Lemon','Produce','4','2026-07-01','2026-07-07'],
];

const headers = ['Item','Category','Qty','Purchase Date','Expiration Date','Days Until Exp','Status'];

(async () => {
  // Write item data (cols A-E only; F and G are formulas)
  const dataRows = items.map((item, i) => {
    const row = i + 3; // rows 3-22
    return [
      item[0], item[1], item[2], item[3], item[4],
      `=E${row}-TODAY()`,
      `=IF(F${row}<0,"Expired",IF(F${row}<=3,"Use Soon","Fresh"))`,
    ];
  });

  const data = [
    { range: "'🥗 Food Inventory'!A1", values: [['🥗']] },
    { range: "'🥗 Food Inventory'!B1", values: [['Food Inventory']] },
    { range: "'🥗 Food Inventory'!A2", values: [headers] },
    { range: "'🥗 Food Inventory'!A3", values: dataRows },
  ];
  await valuesBatchUpdate(id, data, 'fin-values');

  const reqs = [];

  // Merge B1:G1
  reqs.push({ mergeCells: { range: gridRange(FIN, 0, 1, 1, 7), mergeType: 'MERGE_ALL' } });

  // A1 standalone
  reqs.push({ repeatCell: {
    range: gridRange(FIN, 0, 1, 0, 1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.mutedSage), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // B1:G1 title
  reqs.push({ repeatCell: {
    range: gridRange(FIN, 0, 1, 1, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.mutedSage), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Header row
  reqs.push({ repeatCell: {
    range: gridRange(FIN, 1, 2, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Alt rows
  for (let r = 0; r < 20; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(FIN, 2 + r, 3 + r, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  // Formula cols F & G
  reqs.push({ repeatCell: {
    range: gridRange(FIN, 2, 22, 5, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
  }});
  // Date format for D and E cols
  const dateFmt = { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } };
  [3, 4].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(FIN, 2, 22, c, c + 1),
    cell: { userEnteredFormat: { ...dateFmt, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
  }}));

  const widths = [160, 100, 80, 110, 120, 110, 90];
  widths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: FIN, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: FIN, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 34 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'fin-format');
  console.log('Food Inventory complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
