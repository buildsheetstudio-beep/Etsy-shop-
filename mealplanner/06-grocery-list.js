'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const GRL = sheetMap['🛒 Grocery List'];

// A=Ingredient B=Category C=Qty D=HaveItAlready(formula) E=EstCost F=Store G=Purchased(checkbox)
// Items matching Food Inventory (col A): Spinach, Chicken Breast, Avocado, Garlic, Tomatoes, Broccoli, Eggs
const groceries = [
  ['Flour','Pantry','5 lbs',null,4.99,'Walmart',false],
  ['Butter','Dairy','1 lb',null,5.49,'Whole Foods',false],
  ['Sugar','Pantry','4 lbs',null,3.79,'Walmart',false],
  ['Parmesan','Dairy','8 oz',null,6.99,'Whole Foods',false],
  ['Caesar Dressing','Pantry','12 oz',null,4.49,'Trader Joes',false],
  ['Ground Beef','Meat','2 lbs',null,10.99,'Costco',false],
  ['Snap Peas','Produce','12 oz',null,3.49,'Whole Foods',false],
  ['Rice Noodles','Grains','8 oz',null,2.99,'Walmart',false],
  ['Peanut Butter','Pantry','16 oz',null,5.99,'Target',false],
  ['Honey','Pantry','12 oz',null,7.49,'Whole Foods',false],
  ['Mango','Produce','2',null,3.99,'Local Market',false],
  ['Pear','Produce','2',null,2.49,'Local Market',false],
  ['Mozzarella','Dairy','8 oz',null,5.99,'Whole Foods',false],
  ['Protein Powder','Pantry','2 lbs',null,34.99,'Costco',false],
  ['Cumin','Pantry','2 oz',null,2.29,'Walmart',false],
  ['Feta Cheese','Dairy','6 oz',null,4.99,'Trader Joes',false],
  ['Corn Tortillas','Grains','30 count',null,3.49,'Target',false],
  ['Corn Starch','Pantry','16 oz',null,2.19,'Walmart',false],
  // Items also in Food Inventory → D formula returns TRUE
  ['Spinach','Produce','2 bags',null,4.99,'Whole Foods',false],
  ['Chicken Breast','Meat','3 lbs',null,12.99,'Costco',false],
  ['Avocado','Produce','4',null,5.99,'Local Market',false],
  ['Garlic','Produce','2 heads',null,1.99,'Walmart',true],
  ['Tomatoes','Produce','6',null,4.49,'Local Market',true],
  ['Broccoli','Produce','2 heads',null,3.99,'Whole Foods',true],
  ['Eggs','Dairy','18 count',null,6.49,'Costco',true],
];

const headers = ['Ingredient','Category','Qty','Have It Already','Est. Cost','Store','Purchased'];

(async () => {
  const dataRows = groceries.map((g, i) => {
    const row = i + 3;
    return [
      g[0], g[1], g[2],
      `=IFERROR(IF(COUNTIF('🥗 Food Inventory'!$A:$A,A${row})>0,TRUE,FALSE),FALSE)`,
      g[4], g[5], g[6],
    ];
  });

  const data = [
    { range: "'🛒 Grocery List'!A1", values: [['🛒']] },
    { range: "'🛒 Grocery List'!B1", values: [['Grocery List']] },
    { range: "'🛒 Grocery List'!A2", values: [headers] },
    { range: "'🛒 Grocery List'!A3", values: dataRows },
  ];
  await valuesBatchUpdate(id, data, 'grl-values');

  const reqs = [];

  // Merge B1:G1
  reqs.push({ mergeCells: { range: gridRange(GRL, 0, 1, 1, 7), mergeType: 'MERGE_ALL' } });

  // A1 standalone
  reqs.push({ repeatCell: {
    range: gridRange(GRL, 0, 1, 0, 1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBasil), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // B1:G1 title
  reqs.push({ repeatCell: {
    range: gridRange(GRL, 0, 1, 1, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBasil), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Header row
  reqs.push({ repeatCell: {
    range: gridRange(GRL, 1, 2, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Alt rows
  for (let r = 0; r < 25; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(GRL, 2 + r, 3 + r, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  // Formula col D: formulaBg
  reqs.push({ repeatCell: {
    range: gridRange(GRL, 2, 27, 3, 4),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
  }});
  // Currency format for Est. Cost col E
  reqs.push({ repeatCell: {
    range: gridRange(GRL, 2, 27, 4, 5),
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
  }});
  // Center cols C, F, G
  [2, 5, 6].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(GRL, 2, 27, c, c + 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  const widths = [160, 100, 80, 120, 90, 120, 90];
  widths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: GRL, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GRL, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 34 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'grl-format');
  console.log('Grocery List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
