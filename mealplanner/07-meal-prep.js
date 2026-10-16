'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const MPR = sheetMap['🍳 Meal Prep Schedule'];

// A=PrepDay B=Recipes C=Time(min) D=StorageMethod E=UseByDate(formula) F=PortionsMade G=Notes
const rows = [
  ['2026-07-13','Overnight Oats (x4)',15,'Fridge',null,4,'Breakfast prep for the week'],
  ['2026-07-13','Quinoa Power Bowl',30,'Fridge',null,4,'Pack in Mason jars'],
  ['2026-07-13','Chicken Tikka Masala',55,'Fridge',null,4,'Great reheated; store sauce separately'],
  ['2026-07-13','Veggie Stir Fry',25,'Fridge',null,4,'Keep sauce on the side'],
  ['2026-07-13','Almond Butter Cookies',27,'Pantry',null,24,'Bake extra for snacks'],
  ['2026-07-12','Salmon Teriyaki',30,'Fridge',null,2,'Marinate fish overnight first'],
  ['2026-07-12','Beef Bulgogi',45,'Freezer',null,4,'Portion into freezer bags'],
  ['2026-07-06','Ratatouille',90,'Fridge',null,6,'Made on Sunday, needs to be eaten'],
  ['2026-07-06','Spaghetti Bolognese',60,'Freezer',null,6,'Freeze flat in ziplock bags'],
  ['2026-07-01','Classic Pancakes',30,'Freezer',null,12,'Stack with parchment between each'],
];

const headers = ['Prep Day','Recipe(s)','Time (min)','Storage Method','Use By Date','Portions Made','Notes'];

(async () => {
  const dataRows = rows.map((r, i) => {
    const row = i + 3;
    return [
      r[0], r[1], r[2], r[3],
      `=A${row}+IF(D${row}="Freezer",90,IF(D${row}="Fridge",4,14))`,
      r[5], r[6],
    ];
  });

  const data = [
    { range: "'🍳 Meal Prep Schedule'!A1", values: [['🍳 Meal Prep Schedule']] },
    { range: "'🍳 Meal Prep Schedule'!A2", values: [headers] },
    { range: "'🍳 Meal Prep Schedule'!A3", values: dataRows },
  ];
  await valuesBatchUpdate(id, data, 'mpr-values');

  const reqs = [];

  // Merge A1:G1 (full-width, no col A freeze needed)
  reqs.push({ mergeCells: { range: gridRange(MPR, 0, 1, 0, 7), mergeType: 'MERGE_ALL' } });

  // Title row
  reqs.push({ repeatCell: {
    range: gridRange(MPR, 0, 1, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.goldenMustard), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Header row
  reqs.push({ repeatCell: {
    range: gridRange(MPR, 1, 2, 0, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Alt rows
  for (let r = 0; r < 10; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(MPR, 2 + r, 3 + r, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  // Formula col E: formulaBg
  reqs.push({ repeatCell: {
    range: gridRange(MPR, 2, 12, 4, 5),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER', numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment,numberFormat)',
  }});
  // Date format col A
  reqs.push({ repeatCell: {
    range: gridRange(MPR, 2, 12, 0, 1),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
  }});
  // Center numeric cols
  [2, 5].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(MPR, 2, 12, c, c + 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  const widths = [110, 200, 80, 120, 110, 100, 200];
  widths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: MPR, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: MPR, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 34 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'mpr-format');
  console.log('Meal Prep Schedule complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
