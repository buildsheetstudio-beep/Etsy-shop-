'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const WMP = sheetMap['📅 Weekly Meal Plan'];

// A=Day B=MealType C=FamilyMember D=Recipe E=Calories(formula) F=PrepAdvance(checkbox)
const rows = [
  ['Monday','Breakfast','Alex','Classic Pancakes','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D3,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Monday','Lunch','Alex','Caesar Salad','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D4,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Monday','Dinner','Alex','Chicken Tikka Masala','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D5,\'📖 Recipe Book\'!$B:$B,0)),0)',true],
  ['Tuesday','Breakfast','Sam','Overnight Oats','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D6,\'📖 Recipe Book\'!$B:$B,0)),0)',true],
  ['Tuesday','Lunch','Sam','Quinoa Power Bowl','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D7,\'📖 Recipe Book\'!$B:$B,0)),0)',true],
  ['Tuesday','Dinner','Sam','Salmon Teriyaki','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D8,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Wednesday','Breakfast','Jamie','Greek Yogurt Bowl','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D9,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Wednesday','Lunch','Jamie','French Onion Soup','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D10,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Wednesday','Dinner','Jamie','Spaghetti Bolognese','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D11,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Thursday','Breakfast','Alex','Avocado Toast','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D12,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Thursday','Lunch','Alex','Quinoa Power Bowl','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D13,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Thursday','Dinner','Alex','Tacos al Pastor','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D14,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Friday','Breakfast','Sam','Classic Pancakes','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D15,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Friday','Lunch','Sam','Caesar Salad','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D16,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Friday','Dinner','Sam','Beef Bulgogi','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D17,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Saturday','Breakfast','Jamie','Shakshuka','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D18,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Saturday','Lunch','Jamie','Quinoa Power Bowl','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D19,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Saturday','Dinner','Jamie','Ratatouille','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D20,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Sunday','Breakfast','Alex','Green Smoothie','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D21,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Sunday','Lunch','Alex','Thai Peanut Noodles','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D22,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
  ['Sunday','Dinner','Alex','Keto Cauliflower Pizza','=IFERROR(INDEX(\'📖 Recipe Book\'!$I:$I,MATCH(D23,\'📖 Recipe Book\'!$B:$B,0)),0)',false],
];

const headers = ['Day','Meal Type','Family Member','Recipe','Calories','Prepped in Advance'];

(async () => {
  const data = [
    { range: "'📅 Weekly Meal Plan'!A1", values: [['📅']] },
    { range: "'📅 Weekly Meal Plan'!B1", values: [['Weekly Meal Plan']] },
    { range: "'📅 Weekly Meal Plan'!A2", values: [headers] },
    { range: "'📅 Weekly Meal Plan'!A3", values: rows },
  ];
  await valuesBatchUpdate(id, data, 'wmp-values');

  const reqs = [];

  // Merge B1:F1
  reqs.push({ mergeCells: { range: gridRange(WMP, 0, 1, 1, 6), mergeType: 'MERGE_ALL' } });

  // A1 standalone emoji
  reqs.push({ repeatCell: {
    range: gridRange(WMP, 0, 1, 0, 1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBasil), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // B1:F1 title
  reqs.push({ repeatCell: {
    range: gridRange(WMP, 0, 1, 1, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepBasil), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Header row
  reqs.push({ repeatCell: {
    range: gridRange(WMP, 1, 2, 0, 6),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Data rows alt
  for (let r = 0; r < 21; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(WMP, 2 + r, 3 + r, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  // Formula col E: formulaBg
  reqs.push({ repeatCell: {
    range: gridRange(WMP, 2, 23, 4, 5),
    cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
  }});
  // Center cols B, E, F
  [1, 4, 5].forEach(c => reqs.push({ repeatCell: {
    range: gridRange(WMP, 2, 23, c, c + 1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }}));

  const widths = [100, 100, 120, 200, 90, 130];
  widths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: WMP, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: WMP, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 34 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'wmp-format');
  console.log('Weekly Meal Plan complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
