'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];

const mealTypes    = ['Breakfast','Lunch','Dinner','Snack','Dessert'];
const cuisines     = ['American','Italian','Mexican','Asian','Mediterranean','Indian','French','Greek','Thai','Japanese'];
const difficulties = ['Easy','Medium','Hard','Expert'];
const allergyTags  = ['Gluten-Free','Dairy-Free','Nut-Free','Egg-Free','Soy-Free','Vegan','Vegetarian','Keto'];
const groceryCats  = ['Produce','Dairy','Meat','Seafood','Grains','Pantry','Frozen','Bakery','Beverages','Deli'];
const stores       = ['Whole Foods',"Trader Joe's",'Costco','Target','Walmart','Local Market'];
const storage      = ['Fridge','Freezer','Pantry'];
const restrictions = ['Gluten-Free','Dairy-Free','Nut Allergy','Vegetarian','Vegan','Kosher','Halal','Diabetic'];

const headers = ['Meal Types','Cuisines','Difficulty','Allergy Tags','Grocery Categories','Stores','Storage Methods','Restriction Types'];
const lists   = [mealTypes, cuisines, difficulties, allergyTags, groceryCats, stores, storage, restrictions];
const maxLen  = Math.max(...lists.map(l => l.length));

const rows = [headers];
for (let i = 0; i < maxLen; i++) rows.push(lists.map(l => l[i] || ''));

(async () => {
  await valuesBatchUpdate(id, [{ range: "'📋 Reference Data'!A1", values: rows }], 'ref');

  const reqs = [];
  reqs.push({ repeatCell: {
    range: gridRange(REF, 0, 1, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gray), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(REF, 1, maxLen + 1, 0, 8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream), textFormat: { foregroundColor: hex(C.darkText) } } },
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  // Alternate rows
  for (let r = 1; r < maxLen + 1; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(REF, r, r + 1, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  const colWidths = [140,140,110,130,150,130,130,130];
  colWidths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'ref-format');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
