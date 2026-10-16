'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

// All dropdown source lists
const LISTS = {
  A: { label: 'MEAL TYPES', values: ['Breakfast','Brunch','Lunch','Dinner','Snack','Appetizer','Side Dish','Dessert','Beverage','Sauce / Dressing','Other'] },
  B: { label: 'RECIPE CATEGORIES', values: ['Breakfast','Soups','Salads','Sandwiches','Pasta','Rice & Grain','Chicken','Beef','Pork','Seafood','Vegetarian','Vegan','Casseroles','Slow Cooker','One Pot','Sheet Pan','Air Fryer','Grilling','Baking','Bread','Desserts','Snacks','Drinks','Sauces & Condiments','Meal Prep','Freezer Meals','Family Favorites','Holiday','Other'] },
  C: { label: 'CUISINES', values: ['American','Italian','Mexican','Mediterranean','Greek','French','Spanish','Indian','Chinese','Japanese','Korean','Thai','Vietnamese','Middle Eastern','Caribbean','Latin American','African','Fusion','Other'] },
  D: { label: 'DIETARY TAGS', values: ['No Restriction','Vegetarian','Vegan','Pescatarian','Gluten-Free','Dairy-Free','Nut-Free','Egg-Free','Soy-Free','Halal','Kosher','Low Sodium','Low Sugar','High Protein','Other'] },
  E: { label: 'ALLERGEN TAGS', values: ['None','Peanuts','Tree Nuts','Dairy','Eggs','Wheat','Soy','Fish','Shellfish','Sesame','Other'] },
  F: { label: 'COOKING METHODS', values: ['Stovetop','Oven','Slow Cooker','Pressure Cooker','Air Fryer','Grill','Smoker','Microwave','No Cook','Blender','Instant Pot','Sous Vide','Other'] },
  G: { label: 'DIFFICULTY', values: ['Very Easy','Easy','Moderate','Advanced'] },
  H: { label: 'RECIPE STATUS', values: ['Draft','Testing','Approved','Favorite','Seasonal','Archived'] },
  I: { label: 'RATINGS', values: ['1','2','3','4','5'] },
  J: { label: 'UNITS', values: ['tsp','tbsp','cup','fl oz','oz','lb','g','kg','mL','L','piece','clove','slice','can','jar','bag','box','package','bunch','pinch','dash','serving','to taste','other'] },
};

// Also need ingredient categories and instruction types
const EXTRA_LISTS = {
  A: { label: 'INGREDIENT CATEGORIES', row: 50, values: ['Produce','Meat & Seafood','Dairy & Eggs','Grains & Pasta','Baking','Canned Goods','Frozen','Condiments & Sauces','Spices & Seasonings','Pantry','Beverages','Other'] },
  B: { label: 'INSTRUCTION TYPES', row: 50, values: ['Prep','Cook','Bake','Chill','Rest','Assemble','Garnish','Serve','Storage','Reheat','Other'] },
  C: { label: 'YES / NO', row: 50, values: ['Yes','No'] },
};

(async () => {
  const reqs = [];
  const vals = [];

  // Title row
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 14, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A1`, values: [['🍴 REFERENCE DATA — Dropdown Lists & Lookup Tables']] });

  // Write each list in column, starting row 3
  const cols = Object.keys(LISTS);
  cols.forEach((col, ci) => {
    const { label, values } = LISTS[col];
    // Header
    reqs.push({ repeatCell: {
      range: gridRange(SID, 2, 3, ci, ci+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${col}3`, values: [[label]] });
    // Values rows 4..
    vals.push({ range: `${S}!${col}4`, values: values.map(v => [v]) });
    // Style value cells
    reqs.push({ repeatCell: {
      range: gridRange(SID, 3, 3 + values.length, ci, ci+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.altRow), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
  });

  // Extra lists starting at row 50
  const extraCols = Object.keys(EXTRA_LISTS);
  extraCols.forEach((col, ci) => {
    const { label, values } = EXTRA_LISTS[col];
    reqs.push({ repeatCell: {
      range: gridRange(SID, 49, 50, ci, ci+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${col}50`, values: [[label]] });
    vals.push({ range: `${S}!${col}51`, values: values.map(v => [v]) });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 50, 50 + values.length, ci, ci+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.altRow), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: 10 },
    properties: { pixelSize: 180 }, fields: 'pixelSize',
  } });

  // Hide the tab
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, hidden: true },
    fields: 'hidden',
  } });

  await batchUpdate(id, reqs, 'ref-fmt');
  await valuesBatchUpdate(id, vals, 'ref-vals');
  console.log('✓ Reference Data tab complete (hidden)');
})().catch(e => { console.error(e.message || e); process.exit(1); });
