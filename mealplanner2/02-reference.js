'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S   = "'Reference Data'";

const LISTS = {
  MealTypes:        ['Breakfast','Morning Snack','Lunch','Afternoon Snack','Dinner','Dessert','Other'],
  RecipeCategories: ['Breakfast','Lunch','Dinner','Snack','Dessert','Soup','Salad','Pasta','Casserole',
                     'Slow Cooker','Sheet Pan','One Pot','Grill','Baking','Freezer Friendly','Meal Prep',
                     'Kid Friendly','Quick Meal','Special Occasion','Other'],
  GroceryCategories:['Produce','Meat & Seafood','Dairy & Eggs','Bakery','Pantry','Grains & Pasta',
                      'Canned Goods','Frozen','Snacks','Beverages','Condiments & Sauces',
                      'Spices & Seasonings','Baking','Household','Other'],
  IngredientUnits:  ['tsp','tbsp','cup','fl oz','oz','lb','g','kg','mL','L','piece','clove',
                      'can','jar','bag','box','package','bunch','slice','serving','to taste','other'],
  HouseholdRoles:   ['Self','Partner','Child','Roommate','Family Member','Guest','Other'],
  DietaryTags:      ['No Restriction','Vegetarian','Vegan','Pescatarian','Gluten-Free','Dairy-Free',
                      'Nut-Free','Egg-Free','Soy-Free','Halal','Kosher','Low Sodium','Low Sugar',
                      'High Protein','Other'],
  AllergenTags:     ['Peanuts','Tree Nuts','Dairy','Eggs','Wheat','Soy','Fish','Shellfish','Sesame','Other','None'],
  Difficulty:       ['Very Easy','Easy','Moderate','Advanced'],
  RecipeStatus:     ['Active','Favorite','Trial','Seasonal','Archived'],
  PantryStatus:     ['In Stock','Low Stock','Out of Stock','Expiring Soon','Frozen','Ignore'],
  GroceryStatus:    ['Needed','Already Have','Purchased','Optional','Review Units','Skip'],
  YesNo:            ['Yes','No'],
  Days:             ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
  Frequency:        ['Daily','Weekly','Biweekly','Monthly','Occasionally','Other'],
  DietaryCompatibility: ['Compatible','Review Restriction','Review Allergen','Missing Tags','Not Checked'],
};

const HEADERS = Object.keys(LISTS);

(async () => {
  const fmt = [];
  const vals = [];

  // Header row
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, HEADERS.length),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.text),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10, fontFamily: 'Arial' },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // Column headers
  vals.push({ range: `${S}!A1`, values: [HEADERS] });

  // List data
  HEADERS.forEach((key, col) => {
    const list = LISTS[key];
    const colLetter = String.fromCharCode(65 + col);
    vals.push({
      range: `${S}!${colLetter}2:${colLetter}${list.length + 1}`,
      values: list.map(v => [v]),
    });
    // Color-code meal-type column
    if (key === 'MealTypes') {
      const mealColors = [C.butter, C.peach, C.mint, C.peach, C.powder, C.lavender, C.gray];
      list.forEach((_, r) => {
        fmt.push({
          repeatCell: {
            range: gridRange(SID, r + 1, r + 2, col, col + 1),
            cell: { userEnteredFormat: { backgroundColor: hex(mealColors[r] || C.altRow) } },
            fields: 'userEnteredFormat.backgroundColor',
          },
        });
      });
    }
  });

  // Column widths
  fmt.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: HEADERS.length },
      properties: { pixelSize: 160 },
      fields: 'pixelSize',
    },
  });

  // Freeze row 1
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });

  // Hide sheet
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, hidden: true }, fields: 'hidden' } });

  await batchUpdate(id, fmt, 'ref-fmt');
  await valuesBatchUpdate(id, vals, 'ref-vals');
  console.log('✓ Reference Data complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
