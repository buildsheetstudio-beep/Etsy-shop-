'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const RCP = sheetMap['📖 Recipe Book'];

// A=# B=Name C=Cuisine D=MealType E=Difficulty F=Prep G=Cook H=Servings I=Cal J=Protein K=Carbs L=Fat M=AllergyTags N=KeyIngredients O=Rating
const recipes = [
  [1,'Classic Pancakes','American','Breakfast','Easy',15,15,4,350,8,45,12,'Vegetarian, Soy-Free','Flour, Eggs, Milk, Butter, Sugar',4],
  [2,'Avocado Toast','American','Breakfast','Easy',10,5,2,280,7,32,14,'Vegan, Gluten-Free, Soy-Free','Avocado, Sourdough, Lemon, Red Pepper Flakes',5],
  [3,'Greek Yogurt Bowl','Greek','Breakfast','Easy',5,0,1,320,18,42,8,'Vegetarian, Gluten-Free, Soy-Free','Greek Yogurt, Berries, Honey, Granola',4],
  [4,'Overnight Oats','American','Breakfast','Easy',10,0,2,310,9,52,7,'Vegan, Gluten-Free, Nut-Free','Oats, Almond Milk, Chia Seeds, Berries',4],
  [5,'Caesar Salad','Italian','Lunch','Easy',15,0,4,320,10,14,26,'Vegetarian, Egg-Free','Romaine, Parmesan, Croutons, Caesar Dressing',3],
  [6,'Chicken Tikka Masala','Indian','Dinner','Medium',20,35,4,420,35,18,22,'Gluten-Free, Nut-Free','Chicken, Tomatoes, Cream, Garam Masala, Ginger',5],
  [7,'Spaghetti Bolognese','Italian','Dinner','Medium',15,45,6,580,32,68,18,'Egg-Free, Dairy-Free','Pasta, Ground Beef, Tomatoes, Onion, Garlic',5],
  [8,'Veggie Stir Fry','Asian','Dinner','Easy',15,10,4,280,8,36,10,'Vegan, Gluten-Free, Nut-Free','Broccoli, Bell Peppers, Snap Peas, Ginger',4],
  [9,'Thai Peanut Noodles','Thai','Dinner','Easy',10,15,4,490,14,62,20,'Vegetarian, Egg-Free','Rice Noodles, Peanut Butter, Lime, Soy Sauce',4],
  [10,'Salmon Teriyaki','Japanese','Dinner','Medium',10,20,2,380,42,22,14,'Gluten-Free, Dairy-Free, Soy-Free','Salmon, Honey, Sesame Seeds, Ginger, Garlic',5],
  [11,'Tacos al Pastor','Mexican','Dinner','Medium',20,25,4,440,28,38,18,'Gluten-Free, Dairy-Free, Egg-Free','Pork, Pineapple, Corn Tortillas, Cilantro, Onion',5],
  [12,'Shakshuka','Mediterranean','Breakfast','Medium',10,20,4,290,14,22,16,'Vegetarian, Gluten-Free, Nut-Free','Eggs, Tomatoes, Bell Peppers, Feta, Cumin',4],
  [13,'French Onion Soup','French','Lunch','Hard',20,60,4,380,14,42,16,'Vegetarian, Egg-Free, Soy-Free','Onions, Beef Broth, Gruyere, Baguette, Thyme',4],
  [14,'Quinoa Power Bowl','Mediterranean','Lunch','Easy',15,20,4,420,16,58,12,'Vegan, Gluten-Free, Nut-Free, Soy-Free','Quinoa, Chickpeas, Cucumber, Cherry Tomatoes, Tahini',5],
  [15,'Almond Butter Cookies','American','Dessert','Easy',15,12,24,180,5,20,10,'Gluten-Free, Dairy-Free, Soy-Free, Keto','Almond Butter, Eggs, Sugar, Vanilla, Salt',4],
  [16,'Mango Lassi','Indian','Snack','Easy',5,0,2,210,6,38,4,'Vegetarian, Gluten-Free, Nut-Free','Mango, Yogurt, Milk, Cardamom, Rose Water',4],
  [17,'Beef Bulgogi','Asian','Dinner','Medium',30,15,4,480,36,28,22,'Gluten-Free, Dairy-Free, Egg-Free, Keto','Beef Sirloin, Soy Sauce, Sesame Oil, Pear, Garlic',5],
  [18,'Ratatouille','French','Dinner','Hard',30,60,6,240,5,32,10,'Vegan, Gluten-Free, Nut-Free, Soy-Free, Egg-Free','Zucchini, Eggplant, Tomatoes, Bell Peppers, Herbs',4],
  [19,'Keto Cauliflower Pizza','American','Dinner','Hard',20,30,4,320,22,12,22,'Gluten-Free, Nut-Free, Keto, Soy-Free','Cauliflower, Mozzarella, Eggs, Marinara, Toppings',4],
  [20,'Green Smoothie','American','Breakfast','Easy',5,0,1,180,4,28,6,'Vegan, Gluten-Free, Soy-Free, Keto, Nut-Free','Spinach, Banana, Almond Milk, Protein Powder, Flax',3],
];

const headers = ['#','Recipe Name','Cuisine','Meal Type','Difficulty','Prep (min)','Cook (min)','Servings','Cal/Serving','Protein (g)','Carbs (g)','Fat (g)','Allergy Tags','Key Ingredients','Rating'];

(async () => {
  const data = [
    { range: "'📖 Recipe Book'!A1", values: [['📖']] },
    { range: "'📖 Recipe Book'!B1", values: [['Recipe Book']] },
    { range: "'📖 Recipe Book'!A2", values: [headers] },
    { range: "'📖 Recipe Book'!A3", values: recipes },
  ];
  await valuesBatchUpdate(id, data, 'recipe-values');

  const reqs = [];

  // Merge title B1:O1
  reqs.push({ mergeCells: { range: gridRange(RCP, 0, 1, 1, 15), mergeType: 'MERGE_ALL' } });

  // A1 standalone: goldenMustard bg
  reqs.push({ repeatCell: {
    range: gridRange(RCP, 0, 1, 0, 1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.goldenMustard), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // B1:O1 merged title
  reqs.push({ repeatCell: {
    range: gridRange(RCP, 0, 1, 1, 15),
    cell: { userEnteredFormat: { backgroundColor: hex(C.goldenMustard), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Header row A2:O2
  reqs.push({ repeatCell: {
    range: gridRange(RCP, 1, 2, 0, 15),
    cell: { userEnteredFormat: { backgroundColor: hex(C.darkText), textFormat: { foregroundColor: hex(C.white), bold: true }, horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  // Data rows alternating
  for (let r = 0; r < 20; r++) {
    reqs.push({ repeatCell: {
      range: gridRange(RCP, 2 + r, 3 + r, 0, 15),
      cell: { userEnteredFormat: { backgroundColor: hex(r % 2 === 0 ? C.warmCream : C.altRow) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }
  // Center numeric cols
  reqs.push({ repeatCell: {
    range: gridRange(RCP, 2, 22, 5, 12),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }});
  // Rating center
  reqs.push({ repeatCell: {
    range: gridRange(RCP, 2, 22, 14, 15),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // Column widths
  const widths = [40,200,100,90,80,75,75,75,90,80,75,70,200,220,70];
  widths.forEach((px, c) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: RCP, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));
  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: RCP, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 34 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'recipe-format');
  console.log('Recipe Book complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
