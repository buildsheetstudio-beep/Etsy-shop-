'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recipe Book'];
const S   = "'Recipe Book'";
const REF = "'Reference Data'";

const RB1 = 6;   // 1-indexed first data row
const RBN = 1005;
const RB0 = RB1 - 1;  // 0-indexed

const COLS = ['Recipe ID','Recipe Name','Recipe Category','Primary Meal Type','Secondary Meal Type',
  'Base\nServings','Prep Time\n(min)','Cook Time\n(min)','Total Time\n(min)','Difficulty',
  'Est. Recipe\nCost','Est. Cost /\nServing','Favorite?','Recipe Status','Freezer\nFriendly?',
  'Meal Prep\nFriendly?','Kid\nFriendly?','Dietary\nTag 1','Dietary\nTag 2','Dietary\nTag 3',
  'Allergen\nTag 1','Allergen\nTag 2','Allergen\nTag 3','Instructions Summary','Source / Notes'];
// A=0..Y=24

const recipes = [
  // [name, category, mealType, secMeal, servings, prep, cook, difficulty, favChar, status, freezer, mealPrep, kidFriendly, dietTag1, dietTag2, dietTag3, allergen1, allergen2, allergen3, instructions, source]
  ['Classic Pancakes','Breakfast','Breakfast','',4,15,20,'Easy',true,'Favorite',false,true,true,'Vegetarian','','','Dairy','Eggs','Wheat','Mix flour, eggs, milk, butter. Cook on griddle until bubbles form, flip once.','Family recipe'],
  ['Avocado Toast with Eggs','Breakfast','Breakfast','',2,10,5,'Very Easy',true,'Favorite',false,false,false,'Vegetarian','','','Dairy','Eggs','Wheat','Toast bread, mash avocado, top with fried eggs and seasoning.','Adapted online'],
  ['Overnight Oats','Breakfast','Breakfast','Morning Snack',1,5,0,'Very Easy',false,'Active',true,true,true,'Vegetarian','Gluten-Free','','Dairy','','None','Combine oats, milk, yogurt, chia seeds. Refrigerate overnight.','Original'],
  ['Berry Smoothie Bowl','Breakfast','Breakfast','',2,10,0,'Very Easy',false,'Active',false,false,true,'Vegan','Gluten-Free','Dairy-Free','None','','None','Blend frozen berries, banana. Pour into bowl, top with granola and fruit.','Original'],
  ['Vegetable Frittata','Breakfast','Breakfast','Lunch',6,10,20,'Easy',false,'Active',false,true,false,'Vegetarian','Gluten-Free','','Dairy','Eggs','None','Sauté vegetables, pour egg mixture over, bake until set.','Adapted from cookbook'],
  ['French Toast Bake','Breakfast','Breakfast','',8,15,35,'Easy',true,'Favorite',true,true,true,'Vegetarian','','','Dairy','Eggs','Wheat','Layer bread in baking dish, pour custard mixture, bake until golden.','Original'],
  ['Granola with Yogurt','Breakfast','Breakfast','Morning Snack',1,5,0,'Very Easy',false,'Active',false,true,true,'Vegetarian','','','Dairy','','None','Layer yogurt, granola, and fresh berries in a bowl or jar.','Original'],
  ['Banana Oat Muffins','Baking','Breakfast','Morning Snack',12,15,22,'Easy',false,'Active',true,true,true,'Vegetarian','','','Dairy','Eggs','Wheat','Mash bananas, mix with oats, eggs, maple syrup. Bake at 350°F.','Original'],
  // Lunch
  ['Caesar Salad','Salad','Lunch','',4,15,0,'Very Easy',false,'Active',false,false,false,'Vegetarian','','','Dairy','Eggs','Wheat','Toss romaine with Caesar dressing, croutons, and Parmesan.','Classic'],
  ['Turkey & Avocado Wrap','Lunch','Lunch','',2,10,0,'Very Easy',true,'Favorite',false,true,true,'','','','Wheat','','None','Layer turkey, avocado, lettuce in a flour tortilla. Roll tightly.','Original'],
  ['Lentil Soup','Soup','Lunch','Dinner',6,15,40,'Easy',false,'Active',true,true,false,'Vegan','Gluten-Free','Dairy-Free','None','','None','Sauté onion, garlic; add lentils, broth, tomatoes, spices. Simmer until tender.','Family recipe'],
  ['Caprese Panini','Lunch','Lunch','',2,5,8,'Very Easy',false,'Active',false,false,true,'Vegetarian','','','Dairy','Wheat','None','Layer mozzarella, tomato, basil on ciabatta. Grill until golden.','Original'],
  ['Chicken Noodle Soup','Soup','Lunch','Dinner',6,20,40,'Easy',true,'Favorite',true,true,true,'','','','Wheat','','None','Simmer chicken in broth with vegetables and egg noodles.','Family recipe'],
  ['Greek Salad','Salad','Lunch','',4,15,0,'Very Easy',false,'Active',false,true,false,'Vegetarian','Gluten-Free','','Dairy','','None','Combine cucumbers, tomatoes, olives, feta, onion with olive oil and herbs.','Classic'],
  ['BLT Sandwich','Lunch','Lunch','',2,10,10,'Very Easy',false,'Active',false,false,true,'','','','Wheat','','None','Cook bacon until crispy. Layer on toasted bread with lettuce, tomato, mayo.','Classic'],
  ['Quinoa Power Bowl','Lunch','Lunch','',2,10,20,'Easy',false,'Active',false,true,false,'Vegan','Gluten-Free','Dairy-Free','None','','None','Cook quinoa; top with roasted vegetables, chickpeas, tahini dressing.','Original'],
  ['Tomato Basil Pasta','Pasta','Lunch','Dinner',4,10,20,'Easy',false,'Active',false,false,true,'Vegetarian','','','Dairy','Wheat','None','Cook pasta; toss with sautéed tomatoes, garlic, basil, olive oil.','Classic'],
  // Dinner
  ['Sheet Pan Chicken & Vegetables','Sheet Pan','Dinner','',4,15,35,'Easy',true,'Favorite',false,true,true,'Gluten-Free','','','None','','None','Toss chicken thighs and vegetables with olive oil and seasoning; roast at 425°F.','Original'],
  ['Beef Tacos','Dinner','Dinner','',4,15,20,'Easy',true,'Favorite',false,false,true,'','','','Wheat','Dairy','None','Brown ground beef with taco seasoning; serve in tortillas with toppings.','Original'],
  ['Vegetable Stir Fry','One Pot','Dinner','',4,20,15,'Easy',false,'Active',false,true,true,'Vegan','Gluten-Free','Dairy-Free','Soy','','None','Stir-fry vegetables in sesame oil with soy sauce and ginger over high heat.','Original'],
  ['Baked Salmon with Lemon','Sheet Pan','Dinner','',4,10,20,'Very Easy',true,'Favorite',false,false,false,'Gluten-Free','Dairy-Free','','Fish','','None','Season salmon with lemon, herbs; bake at 400°F until flaky.','Original'],
  ['Pasta Bolognese','Pasta','Dinner','',6,15,45,'Moderate','','Active',true,true,true,'','','','Wheat','Dairy','None','Brown beef, add tomatoes and herbs, simmer; serve over pasta.','Family recipe'],
  ['Chicken Tikka Masala','Dinner','Dinner','',4,20,35,'Moderate',false,'Favorite',true,true,false,'Gluten-Free','','','Dairy','','None','Marinate chicken in yogurt and spices; simmer in tomato-cream sauce.','Adapted cookbook'],
  ['Vegetable Curry','One Pot','Dinner','',4,15,30,'Easy',false,'Active',true,true,false,'Vegan','Gluten-Free','Dairy-Free','None','','None','Simmer mixed vegetables in coconut milk with curry paste and spices.','Original'],
  ['Mushroom Risotto','One Pot','Dinner','',4,10,40,'Moderate',false,'Active',false,false,false,'Vegetarian','Gluten-Free','','Dairy','','None','Toast arborio rice, add warm broth ladle by ladle, finish with Parmesan.','Classic'],
  ['BBQ Pulled Pork','Slow Cooker','Dinner','',8,15,480,'Easy',true,'Favorite',true,true,true,'Gluten-Free','Dairy-Free','','None','','None','Slow-cook pork shoulder with BBQ sauce on low 8 hours; shred and serve.','Original'],
  ['Shrimp Tacos','Dinner','Dinner','',4,15,10,'Easy',false,'Active',false,false,false,'','','','Shellfish','Wheat','Dairy','Season and sauté shrimp; serve in tortillas with slaw and lime crema.','Original'],
  ['Eggplant Parmesan','Casserole','Dinner','',6,20,45,'Moderate',false,'Active',true,true,false,'Vegetarian','','','Dairy','Eggs','Wheat','Bread and fry eggplant slices; layer with marinara and cheese; bake.','Classic'],
  ['Chicken Caesar Pasta','Pasta','Dinner','Lunch',4,15,25,'Easy',false,'Active',false,true,true,'','','','Dairy','Eggs','Wheat','Cook pasta; toss with Caesar dressing, grilled chicken, and Parmesan.','Original'],
  ['Black Bean Enchiladas','Casserole','Dinner','',6,20,30,'Easy',false,'Active',true,true,true,'Vegetarian','','','Dairy','Wheat','None','Fill tortillas with black beans and cheese; bake with enchilada sauce.','Original'],
  // Snacks
  ['Hummus with Veggies','Snack','Afternoon Snack','Morning Snack',4,10,0,'Very Easy',false,'Active',false,true,true,'Vegan','Gluten-Free','Dairy-Free','Sesame','','None','Serve store-bought or homemade hummus with sliced vegetables.','Original'],
  ['Guacamole & Chips','Snack','Afternoon Snack','',4,10,0,'Very Easy',true,'Favorite',false,false,true,'Vegan','Gluten-Free','Dairy-Free','None','','None','Mash avocado with lime juice, cilantro, tomato, onion; serve with chips.','Original'],
  ['Apple Slices with Peanut Butter','Snack','Morning Snack','',2,5,0,'Very Easy',false,'Active',false,false,true,'Vegan','Gluten-Free','Dairy-Free','Peanuts','','None','Slice apple; serve with peanut butter for dipping.','Original'],
  ['Cheese & Crackers','Snack','Morning Snack','Afternoon Snack',4,5,0,'Very Easy',false,'Active',false,false,true,'Vegetarian','','','Dairy','Wheat','None','Arrange assorted cheeses with crackers and grapes on a board.','Original'],
  ['Energy Bites','Snack','Morning Snack','',12,15,0,'Very Easy',false,'Active',true,true,true,'Vegetarian','Gluten-Free','','Dairy','','None','Mix oats, peanut butter, honey, chocolate chips; roll into balls; chill.','Original'],
  ['Caprese Skewers','Snack','Afternoon Snack','',6,10,0,'Very Easy',false,'Active',false,false,false,'Vegetarian','Gluten-Free','','Dairy','','None','Thread mozzarella, tomato, basil on skewers; drizzle balsamic.','Original'],
  // Desserts
  ['Chocolate Chip Cookies','Baking','Dessert','',24,15,12,'Easy',true,'Favorite',true,true,true,'Vegetarian','','','Dairy','Eggs','Wheat','Cream butter and sugar, add eggs and vanilla, fold in flour and chips; bake.','Family recipe'],
  ['Fruit Salad','Dessert','Dessert','Morning Snack',6,15,0,'Very Easy',false,'Active',false,true,true,'Vegan','Gluten-Free','Dairy-Free','None','','None','Combine seasonal fruits; toss with lime juice and honey.','Original'],
  ['Brownies','Baking','Dessert','',16,15,30,'Easy',true,'Favorite',true,false,true,'Vegetarian','','','Dairy','Eggs','Wheat','Melt chocolate and butter; mix with sugar, eggs, flour; bake.','Family recipe'],
  ['Rice Pudding','Dessert','Dessert','',4,5,30,'Easy',false,'Active',false,false,true,'Vegetarian','Gluten-Free','','Dairy','','None','Simmer rice in milk with sugar and vanilla until creamy.','Classic'],
  ['Banana Ice Cream','Dessert','Dessert','',4,5,0,'Very Easy',false,'Active',false,true,true,'Vegan','Gluten-Free','Dairy-Free','None','','None','Freeze bananas; blend until smooth. Serve immediately or freeze.','Original'],
  // More varied
  ['Minestrone Soup','Soup','Lunch','Dinner',8,20,40,'Easy',false,'Active',true,true,false,'Vegan','','','Wheat','','None','Simmer vegetables, beans, pasta in tomato broth with Italian herbs.','Classic'],
  ['Grilled Cheese Sandwich','Lunch','Lunch','',2,5,10,'Very Easy',false,'Active',false,false,true,'Vegetarian','','','Dairy','Wheat','None','Butter bread, fill with cheese, grill until golden and melty.','Classic'],
  ['Shakshuka','Breakfast','Breakfast','Dinner',4,10,25,'Easy',false,'Active',false,false,false,'Vegetarian','Gluten-Free','','Dairy','Eggs','None','Simmer tomato sauce with spices; poach eggs directly in sauce.','Middle Eastern classic'],
  ['Loaded Sweet Potato','Dinner','Dinner','Lunch',4,10,50,'Easy',false,'Active',false,true,false,'Vegetarian','Gluten-Free','','Dairy','','None','Bake sweet potato; top with black beans, cheese, sour cream, salsa.','Original'],
  ['Thai Peanut Noodles','Pasta','Dinner','Lunch',4,20,15,'Easy',false,'Active',false,true,false,'Vegan','','','Peanuts','Soy','Wheat','Cook noodles; toss with peanut sauce, cucumber, carrots, cilantro.','Adapted recipe'],
  ['White Bean & Kale Soup','Soup','Dinner','Lunch',6,15,35,'Easy',false,'Active',true,true,false,'Vegan','Gluten-Free','Dairy-Free','None','','None','Sauté aromatics; add white beans, kale, broth; simmer until tender.','Original'],
  ['Tuna Pasta Salad','Salad','Lunch','',4,15,15,'Very Easy',false,'Active',false,true,true,'','','','Wheat','','Fish','Cook pasta; toss with tuna, celery, mayo, lemon, herbs.','Classic'],
  ['Homemade Pizza','Dinner','Dinner','',4,20,20,'Moderate',true,'Favorite',false,false,true,'Vegetarian','','','Dairy','Wheat','None','Spread sauce on dough; top with cheese and veggies; bake until crisp.','Family recipe'],
  ['Oatmeal with Toppings','Breakfast','Breakfast','Morning Snack',2,5,10,'Very Easy',false,'Active',false,true,true,'Vegetarian','Gluten-Free','','Dairy','','None','Cook oats with milk; top with banana, berries, honey, and nuts.','Original'],
];

(async () => {
  const fmt = [];
  const vals = [];
  const NR = recipes.length;

  // ─── Title ────────────────────────────────────────────────────────────────
  const NC = COLS.length; // 25
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.peach), textFormat: { bold: true, fontSize: 20, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['📖 Recipe Book']] });

  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.butter), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A2`, values: [['Your complete recipe library — add, organize, and tag all your recipes here']] });

  // Separator row 3
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.peach) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // Flag row 4
  fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 3, 4, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warning), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A4`, values: [['  ⚑ RECIPE FLAGS: Red = Active recipe with no ingredients | Orange = Missing cost | Yellow = Missing dietary/allergen tags | Gray = Archived']] });

  // Column header row (row 5, 0-indexed 4)
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 4, 5, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A5`, values: [COLS] });

  // ─── Data rows ────────────────────────────────────────────────────────────
  const recipeVals = [];
  for (let i = 0; i < NR; i++) {
    const r1 = RB1 + i; // 1-indexed row
    const [name, cat, mealType, secMeal, servings, prep, cook, diff, fav, status,
           freezer, mealPrep, kidFriendly, diet1, diet2, diet3, all1, all2, all3,
           instructions, source] = recipes[i];
    recipeVals.push([
      `=IF(B${r1}="","","REC-"&TEXT(ROW()-5,"0000"))`,  // A: Recipe ID
      name,                                               // B
      cat,                                                // C
      mealType,                                           // D
      secMeal,                                            // E
      servings,                                           // F
      prep,                                               // G
      cook,                                               // H
      `=IFERROR(G${r1}+H${r1},"")`,                      // I: Total time
      diff,                                               // J
      `=IFERROR(SUMIFS('Recipe Ingredients'!$I$6:$I$8005,'Recipe Ingredients'!$B$6:$B$8005,A${r1}),0)`, // K
      `=IFERROR(K${r1}/F${r1},0)`,                       // L
      fav,                                                // M
      status,                                             // N
      freezer,                                            // O
      mealPrep,                                           // P
      kidFriendly,                                        // Q
      diet1,                                              // R
      diet2,                                              // S
      diet3,                                              // T
      all1,                                               // U
      all2,                                               // V
      all3,                                               // W
      instructions,                                       // X
      source,                                             // Y
    ]);

    // Row background alternating
    const bgColor = i % 2 === 0 ? C.panel : C.altRow;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, RB0+i, RB0+i+1, 0, NC),
        cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 10, fontFamily: 'Arial' } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    // Formula cells: A, I, K, L
    [0, 8, 10, 11].forEach(col => {
      fmt.push({
        repeatCell: {
          range: gridRange(SID, RB0+i, RB0+i+1, col, col+1),
          cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      });
    });
    // Input cells: B-H, J, M-Y
    [1,2,3,4,5,6,7,9,12,13,14,15,16,17,18,19,20,21,22,23,24].forEach(col => {
      fmt.push({
        repeatCell: {
          range: gridRange(SID, RB0+i, RB0+i+1, col, col+1),
          cell: { userEnteredFormat: { backgroundColor: hex(C.input) } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      });
    });
  }

  vals.push({ range: `${S}!A${RB1}:Y${RB1+NR-1}`, values: recipeVals });

  // Checkboxes for columns M, O, P, Q (indices 12, 14, 15, 16)
  [12, 14, 15, 16].forEach(col => {
    fmt.push({
      setDataValidation: {
        range: gridRange(SID, RB0, RB0+NR, col, col+1),
        rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
      },
    });
  });

  // Dropdowns for Category (C=2), MealType (D=3), SecMeal (E=4), Difficulty (J=9), Status (N=13), DietaryTag (R,S,T=17,18,19), Allergen (U,V,W=20,21,22)
  const dropDefs = [
    { col: 2,  refCol: 'C' },  // Recipe Category
    { col: 3,  refCol: 'A' },  // Primary Meal Type
    { col: 4,  refCol: 'A' },  // Secondary Meal Type
    { col: 9,  refCol: 'H' },  // Difficulty
    { col: 13, refCol: 'I' },  // Recipe Status
    { col: 17, refCol: 'I' },  // Dietary Tag 1 (col I in ref = DietaryTags? let's check order)
    { col: 18, refCol: 'I' },
    { col: 19, refCol: 'I' },
    { col: 20, refCol: 'J' },  // Allergen Tag (col J)
    { col: 21, refCol: 'J' },
    { col: 22, refCol: 'J' },
  ];
  // Reference Data column mapping:
  // A=MealTypes, B=RecipeCategories, C=GroceryCategories, D=IngredientUnits
  // E=HouseholdRoles, F=DietaryTags, G=AllergenTags, H=Difficulty, I=RecipeStatus
  // J=PantryStatus, K=GroceryStatus, L=YesNo, M=Days, N=Frequency, O=DietaryCompatibility
  const refColMap = {
    2:  'B', // Recipe Category
    3:  'A', // Meal Type
    4:  'A', // Meal Type
    9:  'H', // Difficulty
    13: 'I', // Recipe Status
    17: 'F', // Dietary Tag
    18: 'F',
    19: 'F',
    20: 'G', // Allergen Tag
    21: 'G',
    22: 'G',
  };
  Object.entries(refColMap).forEach(([col, refLetter]) => {
    fmt.push({
      setDataValidation: {
        range: gridRange(SID, RB0, RB0+NR+500, parseInt(col), parseInt(col)+1),
        rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$${refLetter}$2:$${refLetter}$50` }] }, showCustomUi: true, strict: false },
      },
    });
  });

  // Currency format for K, L (cols 10, 11)
  fmt.push({
    repeatCell: {
      range: gridRange(SID, RB0, RB0+NR+500, 10, 12),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ─── Column widths ────────────────────────────────────────────────────────
  const colWidths = [75,160,110,100,100,60,65,65,65,90,80,70,60,100,65,75,60,100,100,100,100,100,100,200,160];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Freeze rows 1:5 (no column freeze — title rows merge across all columns)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'rb-fmt');
  await valuesBatchUpdate(id, vals, 'rb-vals');
  console.log(`✓ Recipe Book complete (${NR} recipes)`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
