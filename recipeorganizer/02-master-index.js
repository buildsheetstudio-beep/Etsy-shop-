'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Master Recipe Index'];
const S = "'Master Recipe Index'";
const REF = "'Reference Data'";

// Column headers (A=0 through AH=33)
const HEADERS = [
  'Recipe ID','Recipe Name','Primary Category','Secondary Category','Meal Type','Cuisine',
  'Cooking Method','Base Servings','Prep Time (min)','Cook Time (min)','Total Time (min)',
  'Difficulty','Personal Rating','Household Rating','Favorite?','Kid Friendly?',
  'Freezer Friendly?','Meal Prep Friendly?','Dietary Tag 1','Dietary Tag 2','Dietary Tag 3',
  'Allergen Tag 1','Allergen Tag 2','Allergen Tag 3','Est. Recipe Cost','Cost per Serving',
  'Recipe Source','Recipe URL','Date Added','Last Cooked','Times Cooked','Recipe Status',
  'Testing Notes','General Notes',
];

// 60 sample recipes
// [Name, PrimaryCategory, SecondaryCategory, MealType, Cuisine, CookingMethod, BaseServings,
//  PrepMins, CookMins, Difficulty, PersonalRating, HouseholdRating, Favorite, KidFriendly,
//  FreezerFriendly, MealPrepFriendly, DietaryTag1, DietaryTag2, DietaryTag3,
//  AllergenTag1, AllergenTag2, AllergenTag3, Source, URL, DateAdded, LastCooked,
//  TimesCooked, Status, TestingNotes, GeneralNotes]
const RECIPES = [
  ['Classic Buttermilk Pancakes','Breakfast','Family Favorites','Breakfast','American','Stovetop',8,10,20,'Easy',5,5,true,true,true,true,'Vegetarian','','','Dairy','Eggs','Wheat','Family Cookbook','www.myfamilycook.com/pancakes','01/05/2024','08/10/2026',12,'Approved','','Top weekend breakfast. Use real buttermilk.'],
  ['Spinach & Feta Omelette','Breakfast','','Breakfast','Mediterranean','Stovetop',2,5,10,'Easy',4,4,false,false,false,true,'Vegetarian','Gluten-Free','','Dairy','Eggs','','Home Kitchen','','02/10/2024','07/28/2026',8,'Approved','','Quick weekday protein-packed breakfast.'],
  ['Blueberry Overnight Oats','Breakfast','Meal Prep','Breakfast','American','No Cook',1,5,0,'Very Easy',4,4,false,false,false,true,'Vegetarian','','','Dairy','','','Oat Lovers Blog','www.oatlovers.com/overnight-oats','03/01/2024','08/15/2026',15,'Approved','','Prep 5 jars Sunday night for the whole week.'],
  ['Avocado Toast with Poached Eggs','Breakfast','','Brunch','American','Stovetop',2,10,10,'Easy',5,4,true,false,false,false,'Vegetarian','Dairy-Free','','Eggs','Wheat','','Brunch Society','www.brunchsociety.com/avo-toast','01/20/2024','08/12/2026',20,'Approved','','Poach eggs in gentle simmer with a splash of vinegar.'],
  ['Banana Walnut Muffins','Breakfast','Baking','Breakfast','American','Oven',12,15,25,'Easy',4,5,false,true,true,false,'Vegetarian','','','Dairy','Eggs','Tree Nuts','Bake It Simple','www.bakeitsimple.com/banana-muffins','02/15/2024','07/20/2026',7,'Approved','','Use very ripe bananas for best flavor.'],
  ['French Toast Casserole','Breakfast','Family Favorites','Brunch','American','Oven',8,20,45,'Easy',5,5,true,true,true,true,'Vegetarian','','','Dairy','Eggs','Wheat','Sunday Brunch Co.','www.sundaybrunch.com/casserole','12/10/2023','08/04/2026',5,'Approved','','Assemble the night before, bake in the morning.'],
  ['Greek Yogurt Parfait','Breakfast','Snacks','Breakfast','Greek','No Cook',2,5,0,'Very Easy',3,3,false,true,false,false,'Vegetarian','Gluten-Free','','Dairy','','','Quick Bites','','04/01/2024','08/18/2026',25,'Approved','','Layer yogurt, granola, fresh berries. Great for kids.'],
  ['Shakshuka','Breakfast','','Brunch','Middle Eastern','Stovetop',4,10,20,'Easy',5,5,true,false,false,false,'Vegetarian','Gluten-Free','','Eggs','','','Spice Route Kitchen','www.spiceroute.com/shakshuka','11/05/2023','08/01/2026',10,'Approved','','Add crumbled feta on top before serving.'],
  ['Classic Caesar Salad','Salads','','Lunch','American','No Cook',4,15,0,'Easy',4,4,false,false,false,false,'','Gluten-Free','','Dairy','Eggs','','Salad Days','www.saladdays.com/caesar','03/15/2024','08/05/2026',6,'Approved','','Make dressing from scratch — worth it.'],
  ['Turkey & Avocado Wrap','Sandwiches','','Lunch','American','No Cook',2,10,0,'Very Easy',4,4,false,true,false,true,'Dairy-Free','','','Wheat','','','Quick Lunch Box','','04/10/2024','08/14/2026',18,'Approved','','Add a thin layer of hummus for extra flavor.'],
  ['Tomato Basil Soup','Soups','','Lunch','Italian','Stovetop',6,10,30,'Easy',5,5,true,true,true,true,'Vegetarian','Gluten-Free','Dairy-Free','','','','Comfort Kitchen','www.comfortkitchen.com/tomato-soup','10/01/2023','08/08/2026',14,'Approved','','Blend half for texture, leave half chunky.'],
  ['Mediterranean Quinoa Bowl','Rice & Grain','Vegetarian','Lunch','Mediterranean','Stovetop',4,15,20,'Easy',4,4,true,false,false,true,'Vegetarian','Gluten-Free','Dairy-Free','','','','Bowl Culture','www.bowlculture.com/quinoa','02/20/2024','07/30/2026',9,'Approved','','Top with roasted chickpeas for crunch.'],
  ['Chicken BLT Sandwich','Sandwiches','Chicken','Lunch','American','Stovetop',2,10,15,'Easy',4,5,false,true,false,false,'','','','Wheat','','','Deli Street','','05/01/2024','08/11/2026',11,'Approved','','Toast the bread and add a garlic aioli.'],
  ['Lentil Soup','Soups','Vegetarian','Lunch','Mediterranean','Stovetop',6,15,35,'Easy',4,4,false,false,true,true,'Vegan','Gluten-Free','','','','','Vegan Pantry','www.veganpantry.com/lentil-soup','09/15/2023','08/06/2026',8,'Approved','','Rinse lentils twice before cooking.'],
  ['Caprese Salad','Salads','','Lunch','Italian','No Cook',4,10,0,'Very Easy',4,4,false,false,false,false,'Vegetarian','Gluten-Free','','Dairy','','','Italian Table','www.italiantable.com/caprese','05/20/2024','08/03/2026',6,'Approved','','Use good quality fresh mozzarella and basil.'],
  ['Spaghetti Bolognese','Pasta','Beef','Dinner','Italian','Stovetop',6,20,60,'Moderate',5,5,true,true,true,true,'','','','Wheat','Dairy','','Nonna Ginas Recipes','www.nonnarecipes.com/bolognese','08/01/2023','08/09/2026',22,'Approved','','Simmer at least 1 hour for deep flavor.'],
  ['Chicken Tikka Masala','Chicken','','Dinner','Indian','Stovetop',4,20,40,'Moderate',5,5,true,false,false,true,'Gluten-Free','','','Dairy','','','Spice Garden','www.spicegarden.com/tikka','07/15/2023','08/07/2026',18,'Approved','','Marinate chicken overnight if possible.'],
  ['Sheet Pan Lemon Herb Salmon','Seafood','','Dinner','Mediterranean','Oven',4,10,20,'Easy',5,4,true,false,false,true,'Gluten-Free','Dairy-Free','','Fish','','','Simple Suppers','www.simplesuppers.com/salmon','09/01/2023','08/13/2026',12,'Approved','','Line pan with foil for easy cleanup.'],
  ['Beef Tacos','Beef','','Dinner','Mexican','Stovetop',6,15,20,'Easy',5,5,true,true,false,false,'','','','Wheat','Dairy','','Taco Tuesdays','www.tacotuesdays.com/beef','06/10/2023','08/16/2026',30,'Approved','','Season beef well with cumin and chili powder.'],
  ['Mushroom Risotto','Rice & Grain','Vegetarian','Dinner','Italian','Stovetop',4,15,35,'Moderate',4,4,false,false,false,false,'Vegetarian','Gluten-Free','','Dairy','','','Risotto Maestro','www.risottomaestro.com','10/20/2023','07/25/2026',7,'Approved','','Add stock warm, one ladle at a time.'],
  ['Slow Cooker Pulled Pork','Pork','','Dinner','American','Slow Cooker',10,15,480,'Easy',5,5,true,true,true,true,'Gluten-Free','Dairy-Free','','','','','BBQ Nation','www.bbqnation.com/pulled-pork','08/20/2023','08/02/2026',9,'Approved','','Rub the night before, cook on low 8-10 hrs.'],
  ['Thai Green Curry','Chicken','','Dinner','Thai','Stovetop',4,15,25,'Moderate',5,5,true,false,false,true,'Gluten-Free','Dairy-Free','','','','','Thai Kitchen','www.thaikitchen.com/green-curry','11/01/2023','07/22/2026',8,'Approved','','Use full-fat coconut milk for creaminess.'],
  ['Grilled Teriyaki Chicken','Chicken','','Dinner','Japanese','Grill',4,10,20,'Easy',4,4,true,true,false,true,'Gluten-Free','Dairy-Free','','Soy','','','Grill Master','www.grillmaster.com/teriyaki','05/15/2024','08/17/2026',13,'Approved','','Marinate at least 2 hours.'],
  ['Beef Stew','Beef','Slow Cooker','Dinner','American','Slow Cooker',6,20,240,'Easy',5,5,true,true,true,true,'Gluten-Free','Dairy-Free','','Wheat','','','Hearty Kitchen','www.heartykitchen.com/stew','12/01/2023','07/15/2026',6,'Approved','','Potatoes and carrots go in last 45 minutes.'],
  ['Shrimp Scampi','Seafood','Pasta','Dinner','Italian','Stovetop',4,10,15,'Easy',5,4,true,false,false,false,'','','','Shellfish','Wheat','Dairy','Coastal Kitchen','www.coastalkitchen.com/scampi','06/20/2023','08/04/2026',10,'Approved','','Use dry white wine, not sweet.'],
  ['Vegetable Stir Fry','Vegetarian','','Dinner','Chinese','Stovetop',4,15,15,'Easy',4,4,false,true,false,true,'Vegan','Gluten-Free','','Soy','','','Wok Star','www.wokstar.com/vegstirfry','03/10/2024','08/12/2026',11,'Approved','','High heat is key — don\'t overcrowd the wok.'],
  ['BBQ Baby Back Ribs','Pork','Grilling','Dinner','American','Oven',4,20,180,'Moderate',5,5,true,true,false,false,'Gluten-Free','Dairy-Free','','','','','Smokehouse','www.smokehouse.com/ribs','07/04/2023','08/01/2026',5,'Approved','','Bake 3 hrs then finish on grill for char.'],
  ['One Pot Chicken Pasta','Chicken','Pasta','Dinner','Italian','Stovetop',4,10,30,'Easy',5,5,true,true,false,true,'','','','Wheat','Dairy','','Easy Dinners','www.easydinnersco.com/chicken-pasta','04/05/2024','08/10/2026',16,'Approved','','Everything cooks in one pot — minimal cleanup!'],
  ['Korean Beef Bowl','Beef','Rice & Grain','Dinner','Korean','Stovetop',4,15,20,'Easy',5,5,true,true,false,true,'Dairy-Free','','','Soy','Wheat','','Seoul Eats','www.seouleats.com/beef-bowl','03/25/2024','08/15/2026',14,'Approved','','Serve over steamed rice with kimchi on the side.'],
  ['Garlic Butter Shrimp','Seafood','','Dinner','American','Stovetop',4,5,10,'Easy',5,5,true,false,false,false,'Gluten-Free','','','Shellfish','Dairy','','Quick Catch','www.quickcatch.com/shrimp','05/30/2024','08/18/2026',20,'Approved','','Don\'t overcook — shrimp turns rubbery fast.'],
  ['Homemade Guacamole','Snacks','','Snack','Mexican','No Cook',6,10,0,'Very Easy',5,5,true,true,false,false,'Vegan','Gluten-Free','','','','','Avocado Everything','www.avocadoeverything.com','06/01/2024','08/19/2026',30,'Approved','','Mash chunky. Add lime and salt gradually.'],
  ['Stuffed Mushrooms','Snacks','Appetizer','Appetizer','American','Oven',6,15,25,'Easy',4,4,false,false,false,false,'Vegetarian','Gluten-Free','','Dairy','Eggs','','Party Bites','www.partybites.com/mushrooms','11/20/2023','07/10/2026',4,'Approved','','Remove stems and stuff with cream cheese mix.'],
  ['Spinach Artichoke Dip','Snacks','','Appetizer','American','Oven',8,10,25,'Easy',5,5,true,false,false,false,'Vegetarian','Gluten-Free','','Dairy','Eggs','','Game Day Grub','www.gamedaygrub.com/dip','01/20/2024','07/31/2026',8,'Approved','','Serve hot with pita chips or sliced bread.'],
  ['Caprese Bruschetta','Snacks','','Appetizer','Italian','Oven',8,15,10,'Easy',4,4,false,false,false,false,'Vegetarian','','','Wheat','Dairy','','Italian Bites','www.italianbites.com/bruschetta','02/14/2024','08/07/2026',6,'Approved','','Toast bread until golden before topping.'],
  ['Crispy Air Fryer Chickpeas','Snacks','','Snack','Mediterranean','Air Fryer',4,5,20,'Very Easy',4,4,false,true,false,true,'Vegan','Gluten-Free','','','','','Crispy Things','www.crispythings.com/chickpeas','05/05/2024','08/16/2026',9,'Approved','','Pat chickpeas very dry before air frying.'],
  ['Chicken Noodle Soup','Soups','Chicken','Dinner','American','Stovetop',8,20,45,'Easy',5,5,true,true,true,true,'Dairy-Free','','','Wheat','','','Grandma\'s Kitchen','www.grandmaskitchen.com/cns','10/15/2023','08/08/2026',11,'Approved','','Use homemade or quality stock for best flavor.'],
  ['French Onion Soup','Soups','','Dinner','French','Stovetop',4,15,60,'Moderate',5,4,false,false,false,false,'Vegetarian','','','Wheat','Dairy','','Bistro Classics','www.bistroclassics.com/fos','11/10/2023','07/18/2026',5,'Approved','','Caramelize onions slowly — at least 40 minutes.'],
  ['Minestrone','Soups','Vegetarian','Lunch','Italian','Stovetop',8,20,40,'Easy',4,4,false,false,true,true,'Vegan','','','Wheat','','','Italian Table','www.italiantable.com/minestrone','10/05/2023','08/05/2026',7,'Approved','','Add pasta in the last 10 min to avoid mushiness.'],
  ['Creamy Butternut Squash Soup','Soups','','Lunch','American','Oven',6,20,50,'Easy',5,4,false,false,true,true,'Vegetarian','Gluten-Free','','Dairy','','','Autumn Kitchen','www.autumnkitchen.com/bss','09/20/2023','07/28/2026',6,'Seasonal','','Roast squash first for deeper flavor.'],
  ['Chocolate Chip Cookies','Desserts','Baking','Dessert','American','Oven',36,15,12,'Easy',5,5,true,true,true,false,'Vegetarian','','','Dairy','Eggs','Wheat','Classic Bakes','www.classicbakes.com/choco-cookies','08/10/2023','08/11/2026',25,'Approved','','Brown the butter first for a nutty, rich cookie.'],
  ['Classic Cheesecake','Desserts','Baking','Dessert','American','Oven',12,30,60,'Moderate',5,5,true,false,true,false,'Vegetarian','','','Dairy','Eggs','Wheat','Bake It Right','www.bakeitright.com/cheesecake','07/15/2023','07/04/2026',4,'Approved','','Use water bath to prevent cracking.'],
  ['Tiramisu','Desserts','','Dessert','Italian','No Cook',9,30,0,'Moderate',5,5,true,false,false,false,'Vegetarian','','','Dairy','Eggs','Wheat','Dolce Vita','www.dolcevita.com/tiramisu','12/20/2023','08/15/2026',3,'Approved','','Refrigerate at least 6 hours before serving.'],
  ['Apple Pie','Desserts','Baking','Dessert','American','Oven',8,30,60,'Moderate',5,5,true,true,true,false,'Vegetarian','','','Dairy','Wheat','','Pie Palace','www.piepalace.com/apple','10/25/2023','11/28/2025',4,'Seasonal','','Add cinnamon and nutmeg to the filling.'],
  ['Chocolate Lava Cake','Desserts','','Dessert','French','Oven',4,15,12,'Moderate',5,5,true,false,false,false,'Vegetarian','','','Dairy','Eggs','Wheat','Decadent Bakes','www.decadentbakes.com/lava','02/10/2024','08/08/2026',6,'Approved','','Serve immediately with vanilla ice cream.'],
  ['Lemon Bars','Desserts','Baking','Dessert','American','Oven',16,20,40,'Easy',4,5,false,false,true,false,'Vegetarian','','','Dairy','Eggs','Wheat','Citrus Sweets','www.citrussweets.com/lemon-bars','03/10/2024','06/20/2026',3,'Approved','','Dust powdered sugar just before serving.'],
  ['Banana Bread','Bread','Baking','Dessert','American','Oven',10,15,65,'Easy',5,5,true,true,true,false,'Vegetarian','','','Dairy','Eggs','Wheat','Bread & Butter','www.breadandbutter.com/banana','01/15/2024','08/14/2026',18,'Approved','','3 overripe bananas is the sweet spot.'],
  ['Strawberry Shortcake','Desserts','','Dessert','American','Oven',8,20,25,'Easy',5,5,true,true,false,false,'Vegetarian','','','Dairy','Eggs','Wheat','Sweet Occasions','www.sweetoccasions.com/shortcake','06/15/2024','07/20/2026',4,'Approved','','Macerate strawberries with sugar 30 min ahead.'],
  ['No-Knead Artisan Bread','Bread','','Dinner','American','Oven',12,10,45,'Easy',5,5,true,true,true,false,'Vegetarian','','','Wheat','','','Bread Lab','www.breadlab.com/no-knead','02/01/2024','08/09/2026',10,'Approved','','Mix dough the night before; let rise 12-18 hrs.'],
  ['Dinner Rolls','Bread','Baking','Dinner','American','Oven',12,30,25,'Easy',4,5,false,true,true,false,'Vegetarian','','','Dairy','Eggs','Wheat','Roll Call','www.rollcall.com/dinner-rolls','11/25/2023','12/25/2025',3,'Seasonal','','Best served warm with salted butter.'],
  ['Zucchini Bread','Bread','Baking','Breakfast','American','Oven',10,20,60,'Easy',4,4,false,false,true,false,'Vegetarian','','','Dairy','Eggs','Wheat','Garden Bakes','www.gardenbakes.com/zucchini','07/20/2023','08/01/2026',7,'Approved','','Squeeze moisture out of grated zucchini.'],
  ['Turkey Meatballs','Meal Prep','Chicken','Dinner','Italian','Oven',6,20,25,'Easy',4,4,false,true,true,true,'','','','Eggs','Wheat','','Meal Prep Pro','www.mealprepro.com/meatballs','02/25/2024','08/12/2026',12,'Approved','','Freeze raw or cooked. Both work well.'],
  ['Chicken Burrito Bowls','Meal Prep','Chicken','Lunch','Mexican','Oven',5,15,30,'Easy',5,5,true,true,true,true,'Gluten-Free','Dairy-Free','','','','','Prep & Eat','www.prepandeat.com/burrito-bowl','03/20/2024','08/17/2026',20,'Approved','','Prep 5 containers on Sunday for the week.'],
  ['Freezer Breakfast Burritos','Meal Prep','Breakfast','Breakfast','Mexican','Stovetop',8,20,20,'Easy',4,4,false,true,true,true,'','','','Dairy','Eggs','Wheat','Freezer Friendly','www.freezerfriendly.com/burritos','04/15/2024','08/08/2026',6,'Approved','','Wrap individually in foil then freeze in a bag.'],
  ['Slow Cooker Chili','Slow Cooker','Beef','Dinner','American','Slow Cooker',8,15,360,'Very Easy',5,5,true,true,true,true,'Gluten-Free','','','','','','Chili Nation','www.chilinationco.com/beef-chili','09/05/2023','08/10/2026',14,'Approved','','Better the next day — make ahead!'],
  ['Meal Prep Veggie Bowls','Meal Prep','Vegetarian','Lunch','Mediterranean','Oven',5,20,35,'Easy',4,4,false,false,false,true,'Vegan','Gluten-Free','','','','','Plant Prep','www.plantprep.com/veggie-bowls','05/10/2024','08/14/2026',8,'Approved','','Roast veggies at 425°F for caramelization.'],
  ['Asian Sesame Salad','Salads','','Lunch','Asian','No Cook',4,15,0,'Easy',4,4,false,false,false,true,'Vegan','','','Soy','Sesame','Wheat','Sesame Street Eats','www.sesamestreateats.com','04/20/2024','08/11/2026',9,'Approved','','Make dressing fresh — it makes a big difference.'],
  ['Waldorf Salad','Salads','','Lunch','American','No Cook',4,15,0,'Easy',3,3,false,false,false,false,'Vegetarian','Gluten-Free','','Dairy','Tree Nuts','Eggs','Classic Table','','03/05/2024','07/22/2026',4,'Approved','','Use Granny Smith apples for tartness.'],
  ['Greek Salad','Salads','','Lunch','Greek','No Cook',4,15,0,'Very Easy',4,4,false,false,false,false,'Vegetarian','Gluten-Free','','Dairy','','','Aegean Eats','www.aegeaneats.com/greek-salad','05/25/2024','08/13/2026',7,'Approved','','Only dress right before serving.'],
  ['Homemade Marinara Sauce','Sauces & Condiments','','Dinner','Italian','Stovetop',6,5,35,'Easy',5,5,false,false,true,true,'Vegan','Gluten-Free','','','','','La Cucina','www.lacucina.com/marinara','06/05/2024','08/06/2026',10,'Approved','','Simmer slowly. San Marzano tomatoes preferred.'],
  ['Mango Lassi','Drinks','Beverages','Beverage','Indian','Blender',4,5,0,'Very Easy',5,5,true,true,false,false,'Vegetarian','Gluten-Free','','Dairy','','','Tropical Sips','www.tropicalsips.com/lassi','07/01/2024','08/18/2026',15,'Approved','','Use frozen mango if fresh is unavailable.'],
];

(async () => {
  const reqs = [];
  const vals = [];

  // ── Title row ────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 34), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 34),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, fontSize: 18, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  } });
  vals.push({ range: `${S}!A1`, values: [['🍴 MASTER RECIPE INDEX']] });

  // ── Subtitle / instruction row ────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 34), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 34),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A2`, values: [['Source of truth for all recipe information — enter each recipe once here and it flows throughout the entire workbook.']] });

  // ── Alert / info rows 3-5 ────────────────────────────────────────────────────
  // Row 3: Alerts header
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 2, 3, 0, 6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.attention),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A3`, values: [['⚠ RECIPE ALERTS — rows flagged below have data issues to review']] });

  // Rows 4-6: collection stats
  ['Total Recipes','Approved Recipes','Favorites'].forEach((lbl, i) => {
    const col = i * 2;
    reqs.push({ repeatCell: {
      range: gridRange(SID, 3, 4, col, col+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.info),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${colL(col)}4`, values: [[lbl]] });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 4, 5, col, col+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.formula),
        textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.primary), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
  });
  vals.push({ range: `${S}!A5`, values: [[`=COUNTA($B$8:$B$2007)`]] });
  vals.push({ range: `${S}!C5`, values: [[`=COUNTIF($AF$8:$AF$2007,"Approved")`]] });
  vals.push({ range: `${S}!E5`, values: [[`=COUNTIF($O$8:$O$2007,TRUE)`]] });

  // ── Stats row heights ─────────────────────────────────────────────────────────
  [1,2,3,4,5,6].forEach(i => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: i < 3 ? 22 : 28 }, fields: 'pixelSize',
    } });
  });

  // ── Section label row 7 ───────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 34),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryDeep),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
  } });
  vals.push({ range: `${S}!A7`, values: [HEADERS] });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 },
    properties: { pixelSize: 36 }, fields: 'pixelSize',
  } });

  // ── Column widths ─────────────────────────────────────────────────────────────
  const COL_WIDTHS = {
    0:80,1:220,2:130,3:130,4:100,5:110,6:120,7:75,8:70,9:70,10:70,
    11:90,12:70,13:70,14:60,15:60,16:70,17:70,18:110,19:110,20:110,
    21:100,22:100,23:100,24:90,25:85,26:140,27:180,28:90,29:90,30:75,
    31:100,32:160,33:200,
  };
  Object.entries(COL_WIDTHS).forEach(([ci, px]) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: +ci, endIndex: +ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    } });
  });

  // ── Data validation dropdowns ─────────────────────────────────────────────────
  const DV_COLS = {
    2: { ref: `${REF}!B4:B32` },   // Primary Category
    3: { ref: `${REF}!B4:B32` },   // Secondary Category
    4: { ref: `${REF}!A4:A14` },   // Meal Type
    5: { ref: `${REF}!C4:C22` },   // Cuisine
    6: { ref: `${REF}!F4:F16` },   // Cooking Method
    11:{ ref: `${REF}!G4:G7`  },   // Difficulty
    12:{ list: ['1','2','3','4','5'] },  // Personal Rating
    13:{ list: ['1','2','3','4','5'] },  // Household Rating
    18:{ ref: `${REF}!D4:D18` },   // Dietary Tag 1
    19:{ ref: `${REF}!D4:D18` },   // Dietary Tag 2
    20:{ ref: `${REF}!D4:D18` },   // Dietary Tag 3
    21:{ ref: `${REF}!E4:E14` },   // Allergen Tag 1
    22:{ ref: `${REF}!E4:E14` },   // Allergen Tag 2
    23:{ ref: `${REF}!E4:E14` },   // Allergen Tag 3
    31:{ ref: `${REF}!H4:H9` },    // Recipe Status
  };
  Object.entries(DV_COLS).forEach(([ci, dv]) => {
    const cond = dv.ref
      ? { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${dv.ref}` }] }
      : { type: 'ONE_OF_LIST', values: dv.list.map(v => ({ userEnteredValue: v })) };
    reqs.push({ setDataValidation: {
      range: gridRange(SID, 7, 2007, +ci, +ci+1),
      rule: { condition: cond, strict: false, showCustomUi: true },
    } });
  });

  // Checkboxes for Favorite, Kid Friendly, Freezer Friendly, Meal Prep Friendly
  [14,15,16,17].forEach(ci => {
    reqs.push({ setDataValidation: {
      range: gridRange(SID, 7, 2007, ci, ci+1),
      rule: { condition: { type: 'BOOLEAN' }, strict: true },
    } });
  });

  // ── Freeze rows 1-7 and cols A-C ─────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } },
    fields: 'gridProperties.frozenRowCount',
  } });

  // ── Write recipe data rows ────────────────────────────────────────────────────
  const dataRows = RECIPES.map((rec, i) => {
    const r = 8 + i; // 1-indexed row number
    const [name, cat, secCat, mealType, cuisine, method, servings,
           prepMins, cookMins, diff, rating, hhRating, fav, kidF, freezF, mealP,
           diet1, diet2, diet3, allerg1, allerg2, allerg3,
           source, url, dateAdded, lastCooked, timesCkd, status, testNotes, genNotes] = rec;

    return [
      `=IF(B${r}="","","REC-"&TEXT(ROW()-7,"0000"))`, // A: Recipe ID
      name, cat, secCat, mealType, cuisine, method,
      servings, prepMins, cookMins,
      `=IFERROR(I${r}+J${r},"")`, // K: Total Time
      diff, rating, hhRating,
      fav, kidF, freezF, mealP,
      diet1, diet2, diet3,
      allerg1, allerg2, allerg3,
      `=IFERROR(SUMIFS('Recipe Ingredients'!$K$8:$K$12007,'Recipe Ingredients'!$B$8:$B$12007,A${r}),0)`, // Y: Est Cost
      `=IFERROR(Y${r}/H${r},0)`, // Z: Cost/serving
      source, url, dateAdded, lastCooked, timesCkd, status, testNotes, genNotes,
    ];
  });

  // Style data rows alternating
  RECIPES.forEach((_, i) => {
    const rowIdx = 7 + i; // 0-indexed
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(SID, rowIdx, rowIdx+1, 0, 34),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
  });

  // Style formula cells (A, K, Y, Z) differently
  RECIPES.forEach((_, i) => {
    const rowIdx = 7 + i;
    [0, 10, 24, 25].forEach(ci => {
      reqs.push({ repeatCell: {
        range: gridRange(SID, rowIdx, rowIdx+1, ci, ci+1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.formula),
          textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.primaryDeep) },
          verticalAlignment: 'MIDDLE',
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      } });
    });
  });

  // Number formats
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 2007, 24, 26), // Y:Z Cost
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00;[Red]-"$"#,##0.00' } } },
    fields: 'userEnteredFormat(numberFormat)',
  } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 2007, 28, 31), // Date cols
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } },
    fields: 'userEnteredFormat(numberFormat)',
  } });

  // Row height for data rows
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 2007 },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  } });

  await batchUpdate(id, reqs, 'mri-fmt');
  await valuesBatchUpdate(id, vals, 'mri-vals');
  await valuesBatchUpdate(id, [{ range: `${S}!A8`, values: dataRows }], 'mri-data');
  console.log('✓ Master Recipe Index — 60 recipes written');
})().catch(e => { console.error(e.message || e); process.exit(1); });
