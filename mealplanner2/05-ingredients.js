'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recipe Ingredients'];
const S   = "'Recipe Ingredients'";
const REF = "'Reference Data'";
const RB  = "'Recipe Book'";

const RI1 = 6;
const RI0 = RI1 - 1;

const COLS = ['Ingredient\nLine ID','Recipe ID','Recipe Name','Ingredient Name','Grocery Category',
  'Base Qty','Unit','Unit Cost','Base\nIngredient Cost','Optional?','Pantry\nStaple?',
  'Dietary Tag','Allergen Tag','Notes'];

// [recipeId 1-indexed, ingredientName, groceryCat, baseQty, unit, unitCost, optional, pantryStaple, dietTag, allergenTag, notes]
// recipeId corresponds to row offset (REC-0001=recipe[0], etc)
const ingredients = [
  // REC-0001 Classic Pancakes
  ['REC-0001','All-purpose flour','Baking',1,'cup',0.25,false,true,'Vegetarian','Wheat',''],
  ['REC-0001','Eggs','Dairy & Eggs',2,'piece',0.30,false,false,'Vegetarian','Eggs',''],
  ['REC-0001','Milk','Dairy & Eggs',1,'cup',0.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0001','Butter','Dairy & Eggs',2,'tbsp',0.15,false,true,'Vegetarian','Dairy',''],
  ['REC-0001','Baking powder','Baking',1,'tsp',0.05,false,true,'Vegetarian','None',''],
  ['REC-0001','Salt','Spices & Seasonings',0.5,'tsp',0.02,false,true,'Vegetarian','None',''],
  ['REC-0001','Sugar','Baking',1,'tbsp',0.05,false,true,'Vegetarian','None',''],
  // REC-0002 Avocado Toast with Eggs
  ['REC-0002','Bread','Bakery',2,'slice',0.40,false,false,'Vegetarian','Wheat','Sourdough preferred'],
  ['REC-0002','Avocado','Produce',1,'piece',1.20,false,false,'Vegan','None',''],
  ['REC-0002','Eggs','Dairy & Eggs',2,'piece',0.30,false,false,'Vegetarian','Eggs',''],
  ['REC-0002','Lemon juice','Produce',0.5,'tbsp',0.10,false,false,'Vegan','None',''],
  ['REC-0002','Red pepper flakes','Spices & Seasonings',0.25,'tsp',0.02,true,true,'Vegan','None',''],
  ['REC-0002','Salt','Spices & Seasonings',0.25,'tsp',0.02,false,true,'Vegan','None',''],
  // REC-0003 Overnight Oats
  ['REC-0003','Rolled oats','Grains & Pasta',0.5,'cup',0.20,false,true,'Vegetarian','None','Certified GF if needed'],
  ['REC-0003','Milk','Dairy & Eggs',0.5,'cup',0.10,false,false,'Vegetarian','Dairy',''],
  ['REC-0003','Greek yogurt','Dairy & Eggs',0.25,'cup',0.30,false,false,'Vegetarian','Dairy',''],
  ['REC-0003','Chia seeds','Pantry',1,'tbsp',0.20,false,false,'Vegan','None',''],
  ['REC-0003','Honey','Pantry',1,'tbsp',0.15,false,true,'Vegetarian','None',''],
  ['REC-0003','Vanilla extract','Baking',0.5,'tsp',0.10,false,true,'Vegetarian','None',''],
  ['REC-0003','Mixed berries','Produce',0.5,'cup',0.50,true,false,'Vegan','None','Fresh or frozen'],
  // REC-0004 Berry Smoothie Bowl
  ['REC-0004','Frozen mixed berries','Frozen',1,'cup',1.00,false,false,'Vegan','None',''],
  ['REC-0004','Banana','Produce',1,'piece',0.20,false,false,'Vegan','None','Frozen preferred'],
  ['REC-0004','Granola','Pantry',0.5,'cup',0.40,false,false,'Vegetarian','None',''],
  ['REC-0004','Coconut milk','Canned Goods',0.25,'cup',0.30,false,false,'Vegan','None',''],
  ['REC-0004','Fresh fruit','Produce',0.5,'cup',0.60,true,false,'Vegan','None','Topping'],
  // REC-0005 Vegetable Frittata
  ['REC-0005','Eggs','Dairy & Eggs',6,'piece',0.30,false,false,'Vegetarian','Eggs',''],
  ['REC-0005','Bell pepper','Produce',1,'piece',0.80,false,false,'Vegan','None',''],
  ['REC-0005','Spinach','Produce',2,'cup',0.50,false,false,'Vegan','None',''],
  ['REC-0005','Onion','Produce',0.5,'piece',0.20,false,true,'Vegan','None',''],
  ['REC-0005','Cherry tomatoes','Produce',0.5,'cup',0.60,false,false,'Vegan','None',''],
  ['REC-0005','Feta cheese','Dairy & Eggs',0.5,'cup',1.00,true,false,'Vegetarian','Dairy',''],
  ['REC-0005','Olive oil','Pantry',1,'tbsp',0.15,false,true,'Vegan','None',''],
  // REC-0006 French Toast Bake
  ['REC-0006','Brioche bread','Bakery',1,'lb',3.50,false,false,'Vegetarian','Wheat',''],
  ['REC-0006','Eggs','Dairy & Eggs',4,'piece',0.30,false,false,'Vegetarian','Eggs',''],
  ['REC-0006','Milk','Dairy & Eggs',1,'cup',0.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0006','Maple syrup','Pantry',3,'tbsp',0.60,false,true,'Vegetarian','None',''],
  ['REC-0006','Vanilla extract','Baking',1,'tsp',0.10,false,true,'Vegetarian','None',''],
  ['REC-0006','Cinnamon','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  ['REC-0006','Butter','Dairy & Eggs',2,'tbsp',0.15,false,true,'Vegetarian','Dairy',''],
  // REC-0007 Granola with Yogurt
  ['REC-0007','Greek yogurt','Dairy & Eggs',1,'cup',0.60,false,false,'Vegetarian','Dairy',''],
  ['REC-0007','Granola','Pantry',0.5,'cup',0.40,false,false,'Vegetarian','None',''],
  ['REC-0007','Mixed berries','Produce',0.5,'cup',0.50,false,false,'Vegan','None',''],
  ['REC-0007','Honey','Pantry',1,'tsp',0.05,true,true,'Vegetarian','None',''],
  // REC-0008 Banana Oat Muffins
  ['REC-0008','Bananas','Produce',3,'piece',0.20,false,false,'Vegan','None','Very ripe'],
  ['REC-0008','Rolled oats','Grains & Pasta',2,'cup',0.40,false,true,'Vegetarian','None',''],
  ['REC-0008','Eggs','Dairy & Eggs',2,'piece',0.30,false,false,'Vegetarian','Eggs',''],
  ['REC-0008','Maple syrup','Pantry',3,'tbsp',0.60,false,true,'Vegetarian','None',''],
  ['REC-0008','Baking powder','Baking',1,'tsp',0.05,false,true,'Vegetarian','None',''],
  ['REC-0008','Vanilla extract','Baking',1,'tsp',0.10,false,true,'Vegetarian','None',''],
  ['REC-0008','Chocolate chips','Baking',0.5,'cup',0.80,true,false,'Vegetarian','Dairy','Optional mix-in'],
  // REC-0009 Caesar Salad
  ['REC-0009','Romaine lettuce','Produce',1,'bunch',1.50,false,false,'Vegetarian','None',''],
  ['REC-0009','Caesar dressing','Condiments & Sauces',0.25,'cup',0.80,false,false,'Vegetarian','Dairy',''],
  ['REC-0009','Parmesan cheese','Dairy & Eggs',0.5,'cup',1.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0009','Croutons','Bakery',1,'cup',0.60,false,false,'Vegetarian','Wheat',''],
  ['REC-0009','Lemon juice','Produce',1,'tbsp',0.10,false,false,'Vegan','None',''],
  // REC-0010 Turkey & Avocado Wrap
  ['REC-0010','Flour tortilla','Bakery',2,'piece',0.40,false,false,'','Wheat',''],
  ['REC-0010','Turkey breast','Meat & Seafood',4,'oz',1.50,false,false,'','None','Deli sliced'],
  ['REC-0010','Avocado','Produce',1,'piece',1.20,false,false,'Vegan','None',''],
  ['REC-0010','Romaine lettuce','Produce',1,'cup',0.30,false,false,'Vegan','None',''],
  ['REC-0010','Tomato','Produce',0.5,'piece',0.30,false,false,'Vegan','None',''],
  ['REC-0010','Dijon mustard','Condiments & Sauces',1,'tbsp',0.10,true,true,'Vegan','None',''],
  // REC-0011 Lentil Soup
  ['REC-0011','Red lentils','Pantry',1.5,'cup',0.60,false,true,'Vegan','None',''],
  ['REC-0011','Onion','Produce',1,'piece',0.40,false,true,'Vegan','None',''],
  ['REC-0011','Garlic','Produce',3,'clove',0.15,false,true,'Vegan','None',''],
  ['REC-0011','Canned tomatoes','Canned Goods',1,'can',0.90,false,false,'Vegan','None',''],
  ['REC-0011','Vegetable broth','Canned Goods',4,'cup',1.00,false,false,'Vegan','None',''],
  ['REC-0011','Cumin','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  ['REC-0011','Olive oil','Pantry',2,'tbsp',0.30,false,true,'Vegan','None',''],
  ['REC-0011','Lemon juice','Produce',2,'tbsp',0.15,false,false,'Vegan','None',''],
  // REC-0012 Caprese Panini
  ['REC-0012','Ciabatta bread','Bakery',1,'piece',1.50,false,false,'Vegetarian','Wheat',''],
  ['REC-0012','Fresh mozzarella','Dairy & Eggs',4,'oz',2.00,false,false,'Vegetarian','Dairy',''],
  ['REC-0012','Tomato','Produce',1,'piece',0.60,false,false,'Vegan','None',''],
  ['REC-0012','Fresh basil','Produce',8,'piece',0.20,false,false,'Vegan','None',''],
  ['REC-0012','Olive oil','Pantry',1,'tbsp',0.15,false,true,'Vegan','None',''],
  // REC-0013 Chicken Noodle Soup
  ['REC-0013','Chicken breast','Meat & Seafood',1,'lb',3.50,false,false,'','None',''],
  ['REC-0013','Egg noodles','Grains & Pasta',2,'cup',0.60,false,false,'','Wheat',''],
  ['REC-0013','Carrots','Produce',3,'piece',0.30,false,true,'Vegan','None',''],
  ['REC-0013','Celery','Produce',3,'piece',0.20,false,false,'Vegan','None',''],
  ['REC-0013','Onion','Produce',1,'piece',0.40,false,true,'Vegan','None',''],
  ['REC-0013','Chicken broth','Canned Goods',6,'cup',1.50,false,false,'','None',''],
  ['REC-0013','Garlic','Produce',2,'clove',0.10,false,true,'Vegan','None',''],
  ['REC-0013','Thyme','Spices & Seasonings',0.5,'tsp',0.05,false,true,'Vegan','None',''],
  // REC-0014 Greek Salad
  ['REC-0014','Cucumber','Produce',1,'piece',0.60,false,false,'Vegan','None',''],
  ['REC-0014','Cherry tomatoes','Produce',1,'cup',1.00,false,false,'Vegan','None',''],
  ['REC-0014','Kalamata olives','Canned Goods',0.5,'cup',1.20,false,false,'Vegan','None',''],
  ['REC-0014','Feta cheese','Dairy & Eggs',4,'oz',1.80,false,false,'Vegetarian','Dairy',''],
  ['REC-0014','Red onion','Produce',0.25,'piece',0.15,false,false,'Vegan','None',''],
  ['REC-0014','Olive oil','Pantry',3,'tbsp',0.45,false,true,'Vegan','None',''],
  ['REC-0014','Dried oregano','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  // REC-0015 BLT Sandwich
  ['REC-0015','Bacon','Meat & Seafood',6,'slice',1.20,false,false,'','None',''],
  ['REC-0015','Bread','Bakery',4,'slice',0.40,false,false,'','Wheat',''],
  ['REC-0015','Romaine lettuce','Produce',2,'cup',0.30,false,false,'Vegan','None',''],
  ['REC-0015','Tomato','Produce',1,'piece',0.60,false,false,'Vegan','None',''],
  ['REC-0015','Mayonnaise','Condiments & Sauces',2,'tbsp',0.20,false,true,'','Eggs',''],
  // REC-0016 Quinoa Power Bowl
  ['REC-0016','Quinoa','Grains & Pasta',0.5,'cup',0.80,false,false,'Vegan','None',''],
  ['REC-0016','Chickpeas','Canned Goods',1,'can',0.90,false,false,'Vegan','None',''],
  ['REC-0016','Sweet potato','Produce',1,'piece',0.80,false,false,'Vegan','None',''],
  ['REC-0016','Kale','Produce',2,'cup',0.60,false,false,'Vegan','None',''],
  ['REC-0016','Tahini','Pantry',2,'tbsp',0.40,false,false,'Vegan','Sesame',''],
  ['REC-0016','Lemon juice','Produce',2,'tbsp',0.15,false,false,'Vegan','None',''],
  ['REC-0016','Olive oil','Pantry',1,'tbsp',0.15,false,true,'Vegan','None',''],
  // REC-0017 Tomato Basil Pasta
  ['REC-0017','Pasta','Grains & Pasta',0.5,'lb',0.60,false,true,'Vegetarian','Wheat',''],
  ['REC-0017','Cherry tomatoes','Produce',2,'cup',2.00,false,false,'Vegan','None',''],
  ['REC-0017','Garlic','Produce',4,'clove',0.20,false,true,'Vegan','None',''],
  ['REC-0017','Fresh basil','Produce',0.25,'cup',0.30,false,false,'Vegan','None',''],
  ['REC-0017','Olive oil','Pantry',3,'tbsp',0.45,false,true,'Vegan','None',''],
  ['REC-0017','Parmesan cheese','Dairy & Eggs',0.5,'cup',1.20,true,false,'Vegetarian','Dairy',''],
  // REC-0018 Sheet Pan Chicken & Vegetables
  ['REC-0018','Chicken thighs','Meat & Seafood',1.5,'lb',4.00,false,false,'Gluten-Free','None','Bone-in preferred'],
  ['REC-0018','Broccoli','Produce',2,'cup',1.00,false,false,'Vegan','None',''],
  ['REC-0018','Bell pepper','Produce',2,'piece',1.60,false,false,'Vegan','None',''],
  ['REC-0018','Zucchini','Produce',1,'piece',0.70,false,false,'Vegan','None',''],
  ['REC-0018','Olive oil','Pantry',3,'tbsp',0.45,false,true,'Vegan','None',''],
  ['REC-0018','Garlic powder','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  ['REC-0018','Italian seasoning','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  // REC-0019 Beef Tacos
  ['REC-0019','Ground beef','Meat & Seafood',1,'lb',5.00,false,false,'','None',''],
  ['REC-0019','Taco seasoning','Spices & Seasonings',1,'package',0.80,false,false,'','None',''],
  ['REC-0019','Corn tortillas','Bakery',8,'piece',0.80,false,false,'Gluten-Free','None',''],
  ['REC-0019','Cheddar cheese','Dairy & Eggs',1,'cup',1.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0019','Sour cream','Dairy & Eggs',0.25,'cup',0.40,false,false,'Vegetarian','Dairy',''],
  ['REC-0019','Salsa','Condiments & Sauces',0.5,'cup',0.60,false,false,'Vegan','None',''],
  ['REC-0019','Romaine lettuce','Produce',1,'cup',0.30,false,false,'Vegan','None','Shredded'],
  // REC-0020 Vegetable Stir Fry
  ['REC-0020','Broccoli','Produce',2,'cup',1.00,false,false,'Vegan','None',''],
  ['REC-0020','Snap peas','Produce',1,'cup',1.20,false,false,'Vegan','None',''],
  ['REC-0020','Carrots','Produce',2,'piece',0.20,false,true,'Vegan','None','Julienned'],
  ['REC-0020','Bell pepper','Produce',1,'piece',0.80,false,false,'Vegan','None',''],
  ['REC-0020','Soy sauce','Condiments & Sauces',3,'tbsp',0.20,false,true,'Vegan','Soy','Use tamari for GF'],
  ['REC-0020','Sesame oil','Condiments & Sauces',1,'tbsp',0.25,false,true,'Vegan','Sesame',''],
  ['REC-0020','Ginger','Produce',1,'tsp',0.10,false,false,'Vegan','None','Fresh grated'],
  ['REC-0020','Garlic','Produce',2,'clove',0.10,false,true,'Vegan','None',''],
  ['REC-0020','Rice','Grains & Pasta',2,'cup',0.40,false,true,'Vegan','None','For serving'],
  // REC-0021 Baked Salmon with Lemon
  ['REC-0021','Salmon fillet','Meat & Seafood',1.5,'lb',9.00,false,false,'Gluten-Free','Fish',''],
  ['REC-0021','Lemon','Produce',1,'piece',0.50,false,false,'Vegan','None',''],
  ['REC-0021','Garlic','Produce',2,'clove',0.10,false,true,'Vegan','None',''],
  ['REC-0021','Fresh dill','Produce',2,'tbsp',0.30,true,false,'Vegan','None',''],
  ['REC-0021','Olive oil','Pantry',2,'tbsp',0.30,false,true,'Vegan','None',''],
  ['REC-0021','Salt','Spices & Seasonings',0.5,'tsp',0.02,false,true,'Vegan','None',''],
  ['REC-0021','Black pepper','Spices & Seasonings',0.25,'tsp',0.02,false,true,'Vegan','None',''],
  // REC-0022 Pasta Bolognese
  ['REC-0022','Pasta','Grains & Pasta',1,'lb',1.20,false,true,'','Wheat',''],
  ['REC-0022','Ground beef','Meat & Seafood',1,'lb',5.00,false,false,'','None',''],
  ['REC-0022','Canned crushed tomatoes','Canned Goods',1,'can',1.20,false,false,'Vegan','None',''],
  ['REC-0022','Onion','Produce',1,'piece',0.40,false,true,'Vegan','None',''],
  ['REC-0022','Garlic','Produce',3,'clove',0.15,false,true,'Vegan','None',''],
  ['REC-0022','Red wine','Beverages',0.25,'cup',0.80,true,false,'Vegan','None','Optional'],
  ['REC-0022','Parmesan cheese','Dairy & Eggs',0.5,'cup',1.20,false,false,'Vegetarian','Dairy','For serving'],
  ['REC-0022','Italian seasoning','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  // REC-0023 Chicken Tikka Masala
  ['REC-0023','Chicken breast','Meat & Seafood',1.5,'lb',5.25,false,false,'Gluten-Free','None',''],
  ['REC-0023','Greek yogurt','Dairy & Eggs',0.5,'cup',0.60,false,false,'Vegetarian','Dairy','For marinade'],
  ['REC-0023','Canned tomatoes','Canned Goods',1,'can',0.90,false,false,'Vegan','None',''],
  ['REC-0023','Heavy cream','Dairy & Eggs',0.5,'cup',1.00,false,false,'Vegetarian','Dairy',''],
  ['REC-0023','Garam masala','Spices & Seasonings',2,'tsp',0.10,false,true,'Vegan','None',''],
  ['REC-0023','Garlic','Produce',4,'clove',0.20,false,true,'Vegan','None',''],
  ['REC-0023','Ginger','Produce',1,'tbsp',0.15,false,false,'Vegan','None','Fresh'],
  ['REC-0023','Onion','Produce',1,'piece',0.40,false,true,'Vegan','None',''],
  ['REC-0023','Olive oil','Pantry',2,'tbsp',0.30,false,true,'Vegan','None',''],
  // REC-0024 Vegetable Curry
  ['REC-0024','Coconut milk','Canned Goods',1,'can',1.50,false,false,'Vegan','None',''],
  ['REC-0024','Chickpeas','Canned Goods',1,'can',0.90,false,false,'Vegan','None',''],
  ['REC-0024','Sweet potato','Produce',1,'piece',0.80,false,false,'Vegan','None',''],
  ['REC-0024','Spinach','Produce',2,'cup',0.50,false,false,'Vegan','None',''],
  ['REC-0024','Red curry paste','Condiments & Sauces',2,'tbsp',0.60,false,false,'Vegan','None',''],
  ['REC-0024','Garlic','Produce',2,'clove',0.10,false,true,'Vegan','None',''],
  ['REC-0024','Ginger','Produce',1,'tsp',0.10,false,false,'Vegan','None',''],
  ['REC-0024','Rice','Grains & Pasta',2,'cup',0.40,false,true,'Vegan','None','For serving'],
  // REC-0025 Mushroom Risotto
  ['REC-0025','Arborio rice','Grains & Pasta',1.5,'cup',2.00,false,false,'Vegetarian','None',''],
  ['REC-0025','Mixed mushrooms','Produce',1,'lb',3.00,false,false,'Vegan','None',''],
  ['REC-0025','Vegetable broth','Canned Goods',4,'cup',1.00,false,false,'Vegan','None','Warm'],
  ['REC-0025','Parmesan cheese','Dairy & Eggs',1,'cup',2.40,false,false,'Vegetarian','Dairy',''],
  ['REC-0025','Butter','Dairy & Eggs',3,'tbsp',0.23,false,true,'Vegetarian','Dairy',''],
  ['REC-0025','Onion','Produce',1,'piece',0.40,false,true,'Vegan','None',''],
  ['REC-0025','Garlic','Produce',2,'clove',0.10,false,true,'Vegan','None',''],
  ['REC-0025','White wine','Beverages',0.5,'cup',1.00,false,false,'Vegan','None',''],
  // REC-0026 BBQ Pulled Pork
  ['REC-0026','Pork shoulder','Meat & Seafood',3,'lb',9.00,false,false,'Gluten-Free','None',''],
  ['REC-0026','BBQ sauce','Condiments & Sauces',1,'cup',1.50,false,false,'Vegan','None',''],
  ['REC-0026','Brown sugar','Baking',2,'tbsp',0.10,false,true,'Vegan','None',''],
  ['REC-0026','Smoked paprika','Spices & Seasonings',1,'tbsp',0.10,false,true,'Vegan','None',''],
  ['REC-0026','Garlic powder','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  ['REC-0026','Onion','Produce',1,'piece',0.40,false,true,'Vegan','None',''],
  ['REC-0026','Hamburger buns','Bakery',8,'piece',1.60,false,false,'Vegan','Wheat','For serving'],
  // REC-0027 Shrimp Tacos
  ['REC-0027','Shrimp','Meat & Seafood',1,'lb',8.00,false,false,'','Shellfish','Peeled and deveined'],
  ['REC-0027','Corn tortillas','Bakery',8,'piece',0.80,false,false,'Gluten-Free','None',''],
  ['REC-0027','Cabbage','Produce',2,'cup',0.50,false,false,'Vegan','None','Shredded'],
  ['REC-0027','Sour cream','Dairy & Eggs',0.25,'cup',0.40,false,false,'Vegetarian','Dairy',''],
  ['REC-0027','Lime','Produce',2,'piece',0.40,false,false,'Vegan','None',''],
  ['REC-0027','Cilantro','Produce',0.25,'cup',0.30,true,false,'Vegan','None',''],
  ['REC-0027','Chili powder','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  // REC-0028 Eggplant Parmesan
  ['REC-0028','Eggplant','Produce',1,'piece',1.20,false,false,'Vegetarian','None',''],
  ['REC-0028','Marinara sauce','Canned Goods',2,'cup',1.80,false,false,'Vegan','None',''],
  ['REC-0028','Mozzarella cheese','Dairy & Eggs',2,'cup',3.00,false,false,'Vegetarian','Dairy',''],
  ['REC-0028','Parmesan cheese','Dairy & Eggs',0.5,'cup',1.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0028','Breadcrumbs','Bakery',1,'cup',0.50,false,false,'Vegetarian','Wheat',''],
  ['REC-0028','Eggs','Dairy & Eggs',2,'piece',0.30,false,false,'Vegetarian','Eggs','For breading'],
  ['REC-0028','Olive oil','Pantry',3,'tbsp',0.45,false,true,'Vegan','None',''],
  // REC-0029 Chicken Caesar Pasta
  ['REC-0029','Pasta','Grains & Pasta',0.5,'lb',0.60,false,true,'','Wheat',''],
  ['REC-0029','Chicken breast','Meat & Seafood',1,'lb',3.50,false,false,'','None','Grilled'],
  ['REC-0029','Caesar dressing','Condiments & Sauces',0.25,'cup',0.80,false,false,'Vegetarian','Dairy',''],
  ['REC-0029','Parmesan cheese','Dairy & Eggs',0.5,'cup',1.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0029','Romaine lettuce','Produce',2,'cup',0.30,false,false,'Vegan','None',''],
  // REC-0030 Black Bean Enchiladas
  ['REC-0030','Black beans','Canned Goods',2,'can',1.80,false,true,'Vegan','None','Drained'],
  ['REC-0030','Flour tortillas','Bakery',8,'piece',1.60,false,false,'Vegetarian','Wheat',''],
  ['REC-0030','Enchilada sauce','Canned Goods',1,'can',1.50,false,false,'Vegan','None',''],
  ['REC-0030','Cheddar cheese','Dairy & Eggs',2,'cup',2.40,false,false,'Vegetarian','Dairy',''],
  ['REC-0030','Sour cream','Dairy & Eggs',0.5,'cup',0.80,true,false,'Vegetarian','Dairy','For serving'],
  ['REC-0030','Cumin','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  // REC-0031 Hummus with Veggies
  ['REC-0031','Hummus','Pantry',1,'cup',1.50,false,false,'Vegan','Sesame','Store-bought OK'],
  ['REC-0031','Carrots','Produce',2,'piece',0.20,false,true,'Vegan','None',''],
  ['REC-0031','Celery','Produce',3,'piece',0.20,false,false,'Vegan','None',''],
  ['REC-0031','Cucumber','Produce',1,'piece',0.60,false,false,'Vegan','None',''],
  ['REC-0031','Bell pepper','Produce',1,'piece',0.80,false,false,'Vegan','None',''],
  // REC-0032 Guacamole & Chips
  ['REC-0032','Avocado','Produce',3,'piece',3.60,false,false,'Vegan','None','Ripe'],
  ['REC-0032','Lime','Produce',1,'piece',0.20,false,false,'Vegan','None',''],
  ['REC-0032','Cilantro','Produce',2,'tbsp',0.10,true,false,'Vegan','None',''],
  ['REC-0032','Red onion','Produce',0.25,'piece',0.15,false,false,'Vegan','None',''],
  ['REC-0032','Tomato','Produce',1,'piece',0.60,false,false,'Vegan','None',''],
  ['REC-0032','Tortilla chips','Snacks',1,'bag',2.00,false,false,'Vegan','None',''],
  // REC-0033 Apple Slices with Peanut Butter — ALLERGEN EXAMPLE (tree nuts/peanuts)
  ['REC-0033','Apple','Produce',2,'piece',0.60,false,false,'Vegan','None',''],
  ['REC-0033','Peanut butter','Pantry',4,'tbsp',0.40,false,false,'Vegan','Peanuts','⚠ Contains peanuts — not safe for tree-nut-allergic household members without substitution'],
  // REC-0034 Cheese & Crackers
  ['REC-0034','Assorted cheese','Dairy & Eggs',8,'oz',4.00,false,false,'Vegetarian','Dairy',''],
  ['REC-0034','Crackers','Snacks',4,'oz',1.00,false,false,'Vegetarian','Wheat',''],
  ['REC-0034','Grapes','Produce',1,'cup',1.00,true,false,'Vegan','None',''],
  // REC-0035 Energy Bites
  ['REC-0035','Rolled oats','Grains & Pasta',1,'cup',0.20,false,true,'Vegetarian','None',''],
  ['REC-0035','Peanut butter','Pantry',0.5,'cup',0.80,false,false,'Vegan','Peanuts','⚠ Contains peanuts'],
  ['REC-0035','Honey','Pantry',3,'tbsp',0.45,false,true,'Vegetarian','None',''],
  ['REC-0035','Chocolate chips','Baking',0.5,'cup',0.80,true,false,'Vegetarian','Dairy',''],
  ['REC-0035','Chia seeds','Pantry',1,'tbsp',0.20,true,false,'Vegan','None',''],
  // REC-0036 Caprese Skewers
  ['REC-0036','Fresh mozzarella','Dairy & Eggs',8,'oz',2.00,false,false,'Vegetarian','Dairy',''],
  ['REC-0036','Cherry tomatoes','Produce',1,'cup',1.00,false,false,'Vegan','None',''],
  ['REC-0036','Fresh basil','Produce',12,'piece',0.20,false,false,'Vegan','None',''],
  ['REC-0036','Balsamic glaze','Condiments & Sauces',2,'tbsp',0.30,false,false,'Vegan','None',''],
  // REC-0037 Chocolate Chip Cookies
  ['REC-0037','All-purpose flour','Baking',2.25,'cup',0.56,false,true,'Vegetarian','Wheat',''],
  ['REC-0037','Butter','Dairy & Eggs',1,'cup',0.60,false,true,'Vegetarian','Dairy',''],
  ['REC-0037','Eggs','Dairy & Eggs',2,'piece',0.30,false,false,'Vegetarian','Eggs',''],
  ['REC-0037','Chocolate chips','Baking',2,'cup',3.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0037','Brown sugar','Baking',0.75,'cup',0.38,false,true,'Vegetarian','None',''],
  ['REC-0037','Sugar','Baking',0.75,'cup',0.38,false,true,'Vegetarian','None',''],
  ['REC-0037','Vanilla extract','Baking',2,'tsp',0.20,false,true,'Vegetarian','None',''],
  ['REC-0037','Baking soda','Baking',1,'tsp',0.05,false,true,'Vegetarian','None',''],
  // REC-0038 Fruit Salad
  ['REC-0038','Strawberries','Produce',1,'cup',1.50,false,false,'Vegan','None',''],
  ['REC-0038','Blueberries','Produce',1,'cup',1.80,false,false,'Vegan','None',''],
  ['REC-0038','Grapes','Produce',1,'cup',1.00,false,false,'Vegan','None',''],
  ['REC-0038','Watermelon','Produce',2,'cup',0.80,false,false,'Vegan','None','Cubed'],
  ['REC-0038','Lime juice','Produce',1,'tbsp',0.10,false,false,'Vegan','None',''],
  ['REC-0038','Honey','Pantry',1,'tbsp',0.15,true,true,'Vegetarian','None',''],
  // REC-0039 Brownies
  ['REC-0039','Butter','Dairy & Eggs',0.5,'cup',0.30,false,true,'Vegetarian','Dairy',''],
  ['REC-0039','Dark chocolate','Baking',4,'oz',2.00,false,false,'Vegetarian','Dairy',''],
  ['REC-0039','Sugar','Baking',1,'cup',0.50,false,true,'Vegetarian','None',''],
  ['REC-0039','Eggs','Dairy & Eggs',2,'piece',0.30,false,false,'Vegetarian','Eggs',''],
  ['REC-0039','All-purpose flour','Baking',0.5,'cup',0.13,false,true,'Vegetarian','Wheat',''],
  ['REC-0039','Cocoa powder','Baking',0.25,'cup',0.50,false,false,'Vegetarian','None',''],
  ['REC-0039','Vanilla extract','Baking',1,'tsp',0.10,false,true,'Vegetarian','None',''],
  // REC-0040 Rice Pudding
  ['REC-0040','Rice','Grains & Pasta',0.5,'cup',0.10,false,true,'Vegetarian','None',''],
  ['REC-0040','Milk','Dairy & Eggs',2,'cup',0.40,false,false,'Vegetarian','Dairy',''],
  ['REC-0040','Sugar','Baking',3,'tbsp',0.15,false,true,'Vegetarian','None',''],
  ['REC-0040','Vanilla extract','Baking',1,'tsp',0.10,false,true,'Vegetarian','None',''],
  ['REC-0040','Cinnamon','Spices & Seasonings',0.5,'tsp',0.05,true,true,'Vegan','None',''],
  // REC-0041 Banana Ice Cream
  ['REC-0041','Banana','Produce',4,'piece',0.80,false,false,'Vegan','None','Frozen'],
  ['REC-0041','Vanilla extract','Baking',0.5,'tsp',0.05,true,true,'Vegan','None',''],
  ['REC-0041','Chocolate chips','Baking',0.25,'cup',0.40,true,false,'Vegetarian','Dairy','Optional topping'],
  // REC-0042 Minestrone Soup
  ['REC-0042','Cannellini beans','Canned Goods',1,'can',0.90,false,false,'Vegan','None',''],
  ['REC-0042','Diced tomatoes','Canned Goods',1,'can',0.90,false,false,'Vegan','None',''],
  ['REC-0042','Pasta','Grains & Pasta',1,'cup',0.30,false,true,'Vegan','Wheat','Small pasta'],
  ['REC-0042','Zucchini','Produce',1,'piece',0.70,false,false,'Vegan','None',''],
  ['REC-0042','Spinach','Produce',2,'cup',0.50,false,false,'Vegan','None',''],
  ['REC-0042','Vegetable broth','Canned Goods',4,'cup',1.00,false,false,'Vegan','None',''],
  ['REC-0042','Onion','Produce',1,'piece',0.40,false,true,'Vegan','None',''],
  ['REC-0042','Garlic','Produce',2,'clove',0.10,false,true,'Vegan','None',''],
  // REC-0043 Grilled Cheese Sandwich
  ['REC-0043','Bread','Bakery',4,'slice',0.40,false,false,'Vegetarian','Wheat','Sourdough preferred'],
  ['REC-0043','Cheddar cheese','Dairy & Eggs',4,'oz',1.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0043','Butter','Dairy & Eggs',2,'tbsp',0.15,false,true,'Vegetarian','Dairy',''],
  // REC-0044 Shakshuka
  ['REC-0044','Canned tomatoes','Canned Goods',1,'can',0.90,false,false,'Vegan','None',''],
  ['REC-0044','Eggs','Dairy & Eggs',4,'piece',0.30,false,false,'Vegetarian','Eggs',''],
  ['REC-0044','Onion','Produce',1,'piece',0.40,false,true,'Vegan','None',''],
  ['REC-0044','Bell pepper','Produce',1,'piece',0.80,false,false,'Vegan','None',''],
  ['REC-0044','Garlic','Produce',3,'clove',0.15,false,true,'Vegan','None',''],
  ['REC-0044','Cumin','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  ['REC-0044','Paprika','Spices & Seasonings',1,'tsp',0.05,false,true,'Vegan','None',''],
  ['REC-0044','Feta cheese','Dairy & Eggs',0.5,'cup',1.00,true,false,'Vegetarian','Dairy','Optional topping'],
  // REC-0045 Loaded Sweet Potato
  ['REC-0045','Sweet potato','Produce',4,'piece',3.20,false,false,'Vegetarian','None',''],
  ['REC-0045','Black beans','Canned Goods',1,'can',0.90,false,true,'Vegan','None',''],
  ['REC-0045','Cheddar cheese','Dairy & Eggs',1,'cup',1.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0045','Sour cream','Dairy & Eggs',0.25,'cup',0.40,false,false,'Vegetarian','Dairy',''],
  ['REC-0045','Salsa','Condiments & Sauces',0.5,'cup',0.60,false,false,'Vegan','None',''],
  ['REC-0045','Olive oil','Pantry',1,'tbsp',0.15,false,true,'Vegan','None',''],
  // REC-0046 Thai Peanut Noodles — ALLERGEN EXAMPLE
  ['REC-0046','Soba noodles','Grains & Pasta',8,'oz',2.00,false,false,'Vegan','Wheat',''],
  ['REC-0046','Peanut butter','Pantry',4,'tbsp',0.80,false,false,'Vegan','Peanuts','⚠ Contains peanuts'],
  ['REC-0046','Soy sauce','Condiments & Sauces',3,'tbsp',0.20,false,true,'Vegan','Soy',''],
  ['REC-0046','Soba noodles','Grains & Pasta',8,'oz',2.00,false,false,'Vegan','Wheat','Duplicate unit test'],
  ['REC-0046','Rice noodles','Grains & Pasta',8,'oz',2.00,false,false,'Vegan','None','Review Units — incompatible unit with soba'],
  ['REC-0046','Lime juice','Produce',2,'tbsp',0.15,false,false,'Vegan','None',''],
  ['REC-0046','Cucumber','Produce',1,'piece',0.60,false,false,'Vegan','None',''],
  ['REC-0046','Carrot','Produce',1,'piece',0.10,false,true,'Vegan','None','Shredded'],
  // REC-0047 White Bean & Kale Soup
  ['REC-0047','White beans','Canned Goods',2,'can',1.80,false,true,'Vegan','None',''],
  ['REC-0047','Kale','Produce',3,'cup',0.90,false,false,'Vegan','None',''],
  ['REC-0047','Vegetable broth','Canned Goods',4,'cup',1.00,false,false,'Vegan','None',''],
  ['REC-0047','Onion','Produce',1,'piece',0.40,false,true,'Vegan','None',''],
  ['REC-0047','Garlic','Produce',3,'clove',0.15,false,true,'Vegan','None',''],
  ['REC-0047','Rosemary','Spices & Seasonings',0.5,'tsp',0.05,true,true,'Vegan','None',''],
  // REC-0048 Tuna Pasta Salad
  ['REC-0048','Pasta','Grains & Pasta',0.5,'lb',0.60,false,true,'','Wheat',''],
  ['REC-0048','Canned tuna','Canned Goods',2,'can',2.00,false,false,'','Fish',''],
  ['REC-0048','Celery','Produce',2,'piece',0.13,false,false,'Vegan','None',''],
  ['REC-0048','Mayonnaise','Condiments & Sauces',3,'tbsp',0.30,false,true,'','Eggs',''],
  ['REC-0048','Lemon juice','Produce',1,'tbsp',0.10,false,false,'Vegan','None',''],
  ['REC-0048','Dijon mustard','Condiments & Sauces',1,'tsp',0.03,false,true,'Vegan','None',''],
  // REC-0049 Homemade Pizza
  ['REC-0049','Pizza dough','Bakery',1,'lb',1.80,false,false,'Vegetarian','Wheat',''],
  ['REC-0049','Marinara sauce','Canned Goods',0.5,'cup',0.45,false,false,'Vegan','None',''],
  ['REC-0049','Mozzarella cheese','Dairy & Eggs',2,'cup',3.00,false,false,'Vegetarian','Dairy',''],
  ['REC-0049','Bell pepper','Produce',1,'piece',0.80,true,false,'Vegan','None',''],
  ['REC-0049','Mushrooms','Produce',1,'cup',1.00,true,false,'Vegan','None',''],
  ['REC-0049','Olives','Canned Goods',0.25,'cup',0.60,true,false,'Vegan','None',''],
  // REC-0050 Oatmeal with Toppings
  ['REC-0050','Rolled oats','Grains & Pasta',0.5,'cup',0.10,false,true,'Vegetarian','None',''],
  ['REC-0050','Milk','Dairy & Eggs',1,'cup',0.20,false,false,'Vegetarian','Dairy',''],
  ['REC-0050','Banana','Produce',0.5,'piece',0.10,false,false,'Vegan','None',''],
  ['REC-0050','Mixed berries','Produce',0.25,'cup',0.25,true,false,'Vegan','None',''],
  ['REC-0050','Honey','Pantry',1,'tbsp',0.15,false,true,'Vegetarian','None',''],
  ['REC-0050','Walnuts','Pantry',1,'tbsp',0.30,true,false,'Vegan','Tree Nuts','⚠ Tree nuts — check household allergens'],
];

(async () => {
  const fmt = [];
  const vals = [];
  const NI = ingredients.length;
  const NC = COLS.length; // 14

  // ─── Title ────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.peach), textFormat: { bold: true, fontSize: 20, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['🥕 Recipe Ingredients']] });

  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.butter), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A2`, values: [['All recipe ingredients with quantities, costs, and dietary/allergen information']] });

  // Separator + info rows
  for (let r = 2; r < 4; r++) {
    fmt.push({ mergeCells: { range: gridRange(SID, r, r+1, 0, NC), mergeType: 'MERGE_ALL' } });
  }
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.peach) } }, fields: 'userEnteredFormat.backgroundColor' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 3, 4, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.warning), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial' }, horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A4`, values: [['  ⚠ Review allergen tags carefully. Ingredient data does not guarantee safety — always check labels for cross-contamination and substitution needs.']] });

  // Header row
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 4, 5, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A5`, values: [COLS] });

  // ─── Data rows ────────────────────────────────────────────────────────────
  const ingVals = [];
  for (let i = 0; i < NI; i++) {
    const r1 = RI1 + i;
    const [recId, ingName, grocCat, qty, unit, unitCost, optional, pantryStaple, dietTag, allergenTag, notes] = ingredients[i];
    ingVals.push([
      `=IF(OR(B${r1}="",D${r1}=""),"","ING-"&TEXT(ROW()-5,"00000"))`,  // A
      recId,                                                               // B
      `=IFERROR(VLOOKUP(B${r1},'Recipe Book'!$A$6:$B$1005,2,FALSE),"")`, // C
      ingName,                                                             // D
      grocCat,                                                             // E
      qty,                                                                 // F
      unit,                                                                // G
      unitCost,                                                            // H
      `=IFERROR(F${r1}*H${r1},0)`,                                       // I
      optional,                                                            // J
      pantryStaple,                                                        // K
      dietTag,                                                             // L
      allergenTag,                                                         // M
      notes,                                                               // N
    ]);

    const bgColor = i % 2 === 0 ? C.panel : C.altRow;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, RI0+i, RI0+i+1, 0, NC),
        cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 10, fontFamily: 'Arial' } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    // Formula cells: A, C, I
    [0, 2, 8].forEach(col => {
      fmt.push({
        repeatCell: {
          range: gridRange(SID, RI0+i, RI0+i+1, col, col+1),
          cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      });
    });
  }

  // Push all data
  for (let start = 0; start < NI; start += 200) {
    const chunk = ingVals.slice(start, start + 200);
    vals.push({ range: `${S}!A${RI1+start}:N${RI1+start+chunk.length-1}`, values: chunk });
  }

  // Checkboxes J, K (cols 9, 10)
  [9, 10].forEach(col => {
    fmt.push({
      setDataValidation: {
        range: gridRange(SID, RI0, RI0+NI+500, col, col+1),
        rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
      },
    });
  });

  // Dropdown: Grocery Category (E=4), Unit (G=6), Dietary Tag (L=11), Allergen (M=12)
  const refColMap2 = { 4: 'C', 6: 'D', 11: 'F', 12: 'G' };
  Object.entries(refColMap2).forEach(([col, refLetter]) => {
    fmt.push({
      setDataValidation: {
        range: gridRange(SID, RI0, RI0+NI+2000, parseInt(col), parseInt(col)+1),
        rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$${refLetter}$2:$${refLetter}$50` }] }, showCustomUi: true, strict: false },
      },
    });
  });

  // Currency format for H, I
  fmt.push({
    repeatCell: {
      range: gridRange(SID, RI0, RI0+NI+2000, 7, 9),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Column widths
  const colWidths = [80, 80, 160, 160, 120, 70, 80, 80, 90, 70, 80, 110, 110, 200];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Freeze rows 1:5 (no col freeze due to full-row title merges)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'ing-fmt');
  await valuesBatchUpdate(id, vals, 'ing-vals');
  console.log(`✓ Recipe Ingredients complete (${NI} rows)`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
