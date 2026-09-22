'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recipe Ingredients'];
const S = "'Recipe Ingredients'";
const REF = "'Reference Data'";
const MRI = "'Master Recipe Index'";

const HEADERS = [
  'Ingredient Line ID','Recipe ID','Recipe Name','Ingredient Name','Ingredient Category',
  'Base Quantity','Unit','Pkg Quantity','Package Cost','Est. Unit Cost',
  'Ingredient Cost','Optional?','Dietary Tag','Allergen Tag',
  'Substitution','Preparation Note','Notes',
];

// Ingredients: [RecipeID, IngredientName, Category, Qty, Unit, PkgQty, PkgCost, Optional, DietaryTag, AllergenTag, Substitution, PrepNote, Notes]
const INGREDIENTS = [
  // REC-0001 Classic Buttermilk Pancakes (8)
  ['REC-0001','All-Purpose Flour','Baking',2,'cup',5,3.49,false,'Vegetarian','Wheat','GF flour blend','',''],
  ['REC-0001','Buttermilk','Dairy & Eggs',2,'cup',0.5,2.29,false,'Vegetarian','Dairy','Plant milk + 2 tbsp vinegar','',''],
  ['REC-0001','Eggs','Dairy & Eggs',2,'piece',12,3.99,false,'Vegetarian','Eggs','Flax eggs','',''],
  ['REC-0001','Butter','Dairy & Eggs',3,'tbsp',1,4.49,false,'Vegetarian','Dairy','Coconut oil','Melted',''],
  ['REC-0001','Sugar','Baking',2,'tbsp',4,2.49,false,'Vegetarian','','','',''],
  ['REC-0001','Baking Powder','Baking',2,'tsp',8,2.99,false,'Vegetarian','','','',''],
  ['REC-0001','Baking Soda','Baking',0.5,'tsp',8,1.79,false,'Vegetarian','','','',''],
  ['REC-0001','Salt','Spices & Seasonings',0.5,'tsp',26,1.29,false,'Vegetarian','','','',''],
  ['REC-0001','Vanilla Extract','Baking',1,'tsp',4,4.99,true,'Vegetarian','','','','Optional but recommended'],
  // REC-0002 Spinach & Feta Omelette (7)
  ['REC-0002','Eggs','Dairy & Eggs',3,'piece',12,3.99,false,'Vegetarian','Eggs','','',''],
  ['REC-0002','Baby Spinach','Produce',1,'cup',5,3.99,false,'Vegetarian','','','Washed',''],
  ['REC-0002','Feta Cheese','Dairy & Eggs',0.25,'cup',8,4.99,false,'Vegetarian','Dairy','Goat cheese','Crumbled',''],
  ['REC-0002','Olive Oil','Condiments & Sauces',1,'tbsp',17,7.99,false,'Vegetarian','','','',''],
  ['REC-0002','Garlic','Produce',1,'clove',3,0.79,false,'Vegetarian','','','Minced',''],
  ['REC-0002','Salt','Spices & Seasonings',0.25,'tsp',26,1.29,false,'Vegetarian','','','',''],
  ['REC-0002','Black Pepper','Spices & Seasonings',0.25,'tsp',2,2.49,false,'Vegetarian','','','',''],
  // REC-0003 Blueberry Overnight Oats (6)
  ['REC-0003','Rolled Oats','Grains & Pasta',0.5,'cup',40,4.99,false,'Vegetarian','','GF oats if needed','',''],
  ['REC-0003','Almond Milk','Beverages',0.5,'cup',32,3.79,false,'Vegetarian','Tree Nuts','Oat milk or dairy milk','',''],
  ['REC-0003','Greek Yogurt','Dairy & Eggs',0.25,'cup',4,5.49,false,'Vegetarian','Dairy','Coconut yogurt','',''],
  ['REC-0003','Blueberries','Produce',0.5,'cup',6,4.99,false,'Vegetarian','','Any berries','','Fresh or frozen'],
  ['REC-0003','Honey','Condiments & Sauces',1,'tbsp',24,6.99,false,'Vegetarian','','Maple syrup','',''],
  ['REC-0003','Chia Seeds','Pantry',1,'tbsp',12,7.99,true,'Vegetarian','','','','Adds thickness'],
  // REC-0004 Avocado Toast with Poached Eggs (7)
  ['REC-0004','Bread','Grains & Pasta',2,'slice',20,4.49,false,'Vegetarian','Wheat','GF bread','Toasted',''],
  ['REC-0004','Avocado','Produce',1,'piece',3,1.29,false,'Vegetarian','','','Ripe',''],
  ['REC-0004','Eggs','Dairy & Eggs',2,'piece',12,3.99,false,'Vegetarian','Eggs','','Poached',''],
  ['REC-0004','Lemon','Produce',0.5,'piece',8,0.99,false,'Vegetarian','','Lime','Juiced',''],
  ['REC-0004','Red Pepper Flakes','Spices & Seasonings',0.25,'tsp',2,2.99,true,'Vegetarian','','','','Optional heat'],
  ['REC-0004','Salt','Spices & Seasonings',0.25,'tsp',26,1.29,false,'Vegetarian','','','',''],
  ['REC-0004','Flaky Sea Salt','Spices & Seasonings',1,'pinch',4,3.99,true,'Vegetarian','','Regular salt','For topping',''],
  // REC-0005 Banana Walnut Muffins (10)
  ['REC-0005','Ripe Bananas','Produce',3,'piece',7,0.89,false,'Vegetarian','','','Mashed',''],
  ['REC-0005','All-Purpose Flour','Baking',1.5,'cup',5,3.49,false,'Vegetarian','Wheat','','',''],
  ['REC-0005','Walnuts','Pantry',0.5,'cup',16,7.99,false,'Vegetarian','Tree Nuts','Pecans','Chopped',''],
  ['REC-0005','Brown Sugar','Baking',0.75,'cup',4,2.99,false,'Vegetarian','','','',''],
  ['REC-0005','Butter','Dairy & Eggs',0.5,'cup',1,4.49,false,'Vegetarian','Dairy','','Softened',''],
  ['REC-0005','Eggs','Dairy & Eggs',2,'piece',12,3.99,false,'Vegetarian','Eggs','','',''],
  ['REC-0005','Baking Soda','Baking',1,'tsp',8,1.79,false,'Vegetarian','','','',''],
  ['REC-0005','Vanilla Extract','Baking',1,'tsp',4,4.99,false,'Vegetarian','','','',''],
  ['REC-0005','Salt','Spices & Seasonings',0.5,'tsp',26,1.29,false,'Vegetarian','','','',''],
  ['REC-0005','Cinnamon','Spices & Seasonings',0.5,'tsp',2,2.99,true,'Vegetarian','','','',''],
  // REC-0006 French Toast Casserole (9)
  ['REC-0006','Brioche Bread','Grains & Pasta',1,'loaf',1,5.99,false,'Vegetarian','Wheat','Challah','Cubed',''],
  ['REC-0006','Eggs','Dairy & Eggs',6,'piece',12,3.99,false,'Vegetarian','Eggs','','',''],
  ['REC-0006','Whole Milk','Dairy & Eggs',2,'cup',0.5,2.49,false,'Vegetarian','Dairy','','',''],
  ['REC-0006','Heavy Cream','Dairy & Eggs',0.5,'cup',1,3.99,false,'Vegetarian','Dairy','Half and half','',''],
  ['REC-0006','Brown Sugar','Baking',0.5,'cup',4,2.99,false,'Vegetarian','','','',''],
  ['REC-0006','Vanilla Extract','Baking',2,'tsp',4,4.99,false,'Vegetarian','','','',''],
  ['REC-0006','Cinnamon','Spices & Seasonings',1,'tsp',2,2.99,false,'Vegetarian','','','',''],
  ['REC-0006','Butter','Dairy & Eggs',2,'tbsp',1,4.49,false,'Vegetarian','Dairy','','Dotted on top',''],
  ['REC-0006','Powdered Sugar','Baking',2,'tbsp',4,2.49,true,'Vegetarian','','','For serving',''],
  // REC-0007 Greek Yogurt Parfait (5)
  ['REC-0007','Greek Yogurt','Dairy & Eggs',1,'cup',4,5.49,false,'Vegetarian','Dairy','Coconut yogurt','',''],
  ['REC-0007','Granola','Pantry',0.25,'cup',8,4.99,false,'Vegetarian','','','','Check for GF'],
  ['REC-0007','Strawberries','Produce',0.5,'cup',6,4.99,false,'Vegetarian','','Any berries','Sliced',''],
  ['REC-0007','Blueberries','Produce',0.25,'cup',6,4.99,false,'Vegetarian','','','',''],
  ['REC-0007','Honey','Condiments & Sauces',1,'tbsp',24,6.99,false,'Vegetarian','','Maple syrup','Drizzled',''],
  // REC-0008 Shakshuka (8)
  ['REC-0008','Canned Crushed Tomatoes','Canned Goods',28,'oz',1,2.49,false,'Vegetarian','','','',''],
  ['REC-0008','Eggs','Dairy & Eggs',4,'piece',12,3.99,false,'Vegetarian','Eggs','','',''],
  ['REC-0008','Red Bell Pepper','Produce',1,'piece',3,1.29,false,'Vegetarian','','','Diced',''],
  ['REC-0008','Onion','Produce',1,'piece',3,0.79,false,'Vegetarian','','','Diced',''],
  ['REC-0008','Garlic','Produce',3,'clove',3,0.79,false,'Vegetarian','','','Minced',''],
  ['REC-0008','Olive Oil','Condiments & Sauces',2,'tbsp',17,7.99,false,'Vegetarian','','','',''],
  ['REC-0008','Cumin','Spices & Seasonings',1,'tsp',2,2.99,false,'Vegetarian','','','',''],
  ['REC-0008','Paprika','Spices & Seasonings',1,'tsp',2,2.99,false,'Vegetarian','','','',''],
  // REC-0009 Classic Caesar Salad (7)
  ['REC-0009','Romaine Lettuce','Produce',1,'piece',3,2.49,false,'','','','Chopped',''],
  ['REC-0009','Parmesan Cheese','Dairy & Eggs',0.5,'cup',8,5.99,false,'','Dairy','Pecorino','Shaved',''],
  ['REC-0009','Caesar Dressing','Condiments & Sauces',0.25,'cup',16,4.49,false,'','Dairy','','','Store-bought ok'],
  ['REC-0009','Croutons','Grains & Pasta',1,'cup',8,2.99,false,'','Wheat','GF croutons','',''],
  ['REC-0009','Anchovies','Meat & Seafood',2,'piece',10,3.99,true,'','Fish','Omit for vegetarian','',''],
  ['REC-0009','Lemon','Produce',0.5,'piece',8,0.99,false,'','','','Juiced',''],
  ['REC-0009','Black Pepper','Spices & Seasonings',0.5,'tsp',2,2.49,false,'','','','Fresh ground',''],
  // REC-0010 Turkey & Avocado Wrap (7)
  ['REC-0010','Flour Tortilla','Grains & Pasta',2,'piece',10,3.49,false,'','Wheat','GF wrap','','Large size'],
  ['REC-0010','Turkey Breast','Meat & Seafood',4,'oz',1,5.99,false,'Dairy-Free','','','Sliced',''],
  ['REC-0010','Avocado','Produce',1,'piece',3,1.29,false,'Dairy-Free','','','Sliced',''],
  ['REC-0010','Lettuce','Produce',2,'leaf',12,1.99,false,'Dairy-Free','','','',''],
  ['REC-0010','Tomato','Produce',1,'piece',3,0.99,false,'Dairy-Free','','','Sliced',''],
  ['REC-0010','Red Onion','Produce',2,'slice',3,0.69,false,'Dairy-Free','','','Thin sliced',''],
  ['REC-0010','Mustard','Condiments & Sauces',1,'tbsp',20,2.99,false,'Dairy-Free','','','',''],
  // REC-0011 Tomato Basil Soup (8)
  ['REC-0011','Canned San Marzano Tomatoes','Canned Goods',28,'oz',2,3.49,false,'Vegan','','Regular canned tomatoes','',''],
  ['REC-0011','Vegetable Broth','Canned Goods',2,'cup',4,2.99,false,'Vegan','','','',''],
  ['REC-0011','Onion','Produce',1,'piece',3,0.79,false,'Vegan','','','Diced',''],
  ['REC-0011','Garlic','Produce',4,'clove',3,0.79,false,'Vegan','','','',''],
  ['REC-0011','Fresh Basil','Produce',0.5,'cup',1,2.99,false,'Vegan','','Dried basil 2 tsp','',''],
  ['REC-0011','Olive Oil','Condiments & Sauces',2,'tbsp',17,7.99,false,'Vegan','','','',''],
  ['REC-0011','Salt','Spices & Seasonings',1,'tsp',26,1.29,false,'Vegan','','','',''],
  ['REC-0011','Sugar','Baking',1,'tsp',4,2.49,true,'Vegan','','','To balance acidity',''],
  // REC-0012 Mediterranean Quinoa Bowl (9)
  ['REC-0012','Quinoa','Grains & Pasta',1,'cup',4,5.99,false,'Vegan','','','',''],
  ['REC-0012','Cucumber','Produce',1,'piece',3,0.99,false,'Vegan','','','Diced',''],
  ['REC-0012','Cherry Tomatoes','Produce',1,'cup',6,3.49,false,'Vegan','','','Halved',''],
  ['REC-0012','Kalamata Olives','Canned Goods',0.25,'cup',12,3.99,false,'Vegan','','','',''],
  ['REC-0012','Feta Cheese','Dairy & Eggs',0.5,'cup',8,4.99,true,'Vegetarian','Dairy','Omit for vegan','Crumbled',''],
  ['REC-0012','Lemon','Produce',1,'piece',8,0.99,false,'Vegan','','','Juiced',''],
  ['REC-0012','Olive Oil','Condiments & Sauces',3,'tbsp',17,7.99,false,'Vegan','','','',''],
  ['REC-0012','Fresh Mint','Produce',2,'tbsp',1,2.49,true,'Vegan','','Fresh parsley','Chopped',''],
  ['REC-0012','Red Onion','Produce',0.25,'piece',3,0.69,false,'Vegan','','','Diced',''],
  // REC-0013 Chicken BLT Sandwich (8)
  ['REC-0013','Chicken Breast','Meat & Seafood',2,'piece',4,5.99,false,'','','','Cooked',''],
  ['REC-0013','Bread','Grains & Pasta',4,'slice',20,4.49,false,'','Wheat','','Toasted',''],
  ['REC-0013','Bacon','Meat & Seafood',4,'slice',12,5.49,false,'','','Turkey bacon','Cooked crispy',''],
  ['REC-0013','Lettuce','Produce',4,'leaf',12,1.99,false,'','','','',''],
  ['REC-0013','Tomato','Produce',1,'piece',3,0.99,false,'','','','Sliced',''],
  ['REC-0013','Mayonnaise','Condiments & Sauces',2,'tbsp',32,4.49,false,'','Eggs','','',''],
  ['REC-0013','Avocado','Produce',1,'piece',3,1.29,true,'','','','Sliced',''],
  ['REC-0013','Garlic Powder','Spices & Seasonings',0.5,'tsp',4,2.49,false,'','','','For chicken',''],
  // REC-0014 Lentil Soup (8)
  ['REC-0014','Green Lentils','Pantry',1.5,'cup',4,3.49,false,'Vegan','','','Rinsed',''],
  ['REC-0014','Carrot','Produce',2,'piece',12,1.49,false,'Vegan','','','Diced',''],
  ['REC-0014','Celery','Produce',2,'piece',12,1.99,false,'Vegan','','','Diced',''],
  ['REC-0014','Onion','Produce',1,'piece',3,0.79,false,'Vegan','','','Diced',''],
  ['REC-0014','Garlic','Produce',3,'clove',3,0.79,false,'Vegan','','','Minced',''],
  ['REC-0014','Vegetable Broth','Canned Goods',4,'cup',4,2.99,false,'Vegan','','','',''],
  ['REC-0014','Cumin','Spices & Seasonings',1.5,'tsp',2,2.99,false,'Vegan','','','',''],
  ['REC-0014','Lemon','Produce',1,'piece',8,0.99,false,'Vegan','','','Juiced',''],
  // REC-0015 Caprese Salad (6)
  ['REC-0015','Fresh Mozzarella','Dairy & Eggs',8,'oz',1,4.99,false,'Vegetarian','Dairy','','Sliced',''],
  ['REC-0015','Tomatoes','Produce',3,'piece',3,0.99,false,'Vegetarian','','','Sliced',''],
  ['REC-0015','Fresh Basil','Produce',1,'bunch',1,2.99,false,'Vegetarian','','','',''],
  ['REC-0015','Extra Virgin Olive Oil','Condiments & Sauces',3,'tbsp',17,8.99,false,'Vegetarian','','','Drizzled',''],
  ['REC-0015','Balsamic Glaze','Condiments & Sauces',1,'tbsp',8,5.99,true,'Vegetarian','','','Drizzled',''],
  ['REC-0015','Flaky Sea Salt','Spices & Seasonings',1,'pinch',4,3.99,false,'Vegetarian','','','',''],
  // REC-0016 Spaghetti Bolognese (11)
  ['REC-0016','Spaghetti','Grains & Pasta',1,'lb',4,2.49,false,'','Wheat','GF pasta','',''],
  ['REC-0016','Ground Beef','Meat & Seafood',1,'lb',1,5.99,false,'','','Ground turkey','',''],
  ['REC-0016','Canned Tomatoes','Canned Goods',28,'oz',1,2.49,false,'','','','',''],
  ['REC-0016','Onion','Produce',1,'piece',3,0.79,false,'','','','Diced',''],
  ['REC-0016','Carrot','Produce',1,'piece',12,1.49,false,'','','','Diced',''],
  ['REC-0016','Celery','Produce',1,'piece',12,1.99,false,'','','','Diced',''],
  ['REC-0016','Garlic','Produce',4,'clove',3,0.79,false,'','','','',''],
  ['REC-0016','Olive Oil','Condiments & Sauces',3,'tbsp',17,7.99,false,'','','','',''],
  ['REC-0016','Red Wine','Beverages',0.5,'cup',1,8.99,true,'','','Beef broth','',''],
  ['REC-0016','Parmesan','Dairy & Eggs',0.5,'cup',8,5.99,false,'','Dairy','','Grated','For serving'],
  ['REC-0016','Bay Leaves','Spices & Seasonings',2,'piece',8,2.99,false,'','','','','Remove before serving'],
  // REC-0017 Chicken Tikka Masala (12)
  ['REC-0017','Chicken Thighs','Meat & Seafood',1.5,'lb',4,5.49,false,'Gluten-Free','','','Cut into pieces',''],
  ['REC-0017','Yogurt','Dairy & Eggs',0.5,'cup',4,5.49,false,'Gluten-Free','Dairy','Dairy-free yogurt','For marinade',''],
  ['REC-0017','Heavy Cream','Dairy & Eggs',1,'cup',1,3.99,false,'Gluten-Free','Dairy','Coconut cream','',''],
  ['REC-0017','Canned Tomatoes','Canned Goods',14,'oz',2,2.49,false,'Gluten-Free','','','',''],
  ['REC-0017','Onion','Produce',1,'piece',3,0.79,false,'Gluten-Free','','','Diced',''],
  ['REC-0017','Garlic','Produce',4,'clove',3,0.79,false,'Gluten-Free','','','Minced',''],
  ['REC-0017','Ginger','Produce',1,'tbsp',1,1.49,false,'Gluten-Free','','Ground ginger 1 tsp','Grated',''],
  ['REC-0017','Garam Masala','Spices & Seasonings',2,'tsp',2,3.49,false,'Gluten-Free','','','',''],
  ['REC-0017','Cumin','Spices & Seasonings',1,'tsp',2,2.99,false,'Gluten-Free','','','',''],
  ['REC-0017','Turmeric','Spices & Seasonings',0.5,'tsp',2,2.99,false,'Gluten-Free','','','',''],
  ['REC-0017','Butter','Dairy & Eggs',2,'tbsp',1,4.49,false,'Gluten-Free','Dairy','Ghee','',''],
  ['REC-0017','Cilantro','Produce',0.25,'cup',1,2.49,true,'Gluten-Free','','Parsley','Chopped, for garnish',''],
  // REC-0018 Sheet Pan Lemon Herb Salmon (7)
  ['REC-0018','Salmon Fillets','Meat & Seafood',4,'piece',4,9.99,false,'Gluten-Free','Fish','Trout','',''],
  ['REC-0018','Lemon','Produce',2,'piece',8,0.99,false,'Gluten-Free','','','Sliced',''],
  ['REC-0018','Asparagus','Produce',1,'bunch',1,3.49,false,'Gluten-Free','','Green beans','Trimmed',''],
  ['REC-0018','Cherry Tomatoes','Produce',1,'cup',6,3.49,false,'Gluten-Free','','','Halved',''],
  ['REC-0018','Olive Oil','Condiments & Sauces',3,'tbsp',17,7.99,false,'Gluten-Free','','','',''],
  ['REC-0018','Fresh Dill','Produce',3,'tbsp',1,2.99,false,'Gluten-Free','','Dried dill 1 tbsp','',''],
  ['REC-0018','Garlic Powder','Spices & Seasonings',1,'tsp',4,2.49,false,'Gluten-Free','','','',''],
  // REC-0019 Beef Tacos (10)
  ['REC-0019','Ground Beef','Meat & Seafood',1.5,'lb',1,5.99,false,'','','Ground turkey','',''],
  ['REC-0019','Taco Shells','Grains & Pasta',12,'piece',12,3.49,false,'','Wheat','Corn shells for GF','',''],
  ['REC-0019','Cheddar Cheese','Dairy & Eggs',1,'cup',16,4.99,false,'','Dairy','','Shredded',''],
  ['REC-0019','Lettuce','Produce',2,'cup',12,1.99,false,'','','','Shredded',''],
  ['REC-0019','Tomato','Produce',2,'piece',3,0.99,false,'','','','Diced',''],
  ['REC-0019','Sour Cream','Dairy & Eggs',0.5,'cup',2,2.99,false,'','Dairy','Dairy-free sour cream','',''],
  ['REC-0019','Taco Seasoning','Spices & Seasonings',2,'tbsp',2,1.99,false,'','','Homemade blend','',''],
  ['REC-0019','Onion','Produce',0.5,'piece',3,0.79,false,'','','','Diced',''],
  ['REC-0019','Lime','Produce',1,'piece',8,0.69,false,'','','','Juiced',''],
  ['REC-0019','Salsa','Condiments & Sauces',0.5,'cup',16,3.99,false,'','','Homemade salsa','',''],
  // REC-0020 Mushroom Risotto (9)
  ['REC-0020','Arborio Rice','Grains & Pasta',1.5,'cup',4,5.99,false,'Vegetarian','','','',''],
  ['REC-0020','Mushrooms','Produce',8,'oz',12,3.99,false,'Vegetarian','','Cremini or portobello','Sliced',''],
  ['REC-0020','Vegetable Broth','Canned Goods',4,'cup',4,2.99,false,'Vegetarian','','','Warm',''],
  ['REC-0020','White Wine','Beverages',0.5,'cup',1,8.99,false,'Vegetarian','','Broth','',''],
  ['REC-0020','Parmesan','Dairy & Eggs',0.75,'cup',8,5.99,false,'Vegetarian','Dairy','','Grated',''],
  ['REC-0020','Butter','Dairy & Eggs',3,'tbsp',1,4.49,false,'Vegetarian','Dairy','','',''],
  ['REC-0020','Onion','Produce',1,'piece',3,0.79,false,'Vegetarian','','Shallots','Diced',''],
  ['REC-0020','Garlic','Produce',3,'clove',3,0.79,false,'Vegetarian','','','Minced',''],
  ['REC-0020','Fresh Thyme','Produce',3,'sprig',1,2.99,true,'Vegetarian','','Dried thyme 1 tsp','',''],
  // REC-0021 Slow Cooker Pulled Pork (7)
  ['REC-0021','Pork Shoulder','Meat & Seafood',4,'lb',1,8.99,false,'Gluten-Free','','','',''],
  ['REC-0021','BBQ Sauce','Condiments & Sauces',1,'cup',20,4.99,false,'Gluten-Free','','Homemade','',''],
  ['REC-0021','Apple Cider Vinegar','Condiments & Sauces',0.25,'cup',16,3.99,false,'Gluten-Free','','White wine vinegar','',''],
  ['REC-0021','Brown Sugar','Baking',2,'tbsp',4,2.99,false,'Gluten-Free','','','',''],
  ['REC-0021','Paprika','Spices & Seasonings',2,'tsp',2,2.99,false,'Gluten-Free','','','',''],
  ['REC-0021','Garlic Powder','Spices & Seasonings',2,'tsp',4,2.49,false,'Gluten-Free','','','',''],
  ['REC-0021','Onion Powder','Spices & Seasonings',2,'tsp',4,2.49,false,'Gluten-Free','','','',''],
  // REC-0022 Thai Green Curry (9)
  ['REC-0022','Chicken Breast','Meat & Seafood',1,'lb',4,5.99,false,'Gluten-Free','','Tofu for vegan','Cut into pieces',''],
  ['REC-0022','Coconut Milk','Canned Goods',14,'oz',2,2.99,false,'Gluten-Free','','','Full-fat preferred',''],
  ['REC-0022','Green Curry Paste','Condiments & Sauces',3,'tbsp',4,3.99,false,'Gluten-Free','','','',''],
  ['REC-0022','Zucchini','Produce',1,'piece',3,0.99,false,'Gluten-Free','','','Sliced',''],
  ['REC-0022','Bell Pepper','Produce',1,'piece',3,1.29,false,'Gluten-Free','','','Sliced',''],
  ['REC-0022','Fish Sauce','Condiments & Sauces',2,'tbsp',6,3.49,false,'Gluten-Free','Fish','Soy sauce or tamari','',''],
  ['REC-0022','Basil','Produce',0.25,'cup',1,2.99,true,'Gluten-Free','','','Thai basil preferred',''],
  ['REC-0022','Lime','Produce',1,'piece',8,0.69,false,'Gluten-Free','','','Juiced',''],
  ['REC-0022','Jasmine Rice','Grains & Pasta',2,'cup',4,3.99,false,'Gluten-Free','','','For serving',''],
  // REC-0023 Grilled Teriyaki Chicken (7)
  ['REC-0023','Chicken Thighs','Meat & Seafood',4,'piece',4,5.49,false,'Gluten-Free','Soy','','Bone-in or boneless',''],
  ['REC-0023','Soy Sauce','Condiments & Sauces',0.25,'cup',20,3.99,false,'Gluten-Free','Soy','Tamari for GF','',''],
  ['REC-0023','Honey','Condiments & Sauces',3,'tbsp',24,6.99,false,'Gluten-Free','','Maple syrup','',''],
  ['REC-0023','Garlic','Produce',3,'clove',3,0.79,false,'Gluten-Free','','','Minced',''],
  ['REC-0023','Ginger','Produce',1,'tbsp',1,1.49,false,'Gluten-Free','','Ground ginger','Grated',''],
  ['REC-0023','Sesame Oil','Condiments & Sauces',1,'tbsp',8,4.99,false,'Gluten-Free','Sesame','','',''],
  ['REC-0023','Green Onions','Produce',3,'piece',4,0.99,true,'Gluten-Free','','','Sliced, for garnish',''],
  // REC-0024 Beef Stew (10)
  ['REC-0024','Beef Chuck','Meat & Seafood',2,'lb',1,7.99,false,'Gluten-Free','','','Cubed',''],
  ['REC-0024','Potatoes','Produce',3,'piece',12,1.49,false,'Gluten-Free','','','Cubed',''],
  ['REC-0024','Carrots','Produce',3,'piece',12,1.49,false,'Gluten-Free','','','Sliced',''],
  ['REC-0024','Beef Broth','Canned Goods',3,'cup',4,2.99,false,'Gluten-Free','','','',''],
  ['REC-0024','Onion','Produce',1,'piece',3,0.79,false,'Gluten-Free','','','Diced',''],
  ['REC-0024','Garlic','Produce',3,'clove',3,0.79,false,'Gluten-Free','','','',''],
  ['REC-0024','Tomato Paste','Canned Goods',2,'tbsp',6,1.49,false,'Gluten-Free','','','',''],
  ['REC-0024','Thyme','Spices & Seasonings',1,'tsp',2,2.99,false,'Gluten-Free','','','Dried',''],
  ['REC-0024','Bay Leaves','Spices & Seasonings',2,'piece',8,2.99,false,'Gluten-Free','','','Remove before serving',''],
  ['REC-0024','Cornstarch','Baking',2,'tbsp',16,2.49,true,'Gluten-Free','','Arrowroot','For thickening',''],
  // REC-0025 Shrimp Scampi (8)
  ['REC-0025','Shrimp','Meat & Seafood',1,'lb',1,9.99,false,'','Shellfish','','Peeled and deveined',''],
  ['REC-0025','Linguine','Grains & Pasta',12,'oz',4,2.49,false,'','Wheat','GF pasta','',''],
  ['REC-0025','Butter','Dairy & Eggs',4,'tbsp',1,4.49,false,'','Dairy','','',''],
  ['REC-0025','Garlic','Produce',5,'clove',3,0.79,false,'','','','Minced',''],
  ['REC-0025','White Wine','Beverages',0.5,'cup',1,8.99,false,'','','Chicken broth','Dry',''],
  ['REC-0025','Lemon','Produce',1,'piece',8,0.99,false,'','','','Juiced',''],
  ['REC-0025','Parsley','Produce',0.25,'cup',1,1.99,false,'','','','Chopped',''],
  ['REC-0025','Red Pepper Flakes','Spices & Seasonings',0.25,'tsp',2,2.99,true,'','','','',''],
  // REC-0026 Vegetable Stir Fry (9)
  ['REC-0026','Broccoli','Produce',2,'cup',3,2.49,false,'Vegan','','','Florets',''],
  ['REC-0026','Bell Peppers','Produce',2,'piece',3,1.29,false,'Vegan','','','Sliced',''],
  ['REC-0026','Snap Peas','Produce',1,'cup',12,3.99,false,'Vegan','','Green beans','',''],
  ['REC-0026','Carrots','Produce',2,'piece',12,1.49,false,'Vegan','','','Julienned',''],
  ['REC-0026','Soy Sauce','Condiments & Sauces',3,'tbsp',20,3.99,false,'Vegan','Soy','Tamari','',''],
  ['REC-0026','Sesame Oil','Condiments & Sauces',1,'tbsp',8,4.99,false,'Vegan','Sesame','','',''],
  ['REC-0026','Garlic','Produce',3,'clove',3,0.79,false,'Vegan','','','Minced',''],
  ['REC-0026','Ginger','Produce',1,'tbsp',1,1.49,false,'Vegan','','','Grated',''],
  ['REC-0026','Cornstarch','Baking',1,'tbsp',16,2.49,false,'Vegan','','','For sauce',''],
  // REC-0027 BBQ Baby Back Ribs (7)
  ['REC-0027','Baby Back Ribs','Meat & Seafood',2,'lb',1,11.99,false,'Gluten-Free','','','Membrane removed',''],
  ['REC-0027','BBQ Sauce','Condiments & Sauces',1.5,'cup',20,4.99,false,'Gluten-Free','','Homemade','',''],
  ['REC-0027','Brown Sugar','Baking',2,'tbsp',4,2.99,false,'Gluten-Free','','','','Dry rub'],
  ['REC-0027','Paprika','Spices & Seasonings',2,'tsp',2,2.99,false,'Gluten-Free','','','','Dry rub'],
  ['REC-0027','Garlic Powder','Spices & Seasonings',1,'tsp',4,2.49,false,'Gluten-Free','','','','Dry rub'],
  ['REC-0027','Onion Powder','Spices & Seasonings',1,'tsp',4,2.49,false,'Gluten-Free','','','','Dry rub'],
  ['REC-0027','Salt','Spices & Seasonings',1,'tsp',26,1.29,false,'Gluten-Free','','','',''],
  // REC-0028 One Pot Chicken Pasta (10)
  ['REC-0028','Penne Pasta','Grains & Pasta',12,'oz',4,2.49,false,'','Wheat','GF pasta','',''],
  ['REC-0028','Chicken Breast','Meat & Seafood',1,'lb',4,5.99,false,'','','','Cubed',''],
  ['REC-0028','Heavy Cream','Dairy & Eggs',1,'cup',1,3.99,false,'','Dairy','Half and half','',''],
  ['REC-0028','Chicken Broth','Canned Goods',2,'cup',4,2.99,false,'','','','',''],
  ['REC-0028','Parmesan','Dairy & Eggs',0.5,'cup',8,5.99,false,'','Dairy','','Grated',''],
  ['REC-0028','Garlic','Produce',4,'clove',3,0.79,false,'','','','',''],
  ['REC-0028','Sun-Dried Tomatoes','Canned Goods',0.5,'cup',4,4.99,false,'','','','',''],
  ['REC-0028','Baby Spinach','Produce',2,'cup',5,3.99,false,'','','','',''],
  ['REC-0028','Italian Seasoning','Spices & Seasonings',1,'tsp',4,2.99,false,'','','','',''],
  ['REC-0028','Olive Oil','Condiments & Sauces',2,'tbsp',17,7.99,false,'','','','',''],
  // REC-0029 Korean Beef Bowl (8)
  ['REC-0029','Ground Beef','Meat & Seafood',1,'lb',1,5.99,false,'Dairy-Free','','Ground turkey','',''],
  ['REC-0029','Jasmine Rice','Grains & Pasta',2,'cup',4,3.99,false,'Dairy-Free','','','',''],
  ['REC-0029','Soy Sauce','Condiments & Sauces',0.25,'cup',20,3.99,false,'Dairy-Free','Soy','Coconut aminos','',''],
  ['REC-0029','Sesame Oil','Condiments & Sauces',1,'tsp',8,4.99,false,'Dairy-Free','Sesame','','',''],
  ['REC-0029','Brown Sugar','Baking',2,'tbsp',4,2.99,false,'Dairy-Free','','Honey','',''],
  ['REC-0029','Garlic','Produce',4,'clove',3,0.79,false,'Dairy-Free','','','Minced',''],
  ['REC-0029','Ginger','Produce',1,'tsp',1,1.49,false,'Dairy-Free','','','Grated',''],
  ['REC-0029','Green Onions','Produce',3,'piece',4,0.99,false,'Dairy-Free','','','Sliced',''],
  // REC-0030 Garlic Butter Shrimp (6)
  ['REC-0030','Shrimp','Meat & Seafood',1,'lb',1,9.99,false,'Gluten-Free','Shellfish','','Peeled, deveined',''],
  ['REC-0030','Butter','Dairy & Eggs',4,'tbsp',1,4.49,false,'Gluten-Free','Dairy','','',''],
  ['REC-0030','Garlic','Produce',5,'clove',3,0.79,false,'Gluten-Free','','','Minced',''],
  ['REC-0030','Lemon','Produce',1,'piece',8,0.99,false,'Gluten-Free','','','Juiced',''],
  ['REC-0030','Parsley','Produce',2,'tbsp',1,1.99,false,'Gluten-Free','','','Chopped',''],
  ['REC-0030','Red Pepper Flakes','Spices & Seasonings',0.5,'tsp',2,2.99,true,'Gluten-Free','','','',''],
  // REC-0031 Homemade Guacamole (6)
  ['REC-0031','Avocados','Produce',3,'piece',3,1.29,false,'Vegan','','','Ripe',''],
  ['REC-0031','Lime','Produce',1,'piece',8,0.69,false,'Vegan','','Lemon','Juiced',''],
  ['REC-0031','Cilantro','Produce',2,'tbsp',1,2.49,false,'Vegan','','Parsley for those who dislike cilantro','Chopped',''],
  ['REC-0031','Red Onion','Produce',0.25,'piece',3,0.69,false,'Vegan','','White onion','Diced fine',''],
  ['REC-0031','Jalapeño','Produce',1,'piece',4,0.49,true,'Vegan','','Omit for mild','Seeded, diced',''],
  ['REC-0031','Salt','Spices & Seasonings',0.5,'tsp',26,1.29,false,'Vegan','','','',''],
  // REC-0032 Stuffed Mushrooms (7)
  ['REC-0032','Large Mushrooms','Produce',18,'piece',12,3.99,false,'Vegetarian','','','Stems removed',''],
  ['REC-0032','Cream Cheese','Dairy & Eggs',8,'oz',1,3.49,false,'Vegetarian','Dairy','','Softened',''],
  ['REC-0032','Parmesan','Dairy & Eggs',0.25,'cup',8,5.99,false,'Vegetarian','Dairy','','Grated',''],
  ['REC-0032','Garlic','Produce',3,'clove',3,0.79,false,'Vegetarian','','','Minced',''],
  ['REC-0032','Italian Seasoning','Spices & Seasonings',1,'tsp',4,2.99,false,'Vegetarian','','','',''],
  ['REC-0032','Parsley','Produce',2,'tbsp',1,1.99,false,'Vegetarian','','','Chopped',''],
  ['REC-0032','Bread Crumbs','Grains & Pasta',0.25,'cup',16,2.49,false,'Vegetarian','Wheat','GF crumbs','',''],
  // REC-0033 Spinach Artichoke Dip (8)
  ['REC-0033','Baby Spinach','Produce',10,'oz',5,3.99,false,'Vegetarian','','','Chopped',''],
  ['REC-0033','Artichoke Hearts','Canned Goods',14,'oz',2,3.49,false,'Vegetarian','','','Drained, chopped',''],
  ['REC-0033','Cream Cheese','Dairy & Eggs',8,'oz',1,3.49,false,'Vegetarian','Dairy','','Softened',''],
  ['REC-0033','Sour Cream','Dairy & Eggs',0.5,'cup',2,2.99,false,'Vegetarian','Dairy','','',''],
  ['REC-0033','Parmesan','Dairy & Eggs',0.5,'cup',8,5.99,false,'Vegetarian','Dairy','','Grated',''],
  ['REC-0033','Mozzarella','Dairy & Eggs',1,'cup',16,4.99,false,'Vegetarian','Dairy','','Shredded',''],
  ['REC-0033','Garlic','Produce',3,'clove',3,0.79,false,'Vegetarian','','','Minced',''],
  ['REC-0033','Red Pepper Flakes','Spices & Seasonings',0.25,'tsp',2,2.99,true,'Vegetarian','','','',''],
  // REC-0034 Caprese Bruschetta (7)
  ['REC-0034','Baguette','Grains & Pasta',1,'piece',1,3.99,false,'Vegetarian','Wheat','GF bread','Sliced',''],
  ['REC-0034','Fresh Mozzarella','Dairy & Eggs',8,'oz',1,4.99,false,'Vegetarian','Dairy','','Diced',''],
  ['REC-0034','Tomatoes','Produce',3,'piece',3,0.99,false,'Vegetarian','','','Diced',''],
  ['REC-0034','Fresh Basil','Produce',0.25,'cup',1,2.99,false,'Vegetarian','','','Chiffonade',''],
  ['REC-0034','Olive Oil','Condiments & Sauces',3,'tbsp',17,7.99,false,'Vegetarian','','','',''],
  ['REC-0034','Balsamic Glaze','Condiments & Sauces',2,'tbsp',8,5.99,false,'Vegetarian','','','Drizzled',''],
  ['REC-0034','Garlic','Produce',1,'clove',3,0.79,false,'Vegetarian','','','Rub on bread',''],
  // REC-0035 Crispy Air Fryer Chickpeas (5)
  ['REC-0035','Canned Chickpeas','Canned Goods',15,'oz',2,1.29,false,'Vegan','','','Drained, dried well',''],
  ['REC-0035','Olive Oil','Condiments & Sauces',1,'tbsp',17,7.99,false,'Vegan','','','',''],
  ['REC-0035','Cumin','Spices & Seasonings',1,'tsp',2,2.99,false,'Vegan','','','',''],
  ['REC-0035','Paprika','Spices & Seasonings',1,'tsp',2,2.99,false,'Vegan','','','',''],
  ['REC-0035','Salt','Spices & Seasonings',0.5,'tsp',26,1.29,false,'Vegan','','','',''],
  // REC-0036 Chicken Noodle Soup (9)
  ['REC-0036','Chicken Thighs','Meat & Seafood',2,'piece',4,5.49,false,'Dairy-Free','','Rotisserie chicken','',''],
  ['REC-0036','Egg Noodles','Grains & Pasta',2,'cup',16,3.49,false,'Dairy-Free','Wheat','GF noodles','',''],
  ['REC-0036','Chicken Broth','Canned Goods',6,'cup',4,2.99,false,'Dairy-Free','','','',''],
  ['REC-0036','Carrots','Produce',3,'piece',12,1.49,false,'Dairy-Free','','','Sliced',''],
  ['REC-0036','Celery','Produce',3,'piece',12,1.99,false,'Dairy-Free','','','Sliced',''],
  ['REC-0036','Onion','Produce',1,'piece',3,0.79,false,'Dairy-Free','','','Diced',''],
  ['REC-0036','Garlic','Produce',3,'clove',3,0.79,false,'Dairy-Free','','','',''],
  ['REC-0036','Thyme','Spices & Seasonings',1,'tsp',2,2.99,false,'Dairy-Free','','','Dried',''],
  ['REC-0036','Parsley','Produce',0.25,'cup',1,1.99,false,'Dairy-Free','','','Chopped',''],
  // REC-0037 French Onion Soup (8)
  ['REC-0037','Yellow Onions','Produce',4,'piece',3,0.79,false,'Vegetarian','','','Thinly sliced',''],
  ['REC-0037','Gruyere Cheese','Dairy & Eggs',2,'cup',16,7.99,false,'Vegetarian','Dairy','Swiss cheese','Shredded',''],
  ['REC-0037','Baguette','Grains & Pasta',1,'piece',1,3.99,false,'Vegetarian','Wheat','GF bread','Sliced',''],
  ['REC-0037','Butter','Dairy & Eggs',3,'tbsp',1,4.49,false,'Vegetarian','Dairy','','',''],
  ['REC-0037','Beef Broth','Canned Goods',4,'cup',4,2.99,false,'Vegetarian','','Veg broth','',''],
  ['REC-0037','Dry White Wine','Beverages',0.5,'cup',1,8.99,false,'Vegetarian','','Dry sherry','',''],
  ['REC-0037','Thyme','Spices & Seasonings',2,'sprig',1,2.99,false,'Vegetarian','','Dried thyme 1 tsp','',''],
  ['REC-0037','Bay Leaf','Spices & Seasonings',1,'piece',8,2.99,false,'Vegetarian','','','Remove before serving',''],
  // REC-0038 Minestrone (11)
  ['REC-0038','Canned Diced Tomatoes','Canned Goods',28,'oz',2,2.49,false,'Vegan','','','',''],
  ['REC-0038','Cannellini Beans','Canned Goods',15,'oz',2,1.29,false,'Vegan','','Any white bean','Drained',''],
  ['REC-0038','Ditalini Pasta','Grains & Pasta',1,'cup',4,2.49,false,'Vegan','Wheat','GF pasta','',''],
  ['REC-0038','Zucchini','Produce',1,'piece',3,0.99,false,'Vegan','','','Diced',''],
  ['REC-0038','Green Beans','Produce',1,'cup',12,2.99,false,'Vegan','','','Trimmed',''],
  ['REC-0038','Onion','Produce',1,'piece',3,0.79,false,'Vegan','','','Diced',''],
  ['REC-0038','Celery','Produce',2,'piece',12,1.99,false,'Vegan','','','Sliced',''],
  ['REC-0038','Carrot','Produce',2,'piece',12,1.49,false,'Vegan','','','Diced',''],
  ['REC-0038','Garlic','Produce',3,'clove',3,0.79,false,'Vegan','','','',''],
  ['REC-0038','Vegetable Broth','Canned Goods',4,'cup',4,2.99,false,'Vegan','','','',''],
  ['REC-0038','Italian Seasoning','Spices & Seasonings',1,'tsp',4,2.99,false,'Vegan','','','',''],
  // REC-0039 Butternut Squash Soup (7)
  ['REC-0039','Butternut Squash','Produce',2,'lb',1,2.49,false,'Vegetarian','','','Cubed',''],
  ['REC-0039','Onion','Produce',1,'piece',3,0.79,false,'Vegetarian','','','Diced',''],
  ['REC-0039','Garlic','Produce',4,'clove',3,0.79,false,'Vegetarian','','','',''],
  ['REC-0039','Vegetable Broth','Canned Goods',3,'cup',4,2.99,false,'Vegetarian','','','',''],
  ['REC-0039','Heavy Cream','Dairy & Eggs',0.5,'cup',1,3.99,false,'Vegetarian','Dairy','Coconut cream','',''],
  ['REC-0039','Olive Oil','Condiments & Sauces',2,'tbsp',17,7.99,false,'Vegetarian','','','',''],
  ['REC-0039','Nutmeg','Spices & Seasonings',0.25,'tsp',2,3.49,false,'Vegetarian','','','',''],
  // REC-0040 Chocolate Chip Cookies (10)
  ['REC-0040','All-Purpose Flour','Baking',2.25,'cup',5,3.49,false,'Vegetarian','Wheat','GF flour','',''],
  ['REC-0040','Butter','Dairy & Eggs',1,'cup',1,4.49,false,'Vegetarian','Dairy','','Browned',''],
  ['REC-0040','Brown Sugar','Baking',0.75,'cup',4,2.99,false,'Vegetarian','','','',''],
  ['REC-0040','Granulated Sugar','Baking',0.75,'cup',4,2.49,false,'Vegetarian','','','',''],
  ['REC-0040','Eggs','Dairy & Eggs',2,'piece',12,3.99,false,'Vegetarian','Eggs','','',''],
  ['REC-0040','Vanilla Extract','Baking',1,'tsp',4,4.99,false,'Vegetarian','','','',''],
  ['REC-0040','Chocolate Chips','Baking',2,'cup',12,4.99,false,'Vegetarian','Dairy','Dairy-free chips','',''],
  ['REC-0040','Baking Soda','Baking',1,'tsp',8,1.79,false,'Vegetarian','','','',''],
  ['REC-0040','Salt','Spices & Seasonings',1,'tsp',26,1.29,false,'Vegetarian','','','',''],
  ['REC-0040','Walnuts','Pantry',1,'cup',16,7.99,true,'Vegetarian','Tree Nuts','Pecans','Chopped',''],
  // REC-0041 Classic Cheesecake (9)
  ['REC-0041','Cream Cheese','Dairy & Eggs',24,'oz',1,3.49,false,'Vegetarian','Dairy','','Softened','3 packages'],
  ['REC-0041','Sugar','Baking',1,'cup',4,2.49,false,'Vegetarian','','','',''],
  ['REC-0041','Eggs','Dairy & Eggs',3,'piece',12,3.99,false,'Vegetarian','Eggs','','Room temperature',''],
  ['REC-0041','Sour Cream','Dairy & Eggs',1,'cup',2,2.99,false,'Vegetarian','Dairy','','',''],
  ['REC-0041','Vanilla Extract','Baking',2,'tsp',4,4.99,false,'Vegetarian','','','',''],
  ['REC-0041','Graham Crackers','Baking',1.5,'cup',12,4.49,false,'Vegetarian','Wheat','GF crackers','Crushed','For crust'],
  ['REC-0041','Butter','Dairy & Eggs',5,'tbsp',1,4.49,false,'Vegetarian','Dairy','','Melted','For crust'],
  ['REC-0041','Lemon Zest','Produce',1,'tsp',8,0.99,true,'Vegetarian','','','',''],
  ['REC-0041','Heavy Cream','Dairy & Eggs',1,'cup',1,3.99,true,'Vegetarian','Dairy','','Whipped, for topping',''],
  // REC-0042 Tiramisu (8)
  ['REC-0042','Ladyfinger Cookies','Baking',24,'piece',24,4.99,false,'Vegetarian','Wheat','GF ladyfingers','',''],
  ['REC-0042','Mascarpone','Dairy & Eggs',16,'oz',1,7.99,false,'Vegetarian','Dairy','','',''],
  ['REC-0042','Eggs','Dairy & Eggs',4,'piece',12,3.99,false,'Vegetarian','Eggs','','Separated',''],
  ['REC-0042','Sugar','Baking',0.75,'cup',4,2.49,false,'Vegetarian','','','',''],
  ['REC-0042','Espresso','Beverages',1,'cup',8,8.99,false,'Vegetarian','','Strong coffee','Cooled',''],
  ['REC-0042','Coffee Liqueur','Beverages',3,'tbsp',1,14.99,true,'Vegetarian','','Omit for alcohol-free','',''],
  ['REC-0042','Cocoa Powder','Baking',2,'tbsp',8,3.99,false,'Vegetarian','','','Sifted, for dusting',''],
  ['REC-0042','Heavy Cream','Dairy & Eggs',0.5,'cup',1,3.99,true,'Vegetarian','Dairy','','',''],
  // REC-0043 Apple Pie (9)
  ['REC-0043','Apples','Produce',6,'piece',12,0.79,false,'Vegetarian','','','Peeled, sliced','Granny Smith or Honeycrisp'],
  ['REC-0043','Sugar','Baking',0.75,'cup',4,2.49,false,'Vegetarian','','','',''],
  ['REC-0043','Brown Sugar','Baking',0.25,'cup',4,2.99,false,'Vegetarian','','','',''],
  ['REC-0043','Cinnamon','Spices & Seasonings',2,'tsp',2,2.99,false,'Vegetarian','','','',''],
  ['REC-0043','Nutmeg','Spices & Seasonings',0.25,'tsp',2,3.49,false,'Vegetarian','','','',''],
  ['REC-0043','All-Purpose Flour','Baking',2,'tbsp',5,3.49,false,'Vegetarian','Wheat','Cornstarch','','For thickening'],
  ['REC-0043','Butter','Dairy & Eggs',2,'tbsp',1,4.49,false,'Vegetarian','Dairy','','Dotted',''],
  ['REC-0043','Pie Crust','Baking',2,'piece',2,4.99,false,'Vegetarian','Wheat','GF crust','','Store-bought ok'],
  ['REC-0043','Egg Wash','Dairy & Eggs',1,'piece',12,3.99,false,'Vegetarian','Eggs','Milk wash','Beaten with water',''],
  // REC-0044 Chocolate Lava Cake (7)
  ['REC-0044','Dark Chocolate','Baking',6,'oz',8,4.99,false,'Vegetarian','Dairy','','Chopped','70% cocoa'],
  ['REC-0044','Butter','Dairy & Eggs',0.5,'cup',1,4.49,false,'Vegetarian','Dairy','','',''],
  ['REC-0044','Eggs','Dairy & Eggs',4,'piece',12,3.99,false,'Vegetarian','Eggs','','Room temperature','2 whole + 2 yolks'],
  ['REC-0044','Sugar','Baking',0.75,'cup',4,2.49,false,'Vegetarian','','','',''],
  ['REC-0044','All-Purpose Flour','Baking',0.5,'cup',5,3.49,false,'Vegetarian','Wheat','GF flour','',''],
  ['REC-0044','Vanilla Extract','Baking',1,'tsp',4,4.99,false,'Vegetarian','','','',''],
  ['REC-0044','Powdered Sugar','Baking',1,'tbsp',4,2.49,false,'Vegetarian','','','For dusting',''],
  // REC-0045 Lemon Bars (8)
  ['REC-0045','All-Purpose Flour','Baking',2,'cup',5,3.49,false,'Vegetarian','Wheat','GF flour','','Crust and filling'],
  ['REC-0045','Butter','Dairy & Eggs',1,'cup',1,4.49,false,'Vegetarian','Dairy','','Softened',''],
  ['REC-0045','Powdered Sugar','Baking',0.5,'cup',4,2.49,false,'Vegetarian','','','','For crust'],
  ['REC-0045','Eggs','Dairy & Eggs',4,'piece',12,3.99,false,'Vegetarian','Eggs','','','For filling'],
  ['REC-0045','Sugar','Baking',1.5,'cup',4,2.49,false,'Vegetarian','','','','For filling'],
  ['REC-0045','Lemon Juice','Produce',0.5,'cup',8,0.99,false,'Vegetarian','','','Fresh','3-4 lemons'],
  ['REC-0045','Lemon Zest','Produce',2,'tbsp',8,0.99,false,'Vegetarian','','','Grated',''],
  ['REC-0045','Powdered Sugar','Baking',2,'tbsp',4,2.49,false,'Vegetarian','','','For dusting',''],
  // REC-0046 Banana Bread (9)
  ['REC-0046','Ripe Bananas','Produce',3,'piece',7,0.89,false,'Vegetarian','','','Mashed','Very ripe'],
  ['REC-0046','All-Purpose Flour','Baking',1.5,'cup',5,3.49,false,'Vegetarian','Wheat','GF flour','',''],
  ['REC-0046','Sugar','Baking',0.75,'cup',4,2.49,false,'Vegetarian','','Brown sugar','',''],
  ['REC-0046','Butter','Dairy & Eggs',0.5,'cup',1,4.49,false,'Vegetarian','Dairy','Coconut oil','Melted',''],
  ['REC-0046','Eggs','Dairy & Eggs',2,'piece',12,3.99,false,'Vegetarian','Eggs','','',''],
  ['REC-0046','Baking Soda','Baking',1,'tsp',8,1.79,false,'Vegetarian','','','',''],
  ['REC-0046','Salt','Spices & Seasonings',0.5,'tsp',26,1.29,false,'Vegetarian','','','',''],
  ['REC-0046','Vanilla Extract','Baking',1,'tsp',4,4.99,false,'Vegetarian','','','',''],
  ['REC-0046','Chocolate Chips','Baking',0.5,'cup',12,4.99,true,'Vegetarian','Dairy','Walnuts','','Stir in'],
  // REC-0047 Strawberry Shortcake (9)
  ['REC-0047','Strawberries','Produce',2,'lb',6,4.99,false,'Vegetarian','','','Sliced','Macerate with sugar'],
  ['REC-0047','All-Purpose Flour','Baking',2,'cup',5,3.49,false,'Vegetarian','Wheat','','',''],
  ['REC-0047','Sugar','Baking',0.25,'cup',4,2.49,false,'Vegetarian','','','','For biscuits'],
  ['REC-0047','Butter','Dairy & Eggs',0.5,'cup',1,4.49,false,'Vegetarian','Dairy','','Cold, cubed',''],
  ['REC-0047','Heavy Cream','Dairy & Eggs',1.5,'cup',1,3.99,false,'Vegetarian','Dairy','','','Whipped for topping'],
  ['REC-0047','Baking Powder','Baking',1,'tbsp',8,2.99,false,'Vegetarian','','','',''],
  ['REC-0047','Eggs','Dairy & Eggs',1,'piece',12,3.99,false,'Vegetarian','Eggs','','',''],
  ['REC-0047','Powdered Sugar','Baking',2,'tbsp',4,2.49,false,'Vegetarian','','','','Whipped cream'],
  ['REC-0047','Vanilla Extract','Baking',1,'tsp',4,4.99,false,'Vegetarian','','','',''],
  // REC-0048 No-Knead Artisan Bread (5)
  ['REC-0048','All-Purpose Flour','Baking',3,'cup',5,3.49,false,'Vegetarian','Wheat','Bread flour','',''],
  ['REC-0048','Instant Yeast','Baking',0.25,'tsp',4,3.49,false,'Vegetarian','','','',''],
  ['REC-0048','Salt','Spices & Seasonings',1.5,'tsp',26,1.29,false,'Vegetarian','','','',''],
  ['REC-0048','Water','Beverages',1.5,'cup',0,0,false,'Vegetarian','','','Room temperature',''],
  ['REC-0048','Olive Oil','Condiments & Sauces',1,'tsp',17,7.99,true,'Vegetarian','','','For drizzling',''],
  // REC-0049 Dinner Rolls (8)
  ['REC-0049','All-Purpose Flour','Baking',3.5,'cup',5,3.49,false,'Vegetarian','Wheat','','',''],
  ['REC-0049','Active Dry Yeast','Baking',2.25,'tsp',4,3.49,false,'Vegetarian','','','',''],
  ['REC-0049','Butter','Dairy & Eggs',4,'tbsp',1,4.49,false,'Vegetarian','Dairy','','Softened',''],
  ['REC-0049','Milk','Dairy & Eggs',1,'cup',0.5,2.49,false,'Vegetarian','Dairy','Oat milk','Warm',''],
  ['REC-0049','Sugar','Baking',3,'tbsp',4,2.49,false,'Vegetarian','','','',''],
  ['REC-0049','Egg','Dairy & Eggs',1,'piece',12,3.99,false,'Vegetarian','Eggs','','',''],
  ['REC-0049','Salt','Spices & Seasonings',1,'tsp',26,1.29,false,'Vegetarian','','','',''],
  ['REC-0049','Butter','Dairy & Eggs',2,'tbsp',1,4.49,false,'Vegetarian','Dairy','','Melted, for tops',''],
  // REC-0050 Zucchini Bread (9)
  ['REC-0050','Zucchini','Produce',2,'piece',3,0.99,false,'Vegetarian','','','Grated, squeezed dry',''],
  ['REC-0050','All-Purpose Flour','Baking',2,'cup',5,3.49,false,'Vegetarian','Wheat','GF flour','',''],
  ['REC-0050','Sugar','Baking',1,'cup',4,2.49,false,'Vegetarian','','Brown sugar','',''],
  ['REC-0050','Vegetable Oil','Condiments & Sauces',0.5,'cup',32,3.99,false,'Vegetarian','','Coconut oil','',''],
  ['REC-0050','Eggs','Dairy & Eggs',3,'piece',12,3.99,false,'Vegetarian','Eggs','','',''],
  ['REC-0050','Baking Soda','Baking',1,'tsp',8,1.79,false,'Vegetarian','','','',''],
  ['REC-0050','Baking Powder','Baking',0.5,'tsp',8,2.99,false,'Vegetarian','','','',''],
  ['REC-0050','Cinnamon','Spices & Seasonings',2,'tsp',2,2.99,false,'Vegetarian','','','',''],
  ['REC-0050','Salt','Spices & Seasonings',0.5,'tsp',26,1.29,false,'Vegetarian','','','',''],
  // REC-0051 Turkey Meatballs (8)
  ['REC-0051','Ground Turkey','Meat & Seafood',1.5,'lb',1,5.99,false,'','','Ground chicken','',''],
  ['REC-0051','Bread Crumbs','Grains & Pasta',0.5,'cup',16,2.49,false,'','Wheat','GF crumbs','',''],
  ['REC-0051','Parmesan','Dairy & Eggs',0.25,'cup',8,5.99,false,'','Dairy','','Grated',''],
  ['REC-0051','Egg','Dairy & Eggs',1,'piece',12,3.99,false,'','Eggs','','',''],
  ['REC-0051','Garlic','Produce',3,'clove',3,0.79,false,'','','','Minced',''],
  ['REC-0051','Italian Seasoning','Spices & Seasonings',1,'tsp',4,2.99,false,'','','','',''],
  ['REC-0051','Onion Powder','Spices & Seasonings',0.5,'tsp',4,2.49,false,'','','','',''],
  ['REC-0051','Olive Oil','Condiments & Sauces',2,'tbsp',17,7.99,false,'','','','For baking sheet',''],
  // REC-0052 Chicken Burrito Bowls (9)
  ['REC-0052','Chicken Breast','Meat & Seafood',1.5,'lb',4,5.99,false,'Gluten-Free','','','',''],
  ['REC-0052','Jasmine Rice','Grains & Pasta',2.5,'cup',4,3.99,false,'Gluten-Free','','','',''],
  ['REC-0052','Black Beans','Canned Goods',15,'oz',2,1.29,false,'Gluten-Free','','','Drained',''],
  ['REC-0052','Corn','Frozen',1,'cup',12,2.99,false,'Gluten-Free','','','',''],
  ['REC-0052','Lime','Produce',2,'piece',8,0.69,false,'Gluten-Free','','','Juiced',''],
  ['REC-0052','Olive Oil','Condiments & Sauces',2,'tbsp',17,7.99,false,'Gluten-Free','','','',''],
  ['REC-0052','Taco Seasoning','Spices & Seasonings',2,'tbsp',2,1.99,false,'Gluten-Free','','Homemade','',''],
  ['REC-0052','Salsa','Condiments & Sauces',0.5,'cup',16,3.99,false,'Gluten-Free','','','','For serving'],
  ['REC-0052','Avocado','Produce',2,'piece',3,1.29,false,'Gluten-Free','','','Sliced, for topping',''],
  // REC-0053 Freezer Breakfast Burritos (9)
  ['REC-0053','Flour Tortillas','Grains & Pasta',8,'piece',10,3.49,false,'','Wheat','','Large size',''],
  ['REC-0053','Eggs','Dairy & Eggs',8,'piece',12,3.99,false,'','Eggs','','Scrambled',''],
  ['REC-0053','Breakfast Sausage','Meat & Seafood',1,'lb',1,4.99,false,'','','Turkey sausage','Cooked, crumbled',''],
  ['REC-0053','Cheddar Cheese','Dairy & Eggs',1,'cup',16,4.99,false,'','Dairy','','Shredded',''],
  ['REC-0053','Potatoes','Produce',2,'piece',12,1.49,false,'','','','Diced, cooked',''],
  ['REC-0053','Bell Pepper','Produce',1,'piece',3,1.29,false,'','','','Diced',''],
  ['REC-0053','Onion','Produce',0.5,'piece',3,0.79,false,'','','','Diced',''],
  ['REC-0053','Salsa','Condiments & Sauces',0.5,'cup',16,3.99,false,'','','','',''],
  ['REC-0053','Salt','Spices & Seasonings',0.5,'tsp',26,1.29,false,'','','','',''],
  // REC-0054 Slow Cooker Chili (11)
  ['REC-0054','Ground Beef','Meat & Seafood',1.5,'lb',1,5.99,false,'Gluten-Free','','Ground turkey','',''],
  ['REC-0054','Kidney Beans','Canned Goods',30,'oz',2,1.29,false,'Gluten-Free','','Black beans','Drained','Two 15-oz cans'],
  ['REC-0054','Canned Tomatoes','Canned Goods',28,'oz',2,2.49,false,'Gluten-Free','','','',''],
  ['REC-0054','Tomato Sauce','Canned Goods',8,'oz',2,1.49,false,'Gluten-Free','','','',''],
  ['REC-0054','Onion','Produce',1,'piece',3,0.79,false,'Gluten-Free','','','Diced',''],
  ['REC-0054','Bell Pepper','Produce',1,'piece',3,1.29,false,'Gluten-Free','','','Diced',''],
  ['REC-0054','Garlic','Produce',4,'clove',3,0.79,false,'Gluten-Free','','','Minced',''],
  ['REC-0054','Chili Powder','Spices & Seasonings',2,'tbsp',2,2.49,false,'Gluten-Free','','','',''],
  ['REC-0054','Cumin','Spices & Seasonings',1,'tsp',2,2.99,false,'Gluten-Free','','','',''],
  ['REC-0054','Smoked Paprika','Spices & Seasonings',1,'tsp',2,2.99,false,'Gluten-Free','','','',''],
  ['REC-0054','Jalapeño','Produce',1,'piece',4,0.49,true,'Gluten-Free','','Omit for mild','Seeded',''],
  // REC-0055 Meal Prep Veggie Bowls (10)
  ['REC-0055','Sweet Potato','Produce',2,'piece',12,1.49,false,'Vegan','','','Cubed',''],
  ['REC-0055','Broccoli','Produce',1,'head',3,2.49,false,'Vegan','','','Florets',''],
  ['REC-0055','Chickpeas','Canned Goods',15,'oz',2,1.29,false,'Vegan','','','Drained',''],
  ['REC-0055','Quinoa','Grains & Pasta',1.5,'cup',4,5.99,false,'Vegan','','Brown rice','',''],
  ['REC-0055','Olive Oil','Condiments & Sauces',3,'tbsp',17,7.99,false,'Vegan','','','',''],
  ['REC-0055','Lemon','Produce',1,'piece',8,0.99,false,'Vegan','','','Juiced',''],
  ['REC-0055','Tahini','Condiments & Sauces',3,'tbsp',16,7.99,false,'Vegan','Sesame','','For dressing',''],
  ['REC-0055','Garlic','Produce',2,'clove',3,0.79,false,'Vegan','','','Minced',''],
  ['REC-0055','Cumin','Spices & Seasonings',1,'tsp',2,2.99,false,'Vegan','','','',''],
  ['REC-0055','Cherry Tomatoes','Produce',1,'cup',6,3.49,false,'Vegan','','','Halved, for topping',''],
  // REC-0056 Asian Sesame Salad (8)
  ['REC-0056','Napa Cabbage','Produce',0.5,'head',3,1.99,false,'Vegan','','','Shredded',''],
  ['REC-0056','Red Cabbage','Produce',1,'cup',3,1.49,false,'Vegan','','','Shredded',''],
  ['REC-0056','Carrots','Produce',2,'piece',12,1.49,false,'Vegan','','','Julienned',''],
  ['REC-0056','Edamame','Frozen',1,'cup',12,3.49,false,'Vegan','Soy','','Shelled',''],
  ['REC-0056','Sesame Oil','Condiments & Sauces',2,'tbsp',8,4.99,false,'Vegan','Sesame','','For dressing',''],
  ['REC-0056','Soy Sauce','Condiments & Sauces',3,'tbsp',20,3.99,false,'Vegan','Soy','Tamari','',''],
  ['REC-0056','Rice Vinegar','Condiments & Sauces',2,'tbsp',16,3.49,false,'Vegan','','','',''],
  ['REC-0056','Sesame Seeds','Spices & Seasonings',2,'tbsp',8,2.99,false,'Vegan','Sesame','','Toasted',''],
  // REC-0057 Waldorf Salad (8)
  ['REC-0057','Granny Smith Apples','Produce',2,'piece',12,0.79,false,'Vegetarian','','','Diced',''],
  ['REC-0057','Celery','Produce',3,'piece',12,1.99,false,'Vegetarian','','','Sliced',''],
  ['REC-0057','Walnuts','Pantry',0.5,'cup',16,7.99,false,'Vegetarian','Tree Nuts','Pecans','Chopped',''],
  ['REC-0057','Red Grapes','Produce',1,'cup',6,3.49,false,'Vegetarian','','','Halved',''],
  ['REC-0057','Mayonnaise','Condiments & Sauces',0.5,'cup',32,4.49,false,'Vegetarian','Eggs','Greek yogurt','',''],
  ['REC-0057','Lemon Juice','Produce',1,'tbsp',8,0.99,false,'Vegetarian','','','Fresh',''],
  ['REC-0057','Lettuce','Produce',4,'leaf',12,1.99,false,'Vegetarian','','','For serving',''],
  ['REC-0057','Salt','Spices & Seasonings',0.25,'tsp',26,1.29,false,'Vegetarian','','','',''],
  // REC-0058 Greek Salad (7)
  ['REC-0058','Cucumbers','Produce',2,'piece',3,0.99,false,'Vegetarian','','','Diced',''],
  ['REC-0058','Tomatoes','Produce',3,'piece',3,0.99,false,'Vegetarian','','','Diced',''],
  ['REC-0058','Kalamata Olives','Canned Goods',0.5,'cup',12,3.99,false,'Vegetarian','','','',''],
  ['REC-0058','Feta Cheese','Dairy & Eggs',1,'cup',8,4.99,false,'Vegetarian','Dairy','Omit for vegan','Crumbled',''],
  ['REC-0058','Red Onion','Produce',0.5,'piece',3,0.69,false,'Vegetarian','','','Thinly sliced',''],
  ['REC-0058','Olive Oil','Condiments & Sauces',3,'tbsp',17,7.99,false,'Vegetarian','','','',''],
  ['REC-0058','Dried Oregano','Spices & Seasonings',1,'tsp',2,2.49,false,'Vegetarian','','','',''],
  // REC-0059 Homemade Marinara Sauce (7)
  ['REC-0059','San Marzano Tomatoes','Canned Goods',28,'oz',2,3.49,false,'Vegan','','Regular tomatoes','',''],
  ['REC-0059','Garlic','Produce',5,'clove',3,0.79,false,'Vegan','','','Minced',''],
  ['REC-0059','Olive Oil','Condiments & Sauces',3,'tbsp',17,7.99,false,'Vegan','','','',''],
  ['REC-0059','Fresh Basil','Produce',0.5,'bunch',1,2.99,false,'Vegan','','Dried basil 2 tsp','',''],
  ['REC-0059','Sugar','Baking',1,'tsp',4,2.49,false,'Vegan','','','Pinch to balance acid',''],
  ['REC-0059','Salt','Spices & Seasonings',1,'tsp',26,1.29,false,'Vegan','','','',''],
  ['REC-0059','Red Pepper Flakes','Spices & Seasonings',0.25,'tsp',2,2.99,true,'Vegan','','','',''],
  // REC-0060 Mango Lassi (5)
  ['REC-0060','Mango','Frozen',2,'cup',12,4.99,false,'Vegetarian','','','Frozen or fresh',''],
  ['REC-0060','Greek Yogurt','Dairy & Eggs',1,'cup',4,5.49,false,'Vegetarian','Dairy','Dairy-free yogurt','','Plain'],
  ['REC-0060','Milk','Dairy & Eggs',0.5,'cup',0.5,2.49,false,'Vegetarian','Dairy','Almond milk','',''],
  ['REC-0060','Honey','Condiments & Sauces',1,'tbsp',24,6.99,false,'Vegetarian','','Maple syrup','To taste',''],
  ['REC-0060','Cardamom','Spices & Seasonings',0.25,'tsp',2,3.49,true,'Vegetarian','','','','Traditional spice'],
];

(async () => {
  const reqs = [];
  const vals = [];

  // ── Title row ────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 17), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  } });
  vals.push({ range: `${S}!A1`, values: [['🧂 RECIPE INGREDIENTS']] });

  // ── Subtitle ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 17), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A2`, values: [['Enter each ingredient on its own row. Link ingredients to a recipe by entering the Recipe ID in column B.']] });

  // Stats row
  const STAT_LABELS = [
    ['Total Ingredients', `=COUNTA($D$8:$D$12007)`],
    ['Recipes with Ingredients', `=SUMPRODUCT(IFERROR((LEN($B$8:$B$12007)>0)*(1/COUNTIF($B$8:$B$12007,$B$8:$B$12007)),0))`],
  ];
  STAT_LABELS.forEach(([lbl, frm], i) => {
    const col = i * 2;
    reqs.push({ repeatCell: {
      range: gridRange(SID, 2, 3, col, col+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.info),
        textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${i===0?'A':'C'}3`, values: [[lbl]] });
    reqs.push({ repeatCell: {
      range: gridRange(SID, 3, 4, col, col+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.formula),
        textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.secondary), fontFamily: 'Arial' },
        horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    } });
    vals.push({ range: `${S}!${i===0?'A':'C'}4`, values: [[frm]] });
  });

  // Rows 4-6 heights
  [2,3,4,5,6].forEach(i => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize',
    } });
  });

  // ── Empty rows 5-6 ─────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 4, 7, 0, 17),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Header row 7 ────────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 17),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondaryDeep),
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

  // ── Column widths ────────────────────────────────────────────────────────────
  const CW = [90,80,160,160,120,70,60,80,85,80,80,65,110,100,150,130,140];
  CW.forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    } });
  });

  // ── Data validation ──────────────────────────────────────────────────────────
  // G: Unit
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 12007, 6, 7),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!J4:J27` }] }, strict: false, showCustomUi: true },
  } });
  // E: Ingredient Category
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 12007, 4, 5),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!A51:A62` }] }, strict: false, showCustomUi: true },
  } });
  // L: Optional checkbox
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 12007, 11, 12),
    rule: { condition: { type: 'BOOLEAN' }, strict: true },
  } });
  // M: Dietary Tag
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 12007, 12, 13),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!D4:D18` }] }, strict: false, showCustomUi: true },
  } });
  // N: Allergen Tag
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 12007, 13, 14),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!E4:E14` }] }, strict: false, showCustomUi: true },
  } });

  // ── Freeze ───────────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } },
    fields: 'gridProperties.frozenRowCount',
  } });

  // ── Row height for data rows ─────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 12007 },
    properties: { pixelSize: 21 }, fields: 'pixelSize',
  } });

  // ── Write data rows ──────────────────────────────────────────────────────────
  const dataRows = INGREDIENTS.map((ing, i) => {
    const r = 8 + i;
    const [recipeId, name, category, qty, unit, pkgQty, pkgCost, optional,
           dietTag, allerTag, sub, prepNote, notes] = ing;
    return [
      `=IF(OR(B${r}="",D${r}=""),"","ING-"&TEXT(ROW()-7,"00000"))`, // A
      recipeId, // B
      `=IFERROR(VLOOKUP(B${r},'Master Recipe Index'!$A$8:$B$2007,2,0),"")`, // C Recipe Name
      name, category, qty, unit, pkgQty, pkgCost,
      `=IFERROR(I${r}/H${r},0)`, // J Est Unit Cost
      `=IFERROR(F${r}*J${r},0)`, // K Ingredient Cost
      optional,
      dietTag, allerTag, sub, prepNote, notes,
    ];
  });

  // Style data rows
  INGREDIENTS.forEach((_, i) => {
    const rowIdx = 7 + i;
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(SID, rowIdx, rowIdx+1, 0, 17),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } });
    // Formula cells A, C, J, K
    [0, 2, 9, 10].forEach(ci => {
      reqs.push({ repeatCell: {
        range: gridRange(SID, rowIdx, rowIdx+1, ci, ci+1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.formula),
          textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secondaryDeep) },
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      } });
    });
  });

  // Currency format for cost columns
  reqs.push({ repeatCell: {
    range: gridRange(SID, 7, 12007, 8, 12), // I, J, K cols
    cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } } },
    fields: 'userEnteredFormat(numberFormat)',
  } });

  await batchUpdate(id, reqs, 'ing-fmt');
  await valuesBatchUpdate(id, vals, 'ing-vals');
  await valuesBatchUpdate(id, [{ range: `${S}!A8`, values: dataRows }], 'ing-data');
  console.log(`✓ Recipe Ingredients — ${INGREDIENTS.length} rows written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
