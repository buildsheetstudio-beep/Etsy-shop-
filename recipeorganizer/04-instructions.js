'use strict';
const { hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Recipe Instructions'];
const S = "'Recipe Instructions'";
const REF = "'Reference Data'";

const HEADERS = [
  'Instruction ID','Recipe ID','Recipe Name','Step #','Instruction Type',
  'Instruction','Est. Minutes','Equipment / Tool','Make Ahead?','Notes',
];

// [RecipeID, StepNum, Type, Instruction, EstMinutes, Tool, MakeAhead, Notes]
const STEPS = [
  // REC-0001 Classic Buttermilk Pancakes
  ['REC-0001',1,'Prep','In a large bowl, whisk together flour, sugar, baking powder, baking soda, and salt.',3,'Large mixing bowl',false,''],
  ['REC-0001',2,'Prep','In a separate bowl, whisk buttermilk, eggs, and melted butter together.',2,'Medium bowl',false,''],
  ['REC-0001',3,'Cook','Pour wet ingredients into dry ingredients and stir until just combined. Do not overmix; lumps are fine.',1,'',false,'Overmixing makes pancakes tough'],
  ['REC-0001',4,'Cook','Heat a non-stick skillet or griddle over medium heat. Lightly butter the surface.',2,'Non-stick skillet or griddle',false,''],
  ['REC-0001',5,'Cook','Pour 1/4 cup batter per pancake. Cook until bubbles form on the surface and edges look set, about 2-3 minutes.',3,'Measuring cup',false,''],
  ['REC-0001',6,'Cook','Flip and cook the other side until golden brown, about 1-2 minutes more.',2,'Spatula',false,''],
  ['REC-0001',7,'Serve','Serve warm with maple syrup, fresh berries, or whipped butter.',1,'',false,''],
  // REC-0002 Spinach & Feta Omelette
  ['REC-0002',1,'Prep','Crack eggs into a bowl, season with salt and pepper, and whisk well.',2,'Bowl, whisk',false,''],
  ['REC-0002',2,'Cook','Heat olive oil in a non-stick pan over medium heat. Add minced garlic and cook 30 seconds.',1,'Non-stick pan',false,''],
  ['REC-0002',3,'Cook','Add spinach and cook until just wilted, about 1 minute.',1,'',false,''],
  ['REC-0002',4,'Cook','Pour egg mixture over spinach. Tilt pan to spread eggs evenly.',1,'',false,''],
  ['REC-0002',5,'Cook','When eggs are mostly set, add crumbled feta to one half. Fold omelette in half.',2,'Spatula',false,''],
  ['REC-0002',6,'Serve','Slide onto plate and serve immediately.',1,'',false,''],
  // REC-0003 Blueberry Overnight Oats
  ['REC-0003',1,'Prep','Add rolled oats, almond milk, Greek yogurt, honey, and chia seeds to a jar or container.',3,'Mason jar or container',true,''],
  ['REC-0003',2,'Prep','Stir well to combine all ingredients.',1,'Spoon',true,''],
  ['REC-0003',3,'Chill','Cover and refrigerate overnight, or at least 4 hours.',0,'Refrigerator',true,''],
  ['REC-0003',4,'Prep','In the morning, stir oats. Add a splash of milk if too thick.',1,'',false,''],
  ['REC-0003',5,'Serve','Top with fresh blueberries and an extra drizzle of honey. Serve cold.',1,'',false,''],
  // REC-0004 Avocado Toast with Poached Eggs
  ['REC-0004',1,'Cook','Bring a pot of water to a gentle simmer. Add a splash of white vinegar.',3,'Small pot',false,''],
  ['REC-0004',2,'Prep','Toast bread until golden. While bread toasts, mash avocado with lemon juice and a pinch of salt.',3,'Toaster',false,''],
  ['REC-0004',3,'Cook','Create a gentle swirl in the water. Crack each egg into a small cup and slide into the water.',4,'Small cups',false,'Vinegar helps whites hold together'],
  ['REC-0004',4,'Cook','Poach eggs 3-4 minutes for runny yolk. Remove with a slotted spoon.',1,'Slotted spoon',false,''],
  ['REC-0004',5,'Assemble','Spread mashed avocado on toast. Top with poached eggs.',1,'',false,''],
  ['REC-0004',6,'Serve','Sprinkle with flaky salt and red pepper flakes. Serve immediately.',1,'',false,''],
  // REC-0005 Banana Walnut Muffins
  ['REC-0005',1,'Prep','Preheat oven to 375°F. Line a 12-cup muffin tin with paper liners.',5,'Muffin tin, liners',false,''],
  ['REC-0005',2,'Prep','In a large bowl, mash bananas until smooth.',3,'Large bowl',true,''],
  ['REC-0005',3,'Prep','Whisk brown sugar, melted butter, eggs, and vanilla into mashed bananas.',2,'Whisk',false,''],
  ['REC-0005',4,'Prep','In a separate bowl, combine flour, baking soda, cinnamon, and salt.',2,'Medium bowl',false,''],
  ['REC-0005',5,'Prep','Fold dry ingredients into wet ingredients until just combined. Stir in walnuts.',2,'Spatula',false,''],
  ['REC-0005',6,'Bake','Divide batter evenly among muffin cups. Bake 20-25 minutes until a toothpick comes out clean.',25,'Toothpick',false,''],
  ['REC-0005',7,'Rest','Cool in pan 5 minutes, then transfer to a wire rack.',5,'Wire rack',false,''],
  // REC-0006 French Toast Casserole
  ['REC-0006',1,'Prep','Butter a 9x13 baking dish. Spread cubed brioche in an even layer.',5,'9x13 baking dish',true,''],
  ['REC-0006',2,'Prep','Whisk together eggs, milk, heavy cream, brown sugar, vanilla, and cinnamon.',3,'Large bowl',true,''],
  ['REC-0006',3,'Chill','Pour custard over bread. Dot with butter, cover and refrigerate overnight.',0,'',true,''],
  ['REC-0006',4,'Bake','Preheat oven to 350°F. Remove casserole from refrigerator 30 minutes before baking.',30,'Oven',false,''],
  ['REC-0006',5,'Bake','Bake uncovered 45-50 minutes until puffed, golden, and set in the center.',50,'',false,''],
  ['REC-0006',6,'Serve','Dust with powdered sugar and serve with maple syrup.',2,'Sifter',false,''],
  // REC-0007 Greek Yogurt Parfait
  ['REC-0007',1,'Prep','Select 2 tall glasses or bowls for serving.',1,'Glasses or bowls',false,''],
  ['REC-0007',2,'Assemble','Spoon half the Greek yogurt evenly into the bottom of each glass.',1,'Spoon',false,''],
  ['REC-0007',3,'Assemble','Add a layer of granola over the yogurt.',1,'',false,''],
  ['REC-0007',4,'Assemble','Add a layer of strawberries and blueberries.',1,'',false,''],
  ['REC-0007',5,'Serve','Drizzle with honey and serve immediately, or refrigerate up to 1 hour before serving.',1,'',false,''],
  // REC-0008 Shakshuka
  ['REC-0008',1,'Cook','Heat olive oil in a large skillet over medium heat. Add diced onion and cook until softened, about 5 minutes.',5,'Large skillet',false,''],
  ['REC-0008',2,'Cook','Add bell pepper and cook 3 more minutes. Add garlic, cumin, and paprika; cook 1 minute until fragrant.',4,'',false,''],
  ['REC-0008',3,'Cook','Pour in crushed tomatoes. Season with salt. Simmer 10 minutes, stirring occasionally.',10,'',false,''],
  ['REC-0008',4,'Cook','Using a spoon, create 4 wells in the sauce. Crack an egg into each well.',2,'Spoon',false,''],
  ['REC-0008',5,'Cook','Cover and cook 5-8 minutes until egg whites are set but yolks are still runny.',7,'Lid',false,''],
  ['REC-0008',6,'Serve','Garnish with fresh herbs and crumbled feta if desired. Serve with crusty bread.',2,'',false,''],
  // REC-0009 Classic Caesar Salad
  ['REC-0009',1,'Prep','Wash and dry romaine lettuce. Chop into bite-sized pieces.',5,'Salad spinner',false,''],
  ['REC-0009',2,'Prep','In a large bowl, whisk together Caesar dressing and lemon juice.',2,'Large bowl, whisk',false,''],
  ['REC-0009',3,'Assemble','Add romaine to the bowl and toss to coat with dressing.',2,'Tongs',false,''],
  ['REC-0009',4,'Serve','Divide onto plates. Top with croutons, shaved parmesan, and freshly ground pepper.',2,'Vegetable peeler for parmesan',false,''],
  // REC-0010 Turkey & Avocado Wrap
  ['REC-0010',1,'Prep','Lay tortilla flat. Spread mustard evenly across the surface.',1,'',false,''],
  ['REC-0010',2,'Assemble','Layer turkey, avocado slices, lettuce, tomato, and red onion down the center.',2,'',false,''],
  ['REC-0010',3,'Assemble','Fold in the sides of the tortilla, then roll up tightly from the bottom.',1,'',false,''],
  ['REC-0010',4,'Serve','Cut in half on a diagonal and serve immediately.',1,'Sharp knife',false,''],
  // REC-0011 Tomato Basil Soup
  ['REC-0011',1,'Cook','Heat olive oil in a large pot. Add diced onion and cook until translucent, about 5 minutes.',5,'Large pot',false,''],
  ['REC-0011',2,'Cook','Add garlic and cook 1 minute until fragrant.',1,'',false,''],
  ['REC-0011',3,'Cook','Pour in crushed tomatoes and vegetable broth. Bring to a boil, reduce heat, simmer 20 minutes.',20,'',true,''],
  ['REC-0011',4,'Prep','Add fresh basil leaves. Using an immersion blender, blend about half the soup for a chunky-smooth texture.',3,'Immersion blender',false,''],
  ['REC-0011',5,'Cook','Season with salt, sugar, and black pepper to taste.',2,'',false,''],
  ['REC-0011',6,'Serve','Ladle into bowls. Drizzle with olive oil and garnish with fresh basil.',2,'',false,''],
  // REC-0012 Mediterranean Quinoa Bowl
  ['REC-0012',1,'Cook','Rinse quinoa under cold water. Add to a pot with 2 cups water and a pinch of salt.',1,'Medium pot, strainer',false,''],
  ['REC-0012',2,'Cook','Bring to a boil, reduce heat, cover and cook 15 minutes until water is absorbed. Let sit 5 minutes, then fluff.',20,'',true,''],
  ['REC-0012',3,'Prep','While quinoa cooks, dice cucumber, halve cherry tomatoes, dice red onion, and drain olives.',10,'',false,''],
  ['REC-0012',4,'Prep','In a small bowl, whisk together olive oil, lemon juice, and a pinch of salt and pepper.',2,'Small bowl',false,''],
  ['REC-0012',5,'Assemble','Combine cooled quinoa, vegetables, and feta in a large bowl. Pour dressing over and toss.',3,'Large bowl, tongs',false,''],
  ['REC-0012',6,'Serve','Garnish with fresh mint. Serve at room temperature or chilled.',1,'',false,''],
  // REC-0013 Chicken BLT Sandwich
  ['REC-0013',1,'Cook','Season chicken breast with garlic powder, salt, and pepper. Cook in a skillet over medium-high heat 6-7 minutes per side until cooked through.',15,'Skillet',false,''],
  ['REC-0013',2,'Cook','While chicken cooks, cook bacon in the same skillet until crispy. Drain on paper towels.',8,'Paper towels',false,''],
  ['REC-0013',3,'Prep','Toast bread slices until golden.',3,'Toaster',false,''],
  ['REC-0013',4,'Assemble','Spread mayonnaise on both bread slices. Layer lettuce, tomato, chicken, bacon, and avocado.',2,'',false,''],
  ['REC-0013',5,'Serve','Close sandwich and cut in half. Serve immediately.',1,'',false,''],
  // REC-0014 Lentil Soup
  ['REC-0014',1,'Prep','Rinse lentils under cold water and drain.',2,'Strainer',false,''],
  ['REC-0014',2,'Cook','Heat olive oil in a large pot. Sauté onion, carrot, and celery until softened, about 8 minutes.',8,'Large pot',false,''],
  ['REC-0014',3,'Cook','Add garlic and cumin; cook 1 minute. Add lentils, broth, and 2 cups water. Bring to boil.',3,'',false,''],
  ['REC-0014',4,'Cook','Reduce heat and simmer 25-30 minutes until lentils are tender.',30,'',true,''],
  ['REC-0014',5,'Prep','Squeeze lemon juice into soup. Season with salt and pepper to taste.',2,'',false,''],
  ['REC-0014',6,'Serve','Serve hot, optionally garnished with a drizzle of olive oil and fresh parsley.',2,'',false,''],
  // REC-0015 Caprese Salad
  ['REC-0015',1,'Prep','Slice mozzarella and tomatoes into 1/4-inch rounds.',5,'Sharp knife',false,''],
  ['REC-0015',2,'Assemble','Alternate slices of mozzarella, tomato, and fresh basil leaves on a serving platter.',3,'Serving platter',false,''],
  ['REC-0015',3,'Serve','Drizzle generously with extra virgin olive oil and balsamic glaze. Season with flaky salt.',2,'',false,''],
  // REC-0016 Spaghetti Bolognese
  ['REC-0016',1,'Cook','Heat olive oil in a large, heavy-bottomed pot. Add diced onion, carrot, and celery. Cook over medium heat until softened, about 10 minutes.',10,'Large Dutch oven',false,''],
  ['REC-0016',2,'Cook','Add garlic and cook 1 minute. Add ground beef and cook, breaking it up, until browned.',8,'Wooden spoon',false,''],
  ['REC-0016',3,'Cook','Add red wine and cook until evaporated, about 3 minutes.',3,'',false,''],
  ['REC-0016',4,'Cook','Add crushed tomatoes, bay leaves, salt, and pepper. Bring to a simmer.',3,'',false,''],
  ['REC-0016',5,'Cook','Reduce heat to very low and cook, partially covered, for at least 1 hour. Stir occasionally.',60,'',true,''],
  ['REC-0016',6,'Cook','Cook spaghetti in salted boiling water according to package directions until al dente. Drain.',12,'Large pot',false,''],
  ['REC-0016',7,'Serve','Toss pasta with sauce. Serve with freshly grated parmesan.',2,'',false,''],
  // REC-0017 Chicken Tikka Masala
  ['REC-0017',1,'Prep','Combine chicken pieces with yogurt, half the garam masala, turmeric, and salt. Marinate at least 30 minutes, or overnight.',30,'Bowl',true,''],
  ['REC-0017',2,'Cook','Heat butter in a large pan over medium-high. Cook marinated chicken until charred and cooked through, about 8 minutes. Set aside.',8,'Large skillet',false,''],
  ['REC-0017',3,'Cook','In the same pan, sauté diced onion until golden, about 8 minutes. Add garlic and ginger; cook 2 minutes.',10,'',false,''],
  ['REC-0017',4,'Cook','Add remaining garam masala and cumin. Cook 1 minute. Add canned tomatoes and simmer 15 minutes.',15,'',false,''],
  ['REC-0017',5,'Cook','Stir in heavy cream and cooked chicken. Simmer 10 minutes until sauce thickens.',10,'',false,''],
  ['REC-0017',6,'Serve','Garnish with cilantro and serve over basmati rice with naan bread.',2,'',false,''],
  // REC-0018 Sheet Pan Lemon Herb Salmon
  ['REC-0018',1,'Prep','Preheat oven to 400°F. Line a sheet pan with foil.',5,'Sheet pan, foil',false,''],
  ['REC-0018',2,'Prep','Toss asparagus and cherry tomatoes with 1 tbsp olive oil, salt, and pepper. Spread on pan.',3,'',false,''],
  ['REC-0018',3,'Prep','Place salmon fillets on the pan. Drizzle remaining olive oil over salmon. Season with salt, pepper, and garlic powder. Top with lemon slices and fresh dill.',3,'',false,''],
  ['REC-0018',4,'Bake','Bake 18-20 minutes until salmon flakes easily with a fork.',20,'',false,''],
  ['REC-0018',5,'Serve','Serve immediately with roasted vegetables and lemon wedges.',2,'',false,''],
  // REC-0019 Beef Tacos
  ['REC-0019',1,'Cook','Heat a large skillet over medium-high heat. Add ground beef and cook, breaking apart, until browned, about 8 minutes. Drain excess fat.',8,'Large skillet',false,''],
  ['REC-0019',2,'Cook','Add diced onion and cook 3 minutes. Add taco seasoning and 1/4 cup water. Stir and simmer 3 minutes.',6,'',false,''],
  ['REC-0019',3,'Prep','While meat cooks, set up toppings: shredded cheese, lettuce, diced tomato, sour cream, and salsa.',5,'',false,''],
  ['REC-0019',4,'Assemble','Warm taco shells in oven at 300°F for 3 minutes.',3,'Oven',false,''],
  ['REC-0019',5,'Assemble','Fill shells with seasoned beef and desired toppings. Squeeze fresh lime juice over filling.',2,'',false,''],
  ['REC-0019',6,'Serve','Serve immediately.',1,'',false,''],
  // REC-0020 Mushroom Risotto
  ['REC-0020',1,'Prep','Heat vegetable broth in a saucepan and keep warm over low heat.',5,'Small saucepan',false,''],
  ['REC-0020',2,'Cook','Melt butter in a large skillet over medium heat. Add sliced mushrooms and cook until golden, about 8 minutes. Set aside.',8,'Large skillet',false,''],
  ['REC-0020',3,'Cook','In the same pan, sauté onion and garlic until softened. Add arborio rice and toast 2 minutes, stirring.',5,'',false,''],
  ['REC-0020',4,'Cook','Add white wine and stir until absorbed. Begin adding warm broth one ladle at a time, stirring constantly and adding more as each portion absorbs.',25,'Ladle',false,''],
  ['REC-0020',5,'Cook','When rice is al dente and creamy, stir in mushrooms, remaining butter, and parmesan.',3,'',false,''],
  ['REC-0020',6,'Serve','Season with salt and pepper. Garnish with fresh thyme and extra parmesan. Serve immediately.',2,'',false,''],
  // REC-0021 Slow Cooker Pulled Pork
  ['REC-0021',1,'Prep','Mix together paprika, garlic powder, onion powder, brown sugar, salt, and pepper to make dry rub.',3,'Small bowl',true,''],
  ['REC-0021',2,'Prep','Rub the spice mixture all over the pork shoulder. For best results, cover and refrigerate overnight.',5,'',true,''],
  ['REC-0021',3,'Cook','Place rubbed pork in slow cooker. Pour in apple cider vinegar and half the BBQ sauce.',2,'Slow cooker',false,''],
  ['REC-0021',4,'Cook','Cook on LOW 8-10 hours, or HIGH 5-6 hours, until pork is very tender.',480,'',false,''],
  ['REC-0021',5,'Rest','Remove pork and shred with two forks. Return to cooker juices.',10,'Two forks',false,''],
  ['REC-0021',6,'Serve','Toss shredded pork with remaining BBQ sauce. Serve on buns with coleslaw.',5,'',false,''],
  // REC-0022 Thai Green Curry
  ['REC-0022',1,'Cook','Heat 1 tbsp coconut milk in a large wok or pan over medium-high heat. Add curry paste and cook 2 minutes until fragrant.',2,'Wok or large pan',false,''],
  ['REC-0022',2,'Cook','Add chicken pieces and stir to coat in paste. Cook 5 minutes until lightly browned.',5,'',false,''],
  ['REC-0022',3,'Cook','Pour in remaining coconut milk and bring to a simmer.',3,'',false,''],
  ['REC-0022',4,'Cook','Add zucchini and bell pepper. Cook 8 minutes until vegetables are tender and chicken is cooked through.',8,'',false,''],
  ['REC-0022',5,'Cook','Season with fish sauce and lime juice. Adjust curry paste to taste.',2,'',false,''],
  ['REC-0022',6,'Serve','Serve over jasmine rice and garnish with fresh basil leaves.',2,'',false,''],
  // REC-0023 Grilled Teriyaki Chicken
  ['REC-0023',1,'Prep','Whisk together soy sauce, honey, garlic, ginger, and sesame oil to make marinade.',3,'Bowl, whisk',true,''],
  ['REC-0023',2,'Prep','Place chicken in marinade. Cover and refrigerate at least 2 hours, preferably overnight.',2,'Bag or shallow dish',true,''],
  ['REC-0023',3,'Cook','Preheat grill to medium-high heat. Clean and oil grates.',5,'Grill',false,''],
  ['REC-0023',4,'Cook','Remove chicken from marinade. Grill 6-8 minutes per side until internal temperature reaches 165°F.',16,'Meat thermometer',false,''],
  ['REC-0023',5,'Rest','Let rest 5 minutes before slicing.',5,'',false,''],
  ['REC-0023',6,'Serve','Serve over rice, garnished with sliced green onions and sesame seeds.',2,'',false,''],
  // REC-0024 Beef Stew
  ['REC-0024',1,'Prep','Cut beef into 1.5-inch cubes. Season generously with salt and pepper.',5,'',false,''],
  ['REC-0024',2,'Cook','Heat oil in a large pot. Brown beef in batches, 2-3 minutes per side. Do not crowd the pan. Set aside.',12,'Large pot or Dutch oven',false,''],
  ['REC-0024',3,'Cook','In same pot, cook onion and garlic until softened. Add tomato paste and cook 2 minutes.',5,'',false,''],
  ['REC-0024',4,'Cook','Return beef to pot. Add broth, thyme, and bay leaves. Bring to a boil.',3,'',false,''],
  ['REC-0024',5,'Cook','Add potatoes and carrots in the last 45 minutes of cooking. Reduce heat and simmer 2 hours total until beef is tender.',120,'',false,''],
  ['REC-0024',6,'Cook','If desired, mix cornstarch with cold water and stir into stew to thicken.',3,'',false,''],
  ['REC-0024',7,'Serve','Remove bay leaves. Serve hot with crusty bread.',2,'',false,''],
  // REC-0025 Shrimp Scampi
  ['REC-0025',1,'Cook','Cook linguine in generously salted boiling water according to package directions. Reserve 1/2 cup pasta water. Drain.',12,'Large pot',false,''],
  ['REC-0025',2,'Cook','While pasta cooks, melt butter in a large skillet over medium-high heat. Add garlic and cook 1 minute.',1,'Large skillet',false,''],
  ['REC-0025',3,'Cook','Add shrimp in a single layer. Cook 1-2 minutes per side until pink and just cooked through. Remove shrimp.',3,'Tongs',false,''],
  ['REC-0025',4,'Cook','Add white wine to the pan, scraping up any bits. Simmer 2 minutes.',2,'',false,''],
  ['REC-0025',5,'Cook','Add lemon juice and a splash of pasta water. Return shrimp to pan. Toss with drained pasta.',2,'',false,''],
  ['REC-0025',6,'Serve','Garnish with parsley and red pepper flakes. Serve immediately.',1,'',false,''],
  // REC-0026 Vegetable Stir Fry
  ['REC-0026',1,'Prep','Mix soy sauce, cornstarch, and a splash of water in a small bowl to make sauce. Set aside.',2,'Small bowl',false,''],
  ['REC-0026',2,'Prep','Cut all vegetables into similar-sized pieces for even cooking.',8,'Sharp knife, cutting board',false,''],
  ['REC-0026',3,'Cook','Heat sesame oil in a wok or large skillet over high heat until smoking.',1,'Wok',false,'High heat is key'],
  ['REC-0026',4,'Cook','Add harder vegetables (carrots) first, cook 2 minutes. Add broccoli and snap peas, cook 2 minutes.',4,'',false,''],
  ['REC-0026',5,'Cook','Add garlic and ginger, stir 30 seconds. Add bell peppers and cook 1 minute.',2,'',false,''],
  ['REC-0026',6,'Cook','Pour sauce over vegetables and toss to coat. Cook 1-2 minutes until sauce thickens.',2,'',false,''],
  ['REC-0026',7,'Serve','Serve immediately over steamed rice.',1,'',false,''],
  // REC-0027 BBQ Baby Back Ribs
  ['REC-0027',1,'Prep','Remove membrane from the back of the ribs. Pat dry with paper towels.',5,'Paper towels',true,''],
  ['REC-0027',2,'Prep','Mix brown sugar, paprika, garlic powder, onion powder, and salt. Rub all over ribs. Wrap in foil and refrigerate overnight.',5,'Foil',true,''],
  ['REC-0027',3,'Bake','Preheat oven to 275°F. Place foil-wrapped ribs on a baking sheet. Bake 3 hours.',180,'Baking sheet',false,''],
  ['REC-0027',4,'Cook','Preheat grill to medium-high. Carefully unwrap ribs and brush with BBQ sauce.',5,'Grill, brush',false,''],
  ['REC-0027',5,'Cook','Grill ribs 5 minutes per side, brushing more sauce and getting a caramelized char.',10,'',false,''],
  ['REC-0027',6,'Rest','Let rest 5 minutes before cutting into individual ribs.',5,'',false,''],
  ['REC-0027',7,'Serve','Serve with remaining BBQ sauce, coleslaw, and corn.',2,'',false,''],
  // REC-0028 One Pot Chicken Pasta
  ['REC-0028',1,'Cook','Heat olive oil in a large pot over medium-high heat. Season chicken with salt, pepper, and Italian seasoning. Cook until golden, about 5 minutes per side. Remove and slice.',12,'Large pot',false,''],
  ['REC-0028',2,'Cook','In the same pot, sauté garlic 1 minute. Add sun-dried tomatoes and cook 2 minutes.',3,'',false,''],
  ['REC-0028',3,'Cook','Add pasta, chicken broth, and heavy cream. Bring to a boil, stirring frequently.',5,'',false,''],
  ['REC-0028',4,'Cook','Reduce heat and simmer 12-15 minutes, stirring occasionally, until pasta is cooked and sauce thickens.',15,'',false,''],
  ['REC-0028',5,'Cook','Stir in parmesan, spinach, and sliced chicken. Cook 2 minutes until spinach wilts.',3,'',false,''],
  ['REC-0028',6,'Serve','Serve immediately with extra parmesan.',1,'',false,''],
  // REC-0029 Korean Beef Bowl
  ['REC-0029',1,'Cook','Cook jasmine rice according to package instructions.',20,'Rice cooker or pot',false,''],
  ['REC-0029',2,'Prep','Whisk together soy sauce, sesame oil, brown sugar, garlic, and ginger in a small bowl.',3,'Small bowl',false,''],
  ['REC-0029',3,'Cook','Heat a large skillet over medium-high heat. Add ground beef and cook, breaking apart, until browned. Drain most of the fat.',8,'Large skillet',false,''],
  ['REC-0029',4,'Cook','Pour sauce over beef and stir to combine. Cook 2-3 minutes until sauce caramelizes slightly.',3,'',false,''],
  ['REC-0029',5,'Assemble','Divide rice into bowls. Top with Korean beef mixture.',2,'',false,''],
  ['REC-0029',6,'Serve','Garnish with sliced green onions and sesame seeds. Serve with kimchi if desired.',1,'',false,''],
  // REC-0030 Garlic Butter Shrimp
  ['REC-0030',1,'Prep','Pat shrimp dry with paper towels. Season with salt and pepper.',2,'Paper towels',false,''],
  ['REC-0030',2,'Cook','Melt 2 tbsp butter in a large skillet over medium-high heat. Add shrimp in a single layer.',1,'Large skillet',false,''],
  ['REC-0030',3,'Cook','Cook shrimp 1-2 minutes per side until pink and curled. Remove from pan.',3,'Tongs',false,''],
  ['REC-0030',4,'Cook','Reduce heat to medium. Add remaining butter and garlic. Cook 1 minute until fragrant.',2,'',false,''],
  ['REC-0030',5,'Cook','Return shrimp to pan. Add lemon juice and red pepper flakes. Toss to coat.',1,'',false,''],
  ['REC-0030',6,'Serve','Garnish with fresh parsley. Serve immediately with rice or bread.',1,'',false,''],
  // REC-0031 Homemade Guacamole
  ['REC-0031',1,'Prep','Cut avocados in half, remove pits, and scoop flesh into a bowl.',3,'Bowl',false,''],
  ['REC-0031',2,'Prep','Mash avocado with a fork to desired texture — leave it slightly chunky.',2,'Fork',false,''],
  ['REC-0031',3,'Prep','Finely dice red onion and jalapeño. Chop cilantro.',3,'Knife, cutting board',false,''],
  ['REC-0031',4,'Assemble','Mix onion, jalapeño, cilantro, and lime juice into avocado. Season with salt to taste.',2,'',false,''],
  ['REC-0031',5,'Serve','Serve immediately. If making ahead, press plastic wrap directly onto surface.',1,'Plastic wrap',false,''],
  // REC-0032 Stuffed Mushrooms
  ['REC-0032',1,'Prep','Preheat oven to 375°F. Remove stems from mushrooms and finely chop the stems.',5,'Oven',false,''],
  ['REC-0032',2,'Prep','Mix cream cheese, parmesan, garlic, Italian seasoning, chopped stems, parsley, and bread crumbs together.',3,'Bowl',false,''],
  ['REC-0032',3,'Assemble','Fill each mushroom cap generously with the cheese mixture.',5,'Small spoon',false,''],
  ['REC-0032',4,'Bake','Place filled mushrooms on a baking sheet. Bake 20-25 minutes until golden and bubbly.',25,'Baking sheet',false,''],
  ['REC-0032',5,'Serve','Serve warm as an appetizer.',1,'',false,''],
  // REC-0033 Spinach Artichoke Dip
  ['REC-0033',1,'Prep','Preheat oven to 375°F. Squeeze excess moisture from chopped spinach.',5,'Kitchen towel',false,''],
  ['REC-0033',2,'Prep','In a large bowl, mix cream cheese, sour cream, parmesan, mozzarella, garlic, and spinach.',3,'Large bowl',false,''],
  ['REC-0033',3,'Prep','Stir in chopped artichoke hearts. Season with salt and pepper.',2,'',false,''],
  ['REC-0033',4,'Bake','Transfer mixture to a baking dish. Bake 25 minutes until bubbly and golden on top.',25,'Baking dish',false,''],
  ['REC-0033',5,'Serve','Serve hot with pita chips, sliced baguette, or vegetable sticks.',2,'',false,''],
  // REC-0034 Caprese Bruschetta
  ['REC-0034',1,'Prep','Preheat broiler. Slice baguette into 1-inch pieces.',3,'Oven',false,''],
  ['REC-0034',2,'Cook','Broil baguette slices on a baking sheet 1-2 minutes per side until golden and crisp.',3,'Baking sheet',false,''],
  ['REC-0034',3,'Prep','Rub each toasted slice with the cut side of a garlic clove.',1,'',false,''],
  ['REC-0034',4,'Prep','Dice tomatoes and mozzarella. Drizzle with olive oil and season with salt.',3,'',false,''],
  ['REC-0034',5,'Assemble','Spoon tomato and mozzarella mixture onto toasted bread. Top with basil chiffonade.',3,'',false,''],
  ['REC-0034',6,'Serve','Drizzle with balsamic glaze just before serving.',1,'',false,''],
  // REC-0035 Crispy Air Fryer Chickpeas
  ['REC-0035',1,'Prep','Drain and rinse chickpeas. Spread on paper towels and pat very dry — this is key for crispiness.',5,'Paper towels',false,''],
  ['REC-0035',2,'Prep','Toss chickpeas with olive oil, cumin, paprika, and salt.',2,'Bowl',false,''],
  ['REC-0035',3,'Cook','Air fry at 400°F for 15-20 minutes, shaking basket every 5 minutes.',20,'Air fryer',false,''],
  ['REC-0035',4,'Serve','Serve immediately for maximum crispiness. Season with extra salt if needed.',1,'',false,''],
  // REC-0036 Chicken Noodle Soup
  ['REC-0036',1,'Cook','Place chicken thighs in a large pot with chicken broth. Add onion, carrots, celery, and garlic. Bring to a boil.',5,'Large pot',false,''],
  ['REC-0036',2,'Cook','Reduce heat and simmer 20-25 minutes until chicken is cooked through.',25,'',false,''],
  ['REC-0036',3,'Prep','Remove chicken, let cool slightly, then shred with two forks. Return to pot.',10,'Two forks',false,''],
  ['REC-0036',4,'Cook','Add egg noodles and cook according to package directions until tender.',10,'',false,''],
  ['REC-0036',5,'Cook','Season with thyme, salt, and pepper to taste.',2,'',false,''],
  ['REC-0036',6,'Serve','Ladle into bowls and garnish with fresh parsley.',2,'',false,''],
  // REC-0037 French Onion Soup
  ['REC-0037',1,'Cook','Melt butter in a large pot over medium heat. Add thinly sliced onions and a pinch of salt.',2,'Large heavy pot',false,''],
  ['REC-0037',2,'Cook','Cook onions, stirring occasionally, until deeply caramelized — 40-50 minutes. Do not rush this step.',50,'',false,'Patience is essential'],
  ['REC-0037',3,'Cook','Add thyme and bay leaf. Pour in white wine and stir, scraping up any browned bits. Simmer 3 minutes.',5,'',false,''],
  ['REC-0037',4,'Cook','Add beef broth. Bring to a boil, reduce heat and simmer 20 minutes. Remove thyme sprigs and bay leaf.',20,'',false,''],
  ['REC-0037',5,'Cook','Preheat broiler. Ladle soup into oven-safe bowls. Float a slice of baguette on top of each.',3,'Oven-safe bowls',false,''],
  ['REC-0037',6,'Cook','Cover bread with shredded gruyere. Broil 2-3 minutes until cheese is melted and bubbling.',3,'',false,''],
  ['REC-0037',7,'Serve','Serve immediately and be careful — the bowls are very hot!',1,'',false,''],
  // REC-0038 Minestrone
  ['REC-0038',1,'Cook','Heat olive oil in a large pot. Sauté onion, celery, and carrot until softened, about 8 minutes.',8,'Large pot',false,''],
  ['REC-0038',2,'Cook','Add garlic and Italian seasoning; cook 1 minute. Add diced tomatoes, broth, and 2 cups water. Bring to boil.',5,'',false,''],
  ['REC-0038',3,'Cook','Add zucchini and green beans. Reduce heat and simmer 10 minutes.',10,'',true,''],
  ['REC-0038',4,'Cook','Add pasta and cannellini beans. Cook 8-10 minutes until pasta is tender.',10,'',false,'Add pasta at the end only'],
  ['REC-0038',5,'Cook','Season with salt and pepper. Add more broth if too thick.',2,'',false,''],
  ['REC-0038',6,'Serve','Serve with a drizzle of olive oil and grated parmesan if desired.',2,'',false,''],
  // REC-0039 Butternut Squash Soup
  ['REC-0039',1,'Prep','Preheat oven to 400°F. Toss cubed squash with 1 tbsp olive oil, salt, and pepper.',5,'Sheet pan',false,''],
  ['REC-0039',2,'Bake','Roast squash for 25-30 minutes until tender and caramelized at the edges.',30,'',true,''],
  ['REC-0039',3,'Cook','While squash roasts, heat remaining oil in a pot. Cook onion and garlic until softened.',8,'Large pot',false,''],
  ['REC-0039',4,'Cook','Add roasted squash and vegetable broth. Bring to a simmer.',5,'',false,''],
  ['REC-0039',5,'Prep','Blend soup with an immersion blender until completely smooth.',3,'Immersion blender',false,''],
  ['REC-0039',6,'Cook','Stir in heavy cream and nutmeg. Season with salt and pepper. Heat through.',3,'',false,''],
  ['REC-0039',7,'Serve','Serve warm with crusty bread and a swirl of cream.',2,'',false,''],
  // REC-0040 Chocolate Chip Cookies
  ['REC-0040',1,'Prep','Brown butter in a light-colored saucepan over medium heat until golden and nutty-smelling. Let cool.',10,'Light saucepan',false,''],
  ['REC-0040',2,'Prep','Whisk both sugars and browned butter together until combined.',2,'Large bowl, whisk',false,''],
  ['REC-0040',3,'Prep','Whisk in eggs and vanilla. Let mixture rest 5 minutes, then whisk again for shininess.',5,'',false,''],
  ['REC-0040',4,'Prep','Fold in flour, baking soda, and salt until just combined. Stir in chocolate chips.',2,'Spatula',false,''],
  ['REC-0040',5,'Chill','Cover dough and refrigerate at least 1 hour for best flavor, or overnight.',60,'',true,''],
  ['REC-0040',6,'Bake','Preheat oven to 375°F. Scoop dough onto parchment-lined baking sheets, 2 inches apart.',5,'Baking sheets, parchment',false,''],
  ['REC-0040',7,'Bake','Bake 10-12 minutes until edges are set but centers look slightly underdone.',12,'',false,'They firm up as they cool'],
  ['REC-0040',8,'Rest','Cool on baking sheet 5 minutes before transferring to a wire rack.',5,'Wire rack',false,''],
  // REC-0041 Classic Cheesecake
  ['REC-0041',1,'Prep','Preheat oven to 325°F. Wrap the outside of a 9-inch springform pan with foil.',5,'Springform pan, foil',false,''],
  ['REC-0041',2,'Prep','Mix graham cracker crumbs with melted butter. Press into the bottom of the pan. Bake 10 minutes.',10,'',false,''],
  ['REC-0041',3,'Prep','Beat cream cheese and sugar until smooth. Add eggs one at a time. Mix in sour cream, vanilla, and lemon zest.',5,'Electric mixer',false,'Do not overmix'],
  ['REC-0041',4,'Bake','Pour filling over crust. Place springform pan in a larger pan with 1 inch of hot water (water bath).',5,'Roasting pan',false,'Water bath prevents cracking'],
  ['REC-0041',5,'Bake','Bake 55-65 minutes until edges are set but center still jiggles slightly.',60,'',false,''],
  ['REC-0041',6,'Chill','Turn oven off, crack door, leave cheesecake inside 1 hour. Then refrigerate at least 4 hours or overnight.',240,'Refrigerator',true,''],
  ['REC-0041',7,'Serve','Remove springform ring. Slice with a hot, clean knife and serve with whipped cream.',5,'Hot knife',false,''],
  // REC-0042 Tiramisu
  ['REC-0042',1,'Prep','Separate eggs. Whisk yolks with sugar until pale and thick, about 5 minutes.',5,'Two bowls, electric mixer',false,''],
  ['REC-0042',2,'Prep','Fold mascarpone into yolk mixture until just combined.',2,'Spatula',false,''],
  ['REC-0042',3,'Prep','Beat egg whites (or cream) to stiff peaks. Fold gently into mascarpone mixture.',5,'Electric mixer',false,''],
  ['REC-0042',4,'Prep','Mix cooled espresso with coffee liqueur in a shallow dish.',1,'Shallow dish',false,''],
  ['REC-0042',5,'Assemble','Dip each ladyfinger briefly in espresso — do not soak. Layer in a dish.',5,'Serving dish',false,''],
  ['REC-0042',6,'Assemble','Spread half the mascarpone cream over ladyfingers. Repeat layers.',5,'Offset spatula',false,''],
  ['REC-0042',7,'Chill','Cover and refrigerate at least 6 hours, preferably overnight.',0,'Refrigerator',true,''],
  ['REC-0042',8,'Serve','Dust generously with cocoa powder just before serving.',2,'Sifter',false,''],
  // REC-0043 Apple Pie
  ['REC-0043',1,'Prep','Peel, core, and slice apples 1/4 inch thick. Toss with sugars, cinnamon, nutmeg, and flour.',10,'Large bowl',true,''],
  ['REC-0043',2,'Prep','Fit one pie crust into a 9-inch pie dish. Refrigerate while preparing filling.',2,'9-inch pie dish',false,''],
  ['REC-0043',3,'Assemble','Pile apple filling into the crust, mounding it slightly in the center. Dot with butter.',3,'',false,''],
  ['REC-0043',4,'Assemble','Place second crust on top. Trim and crimp edges to seal. Cut vents in top crust. Brush with egg wash.',5,'Pastry brush',false,''],
  ['REC-0043',5,'Bake','Preheat oven to 425°F. Bake 20 minutes, then reduce to 375°F and bake 35-40 more minutes.',60,'Oven',false,''],
  ['REC-0043',6,'Rest','Cool at least 2 hours before slicing to allow filling to set.',120,'',false,''],
  // REC-0044 Chocolate Lava Cake
  ['REC-0044',1,'Prep','Preheat oven to 425°F. Butter four 6-oz ramekins, dust with cocoa powder or flour.',5,'Ramekins',false,''],
  ['REC-0044',2,'Prep','Melt chocolate and butter together in a double boiler or microwave in 30-second intervals. Stir smooth.',5,'Double boiler or microwave',false,''],
  ['REC-0044',3,'Prep','Whisk eggs, yolks, and sugar until combined. Fold in chocolate mixture, then flour and vanilla.',3,'Whisk',false,''],
  ['REC-0044',4,'Prep','Pour batter into prepared ramekins. Refrigerate up to 24 hours if making ahead.',1,'Refrigerator',true,''],
  ['REC-0044',5,'Bake','Bake at 425°F for exactly 12-13 minutes — edges should be firm but center still soft.',13,'',false,'Timing is critical'],
  ['REC-0044',6,'Serve','Run a knife around edges and immediately invert onto plate. Serve with vanilla ice cream.',2,'',false,'Serve at once!'],
  // REC-0045 Lemon Bars
  ['REC-0045',1,'Prep','Preheat oven to 350°F. Line a 9x13 pan with parchment paper.',5,'9x13 pan, parchment',false,''],
  ['REC-0045',2,'Prep','Mix flour, softened butter, and powdered sugar until crumbly. Press evenly into pan.',5,'',false,''],
  ['REC-0045',3,'Bake','Bake crust 20 minutes until just golden.',20,'',false,''],
  ['REC-0045',4,'Prep','Whisk eggs, sugar, lemon juice, lemon zest, and remaining flour together for filling.',3,'Whisk',false,''],
  ['REC-0045',5,'Bake','Pour filling over hot crust and bake 18-22 more minutes until filling is set.',20,'',false,''],
  ['REC-0045',6,'Chill','Cool completely, then refrigerate at least 1 hour.',60,'Refrigerator',true,''],
  ['REC-0045',7,'Serve','Dust with powdered sugar just before serving. Cut into bars.',3,'Sifter, sharp knife',false,''],
  // REC-0046 Banana Bread
  ['REC-0046',1,'Prep','Preheat oven to 350°F. Grease a 9x5 loaf pan.',5,'Loaf pan',false,''],
  ['REC-0046',2,'Prep','In a large bowl, mash bananas well. Stir in melted butter, sugar, eggs, and vanilla.',3,'Large bowl, fork',false,''],
  ['REC-0046',3,'Prep','Fold in flour, baking soda, and salt until just combined. Stir in optional add-ins.',2,'Spatula',false,''],
  ['REC-0046',4,'Bake','Pour into prepared loaf pan. Bake 60-65 minutes until a toothpick comes out clean.',65,'Toothpick',false,''],
  ['REC-0046',5,'Rest','Cool in pan 10 minutes, then transfer to a wire rack. Slice when completely cool.',20,'Wire rack',false,''],
  // REC-0047 Strawberry Shortcake
  ['REC-0047',1,'Prep','Slice strawberries and toss with 2 tbsp sugar. Let macerate at room temperature 30 minutes.',30,'Bowl',true,''],
  ['REC-0047',2,'Prep','Preheat oven to 400°F. Mix flour, remaining sugar, and baking powder. Cut in cold butter until crumbly.',8,'Pastry cutter or food processor',false,''],
  ['REC-0047',3,'Prep','Whisk egg with 1 cup heavy cream. Pour over flour mixture and stir until dough just comes together.',2,'',false,''],
  ['REC-0047',4,'Bake','Drop dough in 8 mounds on a parchment-lined baking sheet. Bake 18-22 minutes until golden.',22,'Baking sheet, parchment',false,''],
  ['REC-0047',5,'Prep','Whip remaining cream with powdered sugar and vanilla to soft peaks.',3,'Electric mixer or whisk',false,''],
  ['REC-0047',6,'Assemble','Split biscuits. Layer strawberries and whipped cream between biscuit halves and on top.',3,'',false,''],
  // REC-0048 No-Knead Artisan Bread
  ['REC-0048',1,'Prep','Combine flour, yeast, and salt in a large bowl. Add water and stir until shaggy dough forms.',3,'Large bowl',true,''],
  ['REC-0048',2,'Rest','Cover bowl with plastic wrap or a damp towel. Rest at room temperature 12-18 hours.',0,'',true,'Do not skip this rise'],
  ['REC-0048',3,'Prep','When dough is bubbly and doubled, turn out onto floured surface. Shape into a ball without kneading.',5,'Floured surface',false,''],
  ['REC-0048',4,'Rest','Cover and let rest 15-30 minutes. Meanwhile, preheat oven to 450°F with a Dutch oven inside.',30,'Dutch oven',false,''],
  ['REC-0048',5,'Bake','Carefully place dough in hot Dutch oven. Cover and bake 30 minutes. Uncover and bake 15 more minutes until deep golden.',45,'Dutch oven with lid',false,''],
  ['REC-0048',6,'Rest','Cool on a wire rack at least 30 minutes before slicing.',30,'Wire rack',false,''],
  // REC-0049 Dinner Rolls
  ['REC-0049',1,'Prep','Warm milk to 110°F. Dissolve yeast in warm milk with a pinch of sugar. Let stand 5 minutes until foamy.',5,'Small bowl, thermometer',false,''],
  ['REC-0049',2,'Prep','In a large bowl, mix flour, remaining sugar, and salt. Add yeast mixture, egg, and softened butter. Mix until dough comes together.',5,'Large bowl',false,''],
  ['REC-0049',3,'Cook','Knead dough on floured surface 8-10 minutes until smooth and elastic.',10,'Floured surface',false,'Or use stand mixer'],
  ['REC-0049',4,'Rest','Place in oiled bowl, cover, and let rise in a warm place 1 hour until doubled.',60,'',false,''],
  ['REC-0049',5,'Prep','Punch down dough and divide into 12 equal pieces. Shape each into a smooth ball.',8,'',false,''],
  ['REC-0049',6,'Rest','Place in buttered 9x13 pan. Cover and let rise 45 minutes until puffed.',45,'9x13 pan',false,''],
  ['REC-0049',7,'Bake','Preheat oven to 375°F. Bake 20-25 minutes until golden brown.',25,'Oven',false,''],
  ['REC-0049',8,'Serve','Brush immediately with melted butter. Serve warm.',2,'Pastry brush',false,''],
  // REC-0050 Zucchini Bread
  ['REC-0050',1,'Prep','Preheat oven to 350°F. Grease two 8x4 loaf pans.',5,'Loaf pans',false,''],
  ['REC-0050',2,'Prep','Grate zucchini and squeeze in a clean towel to remove as much moisture as possible.',5,'Box grater, clean towel',false,''],
  ['REC-0050',3,'Prep','Whisk together eggs, oil, and sugar in a large bowl.',2,'Large bowl',false,''],
  ['REC-0050',4,'Prep','Stir in grated zucchini. Add flour, baking soda, baking powder, cinnamon, and salt. Stir until combined.',3,'',false,''],
  ['REC-0050',5,'Bake','Divide between prepared pans. Bake 55-65 minutes until toothpick comes out clean.',60,'Toothpick',false,''],
  ['REC-0050',6,'Rest','Cool in pan 10 minutes, then cool completely on wire rack.',30,'Wire rack',false,''],
  // REC-0051 Turkey Meatballs
  ['REC-0051',1,'Prep','Preheat oven to 400°F. Line a baking sheet with parchment and drizzle with olive oil.',5,'Baking sheet, parchment',false,''],
  ['REC-0051',2,'Prep','In a large bowl, combine ground turkey, bread crumbs, parmesan, egg, garlic, Italian seasoning, onion powder, salt, and pepper. Mix until just combined.',3,'Large bowl',false,'Do not overmix'],
  ['REC-0051',3,'Prep','Roll into 1.5-inch balls (makes about 24).',5,'',false,''],
  ['REC-0051',4,'Bake','Bake 20-25 minutes until cooked through and lightly golden.',25,'',false,''],
  ['REC-0051',5,'Serve','Serve with marinara sauce, over pasta, or freeze for later.',2,'',false,''],
  // REC-0052 Chicken Burrito Bowls
  ['REC-0052',1,'Prep','Season chicken with taco seasoning, salt, and olive oil.',3,'Bowl',true,''],
  ['REC-0052',2,'Cook','Bake chicken at 400°F for 25-30 minutes, or cook in a skillet until done.',30,'Oven or skillet',false,''],
  ['REC-0052',3,'Cook','While chicken cooks, prepare jasmine rice according to package. Stir in lime juice.',20,'Rice cooker or pot',false,''],
  ['REC-0052',4,'Cook','Heat black beans with a pinch of cumin and salt. Warm corn.',5,'Small pot',false,''],
  ['REC-0052',5,'Prep','Slice cooked chicken.',2,'Cutting board',false,''],
  ['REC-0052',6,'Assemble','Divide rice, chicken, beans, corn, and avocado among 5 containers or bowls.',5,'Meal prep containers',false,''],
  ['REC-0052',7,'Serve','Top with salsa, sour cream, and fresh cilantro. Refrigerate up to 4 days.',1,'',false,''],
  // REC-0053 Freezer Breakfast Burritos
  ['REC-0053',1,'Cook','Dice potatoes and cook in a skillet with oil, bell pepper, and onion until tender and golden.',15,'Large skillet',false,''],
  ['REC-0053',2,'Cook','Cook breakfast sausage in another skillet until browned and cooked through.',10,'Skillet',false,''],
  ['REC-0053',3,'Cook','Scramble eggs in the sausage pan until just cooked. Season with salt and pepper.',5,'',false,''],
  ['REC-0053',4,'Assemble','Warm tortillas. Layer potatoes, eggs, sausage, shredded cheese, and salsa down the center.',5,'Microwave or dry pan',false,''],
  ['REC-0053',5,'Assemble','Roll burritos tightly, folding in the sides.',3,'',false,''],
  ['REC-0053',6,'Storage','Let cool completely. Wrap each burrito in foil, then place in freezer bags. Freeze up to 3 months.',10,'Foil, freezer bags',false,''],
  ['REC-0053',7,'Reheat','To reheat: unwrap foil, wrap in a damp paper towel, microwave 2-3 minutes.',1,'',false,''],
  // REC-0054 Slow Cooker Chili
  ['REC-0054',1,'Cook','In a skillet, brown ground beef until cooked through. Drain excess fat.',8,'Skillet',false,''],
  ['REC-0054',2,'Cook','Add diced onion, bell pepper, and garlic to skillet. Cook 3 minutes.',3,'',false,''],
  ['REC-0054',3,'Cook','Transfer everything to the slow cooker. Add kidney beans, canned tomatoes, tomato sauce, chili powder, cumin, smoked paprika, and jalapeño.',3,'Slow cooker',false,''],
  ['REC-0054',4,'Cook','Cook on LOW for 6 hours, or HIGH for 3 hours. Stir occasionally.',360,'',false,''],
  ['REC-0054',5,'Cook','Season with salt and pepper to taste. Add more chili powder for heat.',2,'',false,''],
  ['REC-0054',6,'Serve','Serve topped with sour cream, shredded cheese, and green onions.',2,'',false,''],
  // REC-0055 Meal Prep Veggie Bowls
  ['REC-0055',1,'Prep','Preheat oven to 425°F. Cut sweet potato, broccoli, and chickpeas into similar sizes.',8,'Sheet pan',false,''],
  ['REC-0055',2,'Cook','Toss vegetables and chickpeas with olive oil, cumin, salt, and pepper. Roast on sheet pan 30-35 minutes, tossing halfway.',35,'Sheet pan',false,''],
  ['REC-0055',3,'Cook','Cook quinoa according to package directions.',20,'Pot',true,''],
  ['REC-0055',4,'Prep','Make tahini dressing: whisk tahini, lemon juice, garlic, water, salt, and pepper until smooth.',3,'Small bowl',false,''],
  ['REC-0055',5,'Assemble','Divide quinoa into 5 meal prep containers. Top with roasted vegetables, cherry tomatoes.',5,'Containers',false,''],
  ['REC-0055',6,'Storage','Drizzle with tahini dressing or store separately. Refrigerate up to 4 days.',1,'',false,''],
  // REC-0056 Asian Sesame Salad
  ['REC-0056',1,'Prep','Shred napa and red cabbage. Julienne carrots. Cook edamame if frozen.',10,'',false,''],
  ['REC-0056',2,'Prep','Make dressing: whisk together sesame oil, soy sauce, rice vinegar, and a pinch of sugar.',2,'Small bowl',false,''],
  ['REC-0056',3,'Assemble','Combine all vegetables in a large bowl. Pour dressing over and toss well.',2,'Large bowl, tongs',false,''],
  ['REC-0056',4,'Serve','Top with toasted sesame seeds and crispy wonton strips if desired. Serve immediately.',1,'',false,''],
  // REC-0057 Waldorf Salad
  ['REC-0057',1,'Prep','Dice apples and toss immediately with lemon juice to prevent browning.',3,'Bowl',false,''],
  ['REC-0057',2,'Prep','Slice celery, halve grapes, and roughly chop walnuts.',4,'',false,''],
  ['REC-0057',3,'Prep','In a small bowl, mix mayonnaise with remaining lemon juice, salt, and pepper for dressing.',2,'Small bowl',false,''],
  ['REC-0057',4,'Assemble','Combine apples, celery, grapes, and walnuts in a large bowl. Fold in dressing.',3,'',false,''],
  ['REC-0057',5,'Serve','Serve on lettuce leaves. Refrigerate if not eating immediately.',1,'',false,''],
  // REC-0058 Greek Salad
  ['REC-0058',1,'Prep','Dice cucumbers and tomatoes into 1-inch chunks. Thinly slice red onion.',5,'Cutting board',false,''],
  ['REC-0058',2,'Assemble','Combine cucumber, tomatoes, red onion, and olives in a large bowl.',2,'Large bowl',false,''],
  ['REC-0058',3,'Prep','In a small bowl, whisk olive oil, oregano, salt, and pepper for dressing.',1,'Small bowl',false,''],
  ['REC-0058',4,'Serve','Pour dressing over salad and toss gently. Top with crumbled feta. Serve immediately.',2,'',false,''],
  // REC-0059 Homemade Marinara Sauce
  ['REC-0059',1,'Cook','Heat olive oil in a saucepan over medium heat. Add garlic and cook 1-2 minutes until golden but not browned.',2,'Medium saucepan',false,''],
  ['REC-0059',2,'Cook','Add crushed tomatoes, sugar, and a pinch of salt. Stir well.',2,'',false,''],
  ['REC-0059',3,'Cook','Bring to a simmer, partially cover, and cook 25-30 minutes, stirring occasionally.',30,'',true,''],
  ['REC-0059',4,'Cook','Add fresh basil in the last 5 minutes of cooking.',5,'',false,''],
  ['REC-0059',5,'Cook','Add red pepper flakes if using. Adjust salt to taste.',1,'',false,''],
  ['REC-0059',6,'Serve','Use immediately or cool and store. Keeps 1 week refrigerated or 3 months frozen.',1,'',false,''],
  // REC-0060 Mango Lassi
  ['REC-0060',1,'Prep','If using fresh mango, peel and cube it. If using frozen, measure out 2 cups.',2,'Blender',false,''],
  ['REC-0060',2,'Prep','Add mango, Greek yogurt, milk, honey, and cardamom (if using) to blender.',1,'',false,''],
  ['REC-0060',3,'Cook','Blend until completely smooth, about 1 minute.',1,'Blender',false,''],
  ['REC-0060',4,'Serve','Taste and adjust sweetness. Pour into glasses over ice and serve immediately.',1,'Glasses',false,''],
];

(async () => {
  const reqs = [];
  const vals = [];

  // ── Title ───────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 0, 1, 0, 10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.sage),
      textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  } });
  vals.push({ range: `${S}!A1`, values: [['📋 RECIPE INSTRUCTIONS']] });

  // ── Subtitle ────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 1, 2, 0, 10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.bg),
      textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A2`, values: [['Enter each preparation step on its own row. Link steps to a recipe by entering the Recipe ID in column B. Use Step # to control display order.']] });

  // ── Stats row ───────────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 2, 3, 0, 2),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.sageDeep),
      textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A3`, values: [['Total Steps']] });
  vals.push({ range: `${S}!B3`, values: [['Recipes with Steps']] });
  reqs.push({ repeatCell: {
    range: gridRange(SID, 3, 4, 0, 2),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
      textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.sageDeep), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  } });
  vals.push({ range: `${S}!A4`, values: [[`=COUNTA($F$8:$F$10007)`]] });
  vals.push({ range: `${S}!B4`, values: [[`=SUMPRODUCT((LEN($B$8:$B$10007)>0)*(1/COUNTIF($B$8:$B$10007,$B$8:$B$10007)))`]] });

  [1,2,3,4,5,6].forEach(i => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: 24 }, fields: 'pixelSize',
    } });
  });

  reqs.push({ repeatCell: {
    range: gridRange(SID, 4, 7, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } },
    fields: 'userEnteredFormat(backgroundColor)',
  } });

  // ── Header row 7 ─────────────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(SID, 6, 7, 0, 10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.sageDeep),
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
  const CW = [90,80,160,55,110,380,70,130,65,160];
  CW.forEach((px, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    } });
  });

  // ── Instruction Type dropdown ─────────────────────────────────────────────────
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 10007, 4, 5),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!B51:B61` }] }, strict: false, showCustomUi: true },
  } });
  // ── Make Ahead checkbox ───────────────────────────────────────────────────────
  reqs.push({ setDataValidation: {
    range: gridRange(SID, 7, 10007, 8, 9),
    rule: { condition: { type: 'BOOLEAN' }, strict: true },
  } });

  // ── Freeze ────────────────────────────────────────────────────────────────────
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } },
    fields: 'gridProperties(frozenRowCount)',
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 10007 },
    properties: { pixelSize: 21 }, fields: 'pixelSize',
  } });

  // ── Write data rows ──────────────────────────────────────────────────────────
  const dataRows = STEPS.map((step, i) => {
    const r = 8 + i;
    const [recipeId, stepNum, type, instruction, estMins, tool, makeAhead, notes] = step;
    return [
      `=IF(OR(B${r}="",D${r}=""),"","STEP-"&TEXT(ROW()-7,"00000"))`,
      recipeId,
      `=IFERROR(VLOOKUP(B${r},'Master Recipe Index'!$A$8:$B$2007,2,0),"")`,
      stepNum, type, instruction, estMins, tool, makeAhead, notes,
    ];
  });

  STEPS.forEach((_, i) => {
    const rowIdx = 7 + i;
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(SID, rowIdx, rowIdx+1, 0, 10),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
        verticalAlignment: 'MIDDLE',
        wrapStrategy: 'WRAP',
      } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,wrapStrategy)',
    } });
    [0, 2].forEach(ci => {
      reqs.push({ repeatCell: {
        range: gridRange(SID, rowIdx, rowIdx+1, ci, ci+1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.formula),
          textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.sageDeep) },
        } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      } });
    });
  });

  await batchUpdate(id, reqs, 'instr-fmt');
  await valuesBatchUpdate(id, [{ range: `${S}!A8`, values: dataRows }], 'instr-data');
  console.log(`✓ Recipe Instructions — ${STEPS.length} steps written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
