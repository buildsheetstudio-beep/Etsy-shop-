'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const FD = sheetMap['Food, Drinks & Cake'];
const S  = "'Food, Drinks & Cake'";

(async () => {
  const data = [];
  const reqs = [];

  const sectionHdr = (row, label) => {
    data.push({ range: `${S}!A${row}`, values: [[label]] });
    reqs.push({ mergeCells: { range: gridRange(FD, row - 1, row, 0, 16), mergeType: 'MERGE_ALL' } });
    reqs.push({ repeatCell: { range: gridRange(FD, row - 1, row, 0, 16), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
      horizontalAlignment: 'LEFT',
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});
  };

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['FOOD, DRINKS & CAKE PLANNER']] });
  reqs.push({ mergeCells: { range: gridRange(FD, 0, 1, 0, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(FD, 0, 1, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 14 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: FD, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── SECTION 1: MENU PLANNING ──────────────────────────────────────────────
  sectionHdr(3, '  SECTION 1 — MENU PLANNING');

  const menuHeaders = [
    'Course / Category', 'Item Name', 'Description', 'Serves',
    'Qty Ordered', 'Cost per Serving', 'Total Cost (auto)',
    'Caterer / Chef', 'GF', 'Nut-Free', 'Dairy-Free', 'Vegan', 'Vegetarian', 'Halal', 'Kosher', 'Notes',
  ];
  data.push({ range: `${S}!A4`, values: [menuHeaders] });
  reqs.push({ repeatCell: { range: gridRange(FD, 3, 4, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: FD, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // Menu items
  const menuItems = [
    // Cocktail hour
    ['Appetizer', 'Caprese Skewers',       'Fresh mozzarella, cherry tomatoes, basil drizzle', 'All', 120, 3.50,   'Harvest Table', true,  true,  false, true,  true,  false, true,  ''],
    ['Appetizer', 'Jerk Chicken Skewers',  'Caribbean jerk marinade, pineapple relish',         'Adults', 100, 4.00, 'Harvest Table', false, true,  false, false, false, false, false, 'Spicy — mild option available'],
    ['Appetizer', 'Smoked Salmon Blini',   'Cream cheese, dill, capers on mini pancakes',       'Adults', 80,  5.00, 'Harvest Table', false, true,  false, false, false, false, false, ''],
    ['Appetizer', 'Spinach Artichoke Cups','Savory spinach-artichoke in phyllo cups',            'All', 100, 3.25,   'Harvest Table', false, true,  true,  true,  true,  false, false, ''],
    // Salad
    ['First Course', 'Garden Harvest Salad', 'Mixed greens, candied pecans, champagne vinaigrette', 'All', 130, 8.00, 'Harvest Table', false, false, true, true, true, false, false, 'GF without croutons'],
    // Entrées
    ['Entrée', 'Herb-Roasted Chicken', 'Free-range chicken, roasted root vegetables, au jus', 'Adults', 55, 38.00, 'Harvest Table', true, true, false, false, false, false, false, 'Most popular choice'],
    ['Entrée', 'Pan-Seared Atlantic Salmon', 'Lemon-dill sauce, asparagus, wild rice', 'Adults', 30, 42.00, 'Harvest Table', true, true, true, false, false, false, false, ''],
    ['Entrée', 'Braised Short Rib', 'Red wine braised, truffle mashed potato, crispy shallots', 'Adults', 20, 52.00, 'Harvest Table', true, true, false, false, false, false, false, 'Premium upgrade'],
    ['Entrée', 'Wild Mushroom Risotto', 'Arborio rice, seasonal mushrooms, truffle oil, parmesan', 'All', 15, 32.00, 'Harvest Table', true, true, false, false, true, false, false, 'Can be made vegan'],
    ['Entrée', 'Vegan Buddha Bowl', 'Roasted chickpeas, quinoa, tahini dressing, roasted veg', 'All', 5, 28.00, 'Harvest Table', true, true, true, true, true, true, true, ''],
    // Sides
    ['Side', 'Truffle Mashed Potatoes', 'Yukon Gold, black truffle oil, cream', 'Adults', 100, 6.00, 'Harvest Table', true, true, false, false, true, false, false, ''],
    ['Side', 'Grilled Asparagus',       'Lemon zest, olive oil, sea salt',      'All',    120, 4.00, 'Harvest Table', true, true, true,  true,  true,  true,  true,  ''],
    // Kids meal
    ['Kids Meal', 'Kids Plate', 'Chicken tenders, mac & cheese, apple slices, juice box', 'Children', 8, 18.00, 'Harvest Table', false, true, false, false, false, false, false, ''],
    // Late night
    ['Late Night', 'Mini Sliders',    'Beef & veggie options on brioche buns', 'All', 120, 5.00, 'Harvest Table', false, true, false, false, false, false, false, 'Served 9:30 PM'],
    ['Late Night', 'Loaded Fries',    'Truffle fries, aioli, parmesan, herbs',  'All', 100, 4.50, 'Harvest Table', true, true, false, true, false, false, false, ''],
  ];

  menuItems.forEach((item, i) => {
    const row = 5 + i;
    const [course, name, desc, serves, qty, costPer, caterer, gf, nutFree, dairyFree, vegan, veg, halal, kosher, notes] = item;
    data.push({ range: `${S}!A${row}`, values: [[course]] });
    data.push({ range: `${S}!B${row}`, values: [[name]] });
    data.push({ range: `${S}!C${row}`, values: [[desc]] });
    data.push({ range: `${S}!D${row}`, values: [[serves]] });
    data.push({ range: `${S}!E${row}`, values: [[qty]] });
    data.push({ range: `${S}!F${row}`, values: [[costPer]] });
    data.push({ range: `${S}!G${row}`, values: [[`=IFERROR(E${row}*F${row},0)`]] });
    data.push({ range: `${S}!H${row}`, values: [[caterer]] });
    data.push({ range: `${S}!I${row}`, values: [[gf]] });
    data.push({ range: `${S}!J${row}`, values: [[nutFree]] });
    data.push({ range: `${S}!K${row}`, values: [[dairyFree]] });
    data.push({ range: `${S}!L${row}`, values: [[vegan]] });
    data.push({ range: `${S}!M${row}`, values: [[veg]] });
    data.push({ range: `${S}!N${row}`, values: [[halal]] });
    data.push({ range: `${S}!O${row}`, values: [[kosher]] });
    data.push({ range: `${S}!P${row}`, values: [[notes]] });
  });

  const menuEnd = 5 + menuItems.length; // row 20

  // Formula + currency format col G (Total Cost)
  reqs.push({ repeatCell: { range: gridRange(FD, 4, menuEnd, 6, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula),
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat(backgroundColor,numberFormat)' }});
  reqs.push({ repeatCell: { range: gridRange(FD, 4, menuEnd, 5, 6), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // Checkboxes I-O (dietary)
  for (let col = 8; col <= 14; col++) {
    reqs.push({ setDataValidation: {
      range: gridRange(FD, 4, menuEnd, col, col + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  }

  // Menu total row
  data.push({ range: `${S}!F${menuEnd}`, values: [['MENU SUBTOTAL']] });
  data.push({ range: `${S}!G${menuEnd}`, values: [[`=IFERROR(SUM(G5:G${menuEnd - 1}),0)`]] });
  reqs.push({ repeatCell: { range: gridRange(FD, menuEnd - 1, menuEnd, 5, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white) },
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' }});

  // Alternating rows for menu
  for (let i = 0; i < menuItems.length; i++) {
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(FD, 4 + i, 5 + i, 0, 16), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9 },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // ── SECTION 2: DRINKS ─────────────────────────────────────────────────────
  const drinkStart = menuEnd + 2;
  sectionHdr(drinkStart, '  SECTION 2 — DRINKS SELECTION');

  const drinkHeaders = ['Type', 'Drink Name', 'Description', 'Bottles / Units', 'Cost Each', 'Total Cost (auto)', 'Alcoholic', 'Notes'];
  data.push({ range: `${S}!A${drinkStart + 1}`, values: [drinkHeaders] });
  reqs.push({ repeatCell: { range: gridRange(FD, drinkStart, drinkStart + 1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});

  const drinks = [
    ['Signature Cocktail', '"Sundown Spritz"',     'Prosecco, aperol, blood orange, rosemary',   50,  12.00, true,  'Pre-batch × 50 servings'],
    ['Signature Cocktail', '"Garden Mule" (mocktail)', 'Ginger beer, cucumber, mint, lime',        30,   6.00, false, 'Zero-proof crowd-pleaser'],
    ['Champagne',          'Toast Champagne',       'For toasts — Moët & Chandon brut',           20, 45.00, true,  '1 bottle per 6 guests'],
    ['Wine (White)',       'Chardonnay',            'La Marca Pinot Grigio — cocktail + dinner', 12,  22.00, true,  ''],
    ['Wine (Red)',         'Cabernet Sauvignon',    'Decoy by Duckhorn — dinner service',        10,  28.00, true,  ''],
    ['Wine (Rosé)',        'Rosé',                  'Whispering Angel — cocktail hour',           8,  38.00, true,  'Popular choice'],
    ['Beer',              'IPA',                   'Local craft IPA — Austin Beerworks',         48,   5.00, true,  '2 cases'],
    ['Beer',              'Lager',                 'Topo Chico Hard Seltzer',                    48,   5.00, true,  '2 cases'],
    ['Soft Drink',        'Sparkling Water',       'San Pellegrino — all tables',               120,   2.50, false, ''],
    ['Soft Drink',        'Assorted Sodas',        'Coke, Diet Coke, Sprite, Ginger Ale',       100,  2.00, false, ''],
    ['Coffee / Tea',      'Coffee Station',        'Regular & decaf drip, cappuccino on request', 1, 500.00, false, 'Per service station'],
    ['Juice',             'Orange & Cranberry Juice', 'Non-alcoholic — brunch & cocktail hour',  20,   8.00, false, ''],
  ];

  drinks.forEach((drink, i) => {
    const row = drinkStart + 2 + i;
    const [type, name, desc, qty, costEach, alcoholic, notes] = drink;
    data.push({ range: `${S}!A${row}`, values: [[type]] });
    data.push({ range: `${S}!B${row}`, values: [[name]] });
    data.push({ range: `${S}!C${row}`, values: [[desc]] });
    data.push({ range: `${S}!D${row}`, values: [[qty]] });
    data.push({ range: `${S}!E${row}`, values: [[costEach]] });
    data.push({ range: `${S}!F${row}`, values: [[`=IFERROR(D${row}*E${row},0)`]] });
    data.push({ range: `${S}!G${row}`, values: [[alcoholic]] });
    data.push({ range: `${S}!H${row}`, values: [[notes]] });
  });

  const drinkEnd = drinkStart + 2 + drinks.length;
  reqs.push({ repeatCell: { range: gridRange(FD, drinkStart + 1, drinkEnd, 5, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula), numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat(backgroundColor,numberFormat)' }});
  reqs.push({ setDataValidation: {
    range: gridRange(FD, drinkStart + 1, drinkEnd, 6, 7),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
  }});

  // Drink total
  data.push({ range: `${S}!E${drinkEnd}`, values: [['DRINKS SUBTOTAL']] });
  data.push({ range: `${S}!F${drinkEnd}`, values: [[`=IFERROR(SUM(F${drinkStart+2}:F${drinkEnd-1}),0)`]] });
  reqs.push({ repeatCell: { range: gridRange(FD, drinkEnd - 1, drinkEnd, 4, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) },
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' }});

  // ── SECTION 3: CAKE & DESSERTS ────────────────────────────────────────────
  const cakeStart = drinkEnd + 2;
  sectionHdr(cakeStart, '  SECTION 3 — CAKE & DESSERTS');

  const cakeHeaders = ['Item', 'Description', 'Flavor / Filling', 'Tiers / Quantity', 'Cost', 'Serves', 'Bakery', 'Notes'];
  data.push({ range: `${S}!A${cakeStart + 1}`, values: [cakeHeaders] });
  reqs.push({ repeatCell: { range: gridRange(FD, cakeStart, cakeStart + 1, 0, 8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  const cakeItems = [
    ['Wedding Cake — Tier 1',    '10" bottom tier — ceremonial',        'Vanilla bean, lemon curd filling',         1, 350, 40, 'Sugar & Spire', 'Main cutting cake'],
    ['Wedding Cake — Tier 2',    '8" middle tier',                       'Champagne cake, strawberry compote',       1, 280, 30, 'Sugar & Spire', ''],
    ['Wedding Cake — Tier 3',    '6" top tier — anniversary keeper',     'Almond, amaretto buttercream',             1, 200, 20, 'Sugar & Spire', 'Freeze for 1-year anniversary'],
    ['Wedding Cake — Tier 4',    '4" display topper',                    'Dummy tier — structural',                  1,  80,  0, 'Sugar & Spire', 'Floral topper'],
    ['Sheet Cake',               'Supplemental cutting cake',            'Chocolate fudge',                          2, 220, 60, 'Sugar & Spire', 'Extra servings'],
    ['Dessert Bar — Macarons',   'Assorted French macarons',             'Various (rose, lemon, chocolate)',        120,  3.50, 1, 'Sugar & Spire', 'Favor + display'],
    ['Dessert Bar — Petit Fours','Bite-size almond cakes with glaze',    'Assorted flavors',                         80,  3.00, 1, 'Sugar & Spire', ''],
    ['Dessert Bar — Cheesecake Mini', 'Individual cheesecake bites',    'Classic + strawberry swirl',               50,  4.00, 1, 'Sugar & Spire', ''],
    ['Dessert Bar — Brownies',   'Fudge brownies + sea salt blondies',   'Dark chocolate + brown butter',            60,  2.50, 1, 'Sugar & Spire', 'GF version available'],
  ];

  cakeItems.forEach((item, i) => {
    const row = cakeStart + 2 + i;
    data.push({ range: `${S}!A${row}`, values: [[item[0]]] });
    data.push({ range: `${S}!B${row}`, values: [[item[1]]] });
    data.push({ range: `${S}!C${row}`, values: [[item[2]]] });
    data.push({ range: `${S}!D${row}`, values: [[item[3]]] });
    data.push({ range: `${S}!E${row}`, values: [[item[4]]] });
    data.push({ range: `${S}!F${row}`, values: [[item[5]]] });
    data.push({ range: `${S}!G${row}`, values: [[item[6]]] });
    data.push({ range: `${S}!H${row}`, values: [[item[7]]] });
  });

  const cakeEnd = cakeStart + 2 + cakeItems.length;
  data.push({ range: `${S}!D${cakeEnd}`, values: [['CAKE SUBTOTAL']] });
  data.push({ range: `${S}!E${cakeEnd}`, values: [[`=IFERROR(SUM(E${cakeStart+2}:E${cakeEnd-1}),0)`]] });
  reqs.push({ repeatCell: { range: gridRange(FD, cakeEnd - 1, cakeEnd, 3, 6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white) },
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' }});

  // ── SECTION 4: DIETARY COVERAGE TABLE ─────────────────────────────────────
  const dietStart = cakeEnd + 2;
  sectionHdr(dietStart, '  SECTION 4 — DIETARY COVERAGE SUMMARY');

  const dietHeaders = ['Dietary Need', 'Guest Count (auto)', 'Menu Items Available', 'Coverage Check'];
  data.push({ range: `${S}!A${dietStart + 1}`, values: [dietHeaders] });
  reqs.push({ repeatCell: { range: gridRange(FD, dietStart, dietStart + 1, 0, 4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  const dietRows = [
    ['Gluten-Free',   `=IFERROR(SUMPRODUCT((ISNUMBER(SEARCH("Gluten-Free",'Guest List & RSVP'!K$6:K$600)))*1),"")`, `=IFERROR(SUMPRODUCT((I5:I${menuEnd-1}=TRUE)*1),"")`, `=IFERROR(IF(C${dietStart+2}>0,"✓ Covered","Review"),"?")`],
    ['Nut-Free',      `=IFERROR(SUMPRODUCT((ISNUMBER(SEARCH("Nut Allergy",'Guest List & RSVP'!K$6:K$600)))*1),"")`,  `=IFERROR(SUMPRODUCT((J5:J${menuEnd-1}=TRUE)*1),"")`, `=IFERROR(IF(C${dietStart+3}>0,"✓ Covered","Review"),"?")`],
    ['Dairy-Free',    `=IFERROR(SUMPRODUCT((ISNUMBER(SEARCH("Dairy-Free",'Guest List & RSVP'!K$6:K$600)))*1),"")`,   `=IFERROR(SUMPRODUCT((K5:K${menuEnd-1}=TRUE)*1),"")`, `=IFERROR(IF(C${dietStart+4}>0,"✓ Covered","Review"),"?")`],
    ['Vegan',         `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!J$6:J$600="Vegan")*1),"")`,                           `=IFERROR(SUMPRODUCT((L5:L${menuEnd-1}=TRUE)*1),"")`, `=IFERROR(IF(C${dietStart+5}>0,"✓ Covered","Review"),"?")`],
    ['Vegetarian',    `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!J$6:J$600="Vegetarian")*1),"")`,                      `=IFERROR(SUMPRODUCT((M5:M${menuEnd-1}=TRUE)*1),"")`, `=IFERROR(IF(C${dietStart+6}>0,"✓ Covered","Review"),"?")`],
    ['Halal',         `=IFERROR(SUMPRODUCT((ISNUMBER(SEARCH("Halal",'Guest List & RSVP'!K$6:K$600)))*1),"")`,        `=IFERROR(SUMPRODUCT((N5:N${menuEnd-1}=TRUE)*1),"")`, `=IFERROR(IF(C${dietStart+7}>0,"✓ Covered","Review"),"?")`],
    ['Kosher',        `=IFERROR(SUMPRODUCT((ISNUMBER(SEARCH("Kosher",'Guest List & RSVP'!K$6:K$600)))*1),"")`,       `=IFERROR(SUMPRODUCT((O5:O${menuEnd-1}=TRUE)*1),"")`, `=IFERROR(IF(C${dietStart+8}>0,"✓ Covered","Review"),"?")`],
    ['Kids Meals',    `=IFERROR(SUMPRODUCT(('Guest List & RSVP'!J$6:J$600="Kids Meal")*1),"")`,                       `=IFERROR(SUMPRODUCT(('Food, Drinks & Cake'!A5:A${menuEnd-1}="Kids Meal")*1),"")`, `=IFERROR(IF(C${dietStart+9}>0,"✓ Covered","Review"),"?")`],
  ];

  dietRows.forEach((dr, i) => {
    const row = dietStart + 2 + i;
    data.push({ range: `${S}!A${row}`, values: [[dr[0]]] });
    data.push({ range: `${S}!B${row}`, values: [[dr[1]]] });
    data.push({ range: `${S}!C${row}`, values: [[dr[2]]] });
    data.push({ range: `${S}!D${row}`, values: [[dr[3]]] });
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(FD, row - 1, row, 0, 4), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9 },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  });
  // Formula cells
  reqs.push({ repeatCell: { range: gridRange(FD, dietStart + 1, dietStart + 2 + dietRows.length, 1, 4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula), textFormat: { fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // CF: Coverage check ✓ → green; Review → amber
  reqs.push({ addConditionalFormatRule: {
    rule: { ranges: [gridRange(FD, dietStart + 1, dietStart + 2 + dietRows.length, 3, 4)],
      booleanRule: { condition: { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: '✓' }] },
        format: { backgroundColor: hex(C.success), textFormat: { bold: true, foregroundColor: hex(C.white) } } },
    }, index: 0,
  }});
  reqs.push({ addConditionalFormatRule: {
    rule: { ranges: [gridRange(FD, dietStart + 1, dietStart + 2 + dietRows.length, 3, 4)],
      booleanRule: { condition: { type: 'TEXT_CONTAINS', values: [{ userEnteredValue: 'Review' }] },
        format: { backgroundColor: hex(C.warning), textFormat: { bold: true } } },
    }, index: 0,
  }});

  // ── GRAND TOTAL ROW ───────────────────────────────────────────────────────
  const grandRow = dietStart + 2 + dietRows.length + 2;
  data.push({ range: `${S}!E${grandRow}`, values: [['FOOD & BEVERAGE TOTAL']] });
  data.push({ range: `${S}!F${grandRow}`, values: [[`=IFERROR(G${menuEnd}+F${drinkEnd}+E${cakeEnd},0)`]] });
  reqs.push({ mergeCells: { range: gridRange(FD, grandRow - 1, grandRow, 4, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(FD, grandRow - 1, grandRow, 4, 7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 11 },
    numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)' }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const widths = [130, 160, 200, 80, 80, 80, 90, 160, 50, 60, 65, 50, 70, 50, 55, 160];
  widths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: FD, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze rows 1-4
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: FD, gridProperties: { frozenRowCount: 4 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: FD, tabColor: hex(C.secondary) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'food-data');
  await batchUpdate(id, reqs, 'food-format');

  console.log('✓ Food, Drinks & Cake built');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
