'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Grocery List'];
const S = "'Grocery List'";

// 72 grocery items across categories
// [item, category, qty, unit, store/brand, needed, got_it, notes]
const ITEMS = [
  // Produce
  ['Bananas',           'Produce',          6,  'count',  '', true,  false, ''],
  ['Spinach',           'Produce',          1,  'bag',    '', true,  false, 'baby spinach'],
  ['Blueberries',       'Produce',          1,  'pint',   '', true,  false, ''],
  ['Sweet potatoes',    'Produce',          3,  'count',  '', true,  false, ''],
  ['Broccoli',          'Produce',          2,  'heads',  '', true,  false, ''],
  ['Avocados',          'Produce',          4,  'count',  '', false, false, ''],
  ['Cherry tomatoes',   'Produce',          1,  'pint',   '', true,  false, ''],
  ['Zucchini',          'Produce',          2,  'count',  '', true,  false, ''],
  ['Asparagus',         'Produce',          1,  'bunch',  '', false, false, ''],
  ['Mixed berries',     'Produce',          1,  'bag',    '', true,  false, 'frozen OK'],
  // Protein
  ['Chicken breast',    'Protein',          2,  'lbs',    '', true,  false, 'boneless skinless'],
  ['Salmon fillets',    'Protein',          4,  'count',  '', true,  false, '6 oz each'],
  ['Lean ground beef',  'Protein',          1,  'lb',     '', true,  false, '93/7'],
  ['Turkey breast',     'Protein',          0.5,'lb',     '', false, false, 'deli'],
  ['Shrimp',            'Protein',          1,  'lb',     '', true,  false, 'peeled, deveined'],
  ['Eggs',              'Protein',          12, 'count',  '', true,  false, 'large'],
  ['Tuna (canned)',     'Protein',          3,  'cans',   '', true,  false, 'in water'],
  ['Pork tenderloin',   'Protein',          1,  'count',  '', false, false, ''],
  // Dairy & Alternatives
  ['Greek yogurt',      'Dairy & Alternatives', 32,'oz',  '', true,  false, 'plain, 0%'],
  ['Cottage cheese',    'Dairy & Alternatives', 16,'oz',  '', true,  false, 'low-fat'],
  ['Almond milk',       'Dairy & Alternatives', 64,'oz',  '', true,  false, 'unsweetened'],
  ['Milk',              'Dairy & Alternatives', 1, 'gal', '', false, false, '2%'],
  ['Cheddar (shredded)','Dairy & Alternatives', 8, 'oz',  '', false, false, 'sharp'],
  // Grains & Bread
  ['Whole wheat bread', 'Grains & Bread',   1,  'loaf',   '', true,  false, ''],
  ['Brown rice',        'Grains & Bread',   2,  'lbs',    '', true,  false, ''],
  ['Rolled oats',       'Grains & Bread',   1,  'container','',true, false, 'old-fashioned'],
  ['Whole wheat pasta', 'Grains & Bread',   1,  'box',    '', true,  false, ''],
  ['Quinoa',            'Grains & Bread',   1,  'bag',    '', true,  false, ''],
  ['Rice cakes',        'Grains & Bread',   1,  'bag',    '', false, false, ''],
  ['Whole wheat tortillas','Grains & Bread',8,  'count',  '', true,  false, ''],
  // Frozen
  ['Frozen edamame',    'Frozen',           1,  'bag',    '', false, false, ''],
  ['Frozen veggie mix', 'Frozen',           2,  'bags',   '', true,  false, ''],
  ['Frozen berries',    'Frozen',           1,  'bag',    '', true,  false, 'mixed'],
  // Canned & Dry
  ['Black beans',       'Canned & Dry',     2,  'cans',   '', true,  false, '15 oz each'],
  ['Chickpeas',         'Canned & Dry',     2,  'cans',   '', true,  false, ''],
  ['Diced tomatoes',    'Canned & Dry',     2,  'cans',   '', true,  false, 'no salt added'],
  ['Lentils (dry)',     'Canned & Dry',     1,  'bag',    '', true,  false, 'green or brown'],
  ['Vegetable broth',   'Canned & Dry',     2,  'cartons','', true,  false, 'low sodium'],
  ['Chicken broth',     'Canned & Dry',     2,  'cartons','', false, false, 'low sodium'],
  // Condiments & Sauces
  ['Olive oil',         'Condiments & Sauces', 1,'bottle','', false, false, 'extra virgin'],
  ['Balsamic vinegar',  'Condiments & Sauces', 1,'bottle','', false, false, ''],
  ['Marinara sauce',    'Condiments & Sauces', 1,'jar',   '', true,  false, ''],
  ['Hummus',            'Condiments & Sauces', 2,'tubs',  '', true,  false, ''],
  ['Almond butter',     'Condiments & Sauces', 1,'jar',   '', true,  false, 'no added sugar'],
  ['Peanut butter',     'Condiments & Sauces', 1,'jar',   '', false, false, 'natural'],
  ['Soy sauce',         'Condiments & Sauces', 1,'bottle','', false, false, 'low sodium'],
  ['Hot sauce',         'Condiments & Sauces', 1,'bottle','', false, false, ''],
  // Snacks
  ['Mixed nuts',        'Snacks',           1,  'bag',    '', true,  false, 'unsalted'],
  ['Protein bars',      'Snacks',           6,  'count',  '', true,  false, '20g+ protein'],
  ['Dark chocolate',    'Snacks',           1,  'bar',    '', false, false, '70%+'],
  ['Rice crackers',     'Snacks',           1,  'box',    '', false, false, ''],
  ['Apple',             'Snacks',           4,  'count',  '', true,  false, ''],
  // Beverages
  ['Water (sparkling)', 'Beverages',        12, 'cans',   '', false, false, 'unsweetened'],
  ['Green tea bags',    'Beverages',        1,  'box',    '', true,  false, ''],
  ['Coffee beans',      'Beverages',        1,  'bag',    '', true,  false, '12 oz'],
  ['Herbal tea',        'Beverages',        1,  'box',    '', false, false, 'chamomile/mint'],
  // Supplements
  ['Protein powder',    'Supplements',      1,  'tub',    '', true,  false, 'whey or plant-based'],
  ['Vitamin D',         'Supplements',      1,  'bottle', '', false, false, '2000 IU'],
  ['Fish oil',          'Supplements',      1,  'bottle', '', false, false, 'omega-3'],
  ['Magnesium',         'Supplements',      1,  'bottle', '', false, false, 'glycinate'],
  ['Creatine',          'Supplements',      1,  'bag',    '', false, false, 'monohydrate'],
  // Household
  ['Meal prep containers','Household',      6,  'count',  '', false, false, '1L size'],
  ['Zip lock bags',     'Household',        1,  'box',    '', true,  false, 'quart size'],
  ['Parchment paper',   'Household',        1,  'roll',   '', true,  false, ''],
  ['Cutting board',     'Household',        1,  'count',  '', false, false, 'replace worn one'],
  ['Food scale',        'Household',        1,  'count',  '', false, false, 'digital'],
  ['Blender',           'Household',        1,  'count',  '', false, false, 'check if needed'],
  ['Reusable water bottle','Household',     1,  'count',  '', false, false, '32 oz'],
  ['Meal prep bags',    'Household',        1,  'pack',   '', false, false, 'reusable'],
  ['Grocery bags',      'Household',        1,  'set',    '', false, false, 'reusable'],
  ['Kitchen scale',     'Household',        1,  'count',  '', false, false, 'for macro tracking'],
  ['Produce keeper',    'Household',        1,  'count',  '', false, false, 'extend shelf life'],
  ['Glass storage set', 'Household',        1,  'set',    '', false, false, 'for meal prep'],
];

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5000,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['GROCERY LIST']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,9), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Check "Needed?" to flag items to buy. Check "Got It?" once purchased. Filter by category to shop by section.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,9), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Stats
  const STATS = [
    ['Items Listed',   `=COUNTA(B8:B5000)`],
    ['Items Needed',   `=COUNTIF(G8:G5000,TRUE)`],
    ['Items Got',      `=COUNTIF(H8:H5000,TRUE)`],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,0,3), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,3,9), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!D${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,3,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 6: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Row 7: headers
  vals.push({ range: `${S}!A7`, values: [['Item ID','Item Name','Category','Quantity','Unit','Store / Brand','Needed?','Got It?','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // A: Item ID formula
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B8="","","GR-"&TEXT(ROW()-7,"000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push grocery data (B=name, C=category, D=qty, E=unit, F=store, G=needed, H=got_it, I=notes)
  vals.push({ range: `${S}!B8`, values: ITEMS });

  // Row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,5100,0,9),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Column widths
  const WIDTHS = [70, 180, 140, 70, 60, 120, 70, 70, 180];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '09-grocery values');
  await batchUpdate(id, fmt, '09-grocery format');
  console.log(`✅  Grocery List done — ${ITEMS.length} items.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
