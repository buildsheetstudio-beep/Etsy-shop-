'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Automated Grocery List'];
const S   = "'Automated Grocery List'";
const REF = "'Reference Data'";
const MP  = "'Weekly Meal Planner'";
const RI  = "'Recipe Ingredients'";
const PI  = "'Pantry Inventory'";
const HS  = "'Household Setup'";

// Controls live at rows 2-8 (1-indexed), columns A-B (0-1)
// Week Start filter  → B2
// Grocery Category   → B3
// Grocery Status     → B4
// Store filter       → B5
// Include Optional?  → B6
// Include Pantry?    → B7
// Show Zero-Need?    → B8
const CTRL_WEEK = `$B$2`;
const CTRL_CAT  = `$B$3`;
const CTRL_INC_OPT = `$B$6`;
const CTRL_INC_PAN = `$B$7`;

// Grocery data starts row 12 (1-indexed)
const GL1 = 12;
const GL0 = GL1 - 1;

const COLS = ['Grocery Item ID','Ingredient Name','Grocery Category','Required Qty','Unit',
  'Pantry Qty','Min Pantry\nReserve','Est. Qty\nNeeded','Unit Cost','Est. Grocery\nCost',
  'Pantry\nCoverage %','Status','Purchased?','Store / Aisle',
  'Optional?','Related Recipes','Notes'];

// Pre-computed grocery list for week of Aug 18 (the sample week)
// These represent the aggregated requirements from the 4-week meal plans
// For the Aug 18 week specifically:
// Each row: [ingredient, category, requiredQty, unit, pantryQty, minReserve, unitCost, optional, store, relatedRecipes, notes, status]
const groceryItems = [
  // Produce
  ['Onion','Produce',3,'piece',4,1,0.40,false,'Green Valley Market','REC-0013, REC-0019','',        'Already Have'],
  ['Garlic','Produce',8,'clove',8,2,0.10,false,'Green Valley Market','Multiple',          '',        'Already Have'],
  ['Banana','Produce',4,'piece',5,2,0.20,false,'Green Valley Market','REC-0008',          '',        'Already Have'],
  ['Carrot','Produce',3,'piece',6,2,0.10,false,'Green Valley Market','REC-0013',          '',        'Already Have'],
  ['Celery','Produce',3,'piece',4,1,0.13,false,'Green Valley Market','REC-0013',          '',        'Already Have'],
  ['Spinach','Produce',2,'cup',1,0,0.50,false,'Green Valley Market','REC-0018, REC-0022', '',        'Needed'],
  ['Bell pepper','Produce',3,'piece',3,1,0.80,false,'Green Valley Market','REC-0018',     '',        'Already Have'],
  ['Cherry tomatoes','Produce',1,'cup',0.5,1,1.00,false,'Green Valley Market','REC-0017','',        'Needed'],
  ['Broccoli','Produce',2,'cup',0,1,1.00,false,'Green Valley Market','REC-0018',          '',        'Needed'],
  ['Zucchini','Produce',1,'piece',0,0,0.70,false,'Green Valley Market','REC-0018',        '',        'Needed'],
  ['Lemon','Produce',3,'piece',3,1,0.50,false,'Green Valley Market','REC-0021',           '',        'Already Have'],
  ['Romaine lettuce','Produce',3,'cup',0,0,0.30,false,'Green Valley Market','REC-0009, REC-0019','','Needed'],
  ['Tomato','Produce',2,'piece',0,0,0.60,false,'Green Valley Market','REC-0015',          '',        'Needed'],
  ['Fresh basil','Produce',0.5,'cup',0,0,0.30,true,'Green Valley Market','REC-0017',       '',       'Optional'],
  ['Mixed berries','Produce',0.5,'cup',0,0,0.50,true,'Green Valley Market','REC-0003',     '',       'Optional'],
  // Meat & Seafood
  ['Chicken thighs','Meat & Seafood',1.5,'lb',0,0,4.00,false,'Green Valley Market','REC-0018',     '','Needed'],
  ['Ground beef','Meat & Seafood',2,'lb',0,0,5.00,false,'Green Valley Market','REC-0019, REC-0022','','Needed'],
  ['Salmon fillet','Meat & Seafood',1.5,'lb',0,0,9.00,false,'Green Valley Market','REC-0021',        '','Needed'],
  ['Turkey breast','Meat & Seafood',4,'oz',0,0,1.50,false,'Green Valley Market','REC-0010',           '','Needed'],
  // Dairy & Eggs
  ['Eggs','Dairy & Eggs',12,'piece',8,6,0.30,false,'Green Valley Market','Multiple',      '','Needed'],
  ['Milk','Dairy & Eggs',2,'cup',0.5,0,0.20,false,'Green Valley Market','REC-0001, REC-0003','','Needed'],
  ['Butter','Dairy & Eggs',1,'cup',0.5,0.25,0.60,false,'Green Valley Market','REC-0001',  '','Needed'],
  ['Greek yogurt','Dairy & Eggs',1.5,'cup',1,0.5,0.60,false,'Green Valley Market','REC-0003','','Needed'],
  ['Parmesan cheese','Dairy & Eggs',1,'cup',0.5,0.25,1.20,false,'Green Valley Market','Multiple','','Needed'],
  ['Cheddar cheese','Dairy & Eggs',2,'cup',0,0.5,1.20,false,'Green Valley Market','REC-0019','','Needed'],
  ['Mozzarella cheese','Dairy & Eggs',2,'cup',0,0,3.00,false,'Green Valley Market','REC-0049','','Needed'],
  ['Sour cream','Dairy & Eggs',0.5,'cup',0,0,0.40,true,'Green Valley Market','REC-0019',   '','Optional'],
  // Bakery
  ['Bread','Bakery',4,'slice',0,0,0.40,false,'Green Valley Market','REC-0002, REC-0015','','Needed'],
  ['Brioche bread','Bakery',1,'lb',0,0,3.50,false,'Green Valley Market','REC-0006',       '','Needed'],
  ['Flour tortilla','Bakery',8,'piece',0,0,0.40,false,'Green Valley Market','REC-0010',   '','Needed'],
  ['Corn tortillas','Bakery',8,'piece',0,0,0.80,false,'Green Valley Market','REC-0019',   '','Needed'],
  ['Croutons','Bakery',1,'cup',0,0,0.60,false,'Green Valley Market','REC-0009',           '','Needed'],
  ['Ciabatta bread','Bakery',1,'piece',0,0,1.50,false,'Green Valley Market','REC-0012',   '','Needed'],
  ['Pizza dough','Bakery',1,'lb',0,0,1.80,false,'Green Valley Market','REC-0049',          '','Needed'],
  // Pantry
  ['Olive oil','Pantry',0.5,'cup',0.75,0.25,0.15,false,'Green Valley Market','Multiple',  '','Already Have'],
  ['Honey','Pantry',3,'tbsp',0.5,0.25,0.15,true,'Green Valley Market','REC-0003',          '','Already Have'],
  ['Maple syrup','Pantry',3,'tbsp',0.25,0.1,0.60,false,'Green Valley Market','REC-0006',   '','Needed'],
  // Grains
  ['Pasta','Grains & Pasta',1.5,'lb',1,0.5,1.20,false,'Green Valley Market','REC-0017, REC-0022','','Needed'],
  ['Egg noodles','Grains & Pasta',2,'cup',0,0,0.60,false,'Green Valley Market','REC-0013', '','Needed'],
  ['Rolled oats','Grains & Pasta',1,'cup',2,1,0.20,false,'Green Valley Market','REC-0003', '','Already Have'],
  // Canned
  ['Canned tomatoes','Canned Goods',2,'can',2,1,0.90,false,'Green Valley Market','Multiple','','Already Have'],
  ['Chicken broth','Canned Goods',6,'cup',4,2,1.50,false,'Green Valley Market','REC-0013', '','Needed'],
  ['Marinara sauce','Canned Goods',1,'cup',1,0,0.90,false,'Green Valley Market','REC-0049','','Already Have'],
  // Spices (pantry-covered)
  ['Salt','Spices & Seasonings',1,'tsp',0.5,0.25,0.02,false,'Green Valley Market','Multiple','','Already Have'],
  ['Black pepper','Spices & Seasonings',1,'tsp',3,1,0.02,false,'Green Valley Market','Multiple','','Already Have'],
  ['Garlic powder','Spices & Seasonings',1,'tsp',3,1,0.05,false,'Green Valley Market','REC-0018','','Already Have'],
  ['Italian seasoning','Spices & Seasonings',1,'tsp',2,0.5,0.05,false,'Green Valley Market','REC-0018','','Already Have'],
  ['Taco seasoning','Spices & Seasonings',1,'package',0,0,0.80,false,'Green Valley Market','REC-0019','','Needed'],
  // Condiments
  ['Caesar dressing','Condiments & Sauces',0.25,'cup',0.25,0,0.80,false,'Green Valley Market','REC-0009','','Already Have'],
  ['Salsa','Condiments & Sauces',0.5,'cup',0.5,0.25,0.60,false,'Green Valley Market','REC-0019','','Already Have'],
  ['Soy sauce','Condiments & Sauces',3,'tbsp',3,1,0.20,false,'Green Valley Market','REC-0020','','Already Have'],
  ['Sesame oil','Condiments & Sauces',1,'tbsp',2,0.5,0.25,false,'Green Valley Market','REC-0020','','Already Have'],
  // Baking
  ['All-purpose flour','Baking',3,'cup',2,1,0.25,false,'Green Valley Market','REC-0001, REC-0037','','Needed'],
  ['Sugar','Baking',1.75,'cup',1,0.5,0.50,false,'Green Valley Market','REC-0037','','Needed'],
  ['Brown sugar','Baking',0.75,'cup',0.5,0.25,0.38,false,'Green Valley Market','REC-0037','','Needed'],
  ['Chocolate chips','Baking',2,'cup',0,0.25,0.80,false,'Green Valley Market','REC-0037','','Needed'],
  ['Vanilla extract','Baking',3,'tsp',1,0.5,0.10,false,'Green Valley Market','Multiple','','Needed'],
  // Unit mismatch example
  ['Soba noodles','Grains & Pasta',8,'oz',0,0,2.00,false,'Green Valley Market','REC-0046','','Review Units'],
  ['Rice noodles','Grains & Pasta',8,'oz',0,0,2.00,false,'Green Valley Market','REC-0046','Same ingredient — different type; confirm which to buy','Review Units'],
];

(async () => {
  const fmt = [];
  const vals = [];
  const NGL = groceryItems.length;
  const NC  = COLS.length; // 17

  // ─── Title ────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.blush), textFormat: { bold: true, fontSize: 20, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['🛒 Automated Grocery List']] });

  // Subtitle
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.butter), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A2`, values: [['Ingredients are auto-aggregated from planned meals, scaled to servings, and reduced by pantry inventory']] });

  // ─── Controls (rows 2-8, 0-indexed 2-8) ──────────────────────────────────
  // 2-column layout: label (A) + value (B)
  const controls = [
    ['Week Start Date',          `=DATE(2026,8,18)`],
    ['Grocery Category / All',   'All'],
    ['Grocery Status / All',     'All'],
    ['Store / All',              'All'],
    ['Include Optional Ingredients?', false],
    ['Include Pantry Inventory?',     true],
    ['Show Zero-Need Items?',         false],
  ];

  controls.forEach(([label, value], i) => {
    const r = i + 2;
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, 2, 6), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!A${r+1}`, values: [[label]] });
    if (typeof value === 'boolean') {
      fmt.push({ setDataValidation: { range: gridRange(SID, r, r+1, 2, 3), rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } } });
      vals.push({ range: `${S}!C${r+1}`, values: [[value]] });
    } else {
      vals.push({ range: `${S}!C${r+1}`, values: [[value]] });
    }
  });

  // Date format for week start
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 2, 3),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ─── Summary Cards (rows 9-10, 0-indexed) ────────────────────────────────
  const cardDefs = [
    ['Unique Grocery Items',        `=IFERROR(COUNTA(B${GL1}:B1505),0)`],
    ['Items Covered by Pantry',     `=IFERROR(COUNTIF(L${GL1}:L1505,"Already Have"),0)`],
    ['Items to Buy',                `=IFERROR(COUNTIF(L${GL1}:L1505,"Needed"),0)`],
    ['Pantry Coverage %',           `=IFERROR(COUNTIF(L${GL1}:L1505,"Already Have")/COUNTA(L${GL1}:L1505),0)`],
    ['Est. Grocery Cost',           `=IFERROR(SUMIF(L${GL1}:L1505,"Needed",J${GL1}:J1505)+SUMIF(L${GL1}:L1505,"Optional",J${GL1}:J1505),0)`],
    ['Weekly Grocery Budget',       `=IFERROR('Household Setup'!F8,200)`],
    ['Budget Remaining',            `=IFERROR(F10-E10,0)`],
    ['Purchased Completion %',      `=IFERROR(COUNTIF(M${GL1}:M1505,TRUE)/COUNTA(L${GL1}:L1505),0)`],
  ];

  cardDefs.forEach(([label, formula], i) => {
    const col = (i % 4) * 4;
    const row = 8 + Math.floor(i / 4) * 2;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+4), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+2, col, col+4), mergeType: 'MERGE_ALL' } });
    const cardColor = i === 6 ? (i === 6 ? C.mint : C.blush) : C.blush;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(cardColor), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row+1, row+2, col, col+4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // Format card values: currency for E10, F10, G10; percent for D10, H10
  fmt.push({ repeatCell: { range: gridRange(SID, 9, 10, 4, 8), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 9, 10, 12, 16), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 9, 10, 12, 16), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // ─── Column headers (row 11, 0-indexed 10) ───────────────────────────────
  const HDR_ROW = GL0 - 1;
  fmt.push({
    repeatCell: {
      range: gridRange(SID, HDR_ROW, HDR_ROW+1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A11`, values: [COLS] });

  // ─── Grocery data rows ────────────────────────────────────────────────────
  const groceryVals = groceryItems.map(([ingName, cat, reqQty, unit, pantryQty, minReserve, unitCost, optional, store, relatedRecipes, notes, status], i) => {
    const r1 = GL1 + i;
    const estQtyNeeded = `=MAX(0,D${r1}-F${r1}+G${r1})`;
    const estCost      = `=IFERROR(H${r1}*I${r1},0)`;
    const pantryPct    = `=IFERROR(MIN(1,F${r1}/D${r1}),IF(D${r1}=0,1,0))`;
    return [
      `=IF(B${r1}="","","GRI-"&TEXT(ROW()-${GL1-1},"0000"))`, // A
      ingName,    // B
      cat,        // C
      reqQty,     // D Required qty
      unit,       // E
      pantryQty,  // F Pantry qty
      minReserve, // G Min pantry reserve
      estQtyNeeded, // H Est qty needed
      unitCost,   // I Unit cost
      estCost,    // J Est grocery cost
      pantryPct,  // K Pantry coverage %
      status,     // L Status
      false,      // M Purchased?
      store,      // N Store/Aisle
      optional,   // O Optional?
      relatedRecipes, // P Related recipes
      notes,      // Q Notes
    ];
  });

  vals.push({ range: `${S}!A${GL1}:Q${GL1+NGL-1}`, values: groceryVals });

  // Row formatting
  for (let i = 0; i < NGL; i++) {
    const bgColor = i % 2 === 0 ? C.panel : C.altRow;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, GL0+i, GL0+i+1, 0, NC),
        cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 10, fontFamily: 'Arial' } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    // Formula cells: A, H, J, K
    [0, 7, 9, 10].forEach(col => {
      fmt.push({ repeatCell: { range: gridRange(SID, GL0+i, GL0+i+1, col, col+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
    });
  }

  // Checkboxes: M(12), O(14)
  [12, 14].forEach(col => {
    fmt.push({ setDataValidation: { range: gridRange(SID, GL0, GL0+NGL+500, col, col+1), rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } } });
  });

  // Dropdown: C(2)=GroceryCategory, L(11)=GroceryStatus
  fmt.push({
    setDataValidation: {
      range: gridRange(SID, GL0, GL0+NGL+500, 2, 3),
      rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$C$2:$C$20` }] }, showCustomUi: true, strict: false },
    },
  });
  fmt.push({
    setDataValidation: {
      range: gridRange(SID, GL0, GL0+NGL+500, 11, 12),
      rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$K$2:$K$10` }] }, showCustomUi: true, strict: false },
    },
  });

  // Number formats
  fmt.push({ repeatCell: { range: gridRange(SID, GL0, GL0+NGL+500, 8, 10), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, GL0, GL0+NGL+500, 10, 11), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, GL0, GL0+NGL+500, 3, 8), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Column widths
  const colWidths = [80,160,120,80,70,80,80,90,80,90,80,100,75,140,70,180,200];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: HDR_ROW, endIndex: HDR_ROW+1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Freeze rows 1:6
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'grocery-fmt');
  await valuesBatchUpdate(id, vals, 'grocery-vals');
  console.log(`✓ Automated Grocery List complete (${NGL} items)`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
