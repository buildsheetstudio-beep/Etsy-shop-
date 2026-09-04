'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Pantry Inventory'];
const S   = "'Pantry Inventory'";
const REF = "'Reference Data'";

const PI1 = 6;
const PI0 = PI1 - 1;

const COLS = ['Pantry Item ID','Ingredient Name','Grocery Category','Qty On Hand','Unit',
  'Min Stock\nLevel','Status','Purchase\nDate','Expiration /\nBest-By Date',
  'Storage\nLocation','Unit Cost','Est. Inventory\nValue','Pantry\nStaple?',
  'Include in\nGrocery Calc?','Notes'];

// [name, cat, qty, unit, minStock, purchaseDate, expirationDate, location, unitCost, pantryStaple, includeGrocery, notes]
const pantryItems = [
  // Produce
  ['Onion','Produce',4,'piece',2,'2026-08-14','2026-08-28','Counter',0.40,true,true,'Yellow onions'],
  ['Garlic','Produce',2,'piece',1,'2026-08-14','2026-08-28','Counter',0.50,true,true,'Bulbs'],
  ['Lemon','Produce',3,'piece',2,'2026-08-15','2026-08-22','Fridge',0.50,false,true,''],
  ['Lime','Produce',2,'piece',2,'2026-08-15','2026-08-22','Fridge',0.20,false,true,''],
  ['Banana','Produce',5,'piece',3,'2026-08-16','2026-08-20','Counter',0.20,false,true,''],
  ['Carrot','Produce',6,'piece',4,'2026-08-14','2026-09-01','Fridge',0.10,true,true,''],
  ['Celery','Produce',4,'piece',3,'2026-08-14','2026-08-25','Fridge',0.13,false,true,''],
  ['Spinach','Produce',1,'cup',0,'2026-08-16','2026-08-19','Fridge',0.50,false,true,'Expiring soon'],
  ['Bell pepper','Produce',3,'piece',2,'2026-08-15','2026-08-22','Fridge',0.80,false,true,'Mixed colors'],
  ['Cherry tomatoes','Produce',0.5,'cup',1,'2026-08-16','2026-08-20','Fridge',1.00,false,true,'Almost out'],
  ['Cucumber','Produce',2,'piece',1,'2026-08-15','2026-08-22','Fridge',0.60,false,true,''],
  ['Avocado','Produce',1,'piece',0,'2026-08-17','2026-08-19','Counter',1.20,false,true,'Just bought — ripe'],
  // Dairy
  ['Eggs','Dairy & Eggs',8,'piece',6,'2026-08-14','2026-08-28','Fridge',0.30,true,true,''],
  ['Milk','Dairy & Eggs',0.5,'cup',0,'2026-08-10','2026-08-18','Fridge',0.20,false,true,'Expiring soon — use first'],
  ['Butter','Dairy & Eggs',0.5,'cup',0.25,'2026-08-01','2026-09-01','Fridge',0.60,true,true,'Unsalted'],
  ['Greek yogurt','Dairy & Eggs',1,'cup',0.5,'2026-08-15','2026-08-25','Fridge',0.60,false,true,'Plain'],
  ['Parmesan cheese','Dairy & Eggs',0.5,'cup',0.25,'2026-08-10','2026-09-10','Fridge',1.20,false,true,'Grated'],
  ['Cheddar cheese','Dairy & Eggs',0,'cup',0.5,'','','Fridge',1.20,false,true,'Out of stock'],
  // Pantry staples
  ['Olive oil','Pantry',0.75,'cup',0.25,'2026-07-01','2027-07-01','Pantry',0.15,true,true,'Extra virgin'],
  ['Honey','Pantry',0.5,'cup',0.25,'2026-06-01','2027-06-01','Pantry',0.15,true,true,'Raw honey'],
  ['Maple syrup','Pantry',0.25,'cup',0.1,'2026-07-15','2027-07-15','Pantry',0.60,true,true,'Pure maple'],
  ['Peanut butter','Pantry',0.5,'cup',0.25,'2026-08-01','2027-01-01','Pantry',0.40,false,true,'⚠ Peanut allergen — keep separate'],
  ['Hummus','Pantry',0.5,'cup',0,'2026-08-14','2026-08-21','Fridge',1.50,false,true,''],
  ['Chia seeds','Pantry',3,'tbsp',2,'2026-07-01','2027-07-01','Pantry',0.20,false,true,''],
  ['Tahini','Pantry',0.25,'cup',0,'2026-07-15','2027-01-15','Pantry',0.40,false,true,'Sesame — check allergens'],
  ['Red curry paste','Condiments & Sauces',2,'tbsp',0,'2026-08-01','2027-02-01','Fridge',0.60,false,true,''],
  // Spices
  ['Salt','Spices & Seasonings',0.5,'cup',0.25,'2026-01-01','2028-01-01','Pantry',0.02,true,true,'Kosher salt'],
  ['Black pepper','Spices & Seasonings',3,'tbsp',1,'2026-01-01','2028-01-01','Pantry',0.02,true,true,''],
  ['Garlic powder','Spices & Seasonings',3,'tbsp',1,'2026-02-01','2027-02-01','Pantry',0.05,true,true,''],
  ['Onion powder','Spices & Seasonings',3,'tbsp',1,'2026-02-01','2027-02-01','Pantry',0.05,true,true,''],
  ['Cumin','Spices & Seasonings',2,'tbsp',0.5,'2026-03-01','2027-03-01','Pantry',0.05,true,true,''],
  ['Paprika','Spices & Seasonings',2,'tbsp',0.5,'2026-03-01','2027-03-01','Pantry',0.05,true,true,''],
  ['Smoked paprika','Spices & Seasonings',1,'tbsp',0,'2026-04-01','2027-04-01','Pantry',0.10,false,true,''],
  ['Cinnamon','Spices & Seasonings',2,'tbsp',0.5,'2026-01-01','2027-01-01','Pantry',0.05,true,true,''],
  ['Italian seasoning','Spices & Seasonings',2,'tbsp',0.5,'2026-02-01','2027-02-01','Pantry',0.05,true,true,''],
  ['Thyme','Spices & Seasonings',1,'tbsp',0,'2026-03-01','2027-03-01','Pantry',0.05,false,true,''],
  ['Garam masala','Spices & Seasonings',1,'tbsp',0,'2026-05-01','2027-05-01','Pantry',0.10,false,true,''],
  ['Chili powder','Spices & Seasonings',2,'tbsp',0.5,'2026-02-01','2027-02-01','Pantry',0.05,true,true,''],
  ['Red pepper flakes','Spices & Seasonings',1,'tbsp',0,'2026-03-01','2027-03-01','Pantry',0.05,false,true,''],
  // Baking
  ['All-purpose flour','Baking',2,'cup',1,'2026-07-01','2026-11-01','Pantry',0.25,true,true,''],
  ['Sugar','Baking',1,'cup',0.5,'2026-06-01','2028-01-01','Pantry',0.50,true,true,'White sugar'],
  ['Brown sugar','Baking',0.5,'cup',0.25,'2026-06-01','2028-01-01','Pantry',0.38,true,true,''],
  ['Baking powder','Baking',4,'tsp',2,'2026-05-01','2027-05-01','Pantry',0.05,true,true,''],
  ['Baking soda','Baking',4,'tsp',2,'2026-05-01','2027-05-01','Pantry',0.05,true,true,''],
  ['Vanilla extract','Baking',1,'tbsp',0.5,'2026-06-01','2027-06-01','Pantry',0.10,true,true,'Pure vanilla'],
  ['Chocolate chips','Baking',0,'cup',0.25,'','','Pantry',0.80,false,true,'Out of stock'],
  ['Cocoa powder','Baking',0.25,'cup',0,'2026-07-01','2027-07-01','Pantry',0.50,false,true,''],
  // Grains
  ['Rolled oats','Grains & Pasta',2,'cup',1,'2026-07-01','2026-12-01','Pantry',0.20,true,true,''],
  ['Rice','Grains & Pasta',3,'cup',1,'2026-06-01','2026-12-01','Pantry',0.40,true,true,'Long grain'],
  ['Pasta','Grains & Pasta',1,'lb',0.5,'2026-06-01','2026-12-01','Pantry',1.20,true,true,'Spaghetti'],
  ['Quinoa','Grains & Pasta',0.25,'cup',0.5,'2026-07-01','2026-12-01','Pantry',0.80,false,true,'Low stock'],
  // Canned
  ['Canned tomatoes','Canned Goods',2,'can',1,'2026-06-01','2027-06-01','Pantry',0.90,true,true,''],
  ['Vegetable broth','Canned Goods',4,'cup',2,'2026-06-01','2027-06-01','Pantry',1.00,true,true,''],
  ['Chicken broth','Canned Goods',4,'cup',2,'2026-06-01','2027-06-01','Pantry',1.50,true,true,''],
  ['Cannellini beans','Canned Goods',1,'can',1,'2026-05-01','2027-05-01','Pantry',0.90,false,true,''],
  ['Black beans','Canned Goods',2,'can',1,'2026-05-01','2027-05-01','Pantry',0.90,true,true,''],
  ['Chickpeas','Canned Goods',1,'can',0,'2026-06-01','2027-06-01','Pantry',0.90,false,true,''],
  ['Marinara sauce','Canned Goods',1,'cup',0,'2026-07-01','2027-07-01','Pantry',0.90,false,true,''],
  // Condiments
  ['Soy sauce','Condiments & Sauces',3,'tbsp',1,'2026-06-01','2027-06-01','Pantry',0.20,true,true,''],
  ['Sesame oil','Condiments & Sauces',2,'tbsp',0.5,'2026-06-01','2027-01-01','Pantry',0.25,false,true,''],
  ['Mayonnaise','Condiments & Sauces',3,'tbsp',1,'2026-08-01','2026-10-01','Fridge',0.20,true,true,''],
  ['Dijon mustard','Condiments & Sauces',2,'tbsp',0.5,'2026-07-01','2027-07-01','Fridge',0.10,true,true,''],
  ['Balsamic glaze','Condiments & Sauces',2,'tbsp',0,'2026-07-01','2027-07-01','Pantry',0.30,false,true,''],
  // Frozen
  ['Frozen mixed berries','Frozen',2,'cup',1,'2026-07-01','2027-07-01','Freezer',1.00,false,true,''],
  ['Frozen corn','Frozen',1,'cup',0,'2026-06-01','2027-06-01','Freezer',0.60,false,true,''],
  ['Frozen broccoli','Frozen',0,'cup',1,'','','Freezer',1.00,false,true,'Out of stock'],
  // Snacks
  ['Granola','Pantry',1,'cup',0.5,'2026-08-01','2026-10-01','Pantry',0.40,false,true,''],
  ['Tortilla chips','Snacks',0.5,'bag',0,'2026-08-10','2026-10-01','Pantry',2.00,false,true,''],
  ['Crackers','Snacks',4,'oz',0,'2026-08-05','2026-10-01','Pantry',1.00,false,true,''],
  // Beverages
  ['White wine','Beverages',0.5,'cup',0,'2026-08-10','2027-08-10','Pantry',1.00,false,false,'For cooking'],
  ['Coconut milk','Canned Goods',0,'can',0,'','','Pantry',1.50,false,true,'Out of stock'],
  ['Caesar dressing','Condiments & Sauces',0.25,'cup',0,'2026-08-14','2026-10-01','Fridge',0.80,false,true,'Low stock'],
  ['Salsa','Condiments & Sauces',0.5,'cup',0.25,'2026-08-01','2026-11-01','Fridge',0.60,false,true,''],
  ['BBQ sauce','Condiments & Sauces',0,'cup',0.5,'','','Fridge',1.50,false,true,'Out of stock'],
  ['Enchilada sauce','Canned Goods',0,'can',0,'','','Pantry',1.50,false,true,'Out of stock'],
  ['Taco seasoning','Spices & Seasonings',0,'package',0,'','','Pantry',0.80,false,true,'Out of stock'],
  ['Diced tomatoes','Canned Goods',1,'can',0,'2026-06-01','2027-06-01','Pantry',0.90,false,true,''],
  ['White beans','Canned Goods',1,'can',0,'2026-06-01','2027-06-01','Pantry',0.90,false,true,''],
  ['Canned tuna','Canned Goods',2,'can',1,'2026-07-01','2027-07-01','Pantry',1.00,false,true,''],
];

(async () => {
  const fmt = [];
  const vals = [];
  const NP = pantryItems.length;
  const NC = COLS.length; // 15

  // Title
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { bold: true, fontSize: 20, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['🫙 Pantry Inventory']] });

  // Subtitle
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.butter), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A2`, values: [['Track what you have on hand — pantry quantities are used by the Automated Grocery List']] });

  // Summary cards area (row 3-4, 0-indexed 2-3)
  // Paint 8 cards in 2 rows of 4
  const cardDefs = [
    ['Pantry Items',           `=IFERROR(COUNTA(B${PI1}:B1505),0)`],
    ['Items Low',              `=IFERROR(COUNTIF(G${PI1}:G1505,"Low Stock"),0)`],
    ['Items Out',              `=IFERROR(COUNTIF(G${PI1}:G1505,"Out of Stock"),0)`],
    ['Expiring Soon',          `=IFERROR(COUNTIF(G${PI1}:G1505,"Expiring Soon"),0)`],
    ['Pantry Value',           `=IFERROR(SUM(L${PI1}:L1505),0)`],
    ['Pantry Staples',         `=IFERROR(COUNTIF(M${PI1}:M1505,TRUE),0)`],
    ['Grocery Calc Items',     `=IFERROR(COUNTIF(N${PI1}:N1505,TRUE),0)`],
    ['Categories Stocked',     `=IFERROR(SUMPRODUCT(1/COUNTIFS(C${PI1}:C1505,C${PI1}:C1505,B${PI1}:B1505,"?*")),0)`],
  ];

  cardDefs.forEach(([label, formula], i) => {
    const col = (i % 4) * 3;
    const row = 2 + Math.floor(i / 4) * 2;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+2, col, col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col, col+3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.mint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row+1, row+2, col, col+3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(i === 4 ? '#BB7C73' : C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // Currency for Pantry Value card
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 5, 6, 0, 3),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Header row (row 5, 0-indexed)
  const HDR_ROW = PI0 - 1; // = 4 (0-indexed row 5)
  fmt.push({
    repeatCell: {
      range: gridRange(SID, HDR_ROW, HDR_ROW+1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A5`, values: [COLS] });

  // Data rows
  const pantryVals = pantryItems.map(([name, cat, qty, unit, minStock, purchDate, expDate, location, unitCost, pantryStaple, includeGrocery, notes], i) => {
    const r1 = PI1 + i;
    return [
      `=IF(B${r1}="","","PAN-"&TEXT(ROW()-5,"0000"))`,   // A
      name,                                                 // B
      cat,                                                  // C
      qty,                                                  // D
      unit,                                                 // E
      minStock,                                             // F
      `=IF(D${r1}<=0,"Out of Stock",IF(D${r1}<=F${r1},"Low Stock",IF(AND(I${r1}<>"",I${r1}-TODAY()<=5),"Expiring Soon","In Stock")))`, // G
      purchDate ? `=DATE(${purchDate.replace(/-/g,',')})` : '',  // H
      expDate   ? `=DATE(${expDate.replace(/-/g,',')})` : '',    // I
      location,                                             // J
      unitCost,                                             // K
      `=IFERROR(D${r1}*K${r1},0)`,                        // L
      pantryStaple,                                         // M
      includeGrocery,                                       // N
      notes,                                                // O
    ];
  });

  vals.push({ range: `${S}!A${PI1}:O${PI1+NP-1}`, values: pantryVals });

  // Alternating row colors
  for (let i = 0; i < NP; i++) {
    const bgColor = i % 2 === 0 ? C.panel : C.altRow;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, PI0+i, PI0+i+1, 0, NC),
        cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 10, fontFamily: 'Arial' } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    // Formula cells: A(0), G(6), L(11)
    [0, 6, 11].forEach(col => {
      fmt.push({ repeatCell: { range: gridRange(SID, PI0+i, PI0+i+1, col, col+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
    });
  }

  // Checkboxes: M(12), N(13)
  [12, 13].forEach(col => {
    fmt.push({ setDataValidation: { range: gridRange(SID, PI0, PI0+NP+500, col, col+1), rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } } });
  });

  // Dropdowns: C(2)=GroceryCategory, E(4)=Units, G(6)=PantryStatus
  const dropMapP = { 2: 'C', 4: 'D', 6: 'J' };
  Object.entries(dropMapP).forEach(([col, refLetter]) => {
    fmt.push({
      setDataValidation: {
        range: gridRange(SID, PI0, PI0+NP+1000, parseInt(col), parseInt(col)+1),
        rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$${refLetter}$2:$${refLetter}$50` }] }, showCustomUi: true, strict: false },
      },
    });
  });

  // Date format for H, I
  fmt.push({
    repeatCell: {
      range: gridRange(SID, PI0, PI0+NP+500, 7, 9),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Currency for K, L
  fmt.push({
    repeatCell: {
      range: gridRange(SID, PI0, PI0+NP+500, 10, 12),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Column widths
  const colWidths = [80, 160, 120, 70, 70, 70, 100, 100, 110, 110, 80, 90, 70, 80, 200];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: HDR_ROW, endIndex: HDR_ROW+1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Freeze rows 1:5
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'pantry-fmt');
  await valuesBatchUpdate(id, vals, 'pantry-vals');
  console.log(`✓ Pantry Inventory complete (${NP} items)`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
