'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const FM = sheetMap['🍱 Freezer Meal Planner'];

// 11 cols, 35 rows
// Row 0: Title, Row 1: Summary labels, Row 2: Summary values (inc. progress bar)
// Row 3: Col headers, Row 4+: meal entries

(async () => {
  const reqs = [];

  // Title
  reqs.push({ mergeCells: { range: gridRange(FM, 0, 1, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(FM, 0, 1, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: FM, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Summary rows 1-2
  reqs.push({ repeatCell: {
    range: gridRange(FM, 1, 2, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(FM, 2, 3, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: FM, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Header row 3
  reqs.push({ repeatCell: {
    range: gridRange(FM, 3, 4, 0, 11),
    cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: FM, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Alt rows 4-34
  for (let r = 4; r < 35; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(FM, r, r+1, 0, 11),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Wrap
  reqs.push({ repeatCell: { range: gridRange(FM, 4, 35, 0, 11), cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' }});

  // Column widths
  const colW = [30, 180, 120, 90, 70, 100, 110, 100, 120, 80, 160];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: FM, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(FM, 3, 35, 0, 11), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border }});

  await batchUpdate(id, reqs, 'fm-format');

  const vdata = [];
  vdata.push({ range: `'🍱 Freezer Meal Planner'!A1`, values: [['🍱 Freezer Meal Planner']] });

  vdata.push({ range: `'🍱 Freezer Meal Planner'!A2:K2`, values: [['Total Meals','Frozen','Used','Still Planned','Target','Progress','','','','','']] });
  vdata.push({ range: `'🍱 Freezer Meal Planner'!A3:K3`, values: [[
    '=COUNTA(B5:B34)',
    '=COUNTIF(G5:G34,"Frozen")',
    '=COUNTIF(G5:G34,"Used")',
    '=A3-B3-C3',
    30,
    `=IFERROR(SPARKLINE({B3,E3-B3},{"charttype","bar";"color1","#6B5B8B";"color2","#DDD5E8"}),"")`,'','','','','',
  ]]});

  vdata.push({ range: `'🍱 Freezer Meal Planner'!A4:K4`, values: [[
    '#','Meal Name','Category','Servings','Prep Time','Cook Method','Status','Dietary Tags','Date Frozen','Shelf Life','Notes / Instructions',
  ]]});

  const meals = [
    [1,'Chicken Tortilla Soup','Soup & Stew',6,'30 min','Slow cooker','Frozen','None','2025-11-01','3 months','Add toppings fresh — cheese, sour cream, tortilla strips'],
    [2,'Beef & Veggie Lasagne','Pasta',8,'45 min','Oven','Frozen','None','2025-11-02','3 months','Bake 375°F covered 45 min, uncovered 15 min'],
    [3,'Butternut Squash Soup','Soup & Stew',6,'20 min','Stovetop','Frozen','Vegetarian','2025-11-03','3 months','Blend smooth. Serve with crusty bread.'],
    [4,'Turkey & Veggie Meatballs','Snack & Bites',8,'30 min','Oven','Frozen','None','2025-11-04','3 months','Great over pasta or with marinara dipping sauce'],
    [5,'Lentil Dahl','Soup & Stew',6,'25 min','Stovetop','Frozen','Vegan','2025-11-05','3 months','Serve over rice. Add naan if available.'],
    [6,'Chicken & Rice Casserole','Casserole',6,'35 min','Oven','Frozen','Gluten-Free','2025-11-06','3 months','350°F 40 min from frozen'],
    [7,'Banana Oat Muffins (x12)','Baked Goods',12,'20 min','Oven','Frozen','None','2025-11-07','2 months','Quick breakfast — microwave 30 sec each'],
    [8,'Shakshuka Sauce','Soup & Stew',4,'20 min','Stovetop','Planned','Vegetarian','','3 months','Cook eggs fresh in thawed sauce'],
    [9,'Pulled Pork (slow cooker)','Casserole',10,'10 min active','Slow cooker','Planned','Dairy-Free','','3 months','Serve on rolls with coleslaw'],
    [10,'Breakfast Burritos (x8)','Breakfast',8,'30 min','Stovetop','Planned','None','','2 months','Wrap in foil. Microwave 2 min each.'],
    [11,'Mushroom Risotto','Grains & Rice',4,'40 min','Stovetop','Planned','Vegetarian','','1 month','Reheat gently with splash of broth'],
    [12,'Greek Lemon Chicken','Casserole',6,'15 min','Oven','Planned','Gluten-Free','','3 months','Roast from frozen 375°F 1 hour'],
    [13,'Veggie Frittata Cups','Breakfast',6,'25 min','Oven','Planned','Vegetarian','','1 month','Individual cups — easy one-hand snack'],
    [14,'Bolognese Sauce (batch)','Pasta',8,'1 hour','Stovetop','Planned','None','','4 months','Versatile — pasta, lasagne, or spaghetti'],
    [15,'Pumpkin Spice Energy Balls','Snack & Bites',20,'15 min','No-cook','Planned','Vegan','','2 months','Oats + almond butter + pumpkin purée'],
  ];
  vdata.push({ range: `'🍱 Freezer Meal Planner'!A5:K19`, values: meals });

  await valuesBatchUpdate(id, vdata, 'fm-values');
  console.log('Freezer Meal Planner complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
