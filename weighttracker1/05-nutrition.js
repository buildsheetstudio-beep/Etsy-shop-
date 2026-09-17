'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Nutrition Log'];
const S = "'Nutrition Log'";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function fmtDate(d) { return d.toISOString().split('T')[0]; }

// Rotating meals: [meal_type, food, qty, unit, cal, protein, carbs, fat, fiber, sugar, sodium, water]
const MEALS = [
  // Breakfasts
  ['Breakfast','Oatmeal with banana',1,'bowl',320,10,58,6,7,18,120,0],
  ['Breakfast','Greek yogurt with berries',1,'cup',220,18,28,5,3,22,90,0],
  ['Breakfast','Eggs and whole wheat toast',1,'plate',380,24,32,14,4,5,520,0],
  ['Breakfast','Smoothie bowl',1,'bowl',310,12,52,7,6,30,95,0],
  ['Breakfast','Avocado toast with egg',1,'slice',340,16,30,18,7,3,480,0],
  ['Breakfast','Cottage cheese with fruit',1,'bowl',260,24,30,4,2,25,520,0],
  ['Breakfast','Whole grain pancakes',2,'pancakes',360,12,54,8,5,14,520,0],
  // Lunches
  ['Lunch','Grilled chicken salad',1,'bowl',380,38,20,14,6,8,680,0],
  ['Lunch','Turkey and veggie wrap',1,'wrap',410,30,42,12,5,6,820,0],
  ['Lunch','Lentil soup with bread',1,'bowl',420,22,62,6,14,8,740,0],
  ['Lunch','Tuna salad sandwich',1,'sandwich',390,32,36,12,4,5,860,0],
  ['Lunch','Quinoa bowl with roasted veggies',1,'bowl',440,16,66,10,8,12,480,0],
  ['Lunch','Black bean tacos',2,'tacos',460,18,64,14,10,8,740,0],
  ['Lunch','Chicken stir-fry with brown rice',1,'plate',480,36,52,10,5,10,920,0],
  // Dinners
  ['Dinner','Baked salmon with sweet potato',1,'plate',520,42,44,16,6,14,580,0],
  ['Dinner','Lean beef and broccoli with rice',1,'plate',560,40,58,14,6,10,840,0],
  ['Dinner','Pasta with marinara and turkey',1,'bowl',540,34,68,12,6,12,680,0],
  ['Dinner','Grilled chicken with roasted veggies',1,'plate',460,44,32,14,8,10,520,0],
  ['Dinner','Shrimp and zucchini noodles',1,'bowl',380,36,24,14,5,8,860,0],
  ['Dinner','Vegetable curry with lentils',1,'bowl',460,20,68,10,14,16,680,0],
  ['Dinner','Pork tenderloin with asparagus',1,'plate',480,44,18,18,4,6,480,0],
  // Snacks
  ['Snack','Apple with almond butter',1,'serving',220,5,28,12,4,20,40,0],
  ['Snack','Mixed nuts',1,'oz',170,5,6,15,2,1,65,0],
  ['Snack','Protein shake',1,'shake',180,25,12,4,1,8,220,0],
  ['Snack','Hummus and veggies',1,'serving',180,7,22,8,5,6,380,0],
  ['Snack','Rice cake with peanut butter',1,'serving',160,6,22,6,2,4,120,0],
  // Beverages
  ['Beverage','Water',16,'oz',0,0,0,0,0,0,0,480],
  ['Beverage','Green tea',8,'oz',2,0,0,0,0,0,5,240],
  ['Beverage','Coffee with milk',8,'oz',45,2,6,2,0,5,40,240],
];

// 103 days: Jun 1 - Sep 11, 2026 → 3 meals per day = 309 rows
const START = new Date('2026-06-01');
const nutritionRows = [];
for (let day = 0; day < 103; day++) {
  const date = fmtDate(addDays(START, day));
  const dayMeals = [
    MEALS[day % 7],           // Breakfast rotation
    MEALS[7 + (day % 7)],     // Lunch rotation
    MEALS[14 + (day % 7)],    // Dinner rotation
  ];
  // Occasionally add a snack or beverage
  if (day % 3 === 0) dayMeals.push(MEALS[21 + (day % 5)]);  // snack every 3 days
  if (day % 2 === 0) dayMeals.push(MEALS[26 + (day % 3)]);  // beverage every 2 days

  dayMeals.forEach(([mealType, food, qty, unit, cal, protein, carbs, fat, fiber, sugar, sodium, water]) => {
    nutritionRows.push([date, mealType, food, qty, unit, cal, protein, carbs, fat, fiber, sugar, sodium, water, '']);
  });
}

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,10100,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['NUTRITION LOG']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,15), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Log food and drinks. Daily totals aggregate automatically in the Macro Tracker tab. Shaded columns are formula-filled.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,15), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Stats
  const STATS = [
    ['Entries Logged',      `=COUNTA(B8:B10000)`],
    ['Days Logged',         `=SUMPRODUCT(('Nutrition Log'!B8:B5000<>"")*IFERROR(1/COUNTIF('Nutrition Log'!B8:B5000,'Nutrition Log'!B8:B5000),0))`],
    ['Total Calories Logged',`=IFERROR(SUM(G8:G10000),"")`],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,0,4), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,4,9), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!E${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,4), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,4,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 6: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Row 7: headers
  vals.push({ range: `${S}!A7`, values: [['Log ID','Date','Meal Type','Food Item','Qty','Unit','Calories','Protein (g)','Carbs (g)','Fat (g)','Fiber (g)','Sugar (g)','Sodium (mg)','Water (ml)','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // A: Log ID formula
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B8="","","NL-"&TEXT(ROW()-7,"00000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push data (B=date, C=mealType, D=food, E=qty, F=unit, G=cal, H=protein, I=carbs, J=fat, K=fiber, L=sugar, M=sodium, N=water, O=notes)
  for (let i = 0; i < nutritionRows.length; i += 200) {
    vals.push({ range: `${S}!B${8+i}`, values: nutritionRows.slice(i, i+200) });
  }

  // Row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,10100,0,15),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format col B
  fmt.push({ repeatCell: { range: gridRange(SID,7,10000,1,2), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [75, 100, 100, 180, 50, 70, 70, 80, 70, 60, 60, 70, 80, 75, 180];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '05-nutrition values');
  await batchUpdate(id, fmt, '05-nutrition format');
  console.log(`✅  Nutrition Log done — ${nutritionRows.length} entries.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
