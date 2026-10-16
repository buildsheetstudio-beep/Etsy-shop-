'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Weekly Meal Planner'];
const S = "'Weekly Meal Planner'";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function fmtDate(d) { return d.toISOString().split('T')[0]; }

// 8 weeks of meal plans: each week = 7 days × 3 main meals = 21 rows per week
// Columns: A=Week Start, B=Day, C=Meal Slot, D=Meal Description, E=Est Calories, F=Prep Notes, G=Done?
const DAYS = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
const WEEK_PLANS = [
  // Week 1 meals (rotating)
  [['Oatmeal with banana & almonds',320,'5 min'],['Grilled chicken salad',380,'15 min'],['Baked salmon with sweet potato',520,'30 min']],
  [['Greek yogurt with berries',220,'2 min'],['Turkey wrap with veggies',410,'10 min'],['Lean beef stir-fry with brown rice',520,'20 min']],
  [['Eggs & whole wheat toast',380,'10 min'],['Lentil soup',420,'prep evening before'],['Grilled chicken & roasted veggies',460,'25 min']],
  [['Smoothie bowl',310,'5 min'],['Tuna sandwich on whole wheat',390,'8 min'],['Pasta with turkey marinara',540,'25 min']],
  [['Avocado toast with poached egg',340,'10 min'],['Quinoa bowl with chickpeas',440,'15 min'],['Shrimp & zucchini noodles',380,'20 min']],
  [['Cottage cheese with fruit',260,'2 min'],['Black bean tacos (2)',460,'15 min'],['Vegetable curry with lentils',460,'30 min']],
  [['Whole grain pancakes',360,'15 min'],['Chicken caesar salad',420,'10 min'],['Pork tenderloin & asparagus',480,'30 min']],
];
const MEAL_SLOTS = ['Breakfast','Lunch','Dinner'];
const START = new Date('2026-03-02'); // 8 weeks ending before nutrition log period

const mealRows = [];
for (let w = 0; w < 8; w++) {
  const weekStart = fmtDate(addDays(START, w * 7));
  DAYS.forEach((day, di) => {
    const plan = WEEK_PLANS[w % WEEK_PLANS.length];
    MEAL_SLOTS.forEach((slot, mi) => {
      const [desc, cal, prep] = plan[mi];
      mealRows.push([weekStart, day, slot, desc, cal, prep, false]);
    });
  });
}

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5000,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['WEEKLY MEAL PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,7), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range: `${S}!A2`, values: [['Plan meals in advance. Enter Week Start date, then plan Breakfast/Lunch/Dinner for each day. Add your own snacks in free rows.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,7), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 3-5: Stats
  const STATS = [
    ['Meals Planned',    `=COUNTA(D8:D5000)`],
    ['Weeks Covered',   `=IFERROR(SUMPRODUCT((A8:A5000<>"")*IFERROR(1/COUNTIF(A8:A5000,A8:A5000),0)),"")`],
    ['Meals Completed', `=COUNTIF(G8:G5000,TRUE)`],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 3 + i;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,0,3), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,3,7), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!A${r}`, values: [[label]] });
    vals.push({ range: `${S}!D${r}`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,0,3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,3,7), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 11, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: r-1, endIndex: r }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});
  });

  // Row 6: spacer
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Row 7: headers
  vals.push({ range: `${S}!A7`, values: [['Week Start','Day','Meal Slot','Meal Description','Est. Calories','Prep Notes','Made It?']] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // Push data
  vals.push({ range: `${S}!A8`, values: mealRows });

  // Row banding
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,7,5100,0,7),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format col A
  fmt.push({ repeatCell: { range: gridRange(SID,7,5000,0,1), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [100, 90, 100, 240, 90, 180, 75];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows 1-7
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '08-mealplan values');
  await batchUpdate(id, fmt, '08-mealplan format');
  console.log(`✅  Weekly Meal Planner done — ${mealRows.length} meal rows (${mealRows.length/21} weeks).`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
