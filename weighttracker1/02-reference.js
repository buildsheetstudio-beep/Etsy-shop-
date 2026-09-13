'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

(async () => {
  const vals = [];
  const fmt = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,5000,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Column headers row 1
  vals.push({ range: `${S}!A1`, values: [[
    'Workout Types','Meal Types','Habit Categories','Goal Types','Goal Status',
    'NSV Categories','Intensity','Weight Units','Grocery Categories','Days of Week','Meal Slots',
  ]] });

  // Data starting row 2
  const columns = {
    A: ['Strength','Cardio','Yoga/Flexibility','Walking','Cycling','Swimming','Sports','Other'],
    B: ['Breakfast','Lunch','Dinner','Snack','Beverage','Pre-Workout','Post-Workout'],
    C: ['Sleep','Hydration','Movement','Nutrition','Mindfulness','Self-Care','Productivity','Social'],
    D: ['Weight','Body Measurements','Fitness','Nutrition','Habit','Other'],
    E: ['Active','Achieved','Paused','Abandoned'],
    F: ['Energy & Vitality','Sleep Quality','Physical Strength','Endurance','Flexibility','Mood','Confidence','Clothing Fit','Social','Health Markers','Other'],
    G: ['Very Light','Light','Moderate','Hard','Very Hard','Max Effort'],
    H: ['lbs','kg'],
    I: ['Produce','Protein','Dairy & Alternatives','Grains & Bread','Frozen','Canned & Dry','Condiments & Sauces','Snacks','Beverages','Supplements','Household','Other'],
    J: ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
    K: ['Breakfast','Morning Snack','Lunch','Afternoon Snack','Dinner','Evening Snack'],
  };

  for (const [col, values] of Object.entries(columns)) {
    vals.push({ range: `${S}!${col}2`, values: values.map(v => [v]) });
  }

  // Header formatting
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' }});

  // Column widths
  const WIDTHS = [120,100,120,110,90,150,100,90,160,90,110];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  await valuesBatchUpdate(id, vals, '02-reference values');
  await batchUpdate(id, fmt, '02-reference format');
  console.log('✅  Reference Data done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
