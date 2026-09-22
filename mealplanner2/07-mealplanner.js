'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Weekly Meal Planner'];
const S   = "'Weekly Meal Planner'";
const REF = "'Reference Data'";
const RB  = "'Recipe Book'";

const MP1 = 12;   // data starts row 12
const MP0 = MP1 - 1;

const COLS = ['Meal Plan ID','Week Start\nDate','Date','Day','Meal Type',
  'Recipe ID','Recipe Name','Recipe\nCategory','Base\nServings',
  'Default HH\nServings','Serving\nOverride','Final\nServings',
  'Serving\nScale Factor','Leftovers\nPlanned?','Leftover\nServings',
  'Dietary\nCompatibility','Est. Meal\nCost','Cost per\nServing',
  'Prep Ahead?','Notes'];

// Weeks: Aug 18, Aug 25, Sep 1, Sep 8 (all Monday starts)
const WEEKS = ['2026-08-18','2026-08-25','2026-09-01','2026-09-08'];
const DAYS  = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
// mealtype: B=Breakfast, L=Lunch, D=Dinner, S=Snack, Ds=Dessert
// recipeId, mealType, servOverride (null=use default), leftover, leftoverSrv, prepAhead, notes
const PLAN = {
  '2026-08-18': [
    // Mon
    ['REC-0001','Breakfast',4,false,0,false,''],
    ['REC-0010','Lunch',4,false,0,false,''],
    ['REC-0018','Dinner',4,true,2,true,'Marinate chicken Sunday night'],
    // Tue
    ['REC-0003','Breakfast',1,false,0,false,'Prep Sunday night'],
    ['REC-0013','Lunch',6,true,2,true,'Make a big batch'],
    ['REC-0018','Dinner',4,false,0,false,'Leftovers from Monday'],
    // Wed
    ['REC-0007','Breakfast',1,false,0,false,''],
    ['REC-0013','Lunch',4,false,0,false,'Leftover soup'],
    ['REC-0019','Dinner',4,false,0,false,'Taco night'],
    // Thu
    ['REC-0006','Breakfast',8,true,4,true,'Make full pan — freeze half'],
    ['REC-0009','Lunch',4,false,0,false,''],
    ['REC-0022','Dinner',6,true,2,false,''],
    // Fri
    ['REC-0006','Breakfast',4,false,0,false,'Frozen leftovers'],
    ['REC-0022','Dinner',4,false,0,false,'Bolognese leftovers'],
    // Sat
    ['REC-0002','Breakfast',2,false,0,false,''],
    ['REC-0015','Lunch',2,false,0,false,''],
    ['REC-0021','Dinner',4,false,0,false,''],
    ['REC-0037','Dessert',24,true,12,false,'Share with neighbors'],
    // Sun
    ['REC-0008','Breakfast',12,true,6,true,'Batch bake for week'],
    ['REC-0011','Lunch',6,false,0,false,''],
    ['REC-0049','Dinner',4,false,0,false,'Family pizza night'],
    ['REC-0038','Dessert',6,false,0,false,''],
  ],
  '2026-08-25': [
    // Mon
    ['REC-0050','Breakfast',2,false,0,false,''],
    ['REC-0016','Lunch',2,false,0,false,''],
    ['REC-0023','Dinner',4,true,2,false,'Tikka masala'],
    // Tue
    ['REC-0003','Breakfast',1,false,0,false,''],
    ['REC-0023','Lunch',4,false,0,false,'Leftovers'],
    ['REC-0020','Dinner',4,false,0,false,''],
    // Wed
    ['REC-0007','Breakfast',2,false,0,false,''],
    ['REC-0017','Lunch',4,false,0,false,''],
    ['REC-0026','Dinner',8,true,4,true,'Slow cooker — start in morning'],
    // Thu
    ['REC-0004','Breakfast',2,false,0,false,''],
    ['REC-0026','Lunch',8,false,0,false,'BBQ leftovers'],
    ['REC-0025','Dinner',4,false,0,false,''],
    // Fri
    ['REC-0001','Breakfast',4,false,0,false,''],
    ['REC-0043','Lunch',2,false,0,false,''],
    ['REC-0021','Dinner',4,false,0,false,''],
    // Sat
    ['REC-0005','Breakfast',6,false,0,false,''],
    ['REC-0014','Lunch',4,false,0,false,''],
    ['REC-0028','Dinner',6,true,2,false,''],
    ['REC-0039','Dessert',16,true,8,false,''],
    // Sun
    ['REC-0006','Breakfast',8,false,0,true,'Prep for week'],
    ['REC-0011','Lunch',6,false,0,false,''],
    ['REC-0024','Dinner',4,false,0,false,''],
    ['REC-0041','Dessert',4,false,0,false,''],
  ],
  '2026-09-01': [
    // Mon - Labor Day, more relaxed cooking
    ['REC-0001','Breakfast',4,false,0,false,''],
    ['REC-0032','Snack',4,false,0,false,''],
    ['REC-0026','Dinner',8,true,4,false,''],
    // Tue
    ['REC-0003','Breakfast',1,false,0,false,''],
    ['REC-0048','Lunch',4,false,0,false,''],
    ['REC-0030','Dinner',6,false,0,false,''],
    // Wed
    ['REC-0007','Breakfast',2,false,0,false,''],
    ['REC-0009','Lunch',4,false,0,false,''],
    ['REC-0018','Dinner',4,false,0,false,''],
    // Thu
    ['REC-0050','Breakfast',2,false,0,false,''],
    ['REC-0042','Lunch',8,true,4,true,'Batch cook'],
    ['REC-0019','Dinner',4,false,0,false,''],
    // Fri
    ['REC-0008','Breakfast',2,false,0,false,'Frozen muffins'],
    ['REC-0042','Lunch',4,false,0,false,'Leftover minestrone'],
    ['REC-0021','Dinner',4,false,0,false,''],
    // Sat
    ['REC-0002','Breakfast',2,false,0,false,''],
    ['REC-0034','Snack',4,false,0,false,''],
    ['REC-0049','Dinner',4,false,0,false,''],
    ['REC-0040','Dessert',4,false,0,false,''],
    // Sun
    ['REC-0006','Breakfast',8,true,4,false,''],
    ['REC-0047','Lunch',6,false,0,false,''],
    ['REC-0022','Dinner',6,true,2,false,''],
  ],
  '2026-09-08': [
    // Mon
    ['REC-0003','Breakfast',1,false,0,false,''],
    ['REC-0016','Lunch',2,false,0,false,''],
    ['REC-0023','Dinner',4,true,2,false,''],
    // Tue - OVER BUDGET WEEK
    ['REC-0002','Breakfast',2,false,0,false,''],
    ['REC-0023','Lunch',4,false,0,false,'Leftovers'],
    ['REC-0025','Dinner',4,false,0,false,''],
    // Wed
    ['REC-0001','Breakfast',4,false,0,false,''],
    ['REC-0009','Lunch',4,false,0,false,''],
    ['REC-0027','Dinner',4,false,0,false,'Shrimp tacos — treat night'],
    // Thu
    ['REC-0007','Breakfast',2,false,0,false,''],
    ['REC-0017','Lunch',4,false,0,false,''],
    ['REC-0022','Dinner',6,true,2,false,''],
    // Fri
    ['REC-0050','Breakfast',2,false,0,false,''],
    ['REC-0022','Lunch',4,false,0,false,'Leftovers'],
    ['REC-0021','Dinner',4,false,0,false,'Baked salmon'],
    // Sat
    ['REC-0005','Breakfast',6,false,0,false,''],
    ['REC-0031','Snack',4,false,0,false,''],
    ['REC-0049','Dinner',4,false,0,false,''],
    ['REC-0037','Dessert',12,false,0,false,''],
    // Sun
    ['REC-0006','Breakfast',8,true,4,true,''],
    ['REC-0013','Lunch',6,false,0,false,''],
    ['REC-0026','Dinner',8,true,4,true,''],
    ['REC-0038','Dessert',6,false,0,false,''],
  ],
};

function addDays(dateStr, n) {
  const [y,m,d] = dateStr.split('-').map(Number);
  const dt = new Date(y, m-1, d+n);
  return `${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,'0')}-${String(dt.getDate()).padStart(2,'0')}`;
}

// Map recipes → mealType from recipe data (simplified lookup)
const RECIPE_MEAL_TYPE = {
  'REC-0001':'Breakfast','REC-0002':'Breakfast','REC-0003':'Breakfast','REC-0004':'Breakfast',
  'REC-0005':'Breakfast','REC-0006':'Breakfast','REC-0007':'Breakfast','REC-0008':'Breakfast',
  'REC-0009':'Lunch','REC-0010':'Lunch','REC-0011':'Lunch','REC-0012':'Lunch',
  'REC-0013':'Lunch','REC-0014':'Lunch','REC-0015':'Lunch','REC-0016':'Lunch',
  'REC-0017':'Lunch','REC-0018':'Dinner','REC-0019':'Dinner','REC-0020':'Dinner',
  'REC-0021':'Dinner','REC-0022':'Dinner','REC-0023':'Dinner','REC-0024':'Dinner',
  'REC-0025':'Dinner','REC-0026':'Dinner','REC-0027':'Dinner','REC-0028':'Dinner',
  'REC-0029':'Dinner','REC-0030':'Dinner','REC-0031':'Afternoon Snack','REC-0032':'Afternoon Snack',
  'REC-0033':'Morning Snack','REC-0034':'Morning Snack','REC-0035':'Morning Snack',
  'REC-0036':'Afternoon Snack','REC-0037':'Dessert','REC-0038':'Dessert','REC-0039':'Dessert',
  'REC-0040':'Dessert','REC-0041':'Dessert','REC-0042':'Lunch','REC-0043':'Lunch',
  'REC-0044':'Breakfast','REC-0045':'Dinner','REC-0046':'Dinner','REC-0047':'Dinner',
  'REC-0048':'Lunch','REC-0049':'Dinner','REC-0050':'Breakfast',
};

(async () => {
  const fmt = [];
  const vals = [];
  const NC = COLS.length; // 20

  // ─── Title ────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.powder), textFormat: { bold: true, fontSize: 20, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['📅 Weekly Meal Planner']] });

  // Subtitle
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.butter), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A2`, values: [['Plan meals for the week — select recipes, set servings, and track costs']] });

  // Summary cards (rows 3-5, 0-indexed 2-4)
  const cardDefs = [
    ['Meals Planned',        `=IFERROR(COUNTA(E${MP1}:E511),0)`],
    ['Unique Recipes',       `=IFERROR(SUMPRODUCT(1/COUNTIFS(F${MP1}:F511,F${MP1}:F511,F${MP1}:F511,"REC-*")),0)`],
    ['Planned Servings',     `=IFERROR(SUM(L${MP1}:L511),0)`],
    ['Leftover Meals',       `=IFERROR(COUNTIF(N${MP1}:N511,TRUE),0)`],
    ['Est. Meal Cost',       `=IFERROR(SUM(Q${MP1}:Q511),0)`],
    ['Dietary Reviews',      `=IFERROR(COUNTIFS(P${MP1}:P511,"Review*"),0)`],
    ['Favorite Recipes Used',`=IFERROR(SUMPRODUCT((F${MP1}:F511<>"")*IFERROR(VLOOKUP(F${MP1}:F511,'Recipe Book'!$A$6:$M$1005,13,FALSE),FALSE)),0)`],
    ['Planning Completion %',`=IFERROR(COUNTA(F${MP1}:F511)/COUNTA(E${MP1}:E511),0)`],
  ];

  cardDefs.forEach(([label, formula], i) => {
    const col = (i % 4) * 5;
    const row = 2 + Math.floor(i / 4) * 2;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+5), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, row+1, row+2, col, col+5), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, col, col+5),
        cell: { userEnteredFormat: { backgroundColor: hex(C.powder), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.secText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row+1, row+2, col, col+5),
        cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+2}`, values: [[formula]] });
  });

  // Currency for cost card
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 5, 6, 10, 15),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });
  // Percent for completion card
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 5, 6, 15, 20),
      cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Separator row 6 (0-indexed 6)
  fmt.push({ mergeCells: { range: gridRange(SID, 6, 7, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.powder) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ─── Visual weekly grid (rows 7-10, 0-indexed) ───────────────────────────
  // Use a simplified 4-row grid: header row + 3 rows (Breakfast, Lunch/Snack, Dinner/Dessert)
  fmt.push({ mergeCells: { range: gridRange(SID, 7, 8, 0, 2), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 7, 8, 0, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  vals.push({ range: `${S}!A8`, values: [['WEEK AT A GLANCE']] });

  // Day columns in grid (cols 2-8 for Mon-Sun, each 2.5 cols wide → 2+7*2+1 = not exact, use 2 cols each day)
  DAYS.forEach((day, di) => {
    const col = 2 + di * 2;
    fmt.push({ mergeCells: { range: gridRange(SID, 7, 8, col, col+2), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, 7, 8, col, col+2),
        cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+col)}8`, values: [[day]] });
  });

  // Row labels and formula rows for grid
  const gridRows = [
    ['Breakfast', C.butter,     'Breakfast'],
    ['Lunch / Snack', C.mint,   'Lunch'],
    ['Dinner / Dessert', C.powder, 'Dinner'],
  ];
  const WEEK_START_CELL = `'Household Setup'!F24`;  // dashboard control week start date
  gridRows.forEach(([label, color, mealType], ri) => {
    const row = 8 + ri;
    fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, 0, 2), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 2),
        cell: { userEnteredFormat: { backgroundColor: hex(color), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    vals.push({ range: `${S}!A${row+1}`, values: [[label]] });

    DAYS.forEach((day, di) => {
      const col = 2 + di * 2;
      const dayOffset = di;
      fmt.push({ mergeCells: { range: gridRange(SID, row, row+1, col, col+2), mergeType: 'MERGE_ALL' } });
      fmt.push({
        repeatCell: {
          range: gridRange(SID, row, row+1, col, col+2),
          cell: { userEnteredFormat: { backgroundColor: hex(color), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
          fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
        },
      });
      // Formula: find recipe planned for this day and meal type
      const dateExpr = `(${WEEK_START_CELL}+${dayOffset})`;
      const formula = `=IFERROR(INDEX($G$${MP1}:$G$511,MATCH(1,(IFERROR($C$${MP1}:$C$511,0)=${dateExpr})*(IFERROR($E$${MP1}:$E$511,"")="${mealType}"),0)),"—")`;
      vals.push({ range: `${S}!${String.fromCharCode(65+col)}${row+1}`, values: [[formula]] });
    });
  });

  // Separator row 11 (0-indexed 11)
  fmt.push({ mergeCells: { range: gridRange(SID, 11, 12, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 11, 12, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.powder) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ─── Column headers (row 12, 0-indexed 11) ───────────────────────────────
  fmt.push({
    repeatCell: {
      range: gridRange(SID, MP0-1, MP0, 0, NC),
      cell: { userEnteredFormat: { backgroundColor: hex(C.text), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  vals.push({ range: `${S}!A11`, values: [COLS] });

  // ─── Sample meal plan data ────────────────────────────────────────────────
  const mealRows = [];
  let rowOffset = 0;
  for (const weekStart of WEEKS) {
    const weekPlan = PLAN[weekStart];
    // We need to assign day index to each row
    // Group by day: iterate through plan entries, track current day
    // plan entries don't have explicit day — we know Mon has first entries, etc.
    // Let's re-structure: each week's plan is just a flat list ordered by day
    // We'll track day by inspecting recipe meal type
    // Actually we stored them day-by-day, so let's track by counting day transitions
    // For simplicity, we'll use a per-week day counter that increments based on meal type
    // The PLAN arrays are structured: Mon entries, Tue entries, etc.
    // We need to know when we move to next day. Let's re-examine the plan structure.
    // Actually each array entry is [recipeId, mealType, ...] and we know Mon has:
    // Breakfast, Lunch, Dinner (3 entries), then Tue: Breakfast, Lunch, Dinner (3), etc.
    // Let's count by grouping consecutive same-day meal slots.

    // We'll use a simple day index tracking:
    let dayIndex = 0;
    const DAY_MEAL_COUNTS = { '2026-08-18': [3,3,2,3,2,4,4], '2026-08-25': [3,3,4,3,3,4,4],
                              '2026-09-01': [3,3,3,3,3,4,3], '2026-09-08': [3,3,3,3,3,4,4] };
    const dayCounts = DAY_MEAL_COUNTS[weekStart];
    let mealInDay = 0;
    let currentDay = 0;

    for (const [recId, mealType, servOverride, leftover, leftoverSrv, prepAhead, notes] of weekPlan) {
      // Track day
      if (mealInDay >= dayCounts[currentDay]) {
        currentDay++;
        mealInDay = 0;
      }
      mealInDay++;

      const [y,m,d] = weekStart.split('-').map(Number);
      const dateStr = addDays(weekStart, currentDay);
      const [dy,dm,dd] = dateStr.split('-').map(Number);
      const day = DAYS[currentDay];
      const r1 = MP1 + rowOffset;

      mealRows.push([
        `=IF(OR(C${r1}="",E${r1}=""),"","MEAL-"&TEXT(ROW()-${MP1-1},"0000"))`,  // A
        `=DATE(${y},${m},${d})`,                                                   // B Week Start
        `=DATE(${dy},${dm},${dd})`,                                                // C Date
        day,                                                                        // D
        mealType,                                                                   // E
        recId,                                                                      // F
        `=IFERROR(VLOOKUP(F${r1},'Recipe Book'!$A$6:$B$1005,2,FALSE),"")`,       // G
        `=IFERROR(VLOOKUP(F${r1},'Recipe Book'!$A$6:$C$1005,3,FALSE),"")`,       // H
        `=IFERROR(VLOOKUP(F${r1},'Recipe Book'!$A$6:$F$1005,6,FALSE),0)`,        // I Base servings
        4,                                                                          // J Default HH servings
        servOverride || '',                                                         // K Override
        `=IF(K${r1}<>"",K${r1},J${r1})`,                                          // L Final servings
        `=IFERROR(L${r1}/I${r1},1)`,                                              // M Scale factor
        leftover,                                                                   // N
        leftoverSrv || 0,                                                           // O
        'Compatible',                                                               // P Dietary compat (default)
        `=IFERROR(VLOOKUP(F${r1},'Recipe Book'!$A$6:$K$1005,11,FALSE)*M${r1},0)`, // Q Est meal cost
        `=IFERROR(Q${r1}/L${r1},0)`,                                              // R Cost per serving
        prepAhead,                                                                  // S
        notes,                                                                      // T
      ]);
      rowOffset++;
    }
  }

  // Push data rows in batches of 100
  for (let start = 0; start < mealRows.length; start += 100) {
    const chunk = mealRows.slice(start, start + 100);
    vals.push({ range: `${S}!A${MP1+start}:T${MP1+start+chunk.length-1}`, values: chunk });
  }

  // Row formatting
  for (let i = 0; i < mealRows.length; i++) {
    const bgColor = i % 2 === 0 ? C.panel : C.altRow;
    fmt.push({
      repeatCell: {
        range: gridRange(SID, MP0+i, MP0+i+1, 0, NC),
        cell: { userEnteredFormat: { backgroundColor: hex(bgColor), textFormat: { fontSize: 10, fontFamily: 'Arial' } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    // Formula cells: A(0), G(6), H(7), I(8), L(11), M(12), Q(16), R(17)
    [0, 6, 7, 8, 11, 12, 16, 17].forEach(col => {
      fmt.push({ repeatCell: { range: gridRange(SID, MP0+i, MP0+i+1, col, col+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
    });
  }

  // Checkboxes: N(13), S(18)
  [13, 18].forEach(col => {
    fmt.push({ setDataValidation: { range: gridRange(SID, MP0, MP0+500, col, col+1), rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true } } });
  });

  // Dropdowns: E(4)=MealType, F(5)=RecipeID(free text), P(15)=DietaryCompat
  fmt.push({
    setDataValidation: {
      range: gridRange(SID, MP0, MP0+500, 4, 5),
      rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$A$2:$A$20` }] }, showCustomUi: true, strict: false },
    },
  });
  fmt.push({
    setDataValidation: {
      range: gridRange(SID, MP0, MP0+500, 15, 16),
      rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF}!$O$2:$O$10` }] }, showCustomUi: true, strict: false },
    },
  });

  // Date formats: B, C
  fmt.push({
    repeatCell: {
      range: gridRange(SID, MP0, MP0+500, 1, 3),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // Currency: Q(16), R(17)
  fmt.push({
    repeatCell: {
      range: gridRange(SID, MP0, MP0+500, 16, 18),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ─── Column widths ────────────────────────────────────────────────────────
  const colWidths = [80,90,90,90,100,75,170,110,60,70,70,60,65,70,65,110,80,75,70,200];
  colWidths.forEach((w, i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: MP0-1, endIndex: MP0 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Freeze rows 1:6
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 6 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'mp-fmt');
  await valuesBatchUpdate(id, vals, 'mp-vals');
  console.log(`✓ Weekly Meal Planner complete (${mealRows.length} meal rows, 4 weeks)`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
