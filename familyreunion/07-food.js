'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const FD = sheetMap['🍖 Food, Drinks & Potluck'];

// [mealSlot, item, type, whoBringing, serves, dietary, status]
const MENU = [
  ['Day 1 Dinner','Lamb chops','Main','BBQ Team',30,'Standard','Confirmed'],
  ['Day 1 Dinner','Beef snags','Main','BBQ Team',30,'Standard','Confirmed'],
  ['Day 1 Dinner','Chicken wings','Main','BBQ Team',25,'Standard','Confirmed'],
  ['Day 1 Dinner','Corn on the cob','Salad & Sides','Organiser',25,'Vegetarian','Confirmed'],
  ['Day 1 Dinner','Green salad','Salad & Sides','Potluck',20,'Vegan','Confirmed'],
  ['Day 1 Dinner','Potato salad','Salad & Sides','Potluck',25,'Vegetarian','Confirmed'],
  ['Day 1 Dinner','Coleslaw','Salad & Sides','Potluck',25,'Dairy Free','Confirmed'],
  ['Day 1 Dinner','Bread rolls','Salad & Sides','Purchase',25,'Standard','Purchased'],
  ['Day 1 Dinner','Pavlova','Dessert','Potluck',20,'Vegetarian','Confirmed'],
  ['Day 1 Dinner','Fruit platter','Dessert','Organiser',25,'Vegan','Confirmed'],
  ['Day 2 Breakfast','Scrambled eggs','Main','Kitchen Team',25,'Vegetarian','Planned'],
  ['Day 2 Breakfast','Toast + spreads','Salad & Sides','Organiser',25,'Standard','Planned'],
  ['Day 2 Breakfast','Fresh fruit bowl','Snack','Potluck',25,'Vegan','Confirmed'],
  ['Day 2 Breakfast','Juice + coffee','Drinks','Purchase',25,'Standard','Purchased'],
  ['Day 2 Lunch','Leftover BBQ meats','Main','Organiser',20,'Standard','Planned'],
  ['Day 2 Lunch','Fresh rolls + fillings','Salad & Sides','Purchase',25,'Standard','Planned'],
  ['Day 2 Lunch','Pasta salad','Salad & Sides','Potluck',20,'Vegetarian','Confirmed'],
  ['Day 2 Lunch','Watermelon','Snack','Organiser',25,'Vegan','Planned'],
  ['Day 1 Drinks','Beer + cider selection','Drinks','Potluck',25,'Standard','Confirmed'],
  ['Day 1 Drinks','Wine (red + white)','Drinks','Potluck',25,'Standard','Confirmed'],
  ['All Day','Soft drinks + sparkling water','Non-Alcoholic','Purchase',30,'Standard','Purchased'],
  ['All Day','Homemade lemonade','Non-Alcoholic','Potluck',20,'Vegan','Confirmed'],
];

// [family, dish, mealSlot, dietary, serves, confirmed]
const POTLUCK = [
  ['Henderson (James & Sarah)','Fruit platter + pavlova','Day 1 Dinner','Vegetarian',20,true],
  ['Thompson (Diana & Robert)','Decorations + bunting — non-food','Day 1 Dinner','Standard',0,true],
  ['Thompson (Lucy)','Potato salad','Day 1 Dinner','Vegetarian',25,true],
  ['Wilson (Peter & Anna)','Coleslaw (dairy free)','Day 1 Dinner','Dairy Free',25,true],
  ['Henderson (Claire)','Green salad','Day 1 Dinner','Vegan',20,true],
  ['Smith (David & Rachel)','Drinks — beer, wine, cider','Day 1 Drinks','Standard',0,true],
  ['Henderson (Tom & Kate)','Vegetarian skewers','Day 1 Dinner','Vegetarian',15,false],
  ['Collins (Brian)','Homemade lemonade','All Day','Vegan',20,true],
  ['Henderson (Margaret)','GF biscuits + GF bread','Day 2 Breakfast','Gluten Free',15,true],
  ['Wilson (Peter & Anna)','Pasta salad','Day 2 Lunch','Vegetarian',20,true],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'🍖 Food, Drinks & Potluck'";
  const NCOLS = 8; // A-H

  // Column widths
  [130,200,120,130,80,130,110,200].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: FD, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner
  reqs.push({ mergeCells: { range: gridRange(FD,0,1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(FD,0,1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: FD, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  🍖  Food, Drinks & Potluck — Henderson Family Reunion 2025']] });

  // Row 1: Menu section header
  reqs.push({ mergeCells: { range: gridRange(FD,1,2,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(FD,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: FD, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['  🍽️  MENU PLANNER']] });

  // Row 2: column headers (GSheets row 3)
  const menuHdrs = ['DAY / MEAL','ITEM','TYPE','WHO\'S BRINGING','SERVES','DIETARY TAGS','STATUS','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(FD,2,3,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: FD, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [menuHdrs] });

  // Menu data rows (0-indexed 3+)
  for (let i = 0; i < MENU.length; i++) {
    const ri  = 3 + i;
    const row = 4 + i;
    const bg  = i % 2 === 0 ? C.white : C.warmCream;
    reqs.push({ repeatCell: {
      range: gridRange(FD, ri, ri+1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    reqs.push({ repeatCell: {
      range: gridRange(FD, ri, ri+1, 4, 5),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
    const [mealSlot, item, type, who, serves, dietary, status] = MENU[i];
    vals.push({ range: `${TAB}!A${row}`, values: [[mealSlot, item, type, who, serves, dietary, status, '']] });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: FD, dimension: 'ROWS', startIndex: 3, endIndex: 3+MENU.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Borders menu
  reqs.push({ updateBorders: {
    range: gridRange(FD,2,3+MENU.length,0,NCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // ── POTLUCK SECTION ──
  const ptHeaderRi = 3 + MENU.length; // 0-indexed
  const ptHeaderGs = ptHeaderRi + 1;

  // Amber divider
  reqs.push({ repeatCell: {
    range: gridRange(FD,ptHeaderRi,ptHeaderRi+1,0,NCOLS),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmAmber) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: FD, dimension: 'ROWS', startIndex: ptHeaderRi, endIndex: ptHeaderRi+1 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});

  // Potluck section header
  const ptSubRi = ptHeaderRi + 1;
  reqs.push({ mergeCells: { range: gridRange(FD,ptSubRi,ptSubRi+1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(FD,ptSubRi,ptSubRi+1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: FD, dimension: 'ROWS', startIndex: ptSubRi, endIndex: ptSubRi+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A${ptSubRi+1}`, values: [['  🤝  POTLUCK SIGN-UP']] });

  // Potluck col headers
  const ptColHdrRi = ptSubRi + 1;
  const ptHdrs = ['FAMILY UNIT','ASSIGNED DISH','DAY / MEAL','DIETARY TAGS','SERVES','CONFIRMED?','NOTES',''];
  reqs.push({ repeatCell: {
    range: gridRange(FD,ptColHdrRi,ptColHdrRi+1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: FD, dimension: 'ROWS', startIndex: ptColHdrRi, endIndex: ptColHdrRi+1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A${ptColHdrRi+1}`, values: [ptHdrs] });

  // Potluck data rows
  const ptDataStartRi = ptColHdrRi + 1;
  for (let i = 0; i < POTLUCK.length; i++) {
    const ri  = ptDataStartRi + i;
    const row = ri + 1;
    const bg  = i % 2 === 0 ? C.white : C.warmCream;
    reqs.push({ repeatCell: {
      range: gridRange(FD, ri, ri+1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    const [family, dish, mealSlot, dietary, serves, confirmed] = POTLUCK[i];
    vals.push({ range: `${TAB}!A${row}`, values: [[family, dish, mealSlot, dietary, serves, confirmed ? 'TRUE' : 'FALSE', '']] });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: FD, dimension: 'ROWS', startIndex: ptDataStartRi, endIndex: ptDataStartRi+POTLUCK.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Checkbox for potluck col F (index 5)
  reqs.push({ setDataValidation: {
    range: gridRange(FD, ptDataStartRi, ptDataStartRi+POTLUCK.length, 5, 6),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true, strict: true },
  }});

  // Borders potluck
  reqs.push({ updateBorders: {
    range: gridRange(FD,ptColHdrRi,ptDataStartRi+POTLUCK.length,0,7),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'food-format');
  await valuesBatchUpdate(id, vals, 'food-values');
  console.log('Food, Drinks & Potluck complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
