'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const FD = sheetMap['Food, Drinks & Cake'];
const S = "'Food, Drinks & Cake'";

// Menu items: [FoodType, Name, Supplier, HomeOrPurchased, Serves, QtyNeeded, Unit, UnitCost, Vegetarian, Vegan, GlutenFree, NutFree, Notes]
const MENU_ITEMS = [
  ['Main Dish','Mini Slider Burgers','Local Deli','Purchased',2,15,'trays',35,'No','No','No','Yes','Confirm qty 10 Aug'],
  ['Main Dish','Veggie Wraps Platter','Local Deli','Purchased',6,1,'platter',35,'Yes','No','No','Yes','Covers vegetarians'],
  ['Main Dish','Mini Quiche Bites','Homemade','Homemade',3,2,'dozen',8,'Yes','No','No','Yes','Made night before'],
  ['Side','Cucumber & Carrot Sticks','Homemade','Homemade',5,1,'tray',5,'Yes','Yes','Yes','Yes','With hummus'],
  ['Side','Cheese & Crackers Platter','Costco','Purchased',8,1,'platter',18,'Yes','No','Yes','No','Check GF crackers available'],
  ['Snack','Fairy Bread','Homemade','Homemade',6,2,'loaves',6,'Yes','No','No','Yes','White bread, sprinkles'],
  ['Snack','Mixed Nuts & Dried Fruit','Homemade','Homemade',5,1,'bowl',8,'Yes','Yes','Yes','No','NOT nut-free — keep separate table'],
  ['Snack','Popcorn — Lightly Salted','Local Store','Purchased',4,2,'bags',5,'Yes','Yes','Yes','Yes','Individual bags'],
  ['Appetizer','Fruit Skewers','Homemade','Homemade',3,2,'trays',12,'Yes','Yes','Yes','Yes','Strawberry, melon, grapes'],
  ['Snack','Chips & Dips','Costco','Purchased',6,2,'bags',9,'Yes','Yes','Yes','Yes','Corn chips are vegan & GF'],
  ['Side','Green Salad','Homemade','Homemade',8,1,'bowl',8,'Yes','Yes','Yes','Yes','Simple garden salad'],
  ['Kids Food','Sandwich Platter (kids)','Local Deli','Purchased',4,2,'trays',28,'Yes','No','No','Yes','Cheese & vegemite, no nuts'],
  ['Kids Food','Mini Sausage Rolls','Homemade','Homemade',4,2,'dozen',12,'No','No','No','Yes','Puff pastry'],
  ['Dessert','Brownie Bites','Homemade','Homemade',2,1,'dozen',6,'Yes','No','No','Yes','Rich chocolate'],
  ['Dessert','Mixed Berries Bowl','Local Store','Purchased',6,1,'bowl',14,'Yes','Yes','Yes','Yes','Strawberries, blueberries'],
  ['Dessert','Fairy Cupcakes (standard)','Sweet Bloom Bakery','Purchased',1,20,'pcs',3,'Yes','No','No','Yes','Vanilla cream'],
  ['Dessert','Gluten-Free Cupcakes','Sweet Bloom Bakery','Purchased',1,4,'pcs',4.50,'Yes','No','Yes','Yes','For GF guests'],
  ['Cake','Enchanted Garden Cake','Sweet Bloom Bakery','Purchased',1,1,'cake',95,'Yes','No','No','No','3-tier, 25 servings; collect Aug 14'],
  ['Beverage','Juice Boxes','Costco','Purchased',1,24,'boxes',18,'Yes','Yes','Yes','Yes','Apple & orange mixed'],
  ['Beverage','Sparkling Water','Costco','Purchased',1,12,'bottles',12,'Yes','Yes','Yes','Yes','Still available too'],
  ['Beverage','Pink Lemonade','Homemade','Homemade',4,2,'jugs',6,'Yes','Yes','Yes','Yes','From concentrate + berries'],
  ['Adult Food','Sparkling Grape Juice','Local Store','Purchased',1,2,'bottles',7,'Yes','Yes','Yes','Yes','For adult guests'],
  ['Adult Food','Mixed Olives & Antipasto','Local Deli','Purchased',6,1,'platter',22,'Yes','Yes','Yes','Yes','Adult grazing board element'],
];

const DIETARY_RESTRICTIONS = ['Vegetarian','Vegan','Gluten-Free','Dairy-Free','Nut-Free','Egg-Free','Halal','Kosher'];

(async () => {
  const data = [];

  data.push({ range:`${S}!A1`, values:[['FOOD, DRINKS & CAKE PLANNER']] });
  data.push({ range:`${S}!A3`, values:[['Plan your full menu, check dietary coverage, and track food costs. Guest counts pull automatically from the Guest List & RSVP tab.']] });

  // Summary cards rows 5-6
  data.push({ range:`${S}!A5`, values:[['Confirmed Guests']] });
  data.push({ range:`${S}!B5`, values:[[`=IFERROR(SUMIF('Guest List & RSVP'!$I$6:$I$205,"Confirmed",'Guest List & RSVP'!$J$6:$J$205),0)`]] });
  data.push({ range:`${S}!C5`, values:[['Children Attending']] });
  data.push({ range:`${S}!D5`, values:[[`=IFERROR(SUMPRODUCT(('Guest List & RSVP'!$I$6:$I$205="Confirmed")*('Guest List & RSVP'!$E$6:$E$205="Child")*('Guest List & RSVP'!$J$6:$J$205)),0)`]] });
  data.push({ range:`${S}!E5`, values:[['Adults Attending']] });
  data.push({ range:`${S}!F5`, values:[[`=IFERROR(SUMPRODUCT(('Guest List & RSVP'!$I$6:$I$205="Confirmed")*('Guest List & RSVP'!$E$6:$E$205="Adult")*('Guest List & RSVP'!$J$6:$J$205)),0)`]] });
  data.push({ range:`${S}!G5`, values:[['Total Menu Items']] });
  data.push({ range:`${S}!H5`, values:[[`=IFERROR(COUNTA($C$8:$C$107),0)`]] });

  data.push({ range:`${S}!A6`, values:[['Est. Food Cost']] });
  data.push({ range:`${S}!B6`, values:[[`=IFERROR(SUMIF($J$8:$J$107,"<>",""),"")`]] }); // placeholder sumif; overwritten below
  data.push({ range:`${S}!A6`, values:[['Est. Food Cost']] });
  data.push({ range:`${S}!B6`, values:[[`=IFERROR(SUM($J$8:$J$107),0)`]] });
  data.push({ range:`${S}!C6`, values:[['Actual Food Cost']] });
  data.push({ range:`${S}!D6`, values:[[`=IFERROR(SUM($K$8:$K$107),0)`]] });
  data.push({ range:`${S}!E6`, values:[['Cost per Guest']] });
  data.push({ range:`${S}!F6`, values:[[`=IFERROR(B6/MAX(1,B5),0)`]] });
  data.push({ range:`${S}!G6`, values:[['Dietary Restrictions']] });
  data.push({ range:`${S}!H6`, values:[[`=IFERROR(SUMPRODUCT(('Guest List & RSVP'!$I$6:$I$205="Confirmed")*('Guest List & RSVP'!$K$6:$K$205<>"None")*('Guest List & RSVP'!$K$6:$K$205<>"")),0)`]] });

  // Menu planner header row 8
  data.push({ range:`${S}!A8`, values:[['MENU PLANNER']] });
  data.push({ range:`${S}!A9:P9`, values:[[
    'Item ID','Food Type','Item Name','Supplier / Store','Home or Purchased',
    'Serves','Qty Needed','Unit','Est. Unit Cost','Est. Total',
    'Actual Total','Vegetarian?','Vegan?','Gluten-Free?','Nut-Free?','Notes'
  ]] });

  // Item ID formulas rows 10-109
  for (let r = 10; r <= 109; r++) {
    data.push({ range:`${S}!A${r}`, values:[[`=IF(C${r}="","","FD-"&TEXT(ROW()-9,"000"))`]] });
    data.push({ range:`${S}!J${r}`, values:[[`=IFERROR(IF(G${r}="","",G${r}*I${r}),"")`]] });
  }

  // Sample menu data rows 10-32
  MENU_ITEMS.forEach(([ftype,name,supplier,homeOrPurchased,serves,qty,unit,unitCost,veg,vegan,gf,nutFree,notes], i) => {
    const r = 10 + i;
    data.push({ range:`${S}!B${r}:P${r}`, values:[[ftype,name,supplier,homeOrPurchased,serves,qty,unit,unitCost,'',veg,vegan,gf,nutFree,notes]] });
  });

  // Dietary needs summary table row 112
  data.push({ range:`${S}!A111`, values:[['DIETARY NEEDS SUMMARY']] });
  data.push({ range:`${S}!A112:D112`, values:[['Dietary Restriction','Guest Count','Suitable Menu Items','Coverage Status']] });
  DIETARY_RESTRICTIONS.forEach((restriction, i) => {
    const r = 113 + i;
    data.push({ range:`${S}!A${r}`, values:[[restriction]] });
    // Guest count from Guest List (col K = dietary restriction)
    data.push({ range:`${S}!B${r}`, values:[[`=IFERROR(SUMPRODUCT(('Guest List & RSVP'!$I$6:$I$205="Confirmed")*('Guest List & RSVP'!$K$6:$K$205="${restriction}")),0)`]] });
    // Suitable menu items count (using matching Yes/No columns)
    let suitCol;
    if (restriction === 'Vegetarian') suitCol = 'L';
    else if (restriction === 'Vegan') suitCol = 'M';
    else if (restriction === 'Gluten-Free') suitCol = 'N';
    else if (restriction === 'Nut-Free') suitCol = 'O';
    else suitCol = null;

    if (suitCol) {
      data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(COUNTIF($${suitCol}$10:$${suitCol}$109,"Yes"),0)`]] });
      data.push({ range:`${S}!D${r}`, values:[[`=IFERROR(IF(B${r}=0,"No Guests",IF(C${r}>0,"Covered","Review Needed")),"—")`]] });
    } else {
      data.push({ range:`${S}!C${r}`, values:[['—']] });
      data.push({ range:`${S}!D${r}`, values:[[`=IFERROR(IF(B${r}=0,"No Guests","Review Needed"),"—")`]] });
    }
  });

  // Cake planner section row 122
  data.push({ range:`${S}!A122`, values:[['CAKE PLANNER']] });
  const cakeFields = [
    ['Cake Type:','Birthday Layer Cake'],
    ['Flavor:','Vanilla with Passionfruit Cream'],
    ['Size:','3-Tier (6", 8", 10")'],
    ['Servings:',25],
    ['Bakery / Supplier:','Sweet Bloom Bakery'],
    ['Order Date:','Jul 25, 2026'],
    ['Pickup / Delivery Date:','Aug 14, 2026'],
    ['Design Notes:','Enchanted Garden — flowers, butterflies, soft greens and pinks'],
    ['Message on Cake:','"Happy Birthday Mia!"'],
    ['Candle Needed?:','Yes'],
    ['Cake Topper Needed?:','Yes'],
    ['Estimated Cost:',95],
    ['Actual Cost:',''],
    ['Paid?:','No'],
  ];
  cakeFields.forEach(([label, value], i) => {
    const r = 123 + i;
    data.push({ range:`${S}!A${r}`, values:[[label]] });
    data.push({ range:`${S}!B${r}`, values:[[value]] });
  });

  // Drinks planner section row 138
  data.push({ range:`${S}!A138`, values:[['DRINKS PLANNER']] });
  data.push({ range:`${S}!A139:H139`, values:[[
    'Drink','Child / Adult','Servings Needed','Quantity','Unit','Est. Cost','Chilled?','Notes'
  ]] });
  const DRINKS = [
    ['Juice Boxes','Child',1,24,'boxes',18,'No','Apple and orange'],
    ['Pink Lemonade','Child',4,2,'jugs',6,'Yes','Homemade from concentrate'],
    ['Sparkling Water','Both',1,12,'bottles',12,'Yes','Still water also available'],
    ['Sparkling Grape Juice','Adult',2,2,'bottles',7,'Yes','For adult guests — non-alcoholic'],
    ['Iced Tea','Adult',4,2,'jugs',5,'Yes','Optional — prepare if weather hot'],
  ];
  DRINKS.forEach(([drink,childAdult,serves,qty,unit,cost,chilled,notes], i) => {
    const r = 140 + i;
    data.push({ range:`${S}!A${r}:H${r}`, values:[[drink,childAdult,serves,qty,unit,cost,chilled,notes]] });
  });

  await valuesBatchUpdate(id, data, 'food-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(FD,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const currF  = { type:'CURRENCY', pattern:'$#,##0.00' };

  // Title rows 0-1
  reqs.push({ mergeCells:{ range:gridRange(FD,0,2,0,16), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Guidance row 2
  reqs.push({ mergeCells:{ range:gridRange(FD,2,4,0,16), mergeType:'MERGE_ALL' } });
  fmt(2,4,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Cards rows 4-5
  for (let i = 0; i < 4; i++) {
    const c = i * 2;
    reqs.push({ mergeCells:{ range:gridRange(FD,4,5,c,c+1), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(FD,4,5,c+1,c+2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(FD,5,6,c,c+1), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(FD,5,6,c+1,c+2), mergeType:'MERGE_ALL' } });
    fmt(4,5,c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.lavender), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(4,5,c+1,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:14, fontFamily:'Arial', foregroundColor:hex(C.primary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(5,6,c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.mutedBlue), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(5,6,c+1,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:14, fontFamily:'Arial', foregroundColor:hex(C.secondary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  }
  // Currency cards (row 5 indices: cols 1,3,5)
  for (const ci of [1,3,5]) {
    reqs.push({ repeatCell:{ range:gridRange(FD,5,6,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } });
  }
  reqs.push({ updateBorders:{ range:gridRange(FD,4,6,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Menu planner section header row 7 (index 7)
  reqs.push({ mergeCells:{ range:gridRange(FD,7,8,0,16), mergeType:'MERGE_ALL' } });
  fmt(7,8,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Menu header row 9 (index 8)
  fmt(8,9,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });
  // Menu data rows 9-108
  for (let r = 9; r < 109; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,16,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols A(0),J(9)
  for (const ci of [0,9]) {
    reqs.push({ repeatCell:{ range:gridRange(FD,9,109,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  // Input cols B-I(1-8), K-P(10-15)
  for (const ci of [1,2,3,4,5,6,7,8,10,11,12,13,14,15]) {
    reqs.push({ repeatCell:{ range:gridRange(FD,9,109,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  }
  // Currency cols I(8),J(9),K(10)
  for (const ci of [8,9,10]) {
    reqs.push({ repeatCell:{ range:gridRange(FD,9,109,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currF, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  }
  // Dropdown: Food Type B(1)
  reqs.push({ setDataValidation:{ range:gridRange(FD,9,109,1,2), rule:{ condition:{ type:'ONE_OF_LIST', values:['Appetizer','Main Dish','Side','Snack','Dessert','Cake','Beverage','Adult Food','Kids Food','Other'].map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true } } });
  // Dropdown: Home or Purchased E(4)
  reqs.push({ setDataValidation:{ range:gridRange(FD,9,109,4,5), rule:{ condition:{ type:'ONE_OF_LIST', values:['Homemade','Purchased'].map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true } } });
  // Dropdowns: Yes/No cols L(11),M(12),N(13),O(14)
  for (const ci of [11,12,13,14]) {
    reqs.push({ setDataValidation:{ range:gridRange(FD,9,109,ci,ci+1), rule:{ condition:{ type:'ONE_OF_LIST', values:[{ userEnteredValue:'Yes' },{ userEnteredValue:'No' }] }, showCustomUi:true, strict:true } } });
  }
  reqs.push({ updateBorders:{ range:gridRange(FD,8,109,0,16), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Dietary summary section header row 110 (index 110)
  reqs.push({ mergeCells:{ range:gridRange(FD,110,111,0,4), mergeType:'MERGE_ALL' } });
  fmt(110,111,0,4,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(111,112,0,4,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 112; r < 121; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,4,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols B(1),C(2),D(3) for dietary summary
  for (const ci of [1,2,3]) {
    reqs.push({ repeatCell:{ range:gridRange(FD,112,121,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Coverage CF: Covered=success, Review=attention
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(FD,112,121,3,4)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'Covered' }] }, format:{ backgroundColor:hex(C.success), textFormat:{ foregroundColor:hex(C.white), bold:true } } } }, index:0 } });
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(FD,112,121,3,4)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'Review Needed' }] }, format:{ backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true } } } }, index:0 } });
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(FD,112,121,3,4)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'No Guests' }] }, format:{ backgroundColor:hex(C.border), textFormat:{ foregroundColor:hex(C.secText) } } } }, index:0 } });
  reqs.push({ updateBorders:{ range:gridRange(FD,111,121,0,4), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Cake planner section header row 121 (index 121)
  reqs.push({ mergeCells:{ range:gridRange(FD,121,122,0,2), mergeType:'MERGE_ALL' } });
  fmt(121,122,0,2,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Cake data rows 122-135
  fmt(122,137,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, verticalAlignment:'MIDDLE' } });
  fmt(122,137,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  // Currency for Est Cost row 134 (index 133) and Actual row 135 (index 134)
  reqs.push({ repeatCell:{ range:gridRange(FD,133,135,1,2), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(FD,121,137,0,2), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Drinks planner section header row 137 (index 137)
  reqs.push({ mergeCells:{ range:gridRange(FD,137,138,0,8), mergeType:'MERGE_ALL' } });
  fmt(137,138,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(138,139,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 139; r < 145; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(FD,139,145,5,6), cell:{ userEnteredFormat:{ numberFormat:currF, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(FD,138,145,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Column widths
  [[0,70],[1,100],[2,180],[3,130],[4,100],[5,60],[6,70],[7,60],[8,80],[9,90],[10,90],[11,80],[12,60],[13,80],[14,70],[15,150]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:FD, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:FD, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:FD, dimension:'ROWS', startIndex:2, endIndex:9 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:FD, dimension:'ROWS', startIndex:4, endIndex:6 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:FD, dimension:'ROWS', startIndex:9, endIndex:145 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  reqs.push({ updateSheetProperties:{ properties:{ sheetId:FD, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'food-format');
  console.log('Food, Drinks & Cake complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
