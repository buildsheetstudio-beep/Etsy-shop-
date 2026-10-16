'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BE = sheetMap['Budget & Expenses'];
const S = "'Budget & Expenses'";

const CATEGORIES = [
  'Venue','Invitations','Decorations','Food','Drinks','Cake & Desserts',
  'Entertainment','Games & Activities','Party Favors','Supplies',
  'Clothing','Photography','Transportation','Setup & Cleanup','Miscellaneous'
];

// Planned budgets per category
const PLANNED = [250,30,120,180,80,100,75,60,50,60,40,0,0,30,25];

// [Date, Category, Vendor, Item, Qty, UnitCost, PayStatus, DueDate, PaidDate, ReceiptSaved, Notes]
const EXPENSES = [
  ['Jul 1, 2026','Venue','Riverside Community','Room hire 3 hrs',1,200,'Deposit Paid','Jul 1, 2026','Jul 1, 2026',true,'50% deposit paid'],
  ['Jul 10, 2026','Venue','Riverside Community','Damage deposit',1,50,'Paid','Jul 10, 2026','Jul 10, 2026',true,'Refundable'],
  ['Jul 3, 2026','Invitations','Paper Dreams Print','Printed invitations x30',30,0.80,'Paid','Jul 3, 2026','Jul 3, 2026',false,'Includes envelopes'],
  ['Jul 5, 2026','Invitations','Crafty Corner','Envelope seals',1,5.50,'Paid','Jul 5, 2026','Jul 5, 2026',false,''],
  ['Jul 15, 2026','Decorations','Party Hut','Balloon arch kit',1,45,'Paid','Jul 15, 2026','Jul 15, 2026',true,'Pink and white'],
  ['Jul 15, 2026','Decorations','Party Hut','Table centrepieces x6',6,8,'Planned','Aug 1, 2026','',false,'Garden theme'],
  ['Jul 20, 2026','Decorations','Craft Store','Ribbon, streamers, bunting',1,22,'Paid','Jul 20, 2026','Jul 20, 2026',false,''],
  ['Aug 1, 2026','Food','Costco','Sandwich platter x2',2,28,'Planned','Aug 13, 2026','',false,'Order by Aug 10'],
  ['Aug 1, 2026','Food','Costco','Fruit platter',1,22,'Planned','Aug 13, 2026','',false,''],
  ['Aug 1, 2026','Food','Local Deli','Mini sliders x30',30,2.50,'Planned','Aug 13, 2026','',false,'Confirm qty Aug 10'],
  ['Aug 1, 2026','Food','Local Deli','Veggie wrap platter',1,35,'Planned','Aug 13, 2026','',false,'Covers vegetarian guests'],
  ['Aug 1, 2026','Drinks','Costco','Sparkling water x12',1,12,'Planned','Aug 13, 2026','',false,''],
  ['Aug 1, 2026','Drinks','Costco','Juice boxes x24',1,18,'Planned','Aug 13, 2026','',false,''],
  ['Aug 1, 2026','Drinks','Costco','Lemonade concentrate x4',4,4,'Planned','Aug 13, 2026','',false,''],
  ['Aug 1, 2026','Drinks','Local Store','Sparkling grape juice',2,7,'Planned','Aug 13, 2026','',false,'For adults'],
  ['Jul 25, 2026','Cake & Desserts','Sweet Bloom Bakery','3-tier birthday cake',1,95,'Deposit Paid','Aug 14, 2026','',true,'Enchanted Garden theme; 25 servings'],
  ['Jul 25, 2026','Cake & Desserts','Sweet Bloom Bakery','Gluten-free cupcakes x4',4,4.50,'Planned','Aug 14, 2026','',false,'For GF guests'],
  ['Aug 1, 2026','Entertainment','Bubbles n Fun','Face painting artist 2hr',1,75,'Planned','Aug 10, 2026','',false,'Confirm by Aug 5'],
  ['Jul 28, 2026','Games & Activities','Craft House','Craft kit per child x20',20,3,'Paid','Jul 28, 2026','Jul 28, 2026',true,'Fairy garden kits'],
  ['Aug 2, 2026','Party Favors','Party Hut','Favor bags x30',30,1.50,'Planned','Aug 10, 2026','',false,''],
  ['Aug 2, 2026','Party Favors','Party Hut','Favor bag fillers assorted',1,20,'Planned','Aug 10, 2026','',false,'Stickers, mini figurines'],
  ['Jul 18, 2026','Supplies','Party Hut','Paper plates, napkins, cups x30',1,18,'Paid','Jul 18, 2026','Jul 18, 2026',false,'Floral pattern'],
  ['Jul 18, 2026','Supplies','Party Hut','Cutlery set disposable x40',1,12,'Paid','Jul 18, 2026','Jul 18, 2026',false,''],
  ['Aug 8, 2026','Supplies','Local Store','Candles, matches',1,5,'Planned','Aug 14, 2026','',false,''],
  ['Aug 5, 2026','Setup & Cleanup','Cleaning Co.','Venue cleanup service',1,30,'Planned','Aug 15, 2026','',false,'Book by Aug 8'],
];

(async () => {
  const data = [];

  data.push({ range:`${S}!A1`, values:[['PARTY BUDGET & EXPENSE TRACKER']] });
  data.push({ range:`${S}!A3`, values:[['Plan your category budgets and record every expense. Totals update automatically as you add entries to the expense log below.']] });

  // Summary cards — row 5-6 (4 cards per row, 2 rows = 8 cards)
  data.push({ range:`${S}!A5`, values:[['Total Budget']] });
  data.push({ range:`${S}!B5`, values:[[`=IFERROR('Party Setup'!B16,0)`]] });
  data.push({ range:`${S}!C5`, values:[['Planned Spending']] });
  data.push({ range:`${S}!D5`, values:[[`=IFERROR(SUM($C$11:$C$25),0)`]] });
  data.push({ range:`${S}!E5`, values:[['Actual Spending']] });
  data.push({ range:`${S}!F5`, values:[[`=IFERROR(SUM($I$31:$I$229),0)`]] });
  data.push({ range:`${S}!G5`, values:[['Remaining Budget']] });
  data.push({ range:`${S}!H5`, values:[[`=IFERROR(B5-F5,0)`]] });
  data.push({ range:`${S}!A6`, values:[['Amount Over/Under']] });
  data.push({ range:`${S}!B6`, values:[[`=IFERROR(F5-C5,0)`]] });
  data.push({ range:`${S}!C6`, values:[['Cost per Guest']] });
  data.push({ range:`${S}!D6`, values:[[`=IFERROR(F5/MAX(1,SUMIF('Guest List & RSVP'!$I$6:$I$205,"Confirmed",'Guest List & RSVP'!$J$6:$J$205)),0)`]] });
  data.push({ range:`${S}!E6`, values:[['Paid Expenses']] });
  data.push({ range:`${S}!F6`, values:[[`=IFERROR(SUMIFS($I$31:$I$229,$J$31:$J$229,"Paid"),0)`]] });
  data.push({ range:`${S}!G6`, values:[['Unpaid Expenses']] });
  data.push({ range:`${S}!H6`, values:[[`=IFERROR(F5-F6,0)`]] });

  // Category budget table header row 9
  data.push({ range:`${S}!A9`, values:[['CATEGORY BUDGET']] });
  data.push({ range:`${S}!A10:F10`, values:[[
    'Expense Category','Planned Budget','Actual Spent','Remaining','Variance %','Status'
  ]] });

  // Category rows 11-25
  CATEGORIES.forEach((cat, i) => {
    const r = 11 + i;
    const planned = PLANNED[i];
    data.push({ range:`${S}!A${r}`, values:[[cat]] });
    data.push({ range:`${S}!B${r}`, values:[[planned]] });
    data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(SUMIF($C$31:$C$229,$A${r},$I$31:$I$229),0)`]] });
    data.push({ range:`${S}!D${r}`, values:[[`=IFERROR(B${r}-C${r},B${r})`]] });
    data.push({ range:`${S}!E${r}`, values:[[`=IFERROR(IF(B${r}=0,0,(C${r}-B${r})/B${r}),0)`]] });
    data.push({ range:`${S}!F${r}`, values:[[`=IFERROR(IF(B${r}=0,"No Budget Set",IF(C${r}>B${r},"Over Budget",IF(C${r}/B${r}>0.85,"Near Budget","Under Budget"))),"—")`]] });
  });

  // Expense log header row 29
  data.push({ range:`${S}!A29`, values:[['EXPENSE LOG']] });
  data.push({ range:`${S}!A30:N30`, values:[[
    'Expense ID','Date','Category','Vendor / Store','Item / Service',
    'Qty','Unit Cost','Estimated Cost','Actual Cost','Payment Status',
    'Due Date','Paid Date','Receipt Saved?','Notes'
  ]] });

  // Expense ID formulas rows 31-229
  for (let r = 31; r <= 229; r++) {
    data.push({ range:`${S}!A${r}`, values:[[`=IF(E${r}="","","EXP-"&TEXT(ROW()-30,"000"))`]] });
    data.push({ range:`${S}!H${r}`, values:[[`=IFERROR(IF(F${r}=""," ",F${r}*G${r}),"")`]] });
  }

  // Sample expense data rows 31-55
  EXPENSES.forEach(([date,cat,vendor,item,qty,unitCost,payStatus,dueDate,paidDate,receipt,notes], i) => {
    const r = 31 + i;
    data.push({ range:`${S}!B${r}`, values:[[date]] });
    data.push({ range:`${S}!C${r}`, values:[[cat]] });
    data.push({ range:`${S}!D${r}`, values:[[vendor]] });
    data.push({ range:`${S}!E${r}`, values:[[item]] });
    data.push({ range:`${S}!F${r}`, values:[[qty]] });
    data.push({ range:`${S}!G${r}`, values:[[unitCost]] });
    data.push({ range:`${S}!I${r}`, values:[[qty * unitCost]] });
    data.push({ range:`${S}!J${r}`, values:[[payStatus]] });
    if (dueDate) data.push({ range:`${S}!K${r}`, values:[[dueDate]] });
    if (paidDate) data.push({ range:`${S}!L${r}`, values:[[paidDate]] });
    data.push({ range:`${S}!M${r}`, values:[[receipt]] });
    if (notes) data.push({ range:`${S}!N${r}`, values:[[notes]] });
  });

  await valuesBatchUpdate(id, data, 'budget-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(BE,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };
  const currF  = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  const pctF   = { type:'PERCENT', pattern:'0%' };

  // Title rows 0-1
  reqs.push({ mergeCells:{ range:gridRange(BE,0,2,0,14), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Guidance row 2
  reqs.push({ mergeCells:{ range:gridRange(BE,2,4,0,14), mergeType:'MERGE_ALL' } });
  fmt(2,4,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Summary cards rows 4-5 (indices)
  // 4 pairs per row across A-H (cols 0-7)
  for (let pair = 0; pair < 4; pair++) {
    const c = pair * 2;
    // Label row index 4, value row index 5
    reqs.push({ mergeCells:{ range:gridRange(BE,4,5,c,c+1), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(BE,4,5,c+1,c+2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(BE,5,6,c,c+1), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(BE,5,6,c+1,c+2), mergeType:'MERGE_ALL' } });
    fmt(4,5,c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.lavender), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(4,5,c+1,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:14, fontFamily:'Arial', foregroundColor:hex(C.primary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', numberFormat:currF } });
    fmt(5,6,c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.mutedBlue), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(5,6,c+1,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:12, fontFamily:'Arial', foregroundColor:hex(C.secondary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', numberFormat:currF } });
  }
  // Row 5 H6 col index 7 remaining budget — color negative red
  reqs.push({ updateBorders:{ range:gridRange(BE,4,6,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Category table section header row 8 (index 8)
  reqs.push({ mergeCells:{ range:gridRange(BE,8,9,0,6), mergeType:'MERGE_ALL' } });
  fmt(8,9,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Category header row 9 (index 9)
  fmt(9,10,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  // Category data rows 10-24
  for (let r = 10; r < 25; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,6,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols C(2),D(3),E(4),F(5) for category table
  for (const ci of [2,3,4,5]) {
    reqs.push({ repeatCell:{ range:gridRange(BE,10,25,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  // Input col B(1) planned budget
  reqs.push({ repeatCell:{ range:gridRange(BE,10,25,1,2), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  // Currency for B,C,D cols of category table
  for (const ci of [1,2,3]) {
    reqs.push({ repeatCell:{ range:gridRange(BE,10,25,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currF, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  }
  // Pct col E(4)
  reqs.push({ repeatCell:{ range:gridRange(BE,10,25,4,5), cell:{ userEnteredFormat:{ numberFormat:pctF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  // Status CF col F(5)
  const statusColors = [
    { val:'Under Budget', bg:C.success, fg:C.white },
    { val:'Near Budget', bg:C.warning, fg:C.mainText },
    { val:'Over Budget', bg:C.attention, fg:C.white },
    { val:'No Budget Set', bg:C.border, fg:C.secText },
  ];
  statusColors.forEach(({ val, bg, fg }) => {
    reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(BE,10,25,5,6)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:val }] }, format:{ backgroundColor:hex(bg), textFormat:{ foregroundColor:hex(fg), bold:true } } } }, index:0 } });
  });
  reqs.push({ updateBorders:{ range:gridRange(BE,9,25,0,6), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Expense log section header row 28 (index 28)
  reqs.push({ mergeCells:{ range:gridRange(BE,28,29,0,14), mergeType:'MERGE_ALL' } });
  fmt(28,29,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Expense log header row 29 (index 29)
  fmt(29,30,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });
  // Data rows 30-228
  for (let r = 30; r < 229; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,14,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols A(0),H(7)
  for (const ci of [0,7]) {
    reqs.push({ repeatCell:{ range:gridRange(BE,30,229,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  // Input cols B,C,D,E,F,G,I,J,K,L,N (1,2,3,4,5,6,8,9,10,11,13)
  for (const ci of [1,2,3,4,5,6,8,9,10,11,13]) {
    reqs.push({ repeatCell:{ range:gridRange(BE,30,229,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  }
  // Currency cols G(6),H(7),I(8)
  for (const ci of [6,7,8]) {
    reqs.push({ repeatCell:{ range:gridRange(BE,30,229,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currF, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  }
  // Date cols B(1),K(10),L(11)
  for (const ci of [1,10,11]) {
    reqs.push({ repeatCell:{ range:gridRange(BE,30,229,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Checkbox M(12) Receipt Saved
  for (let r = 30; r < 229; r++) {
    reqs.push({ setDataValidation:{ range:gridRange(BE,r,r+1,12,13), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
  }
  // Dropdown Category C(2), Payment Status J(9)
  const catRule = { condition:{ type:'ONE_OF_LIST', values:CATEGORIES.map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true };
  const payRule = { condition:{ type:'ONE_OF_LIST', values:['Planned','Deposit Paid','Partially Paid','Paid','Refunded','Cancelled'].map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true };
  reqs.push({ setDataValidation:{ range:gridRange(BE,30,229,2,3), rule:catRule } });
  reqs.push({ setDataValidation:{ range:gridRange(BE,30,229,9,10), rule:payRule } });

  // Payment Status CF col J(9)
  const payColors = [
    { val:'Paid', bg:C.success, fg:C.white },
    { val:'Deposit Paid', bg:C.mutedBlue, fg:C.mainText },
    { val:'Partially Paid', bg:C.warning, fg:C.mainText },
    { val:'Planned', bg:C.bg, fg:C.secText },
    { val:'Cancelled', bg:C.border, fg:C.secText },
  ];
  payColors.forEach(({ val, bg, fg }) => {
    reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(BE,30,229,9,10)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:val }] }, format:{ backgroundColor:hex(bg), textFormat:{ foregroundColor:hex(fg) } } } }, index:0 } });
  });

  reqs.push({ updateBorders:{ range:gridRange(BE,29,229,0,14), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  [[0,80],[1,100],[2,120],[3,130],[4,200],[5,50],[6,80],[7,90],[8,90],[9,100],[10,100],[11,100],[12,90],[13,180]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:BE, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:BE, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:BE, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:BE, dimension:'ROWS', startIndex:4, endIndex:6 }, properties:{ pixelSize:42 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:BE, dimension:'ROWS', startIndex:6, endIndex:230 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  reqs.push({ updateSheetProperties:{ properties:{ sheetId:BE, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' } });

  // Charts — build summary data first for charts
  // Chart 1: Donut — Actual by Category (use category table C col)
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Actual Spending by Category',
    pieChart:{
      legendPosition:'LABELED_LEGEND',
      domain:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:9, endRowIndex:24, startColumnIndex:0, endColumnIndex:1 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:9, endRowIndex:24, startColumnIndex:2, endColumnIndex:3 }] } },
      pieHole:0.4
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:BE, rowIndex:9, columnIndex:7 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:260 } } } } });

  // Chart 2: Grouped column — Planned vs Actual
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Planned vs Actual Budget by Category',
    basicChart:{
      chartType:'COLUMN',
      legendPosition:'BOTTOM_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Category' },
        { position:'LEFT_AXIS', title:'Amount ($)' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:9, endRowIndex:24, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:9, endRowIndex:24, startColumnIndex:1, endColumnIndex:2 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:9, endRowIndex:24, startColumnIndex:2, endColumnIndex:3 }] } }, targetAxis:'LEFT_AXIS' },
      ],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:BE, rowIndex:9, columnIndex:11 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:260 } } } } });

  // Chart 3: Bar — Paid vs Unpaid (summary in a helper range)
  const data2 = [];
  data2.push({ range:`${S}!I9`, values:[['Payment Status']] });
  data2.push({ range:`${S}!J9`, values:[['Amount ($)']] });
  const payStatuses = ['Paid','Deposit Paid','Partially Paid','Planned','Cancelled'];
  payStatuses.forEach((ps, i) => {
    const r = 10 + i;
    data2.push({ range:`${S}!I${r}`, values:[[ps]] });
    data2.push({ range:`${S}!J${r}`, values:[[`=IFERROR(SUMIF($J$31:$J$229,"${ps}",$I$31:$I$229),0)`]] });
  });
  await valuesBatchUpdate(id, data2, 'budget-values2');

  reqs.push({ addChart:{ chart:{ spec:{
    title:'Expenses by Payment Status',
    basicChart:{
      chartType:'BAR',
      legendPosition:'NO_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Amount ($)' },
        { position:'LEFT_AXIS', title:'Payment Status' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:8, endRowIndex:14, startColumnIndex:8, endColumnIndex:9 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:BE, startRowIndex:8, endRowIndex:14, startColumnIndex:9, endColumnIndex:10 }] } }, targetAxis:'BOTTOM_AXIS' }],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:BE, rowIndex:16, columnIndex:7 }, offsetXPixels:0, offsetYPixels:0, widthPixels:760, heightPixels:200 } } } } });

  await batchUpdate(id, reqs, 'budget-format');
  console.log('Budget & Expenses complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
