'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SL = sheetMap['Shopping List'];
const S = "'Shopping List'";

// [Category, Item, SourceTab, Qty, Unit, Store, EstCost, ActualCost, NeededBy, Purchased, PurchaseDate, Packed, AssignedTo, Notes]
const SHOPPING = [
  // Food items
  ['Food','Mini Slider Burgers (2 trays)','Food, Drinks & Cake',2,'trays','Local Deli',35,'','Aug 13, 2026',false,'','',false,'Sarah','Order by Aug 10'],
  ['Food','Veggie Wraps Platter','Food, Drinks & Cake',1,'platter','Local Deli',35,'','Aug 13, 2026',false,'','',false,'Sarah',''],
  ['Food','Cheese & Crackers Platter','Food, Drinks & Cake',1,'platter','Costco',18,'','Aug 13, 2026',false,'','',false,'Sarah','Check GF crackers'],
  ['Food','Popcorn Bags','Food, Drinks & Cake',2,'bags','Local Store',5,'','Aug 13, 2026',false,'','',false,'Sarah',''],
  ['Food','Sandwich Platter (kids, 2 trays)','Food, Drinks & Cake',2,'trays','Local Deli',28,'','Aug 13, 2026',false,'','',false,'Sarah',''],
  ['Food','Chips & Dips','Food, Drinks & Cake',2,'bags','Costco',9,'','Aug 13, 2026',false,'','',false,'Sarah',''],
  // Drinks
  ['Drinks','Juice Boxes x24','Food, Drinks & Cake',24,'boxes','Costco',18,'','Aug 13, 2026',false,'','',false,'Sarah',''],
  ['Drinks','Sparkling Water x12','Food, Drinks & Cake',12,'bottles','Costco',12,'','Aug 13, 2026',false,'','',false,'Sarah',''],
  ['Drinks','Lemonade Concentrate x4','Food, Drinks & Cake',4,'tins','Local Store',6,'','Aug 13, 2026',false,'','',false,'Sarah',''],
  ['Drinks','Sparkling Grape Juice x2','Food, Drinks & Cake',2,'bottles','Local Store',7,'','Aug 13, 2026',false,'','',false,'Sarah',''],
  // Cake & Desserts
  ['Cake & Desserts','Birthday Cake — pick up Sweet Bloom','Food, Drinks & Cake',1,'cake','Sweet Bloom Bakery',95,'','Aug 14, 2026',false,'','',false,'Sarah','Call to confirm Aug 12'],
  ['Cake & Desserts','GF Cupcakes x4','Food, Drinks & Cake',4,'pcs','Sweet Bloom Bakery',18,'','Aug 14, 2026',false,'','',false,'Sarah','For gluten-free guests'],
  // Decorations
  ['Decorations','Table Centrepieces x6','Budget & Expenses',6,'sets','Party Hut',48,'','Aug 10, 2026',false,'','',false,'Sarah','Garden theme'],
  ['Decorations','Extra Balloons','Manual Entry',1,'bag','Party Hut',12,'','Aug 10, 2026',false,'','',false,'Sarah','Backup supply'],
  ['Decorations','Fairy Lights','Manual Entry',2,'strings','Hardware Store',15,'','Aug 10, 2026',false,'','',false,'Sarah','For venue entrance'],
  // Tableware
  ['Tableware','Extra Plates (20 pack)','Manual Entry',1,'pack','Party Hut',8,'','Aug 12, 2026',false,'','',false,'Sarah','In case of overflow'],
  ['Tableware','Serving Bowls x4','Manual Entry',4,'bowls','Kitchen Supplies',12,'','Aug 12, 2026',false,'','',false,'Sarah',''],
  ['Tableware','Cake Plates (small) x30','Manual Entry',30,'pcs','Party Hut',6,'','Aug 12, 2026',false,'','',false,'Sarah',''],
  // Games & Activities
  ['Activities','Fairy Garden Craft Kits x20','Budget & Expenses',20,'kits','Craft House',60,'60.00','Jul 28, 2026',true,'Jul 28, 2026','',true,'Sarah','Purchased and packed'],
  ['Activities','Pass-the-Parcel Prizes','Manual Entry',1,'set','Party Hut',15,'','Aug 10, 2026',false,'','',false,'Sarah','7 wrapped layers'],
  // Favors
  ['Favors','Party Favor Bags x30','Budget & Expenses',30,'bags','Party Hut',45,'','Aug 10, 2026',false,'','',false,'Sarah','Includes fillers'],
  ['Favors','Thank-You Cards x30','Manual Entry',30,'cards','Stationery Store',8,'','Aug 12, 2026',false,'','',false,'Sarah','Write after party'],
  // Cleaning & Safety
  ['Cleaning','Garbage Bags x2 packs','Manual Entry',2,'packs','Supermarket',5,'','Aug 14, 2026',false,'','',false,'Sarah',''],
  ['Cleaning','Paper Towels x2 rolls','Manual Entry',2,'rolls','Supermarket',4,'','Aug 14, 2026',false,'','',false,'Sarah',''],
  ['Safety','First Aid Kit','Manual Entry',1,'kit','Pharmacy',12,'','Aug 10, 2026',false,'','',false,'Sarah','Check existing kit is stocked'],
];

(async () => {
  const data = [];

  data.push({ range:`${S}!A1`, values:[['PARTY SHOPPING LIST']] });
  data.push({ range:`${S}!A3`, values:[['Consolidate all items to buy here. Mark items as purchased as you shop. The Suggested Items section below lists food items not yet on this list.']] });

  // Summary cards rows 5-6
  data.push({ range:`${S}!A5`, values:[['Total Items']] });
  data.push({ range:`${S}!B5`, values:[[`=IFERROR(COUNTA($C$8:$C$207),0)`]] });
  data.push({ range:`${S}!C5`, values:[['Purchased']] });
  data.push({ range:`${S}!D5`, values:[[`=IFERROR(COUNTIF($K$8:$K$207,TRUE),0)`]] });
  data.push({ range:`${S}!E5`, values:[['Remaining']] });
  data.push({ range:`${S}!F5`, values:[[`=IFERROR(B5-D5,0)`]] });
  data.push({ range:`${S}!G5`, values:[['Est. Total Cost']] });
  data.push({ range:`${S}!H5`, values:[[`=IFERROR(SUM($H$8:$H$207),0)`]] });
  data.push({ range:`${S}!A6`, values:[['Actual Cost']] });
  data.push({ range:`${S}!B6`, values:[[`=IFERROR(SUMIF($I$8:$I$207,">"&0,$I$8:$I$207),0)`]] });
  data.push({ range:`${S}!C6`, values:[['Packed / Ready %']] });
  data.push({ range:`${S}!D6`, values:[[`=IFERROR(IF(B5=0,0,COUNTIF($M$8:$M$207,TRUE)/B5),0)`]] });
  data.push({ range:`${S}!E6`, values:[['Assigned To']] });
  data.push({ range:`${S}!F6`, values:[['Multiple']] });
  data.push({ range:`${S}!G6`, values:[['Cost Saved']] });
  data.push({ range:`${S}!H6`, values:[[`=IFERROR(H5-B6,0)`]] });

  // Section header row 7 (index 7)
  data.push({ range:`${S}!A8`, values:[['SHOPPING LIST']] });
  // Headers row 8 (index 8, but in 1-indexed it's row 9)
  // Wait: A8 says SHOPPING LIST, A9 would be headers
  // Let me clarify: section header at row 8, data at row 9+
  // But per spec: header row at A9, data from A10
  // Actually I set data rows at 8 in valuesBatchUpdate below — let me restart with correct rows

  // Section header row 8, table header row 9, data 10-209 → but spec says rows 8:207
  // Let me adjust: section header A8, headers A9:O9, data rows 10-209
  data.push({ range:`${S}!A9:O9`, values:[[
    'Shopping ID','Category','Item','Source Tab','Quantity','Unit',
    'Store / Vendor','Est. Cost','Actual Cost','Needed By',
    'Purchased?','Purchase Date','Packed / Ready?','Assigned To','Notes'
  ]] });

  // Shopping ID formula rows 10-209
  for (let r = 10; r <= 209; r++) {
    data.push({ range:`${S}!A${r}`, values:[[`=IF(C${r}="","","SHP-"&TEXT(ROW()-9,"000"))`]] });
  }

  // Sample data rows 10-34
  SHOPPING.forEach(([cat,item,source,qty,unit,store,estCost,actCost,neededBy,purchased,purDate,_packed1,packed,assignedTo,notes], i) => {
    const r = 10 + i;
    data.push({ range:`${S}!B${r}`, values:[[cat]] });
    data.push({ range:`${S}!C${r}`, values:[[item]] });
    data.push({ range:`${S}!D${r}`, values:[[source]] });
    data.push({ range:`${S}!E${r}`, values:[[qty]] });
    data.push({ range:`${S}!F${r}`, values:[[unit]] });
    data.push({ range:`${S}!G${r}`, values:[[store]] });
    data.push({ range:`${S}!H${r}`, values:[[estCost]] });
    if (actCost) data.push({ range:`${S}!I${r}`, values:[[actCost]] });
    if (neededBy) data.push({ range:`${S}!J${r}`, values:[[neededBy]] });
    data.push({ range:`${S}!K${r}`, values:[[purchased]] });
    if (purDate) data.push({ range:`${S}!L${r}`, values:[[purDate]] });
    data.push({ range:`${S}!M${r}`, values:[[packed]] });
    data.push({ range:`${S}!N${r}`, values:[[assignedTo]] });
    if (notes) data.push({ range:`${S}!O${r}`, values:[[notes]] });
  });

  // Suggested items section row 212
  data.push({ range:`${S}!A212`, values:[['SUGGESTED ITEMS TO ADD']] });
  data.push({ range:`${S}!A213`, values:[['The items below are planned in Food, Drinks & Cake as "Purchased" but have not yet been added to the shopping list above. Add them manually as needed.']] });
  data.push({ range:`${S}!A214:D214`, values:[['Item Name','Food Type','Qty Needed','Note']] });
  // Helper formulas — SMALL/INDEX/MATCH to list food items where Purchased = "Purchased"
  // Food, Drinks & Cake: E col (HomeOrPurchased) = col 4 (0-indexed), C col = item name
  for (let i = 1; i <= 10; i++) {
    const r = 214 + i;
    data.push({ range:`${S}!A${r}`, values:[[`=IFERROR(INDEX('Food, Drinks & Cake'!$C$10:$C$109,SMALL(IF('Food, Drinks & Cake'!$E$10:$E$109="Purchased",ROW('Food, Drinks & Cake'!$C$10:$C$109)-9),${i})),"")` ]] });
    data.push({ range:`${S}!B${r}`, values:[[`=IFERROR(INDEX('Food, Drinks & Cake'!$B$10:$B$109,SMALL(IF('Food, Drinks & Cake'!$E$10:$E$109="Purchased",ROW('Food, Drinks & Cake'!$B$10:$B$109)-9),${i})),"")` ]] });
    data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(INDEX('Food, Drinks & Cake'!$G$10:$G$109,SMALL(IF('Food, Drinks & Cake'!$E$10:$E$109="Purchased",ROW('Food, Drinks & Cake'!$G$10:$G$109)-9),${i})),"")` ]] });
    data.push({ range:`${S}!D${r}`, values:[['Copy row above to shopping list to add it']] });
  }

  await valuesBatchUpdate(id, data, 'shopping-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(SL,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };
  const currF  = { type:'CURRENCY', pattern:'$#,##0.00' };
  const pctF   = { type:'PERCENT', pattern:'0%' };

  // Title rows 0-1
  reqs.push({ mergeCells:{ range:gridRange(SL,0,2,0,15), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,15,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Guidance rows 2-3
  reqs.push({ mergeCells:{ range:gridRange(SL,2,4,0,15), mergeType:'MERGE_ALL' } });
  fmt(2,4,0,15,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Cards rows 4-5
  for (let i = 0; i < 4; i++) {
    const c = i * 2;
    reqs.push({ mergeCells:{ range:gridRange(SL,4,5,c,c+1), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(SL,4,5,c+1,c+2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(SL,5,6,c,c+1), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(SL,5,6,c+1,c+2), mergeType:'MERGE_ALL' } });
    fmt(4,5,c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.lavender), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(4,5,c+1,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:14, fontFamily:'Arial', foregroundColor:hex(C.primary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(5,6,c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.mutedBlue), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(5,6,c+1,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:14, fontFamily:'Arial', foregroundColor:hex(C.secondary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  }
  // Currency for H5, H6 (indices row 4 col 7, row 5 col 7)
  reqs.push({ repeatCell:{ range:gridRange(SL,4,5,7,8), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(SL,5,6,1,2), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(SL,5,6,3,4), cell:{ userEnteredFormat:{ numberFormat:pctF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(SL,5,6,7,8), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(SL,4,6,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Section header row 7 (index 7), header row 8 (index 8)
  reqs.push({ mergeCells:{ range:gridRange(SL,7,8,0,15), mergeType:'MERGE_ALL' } });
  fmt(7,8,0,15,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(8,9,0,15,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });

  // Data rows 9-208 (indices)
  for (let r = 9; r < 209; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,15,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula col A (0)
  reqs.push({ repeatCell:{ range:gridRange(SL,9,209,0,1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  // Input cols B-O (1-14) except K(10) and M(12) which are checkboxes
  for (const ci of [1,2,3,4,5,6,7,8,9,11,13,14]) {
    reqs.push({ repeatCell:{ range:gridRange(SL,9,209,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  }
  // Currency cols H(7),I(8)
  for (const ci of [7,8]) {
    reqs.push({ repeatCell:{ range:gridRange(SL,9,209,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currF, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  }
  // Date cols J(9),L(11)
  for (const ci of [9,11]) {
    reqs.push({ repeatCell:{ range:gridRange(SL,9,209,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Checkboxes K(10) Purchased, M(12) Packed
  for (let r = 9; r < 209; r++) {
    reqs.push({ setDataValidation:{ range:gridRange(SL,r,r+1,10,11), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
    reqs.push({ setDataValidation:{ range:gridRange(SL,r,r+1,12,13), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
  }
  // Dropdown: Category B(1)
  reqs.push({ setDataValidation:{ range:gridRange(SL,9,209,1,2), rule:{ condition:{ type:'ONE_OF_LIST', values:['Food','Drinks','Cake & Desserts','Decorations','Tableware','Games','Activities','Favors','Cleaning','Safety','Miscellaneous'].map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true } } });
  // Source Tab D(3)
  reqs.push({ setDataValidation:{ range:gridRange(SL,9,209,3,4), rule:{ condition:{ type:'ONE_OF_LIST', values:['Food, Drinks & Cake','Budget & Expenses','Party Setup','Manual Entry'].map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true } } });

  // CF: Purchased = true → sage tint
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(SL,9,209,0,15)], booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=$K10=TRUE` }] }, format:{ backgroundColor:{ red:0.568, green:0.706, blue:0.604, alpha:0.2 } } } }, index:0 } });

  reqs.push({ updateBorders:{ range:gridRange(SL,8,209,0,15), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Suggested items section header row 211 (index 211)
  reqs.push({ mergeCells:{ range:gridRange(SL,211,212,0,4), mergeType:'MERGE_ALL' } });
  fmt(211,212,0,4,{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  reqs.push({ mergeCells:{ range:gridRange(SL,212,213,0,4), mergeType:'MERGE_ALL' } });
  fmt(212,213,0,4,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });
  fmt(213,214,0,4,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 214; r < 225; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,4,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula style for suggested items
  for (const ci of [0,1,2]) {
    reqs.push({ repeatCell:{ range:gridRange(SL,214,225,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(SL,214,225,3,4), cell:{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), fontSize:9, italic:true } } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(SL,213,225,0,4), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  [[0,80],[1,110],[2,200],[3,120],[4,60],[5,60],[6,130],[7,80],[8,90],[9,100],[10,80],[11,100],[12,90],[13,110],[14,170]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:SL, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:SL, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:SL, dimension:'ROWS', startIndex:2, endIndex:9 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:SL, dimension:'ROWS', startIndex:4, endIndex:6 }, properties:{ pixelSize:42 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:SL, dimension:'ROWS', startIndex:9, endIndex:225 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  reqs.push({ updateSheetProperties:{ properties:{ sheetId:SL, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'shopping-format');
  console.log('Shopping List complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
