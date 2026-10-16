'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const IL = sheetMap['Income Log'];
const S = "'Income Log'";

const QUARTER = d => { const m=new Date(d).getMonth()+1; return m<=3?'Q1':m<=6?'Q2':m<=9?'Q3':'Q4'; };

// 30 sample income rows [date, enterprise, category, customer, description, qty, unit, amount, status]
const INCOME = [
  ['2026-01-08','Crops','Crop Sales','Morrison Grain Co.','Winter wheat — 320 bu',320,'bushels',2800,'Received'],
  ['2026-01-15','Livestock','Livestock Sales','Lakeside Stockyards','Feeder calves x6',6,'head',3500,'Received'],
  ['2026-02-05','Dairy','Dairy Sales','Valley Dairy Coop','February milk delivery',890,'lbs',1200,'Received'],
  ['2026-02-18','Poultry','Egg & Poultry Sales','Local Market','Eggs — 32 dozen',32,'dozen',320,'Received'],
  ['2026-03-03','Hay & Forage','Hay & Forage Sales','Ridgewood Stables','Round bales x18',18,'bales',1800,'Received'],
  ['2026-03-12','Livestock','Livestock Sales','Lakeside Stockyards','Weaned calves x5',5,'head',4200,'Pending'],
  ['2026-03-25','Farm Stand','Farm Stand Sales','Direct Customers','Early spring starts',1,'season',480,'Received'],
  ['2026-04-10','Produce','Produce Sales','Green Valley Market','Spring lettuce & radish',85,'lbs',950,'Received'],
  ['2026-04-22','Crops','Crop Sales','Morrison Grain Co.','Cover crop seed resale',100,'bags',1400,'Received'],
  ['2026-04-28','Apiary','Value-Added Product Sales','Local Market','Spring honey — 8 lbs',8,'lbs',380,'Received'],
  ['2026-05-06','Produce','Produce Sales','Green Valley Market','Spinach, kale, chard',210,'lbs',2100,'Received'],
  ['2026-05-14','Farm Stand','Farm Stand Sales','Direct Customers','Spring vegetable mix',1,'lot',1650,'Received'],
  ['2026-05-20','Livestock','Livestock Sales','Tri-County Auction','Finished steers x4',4,'head',5800,'Received'],
  ['2026-06-09','Dairy','Dairy Sales','Valley Dairy Coop','June milk delivery',1080,'lbs',1450,'Received'],
  ['2026-06-18','Produce','Produce Sales','Restaurant Direct','Zucchini, peas, beans',280,'lbs',2800,'Partially Received'],
  ['2026-07-08','Farm Stand','Farm Stand Sales','Direct Customers','Summer peak — week 1',1,'week',3200,'Received'],
  ['2026-07-15','Produce','Produce Sales','Green Valley Market','Tomatoes, peppers, corn',350,'lbs',3500,'Received'],
  ['2026-07-22','Apiary','Value-Added Product Sales','Farmers Market','Summer honey harvest',24,'lbs',720,'Received'],
  ['2026-08-05','Farm Stand','Farm Stand Sales','Direct Customers','Summer peak — week 2',1,'week',4100,'Received'],
  ['2026-08-12','Crops','Crop Sales','Morrison Grain Co.','Corn harvest — 680 bu',680,'bushels',8500,'Received'],
  ['2026-08-20','Livestock','Livestock Sales','Tri-County Auction','Market hogs x8',8,'head',6200,'Received'],
  ['2026-09-04','Poultry','Egg & Poultry Sales','Local Market','Eggs — 64 dozen',64,'dozen',580,'Received'],
  ['2026-09-16','Hay & Forage','Hay & Forage Sales','Ridgewood Stables','Second cutting x30 bales',30,'bales',2400,'Pending'],
  ['2026-10-07','Crops','Crop Sales','Morrison Grain Co.','Soybean harvest — 800 bu',800,'bushels',12000,'Received'],
  ['2026-10-15','Farm Stand','Farm Stand Sales','Direct Customers','Fall harvest sales',1,'season',2800,'Received'],
  ['2026-10-22','Apiary','Value-Added Product Sales','Local Market','Fall honey — 12 lbs',12,'lbs',460,'Received'],
  ['2026-11-05','Livestock','Livestock Sales','Lakeside Stockyards','Finished steers x5',5,'head',7500,'Received'],
  ['2026-11-18','Dairy','Dairy Sales','Valley Dairy Coop','November milk delivery',1000,'lbs',1350,'Received'],
  ['2026-12-09','Hay & Forage','Hay & Forage Sales','Ridgewood Stables','Stored hay x20 bales',20,'bales',1600,'Received'],
  ['2026-12-17','Poultry','Egg & Poultry Sales','Local Market','December eggs — 20 dozen',20,'dozen',240,'Received'],
];

(async () => {
  const data = [];

  // ── Header rows 1-5 ────────────────────────────────────────────────────────
  data.push({ range:`${S}!A1`, values:[['INCOME LOG — Cedar Ridge Farm']] });
  data.push({ range:`${S}!A2`, values:[['Record all farm income here. Use the reporting tabs (Monthly View, Annual Summary, Dashboard) to analyze results.']] });
  data.push({ range:`${S}!A4:N4`, values:[['Income ID','Date','Year','Month','Quarter','Enterprise','Income Category','Customer / Buyer','Description','Quantity','Unit','Gross Income','Payment Status','Notes']] });

  // ── Auto formulas for rows 6:505 ──────────────────────────────────────────
  for (let r = 6; r <= 505; r++) {
    data.push({ range:`${S}!A${r}`, values:[[`=IF(B${r}="","","INC-"&TEXT(ROW()-5,"0000"))`]] });
    data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(YEAR(B${r}),"")`]] });
    data.push({ range:`${S}!D${r}`, values:[[`=IFERROR(TEXT(B${r},"mmmm"),"")`]] });
    data.push({ range:`${S}!E${r}`, values:[[`=IFERROR(IF(MONTH(B${r})<=3,"Q1",IF(MONTH(B${r})<=6,"Q2",IF(MONTH(B${r})<=9,"Q3","Q4"))),"")`]] });
  }

  // ── Sample data ────────────────────────────────────────────────────────────
  INCOME.forEach(([date, ent, cat, cust, desc, qty, unit, amt, status], i) => {
    const r = 6 + i;
    data.push({ range:`${S}!B${r}:N${r}`, values:[[date,null,null,null,ent,cat,cust,desc,qty,unit,amt,status,'']] });
  });

  await valuesBatchUpdate(id, data, 'income-values');

  // ── Formatting ──────────────────────────────────────────────────────────────
  const reqs = [];
  const merge = (r1,r2,c1,c2) => reqs.push({ mergeCells:{ range:gridRange(IL,r1,r2,c1,c2), mergeType:'MERGE_ALL' } });
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(IL,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });

  // Title
  merge(0,1,0,13); merge(1,2,0,13);
  fmt(0,1,0,13,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:16, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,2,0,13,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Row 3 spacer
  fmt(2,3,0,13,{ userEnteredFormat:{ backgroundColor:hex(C.bg) } });

  // Header row 4
  fmt(3,4,0,13,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Row 5 spacer
  fmt(4,5,0,13,{ userEnteredFormat:{ backgroundColor:hex(C.primary) } });

  // Data rows 6-505
  // Formula cols A,C,D,E = pale blue (formula)
  fmt(5,505,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(5,505,2,5,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  // Input cols B,F-N = white
  fmt(5,505,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:9, fontFamily:'Arial' }, numberFormat:{ type:'DATE', pattern:'MMM D, YYYY' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(5,505,5,13,{ userEnteredFormat:{ backgroundColor:hex(C.white), textFormat:{ foregroundColor:hex(C.mainText), fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  // Amount col L (index 11) = currency
  fmt(5,505,11,12,{ userEnteredFormat:{ backgroundColor:hex(C.white), textFormat:{ fontSize:9, fontFamily:'Arial' }, numberFormat:{ type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' }, horizontalAlignment:'RIGHT', verticalAlignment:'MIDDLE' } });

  // Alternating rows
  for (let r = 5; r < 505; r+=2) {
    fmt(r,r+1,0,13,{ userEnteredFormat:{ backgroundColor:hex(C.altRow) } });
  }

  // Borders
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID',        color:hex(C.border) };
  reqs.push({ updateBorders:{ range:gridRange(IL,3,505,0,13), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Col widths: A=90, B=100, C=50, D=90, E=40, F=110, G=140, H=130, I=180, J=60, K=70, L=100, M=110, N=130
  [[0,90],[1,100],[2,50],[3,90],[4,40],[5,110],[6,140],[7,130],[8,180],[9,60],[10,70],[11,100],[12,110],[13,130]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:IL, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:IL, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:32 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:IL, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:IL, dimension:'ROWS', startIndex:5, endIndex:505 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Unmerge title rows before freezing columns (merge spans A:N, blocking col freeze)
  reqs.push({ unmergeCells:{ range:gridRange(IL,0,1,0,13) } });
  reqs.push({ unmergeCells:{ range:gridRange(IL,1,2,0,13) } });

  // Freeze rows 1:5, cols A:B
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:IL, gridProperties:{ frozenRowCount:5, frozenColumnCount:2 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'income-format');
  console.log('Income Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
