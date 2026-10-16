'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const EL = sheetMap['Expense Log'];
const S = "'Expense Log'";

// 40 sample expense rows [date, enterprise, category, vendor, description, method, subtotal, tax, status, deductible]
const EXPENSES = [
  ['2026-01-05','Crops','Seed & Plants','AgriSupply Co.','Winter wheat seed — 500 lbs','Checking Account',1800,0,'Paid','Yes'],
  ['2026-01-10','Livestock','Feed','Heritage Feed Mill','Cattle feed — January','Checking Account',2400,0,'Paid','Yes'],
  ['2026-01-20','General Farm','Utilities','Rural Electric Co.','January electric & water','Checking Account',380,0,'Paid','Yes'],
  ['2026-02-03','Livestock','Veterinary','Rolling Hills Vet','Herd health check + vaccines','Checking Account',650,0,'Paid','Yes'],
  ['2026-02-10','General Farm','Insurance','FarmGuard Insurance','Annual farm liability policy','Checking Account',1200,0,'Paid','Yes'],
  ['2026-02-18','Crops','Fertilizer','AgriSupply Co.','Pre-plant fertilizer — 4 tons','Checking Account',2200,0,'Paid','Yes'],
  ['2026-03-05','General Farm','Fuel','Rural Co-op','Diesel — 200 gal','Cash',420,0,'Paid','Yes'],
  ['2026-03-12','General Farm','Equipment Repairs','Cedar County Equipment','Tractor oil change + filters','Checking Account',890,0,'Paid','Yes'],
  ['2026-03-20','Livestock','Feed','Heritage Feed Mill','Cattle feed — March','Checking Account',2100,0,'Paid','Yes'],
  ['2026-03-28','Crops','Chemicals & Pest Control','AgriSupply Co.','Herbicide pre-emergent','Checking Account',560,0,'Paid','Yes'],
  ['2026-04-06','Crops','Seed & Plants','AgriSupply Co.','Corn seed — 25 bags','Checking Account',950,0,'Paid','Yes'],
  ['2026-04-14','Produce','Soil Amendments','Green Earth Supply','Compost & lime — 2 tons','Check',280,0,'Paid','Yes'],
  ['2026-04-22','General Farm','Labor','Seasonal Workers LLC','Spring planting crew — 3 days','Cash',1800,0,'Paid','Yes'],
  ['2026-05-02','Farm Stand','Packaging','PackRight Supply','Produce bags & boxes','Credit Card',180,14.4,'Paid','No'],
  ['2026-05-08','General Farm','Fuel','Rural Co-op','Diesel — 180 gal','Cash',380,0,'Paid','Yes'],
  ['2026-05-16','Livestock','Livestock Purchases','Tri-County Auction','Stocker calves x6','Checking Account',4500,0,'Paid','Yes'],
  ['2026-05-22','General Farm','Utilities','Rural Electric Co.','May irrigation electric','Checking Account',340,0,'Paid','Yes'],
  ['2026-06-04','Produce','Irrigation','Valley Irrigation','Drip tape repair & fittings','Checking Account',520,0,'Paid','Yes'],
  ['2026-06-11','General Farm','Contract Services','Ridge Top Fencing','Fence repair — east pasture','Checking Account',1200,0,'Paid','Yes'],
  ['2026-06-19','Farm Stand','Market & Selling Fees','County Farmers Market','Summer season booth fee','Check',320,0,'Paid','Yes'],
  ['2026-07-07','General Farm','Labor','Seasonal Workers LLC','Hay harvest crew','Cash',2400,0,'Paid','Yes'],
  ['2026-07-15','Farm Stand','Packaging','PackRight Supply','Summer produce packaging','Credit Card',240,19.2,'Paid','No'],
  ['2026-07-22','Apiary','Supplies','Beekeeper Supply Co.','Hive equipment & canning jars','Credit Card',150,12,'Paid','Yes'],
  ['2026-08-04','Crops','Fuel','Rural Co-op','Diesel — combine harvest','Cash',480,0,'Paid','Yes'],
  ['2026-08-12','General Farm','Equipment Repairs','Cedar County Equipment','Combine head repair','Checking Account',1450,0,'Paid','Yes'],
  ['2026-08-20','Livestock','Veterinary','Rolling Hills Vet','Fall pregnancy check','Checking Account',380,0,'Paid','Yes'],
  ['2026-08-27','General Farm','Small Tools','AgriSupply Co.','Hand tools & sprayer parts','Credit Card',290,23.2,'Paid','Yes'],
  ['2026-09-04','General Farm','Labor','Seasonal Workers LLC','Fall harvest labor','Cash',1800,0,'Paid','Yes'],
  ['2026-09-12','Hay & Forage','Equipment Rental','Riverside Equipment','Round baler rental — 2 days','Checking Account',750,0,'Paid','Yes'],
  ['2026-09-22','General Farm','Repairs & Maintenance','Ridge Top Fencing','Barn roof patch','Check',620,0,'Paid','Yes'],
  ['2026-10-06','General Farm','Property Taxes','Cedar County Tax','Farm property tax — annual','Check',1800,0,'Paid','Yes'],
  ['2026-10-14','Crops','Transportation','Sunrise Trucking','Soybean delivery to elevator','Checking Account',450,0,'Paid','Yes'],
  ['2026-10-22','Farm Stand','Packaging','PackRight Supply','Fall packaging supplies','Credit Card',180,14.4,'Paid','No'],
  ['2026-10-29','General Farm','Insurance','FarmGuard Insurance','Farm equipment rider','Checking Account',1200,0,'Paid','Yes'],
  ['2026-11-05','Livestock','Feed','Heritage Feed Mill','Cattle feed — November','Checking Account',2100,0,'Paid','Yes'],
  ['2026-11-13','General Farm','Professional Fees','Harvest Accounting','Annual bookkeeping review','Check',850,0,'Paid','Yes'],
  ['2026-11-20','General Farm','Office & Software','TechFarm Software','Farm management software','Credit Card',180,14.4,'Paid','No'],
  ['2026-12-03','Livestock','Livestock Purchases','Tri-County Auction','Replacement heifers x2','Checking Account',3200,0,'Pending','Yes'],
  ['2026-12-10','General Farm','Education & Training','AgriLearn Online','Online crop management course','Credit Card',350,0,'Paid','No'],
  ['2026-12-18','General Farm','Rent & Lease','Hillside Properties','Annual cropland lease','Checking Account',2400,0,'Paid','Yes'],
];

(async () => {
  const data = [];

  // ── Header rows 1-5 ────────────────────────────────────────────────────────
  data.push({ range:`${S}!A1`, values:[['EXPENSE LOG — Cedar Ridge Farm']] });
  data.push({ range:`${S}!A2`, values:[['Record all farm expenses here. Column O flags potential deductions for the Tax Deduction Summary tab.']] });
  data.push({ range:`${S}!A4:Q4`, values:[['Expense ID','Date','Year','Month','Quarter','Enterprise','Expense Category','Vendor','Description','Payment Method','Subtotal','Sales Tax','Total Expense','Payment Status','Potentially Deductible?','Tax Group','Notes']] });

  // ── Auto formulas for rows 6:505 ──────────────────────────────────────────
  for (let r = 6; r <= 505; r++) {
    data.push({ range:`${S}!A${r}`, values:[[`=IF(B${r}="","","EXP-"&TEXT(ROW()-5,"0000"))`]] });
    data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(YEAR(B${r}),"")`]] });
    data.push({ range:`${S}!D${r}`, values:[[`=IFERROR(TEXT(B${r},"mmmm"),"")`]] });
    data.push({ range:`${S}!E${r}`, values:[[`=IFERROR(IF(MONTH(B${r})<=3,"Q1",IF(MONTH(B${r})<=6,"Q2",IF(MONTH(B${r})<=9,"Q3","Q4"))),"")`]] });
    data.push({ range:`${S}!M${r}`, values:[[`=IFERROR(K${r}+L${r},"")`]] });
    data.push({ range:`${S}!P${r}`, values:[[`=IFERROR(VLOOKUP(G${r},'Reference Data'!$P$2:$Q$31,2,FALSE),"")`]] });
  }

  // ── Sample data ────────────────────────────────────────────────────────────
  EXPENSES.forEach(([date,ent,cat,vendor,desc,method,sub,tax,status,deduct], i) => {
    const r = 6 + i;
    data.push({ range:`${S}!B${r}:Q${r}`, values:[[date,null,null,null,ent,cat,vendor,desc,method,sub,tax,null,status,deduct,null,'']] });
  });

  await valuesBatchUpdate(id, data, 'expense-values');

  // ── Formatting ──────────────────────────────────────────────────────────────
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(EL,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });

  // Title rows 1-2
  fmt(0,1,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:16, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,2,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(2,3,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.bg) } });

  // Header row 4
  fmt(3,4,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(4,5,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.primary) } });

  // Data rows 6-505: formula cols = pale blue, input = white/input
  fmt(5,505,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(5,505,2,5,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(5,505,12,13,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial' }, numberFormat:{ type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' }, horizontalAlignment:'RIGHT', verticalAlignment:'MIDDLE' } });
  fmt(5,505,15,16,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Input cols B,F-L,N,O = white/input
  fmt(5,505,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:9, fontFamily:'Arial' }, numberFormat:{ type:'DATE', pattern:'MMM D, YYYY' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(5,505,5,12,{ userEnteredFormat:{ backgroundColor:hex(C.white), textFormat:{ foregroundColor:hex(C.mainText), fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  // Currency cols K,L (10,11)
  fmt(5,505,10,12,{ userEnteredFormat:{ backgroundColor:hex(C.white), textFormat:{ fontSize:9, fontFamily:'Arial' }, numberFormat:{ type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' }, horizontalAlignment:'RIGHT', verticalAlignment:'MIDDLE' } });
  fmt(5,505,13,15,{ userEnteredFormat:{ backgroundColor:hex(C.white), textFormat:{ foregroundColor:hex(C.mainText), fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Alternating rows
  for (let r = 5; r < 505; r+=2) {
    fmt(r,r+1,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.altRow) } });
  }

  // Borders
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID',        color:hex(C.border) };
  reqs.push({ updateBorders:{ range:gridRange(EL,3,505,0,16), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Col widths
  [[0,90],[1,100],[2,50],[3,90],[4,40],[5,110],[6,150],[7,130],[8,180],[9,120],[10,90],[11,75],[12,100],[13,100],[14,100],[15,160],[16,130]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:EL, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:EL, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:32 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:EL, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:EL, dimension:'ROWS', startIndex:5, endIndex:505 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Freeze rows 1:5, cols A:B
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:EL, gridProperties:{ frozenRowCount:5, frozenColumnCount:2 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'expense-format');
  console.log('Expense Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
