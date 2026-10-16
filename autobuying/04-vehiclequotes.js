'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const VQ = sheetMap['Vehicle Details & Quotes'];
const S = "'Vehicle Details & Quotes'";
const GL = "'Reference Data'";

(async () => {
  const data = [];

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['VEHICLE DETAILS & QUOTES']] });
  data.push({ range: `${S}!A2`, values: [['Log every quote here. Yellow cells = your inputs. Blue cells = calculated.']] });
  data.push({ range: `${S}!A3`, values: [['Out-the-Door Price = Negotiated Price + Fees + Taxes + Add-Ons − Rebates. Down payment and trade-in reduce Amount Financed, not OTD Price.']] });

  // ── Row 5: Column headers ─────────────────────────────────────────────────
  const headers = [
    'Vehicle ID','Quote ID','Make','Model','Trim','Model Year','Condition','Vehicle Type',
    'Fuel Type','Mileage','Dealer / Seller','Location','Listing Price','Negotiated Price',
    'Destination / Delivery','Documentation Fee','Registration / Title','Est. Sales Tax',
    'Add-Ons / Packages','Rebates / Incentives','Trade-In Credit Applied','Down Payment Applied',
    'Out-the-Door Price','APR','Loan Term (Mo.)','Est. MPG / MPGe',
    'Insurance Est. (Mo.)','Maintenance Est. (Mo.)','Warranty Remaining',
    'Quote Status','Quote Date','Quote Expiration','Notes'
  ];
  data.push({ range: `${S}!A5`, values: [headers] });

  // ── Sample quotes (12 vehicles) ───────────────────────────────────────────
  // Columns: A VehicleID, B=formula, C Make, D Model, E Trim, F Year, G Condition, H Type,
  //          I Fuel, J Mileage, K Dealer, L Location, M ListPrice, N NegPrice,
  //          O Dest, P Doc, Q Reg, R Tax, S AddOns, T Rebates, U TradeIn, V DownPmt,
  //          W=formula OTD, X APR, Y Term, Z MPG, AA Insurance, AB Maintenance, AC Warranty,
  //          AD Status, AE QuoteDate, AF Expiration, AG Notes

  const quotes = [
    // V-001: 2025 Toyota Camry SE, New, Gas, Finance, positive trade-in equity
    ['V-001','','Toyota','Camry','SE',2025,'New','Sedan','Gasoline',5,
     'Westside Toyota','Portland, OR',28990,27800,1095,495,285,2318,0,500,5500,6000,'','0.069',60,32,
     135,60,'3 yr/36k bumper-to-bumper','Selected','Jul 15, 2026','Aug 15, 2026','Best offer — rate match requested'],
    // V-002: 2023 Honda CR-V EX, CPO, Hybrid
    ['V-002','','Honda','CR-V','EX',2023,'Certified Pre-Owned','Crossover','Hybrid',28400,
     'Metro Honda','Portland, OR',31500,30200,0,299,285,2447,0,0,0,4000,'','0.054',48,40,
     145,70,'1 yr/12k remaining powertrain','Quote Received','Jul 10, 2026','Aug 10, 2026','CPO warranty included'],
    // V-003: 2022 Tesla Model 3, Used, EV
    ['V-003','','Tesla','Model 3','Long Range',2022,'Used','Sedan','Electric',34200,
     'Summit Motors','Beaverton, OR',36500,35000,0,499,285,2838,0,3500,5500,6000,'','0.072',60,0,
     155,40,'No factory warranty remaining','Negotiating','Jul 12, 2026','Aug 12, 2026','EV tax credit may apply — verify eligibility'],
    // V-004: 2025 Ford F-150 XLT, New, Gas, over-budget
    ['V-004','','Ford','F-150','XLT',2025,'New','Pickup Truck','Gasoline',8,
     'Valley Ford','Hillsboro, OR',46995,45500,1495,699,285,3687,1200,0,5500,6000,'','0.079',72,20,
     195,110,'3 yr/36k bumper-to-bumper','Quote Received','Jul 8, 2026','Aug 8, 2026','Monthly payment exceeds budget — included for comparison'],
    // V-005: 2024 Toyota RAV4 Hybrid, New, Lease candidate
    ['V-005','','Toyota','RAV4','Hybrid XSE',2024,'New','Crossover','Plug-In Hybrid',9,
     'Eastside Toyota','Gresham, OR',36500,34800,1095,495,285,2815,0,750,0,2000,'','0.0',36,38,
     150,60,'3 yr/36k bumper-to-bumper','Final Offer','Jul 14, 2026','Aug 14, 2026','Lease candidate — money factor 0.00175 residual 58%'],
    // V-006: 2021 Chevy Equinox, Used, affordable monthly / expensive long-term
    ['V-006','','Chevrolet','Equinox','LT',2021,'Used','Crossover','Gasoline',52400,
     'Budget Auto Group','Vancouver, WA',19900,18500,0,399,285,1499,0,0,0,3000,'','0.095',72,28,
     115,90,'Expired','Quote Received','Jul 5, 2026','Aug 5, 2026','Low monthly — high rate and age mean more interest over term'],
    // V-007: 2023 Hyundai Ioniq 6, Used, EV
    ['V-007','','Hyundai','Ioniq 6','SE',2023,'Used','Sedan','Electric',18700,
     'Metro EV Motors','Portland, OR',34000,32500,0,499,285,2633,0,0,0,5000,'','0.065',60,0,
     145,35,'1 yr/12k remaining powertrain','Negotiating','Jul 11, 2026','Aug 11, 2026','Strong EV range — verify charging infrastructure'],
    // V-008: 2024 Kia Telluride SX, New, expensive monthly / better long-term
    ['V-008','','Kia','Telluride','SX',2024,'New','SUV','Gasoline',15,
     'Pacific Kia','Portland, OR',52900,51000,1095,699,285,4133,800,0,0,6000,'','0.075',84,23,
     200,120,'5 yr/60k bumper-to-bumper','Quote Received','Jul 13, 2026','Aug 13, 2026','Premium SUV — long term but very high monthly'],
    // V-009: 2019 Honda Accord LX, Private Party, Cash
    ['V-009','','Honda','Accord','LX',2019,'Private Party','Sedan','Gasoline',74800,
     'Private Seller (Sarah K.)','Portland, OR',17500,16800,0,0,285,1362,0,0,0,0,'','0',0,32,
     110,55,'Expired','Researching','Jul 6, 2026','','Cash purchase — pre-purchase inspection recommended'],
    // V-010: 2024 Toyota Camry XLE, New, Lease — initially cheaper, expensive with mileage
    ['V-010','','Toyota','Camry','XLE',2024,'New','Sedan','Gasoline',11,
     'Westside Toyota','Portland, OR',33990,32500,1095,495,285,2633,0,500,0,2000,'','0.0',36,32,
     140,55,'3 yr/36k bumper-to-bumper','Quote Received','Jul 15, 2026','Aug 15, 2026','Lease — 10k miles/yr cap; excess at $0.25/mi — budget for overages'],
    // V-011: 2022 Subaru Outback, Used, negative trade-in equity
    ['V-011','','Subaru','Outback','Premium',2022,'Used','Wagon','Gasoline',41200,
     'Northwest Subaru','Portland, OR',28500,27200,0,499,285,2205,0,0,8500,6000,'','0.081',60,30,
     130,75,'1 yr/12k remaining powertrain','Negotiating','Jul 9, 2026','Aug 9, 2026','Trade-in has $3k loan balance — negative equity rolls into loan'],
    // V-012: 2023 BMW 330i, Used, reduces emergency fund below min
    ['V-012','','BMW','3 Series','330i xDrive',2023,'Used','Luxury','Gasoline',22000,
     'Prestige Imports','Lake Oswego, OR',48500,46000,0,799,285,3729,1500,0,5500,10000,'','0.089',60,27,
     210,85,'Expired','Researching','Jul 7, 2026','','High down payment reduces emergency fund below $12k minimum'],
  ];

  quotes.forEach((row, i) => {
    const r = i + 6;
    // Write all non-formula cols: A,C-V,X-AC,AD-AG (B and W are formulas)
    data.push({ range: `${S}!A${r}`, values: [[row[0]]] }); // Vehicle ID
    // B is formula — skip for now, written below
    data.push({ range: `${S}!C${r}:V${r}`, values: [[
      row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],
      row[12],row[13],row[14],row[15],row[16],row[17],row[18],row[19],row[20],row[21]
    ]] });
    // W is formula — skip
    data.push({ range: `${S}!X${r}:AC${r}`, values: [[row[23],row[24],row[25],row[26],row[27],row[28]]] });
    data.push({ range: `${S}!AD${r}:AG${r}`, values: [[row[29],row[30],row[31],row[32]]] });
  });

  // ── Formulas for rows 6-17 ────────────────────────────────────────────────
  for (let r = 6; r <= 17; r++) {
    // Quote ID formula (col B = index 1)
    data.push({ range: `${S}!B${r}`, values: [[`=IF(K${r}="","","Q-"&TEXT(ROW()-5,"000"))`]] });
    // Out-the-Door Price (col W = index 22)
    data.push({ range: `${S}!W${r}`, values: [[`=IFERROR(N${r}+O${r}+P${r}+Q${r}+R${r}+S${r}-T${r},"")`]] });
  }

  // Extend blank formula rows 18-105 for user entry
  for (let r = 18; r <= 105; r++) {
    data.push({ range: `${S}!B${r}`, values: [[`=IF(K${r}="","","Q-"&TEXT(ROW()-5,"000"))`]] });
    data.push({ range: `${S}!W${r}`, values: [[`=IFERROR(N${r}+O${r}+P${r}+Q${r}+R${r}+S${r}-T${r},"")`]] });
  }

  await valuesBatchUpdate(id, data, 'vehicle-quotes-values');

  // ── Formatting ─────────────────────────────────────────────────────────────
  const reqs = [];

  // Sheet background
  reqs.push({ repeatCell: { range: gridRange(VQ,0,500,0,35), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Title row (A1, index 0)
  reqs.push({ mergeCells: { range: gridRange(VQ,0,1,0,33), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(VQ,0,1,0,33), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:15 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:VQ, dimension:'ROWS', startIndex:0, endIndex:1 }, properties: { pixelSize:44 }, fields:'pixelSize' }});

  // Row 2 (index 1)
  reqs.push({ mergeCells: { range: gridRange(VQ,1,2,0,33), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(VQ,1,2,0,33), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { italic:true, foregroundColor: hex(C.white), fontSize:10 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Row 3 (index 2)
  reqs.push({ mergeCells: { range: gridRange(VQ,2,3,0,33), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(VQ,2,3,0,33), cell: { userEnteredFormat: {
    backgroundColor: hex(C.mutedBlue), textFormat: { italic:true, foregroundColor: hex(C.mainText), fontSize:9 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:VQ, dimension:'ROWS', startIndex:2, endIndex:3 }, properties: { pixelSize:36 }, fields:'pixelSize' }});

  // Header row (row 5, index 4)
  reqs.push({ repeatCell: { range: gridRange(VQ,4,5,0,33), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:9 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:VQ, dimension:'ROWS', startIndex:4, endIndex:5 }, properties: { pixelSize:36 }, fields:'pixelSize' }});

  // Data rows alternating background
  for (let r = 5; r < 105; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(VQ,r,r+1,0,33), cell: { userEnteredFormat: { backgroundColor: hex(C.white) } }, fields:'userEnteredFormat.backgroundColor' }});
  }
  for (let r = 6; r < 105; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(VQ,r,r+1,0,33), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Input cols (C-V, AA-AG except B and W which are formula): pale yellow
  // Cols C-V = index 2-21, X,Y,Z,AA,AB,AC = 23-28, AD-AG = 29-32
  const inputCols = [[2,22],[23,29],[29,33]];
  inputCols.forEach(([s,e]) => {
    reqs.push({ repeatCell: { range: gridRange(VQ,5,105,s,e), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  });

  // Formula cols B (1) and W (22) — pale blue
  reqs.push({ repeatCell: { range: gridRange(VQ,5,105,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});
  reqs.push({ repeatCell: { range: gridRange(VQ,5,105,22,23), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Currency formats: M,N,O,P,Q,R,S,T,U,V,W (12-22), AA,AB (26-27)
  const currFmt = { type:'CURRENCY', pattern:'$#,##0.00' };
  const redCurrFmt = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  [12,13,14,15,16,17,18,19,20,21,22,26,27].forEach(col => {
    reqs.push({ repeatCell: { range: gridRange(VQ,5,105,col,col+1), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});
  });

  // APR format (col X = 23)
  reqs.push({ repeatCell: { range: gridRange(VQ,5,105,23,24), cell: { userEnteredFormat: { numberFormat: { type:'PERCENT', pattern:'0.00%' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Mileage format (col J = 9)
  reqs.push({ repeatCell: { range: gridRange(VQ,5,105,9,10), cell: { userEnteredFormat: { numberFormat: { type:'NUMBER', pattern:'#,##0' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Column widths
  const colWidths = [
    [0,1,75],[1,2,75],[2,3,80],[3,4,90],[4,5,80],[5,6,60],[6,7,100],[7,8,100],[8,9,90],
    [9,10,80],[10,11,140],[11,12,110],[12,13,110],[13,14,110],[14,15,90],[15,16,90],
    [16,17,90],[17,18,90],[18,19,90],[19,20,90],[20,21,90],[21,22,90],[22,23,110],
    [23,24,70],[24,25,80],[25,26,80],[26,27,100],[27,28,110],[28,29,130],[29,30,110],
    [30,31,100],[31,32,110],[32,33,200]
  ];
  colWidths.forEach(([s,e,px]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId:VQ, dimension:'COLUMNS', startIndex:s, endIndex:e }, properties: { pixelSize:px }, fields:'pixelSize' }});
  });

  // Freeze rows 1:5 only — no column freeze (title merge spans full width, conflicts with frozenColumnCount)
  reqs.push({ updateSheetProperties: { properties: { sheetId:VQ, gridProperties: { frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' }});

  await batchUpdate(id, reqs, 'vehicle-quotes-format');
  console.log('✓ Vehicle Details & Quotes');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
