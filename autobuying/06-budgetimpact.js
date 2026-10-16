'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const MB = sheetMap['Monthly Budget Impact'];
const S  = "'Monthly Budget Impact'";
const VQ = "'Vehicle Details & Quotes'";
const BS = "'Buyer Setup'";
const RD = "'Reference Data'";

// Vehicle IDs we populate columns for
const VEHICLES = ['V-001','V-002','V-003','V-004','V-005','V-006','V-007','V-008'];
// Data cols: D,E,F,G,H,I,J,K (index 3-10)
const colLetters = ['D','E','F','G','H','I','J','K'];

(async () => {
  const data = [];

  data.push({ range: `${S}!A1`, values: [['MONTHLY BUDGET IMPACT']] });
  data.push({ range: `${S}!A2`, values: [['Side-by-side monthly cost comparison across vehicles. All currency values in buyer\'s currency.']] });
  data.push({ range: `${S}!A3`, values: [['Affordability thresholds are planning guidelines only — not lender criteria. Edit in Reference Data.']] });

  // ── Threshold inputs (row 5) ───────────────────────────────────────────────
  data.push({ range: `${S}!A5:B5`, values: [['BUDGET CONTROLS','']] });
  data.push({ range: `${S}!A6:B6`, values: [['Manageable Threshold (% of income)', 0.15]] });
  data.push({ range: `${S}!A7:B7`, values: [['Tight Threshold (% of income)', 0.20]] });
  data.push({ range: `${S}!A8:B8`, values: [['Fuel Price ($/gal)', 3.50]] });
  data.push({ range: `${S}!A9:B9`, values: [['Electricity Price ($/kWh)', 0.16]] });
  data.push({ range: `${S}!A10:B10`, values: [['EV Efficiency (miles/kWh)', 3.5]] });

  // Vehicle ID row
  data.push({ range: `${S}!A12`, values: [['Vehicle ID']] });
  VEHICLES.forEach((vid, i) => {
    data.push({ range: `${S}!${colLetters[i]}12`, values: [[vid]] });
  });

  // Make / Model row
  data.push({ range: `${S}!A13`, values: [['Make / Model']] });
  VEHICLES.forEach((vid, i) => {
    const col = colLetters[i];
    data.push({ range: `${S}!${col}13`, values: [[
      `=IFERROR(INDEX(${VQ}!$C$6:$C$105,MATCH(${col}12,${VQ}!$A$6:$A$105,0))&" "&INDEX(${VQ}!$D$6:$D$105,MATCH(${col}12,${VQ}!$A$6:$A$105,0)),"")`
    ]] });
  });

  // ── Monthly cost categories (rows 15-24) ──────────────────────────────────
  const costLabels = [
    'Loan / Lease Payment',
    'Insurance (Monthly)',
    'Fuel / Electricity',
    'Maintenance',
    'Registration / Taxes (Monthly)',
    'Parking',
    'Tolls',
    'Warranty / Service Plan',
    'Other',
    'TOTAL Monthly Vehicle Cost',
  ];

  costLabels.forEach((label, i) => {
    data.push({ range: `${S}!A${i+15}`, values: [[label]] });
  });

  // Formulas for each vehicle column
  VEHICLES.forEach((vid, vi) => {
    const col = colLetters[vi];
    const matchBase = `MATCH(${col}12,${VQ}!$A$6:$A$105,0)`;

    // Row 15: Loan/Lease Payment — use PMT approximation from OTD, APR, Term, DP
    const otd    = `IFERROR(INDEX(${VQ}!$W$6:$W$105,${matchBase}),0)`;
    const apr    = `IFERROR(INDEX(${VQ}!$X$6:$X$105,${matchBase}),0)`;
    const term   = `IFERROR(INDEX(${VQ}!$Y$6:$Y$105,${matchBase}),60)`;
    const dp     = `IFERROR(INDEX(${VQ}!$V$6:$V$105,${matchBase}),0)`;
    const tradeV = `IFERROR('Buyer Setup'!$B$19,0)`;
    const tradeLB = `IFERROR('Buyer Setup'!$B$20,0)`;
    const amtFin  = `MAX(0,${otd}-${dp}-${tradeV}+${tradeLB})`;
    data.push({ range: `${S}!${col}15`, values: [[`=IFERROR(IF(${apr}=0,0,-PMT(${apr}/12,${term},${amtFin})),0)`]] });

    // Row 16: Insurance
    data.push({ range: `${S}!${col}16`, values: [[`=IFERROR(INDEX(${VQ}!$AA$6:$AA$105,${matchBase}),0)`]] });

    // Row 17: Fuel / Electricity
    const mpg    = `IFERROR(INDEX(${VQ}!$Z$6:$Z$105,${matchBase}),1)`;
    const fuel   = `IFERROR(INDEX(${VQ}!$I$6:$I$105,${matchBase}),"Gasoline")`;
    const miles  = `IFERROR('Buyer Setup'!$B$21,12000)`;
    const fuelPr = `$B$8`;
    const elecPr = `$B$9`;
    const elecEff = `$B$10`;
    data.push({ range: `${S}!${col}17`, values: [[
      `=IFERROR(IF(${fuel}="Electric",(${miles}/${elecEff})*${elecPr}/12,` +
      `IF(${fuel}="Plug-In Hybrid",(${miles}/${elecEff})*${elecPr}/24+(${miles}/${mpg})*${fuelPr}/24,` +
      `(${miles}/${mpg})*${fuelPr}/12)),0)`
    ]] });

    // Row 18: Maintenance
    data.push({ range: `${S}!${col}18`, values: [[`=IFERROR(INDEX(${VQ}!$AB$6:$AB$105,${matchBase}),0)`]] });

    // Row 19: Registration (annual default / 12)
    data.push({ range: `${S}!${col}19`, values: [[`=IFERROR(${RD}!$M$7/12,25)`]] });

    // Row 20: Parking (input — default 0)
    data.push({ range: `${S}!${col}20`, values: [[0]] });

    // Row 21: Tolls (input — default 0)
    data.push({ range: `${S}!${col}21`, values: [[0]] });

    // Row 22: Warranty / Service Plan (input — default 0)
    data.push({ range: `${S}!${col}22`, values: [[0]] });

    // Row 23: Other (input — default 0)
    data.push({ range: `${S}!${col}23`, values: [[0]] });

    // Row 24: Total
    data.push({ range: `${S}!${col}24`, values: [[`=SUM(${col}15:${col}23)`]] });
  });

  // ── Budget metrics (rows 26-33) ───────────────────────────────────────────
  const budgetLabels = [
    'Total Monthly Income',
    'Existing Expenses (excl. vehicle)',
    'Monthly Savings Goal',
    'Monthly Vehicle Cost',
    'Remaining Discretionary Income',
    'Vehicle Cost as % of Income',
    'Total Transportation % of Income',
    'Affordability Status',
  ];

  budgetLabels.forEach((label, i) => {
    data.push({ range: `${S}!A${i+26}`, values: [[label]] });
  });

  VEHICLES.forEach((vid, vi) => {
    const col = colLetters[vi];
    const income   = `IFERROR('Buyer Setup'!$B$31,0)`;
    const expenses = `IFERROR('Buyer Setup'!$B$32,0)`;
    const savings  = `IFERROR('Buyer Setup'!$B$13,0)`;
    const vehCost  = `${col}24`;

    data.push({ range: `${S}!${col}26`, values: [[`=${income}`]] });
    data.push({ range: `${S}!${col}27`, values: [[`=${expenses}`]] });
    data.push({ range: `${S}!${col}28`, values: [[`=${savings}`]] });
    data.push({ range: `${S}!${col}29`, values: [[`=${vehCost}`]] });
    data.push({ range: `${S}!${col}30`, values: [[`=IFERROR(${col}26-${col}27-${col}28-${col}29,0)`]] });
    data.push({ range: `${S}!${col}31`, values: [[`=IFERROR(${col}24/${col}26,0)`]] });
    data.push({ range: `${S}!${col}32`, values: [[`=IFERROR((${col}24)/${col}26,0)`]] });

    // Affordability status
    data.push({ range: `${S}!${col}33`, values: [[
      `=IF(${col}24="","Insufficient Data",` +
      `IF(${col}30<0,"Over Budget",` +
      `IF(${col}31>$B$7,"Tight",` +
      `IF(${col}31>$B$6,"Manageable","Comfortable"))))`
    ]] });
  });

  await valuesBatchUpdate(id, data, 'budget-impact-values');

  // ── Formatting ─────────────────────────────────────────────────────────────
  const reqs = [];

  reqs.push({ repeatCell: { range: gridRange(MB,0,500,0,12), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Title
  reqs.push({ mergeCells: { range: gridRange(MB,0,1,0,11), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(MB,0,1,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:15 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:MB, dimension:'ROWS', startIndex:0, endIndex:1 }, properties: { pixelSize:44 }, fields:'pixelSize' }});

  [1,2].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(MB,ri,ri+1,0,11), mergeType:'MERGE_ALL' }});
    reqs.push({ repeatCell: { range: gridRange(MB,ri,ri+1,0,11), cell: { userEnteredFormat: {
      backgroundColor: ri===1 ? hex(C.secondary) : hex(C.mutedBlue),
      textFormat: { italic:true, foregroundColor: hex(C.white), fontSize:9 },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});
  });

  // Controls header (row 5, index 4)
  reqs.push({ repeatCell: { range: gridRange(MB,4,5,0,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Input cells B6:B10 (index 5-9)
  reqs.push({ repeatCell: { range: gridRange(MB,5,10,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  reqs.push({ repeatCell: { range: gridRange(MB,5,7,1,2), cell: { userEnteredFormat: { numberFormat: { type:'PERCENT', pattern:'0.0%' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Vehicle ID header row (row 12, index 11)
  reqs.push({ repeatCell: { range: gridRange(MB,11,12,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white) },
    horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  // Make/Model row (index 12)
  reqs.push({ repeatCell: { range: gridRange(MB,12,13,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.mutedBlue), textFormat: { bold:true, foregroundColor: hex(C.mainText) },
    horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  // Cost section labels (A15:A24, index 14-23)
  reqs.push({ repeatCell: { range: gridRange(MB,14,24,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { foregroundColor: hex(C.mainText) }, wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)' }});

  // Total row (A24, index 23)
  reqs.push({ repeatCell: { range: gridRange(MB,23,24,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold:true, foregroundColor: hex(C.white) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Vehicle data cells — formula rows: 15,17,18,24 (index 14,16,17,23) — pale blue
  [14,16,17,23].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(MB,ri,ri+1,3,11), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});
  });

  // Input rows (20-23 = parking, tolls, warranty, other, index 19-22) — pale yellow
  [19,20,21,22].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(MB,ri,ri+1,3,11), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  });

  // Currency format for cost rows (index 14-23) and budget rows (25-30)
  const currFmt = { type:'CURRENCY', pattern:'$#,##0.00' };
  reqs.push({ repeatCell: { range: gridRange(MB,14,24,3,11), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(MB,25,31,3,11), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});

  // % format rows 31-32 (index 30-31)
  reqs.push({ repeatCell: { range: gridRange(MB,30,32,3,11), cell: { userEnteredFormat: { numberFormat: { type:'PERCENT', pattern:'0.0%' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Budget metrics label col
  reqs.push({ repeatCell: { range: gridRange(MB,25,34,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { foregroundColor: hex(C.mainText) }, wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)' }});

  // Column widths
  reqs.push({ updateDimensionProperties: { range: { sheetId:MB, dimension:'COLUMNS', startIndex:0, endIndex:1 }, properties: { pixelSize:240 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:MB, dimension:'COLUMNS', startIndex:1, endIndex:2 }, properties: { pixelSize:120 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:MB, dimension:'COLUMNS', startIndex:3, endIndex:11 }, properties: { pixelSize:130 }, fields:'pixelSize' }});

  // Freeze rows 1:6
  reqs.push({ updateSheetProperties: { properties: { sheetId:MB, gridProperties: { frozenRowCount:6 } }, fields:'gridProperties.frozenRowCount' }});

  await batchUpdate(id, reqs, 'budget-impact-format');
  console.log('✓ Monthly Budget Impact');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
