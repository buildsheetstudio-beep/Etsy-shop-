'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const VC = sheetMap['Vehicle Comparison'];
const S  = "'Vehicle Comparison'";
const VQ = "'Vehicle Details & Quotes'";
const MB = "'Monthly Budget Impact'";
const TCO= "'Total Cost of Ownership'";

// 10 vehicles
const VEHICLES = ['V-001','V-002','V-003','V-004','V-005','V-006','V-007','V-008','V-009','V-011'];

(async () => {
  const data = [];

  data.push({ range: `${S}!A1`, values: [['VEHICLE COMPARISON & SCORING']] });
  data.push({ range: `${S}!A2`, values: [['Compare up to 10 vehicles on cost, practicality, and personal preference. Adjust weights below to reflect your priorities.']] });
  data.push({ range: `${S}!A3`, values: [['Lower cost = higher score for financial metrics. All weights must total 100%. Scores are relative — not absolute ratings.']] });

  // ── Scoring Weights (row 5-14) ─────────────────────────────────────────────
  data.push({ range: `${S}!A5`, values: [['SCORING WEIGHTS']] });
  const weights = [
    ['Monthly Affordability',    0.25],
    ['Total Cost of Ownership',  0.25],
    ['Purchase Price',           0.15],
    ['Fuel Cost',                0.10],
    ['Insurance',                0.05],
    ['Resale Value',             0.10],
    ['Practicality',             0.05],
    ['Personal Preference',      0.05],
  ];
  weights.forEach(([label, w], i) => {
    data.push({ range: `${S}!A${i+6}:B${i+6}`, values: [[label, w]] });
  });
  data.push({ range: `${S}!A14:B14`, values: [['WEIGHT TOTAL', `=SUM(B6:B13)`]] });

  // ── Column headers (row 16) ────────────────────────────────────────────────
  const headers = [
    'Vehicle ID','Make / Model','OTD Price','Down Payment','Trade-In Equity','Amt Financed',
    'Monthly Payment','Total Monthly Cost','5-Yr Ownership Cost','Cost per Mile',
    'Est. Resale Value','Insurance (Mo.)','Fuel/Elec (Mo.)','Maintenance (Mo.)',
    'Warranty','Passenger Cap.','Cargo/Utility Score','Safety/Comfort Score',
    'Preference Score','Affordability','Weighted Score','Recommendation'
  ];
  data.push({ range: `${S}!A16`, values: [headers] });

  // Sample vehicle IDs (rows 17-26)
  VEHICLES.forEach((vid, i) => {
    const r = i + 17;
    data.push({ range: `${S}!A${r}`, values: [[vid]] });
  });

  // Populate formulas for rows 17-26
  VEHICLES.forEach((vid, vi) => {
    const r = vi + 17;
    const mBase = `MATCH(A${r},${VQ}!$A$6:$A$105,0)`;
    const mMB   = `MATCH(A${r},'Monthly Budget Impact'!$D$15:'Monthly Budget Impact'!$K$15,0)+3`; // offset col D

    // B: Make/Model
    data.push({ range: `${S}!B${r}`, values: [[
      `=IFERROR(INDEX(${VQ}!$C$6:$C$105,${mBase})&" "&INDEX(${VQ}!$D$6:$D$105,${mBase}),"")`
    ]] });

    // C: OTD Price
    data.push({ range: `${S}!C${r}`, values: [[`=IFERROR(INDEX(${VQ}!$W$6:$W$105,${mBase}),"")`]] });

    // D: Down Payment
    data.push({ range: `${S}!D${r}`, values: [[`=IFERROR(INDEX(${VQ}!$V$6:$V$105,${mBase}),"")`]] });

    // E: Trade-In Equity
    data.push({ range: `${S}!E${r}`, values: [[`=IFERROR('Buyer Setup'!$B$35,0)`]] });

    // F: Amt Financed
    data.push({ range: `${S}!F${r}`, values: [[
      `=IFERROR(MAX(0,C${r}-D${r}-IFERROR('Buyer Setup'!$B$19,0)+IFERROR('Buyer Setup'!$B$20,0)),0)`
    ]] });

    // G: Monthly Payment
    const apr  = `IFERROR(INDEX(${VQ}!$X$6:$X$105,${mBase}),0)`;
    const term = `IFERROR(INDEX(${VQ}!$Y$6:$Y$105,${mBase}),60)`;
    data.push({ range: `${S}!G${r}`, values: [[
      `=IFERROR(IF(${apr}=0,0,-PMT(${apr}/12,${term},F${r})),0)`
    ]] });

    // H: Total Monthly Cost — G + insurance + fuel + maint + reg/12
    const insM  = `IFERROR(INDEX(${VQ}!$AA$6:$AA$105,${mBase}),0)`;
    const maintM= `IFERROR(INDEX(${VQ}!$AB$6:$AB$105,${mBase}),0)`;
    const mpg   = `IFERROR(INDEX(${VQ}!$Z$6:$Z$105,${mBase}),28)`;
    const fuel  = `IFERROR(INDEX(${VQ}!$I$6:$I$105,${mBase}),"Gasoline")`;
    const fuelCost = `IF(${fuel}="Electric",(12000/3.5)*0.16/12,(12000/${mpg})*3.50/12)`;
    data.push({ range: `${S}!H${r}`, values: [[
      `=IFERROR(G${r}+${insM}+${fuelCost}+${maintM}+25,0)`
    ]] });

    // I: 5-yr ownership cost — simplified
    const otd  = `IFERROR(INDEX(${VQ}!$W$6:$W$105,${mBase}),0)`;
    const cond = `IFERROR(INDEX(${VQ}!$G$6:$G$105,${mBase}),"Used")`;
    const resale5 = `${otd}*IF(${cond}="New",0.88^5,0.90^5)`;
    const totInt5 = `IF(${apr}=0,0,IFERROR(-PMT(${apr}/12,${term},F${r})*MIN(${term},60)-F${r},0))`;
    data.push({ range: `${S}!I${r}`, values: [[
      `=IFERROR(${otd}+${totInt5}+(${insM}*60)+(${fuelCost}*60)+(${maintM}*60)+25*5-${resale5},0)`
    ]] });

    // J: Cost per mile
    data.push({ range: `${S}!J${r}`, values: [[`=IFERROR(I${r}/(12000*5),0)`]] });

    // K: Resale value (5yr)
    data.push({ range: `${S}!K${r}`, values: [[`=IFERROR(${resale5},0)`]] });

    // L: Insurance
    data.push({ range: `${S}!L${r}`, values: [[`=IFERROR(INDEX(${VQ}!$AA$6:$AA$105,${mBase}),0)`]] });

    // M: Fuel/Elec
    data.push({ range: `${S}!M${r}`, values: [[`=IFERROR(${fuelCost},0)`]] });

    // N: Maintenance
    data.push({ range: `${S}!N${r}`, values: [[`=IFERROR(INDEX(${VQ}!$AB$6:$AB$105,${mBase}),0)`]] });

    // O: Warranty
    data.push({ range: `${S}!O${r}`, values: [[`=IFERROR(INDEX(${VQ}!$AC$6:$AC$105,${mBase}),"")`]] });

    // P: Passenger capacity (editable default 5)
    data.push({ range: `${S}!P${r}`, values: [[5]] });

    // Q: Cargo/utility score (editable default 5)
    data.push({ range: `${S}!Q${r}`, values: [[5]] });

    // R: Safety/comfort score (editable default 5)
    data.push({ range: `${S}!R${r}`, values: [[5]] });

    // S: Preference score (editable default 5)
    data.push({ range: `${S}!S${r}`, values: [[5]] });

    // T: Affordability
    data.push({ range: `${S}!T${r}`, values: [[
      `=IF(H${r}="","Insufficient Data",` +
      `IF(H${r}/IFERROR('Buyer Setup'!$B$31,5800)>0.20,"Over Budget",` +
      `IF(H${r}/IFERROR('Buyer Setup'!$B$31,5800)>0.15,"Tight",` +
      `IF(H${r}/IFERROR('Buyer Setup'!$B$31,5800)>0.12,"Manageable","Comfortable"))))`
    ]] });

    // U: Weighted Score
    // Normalize cost metrics: lower = better. Score = 1 - (val - MIN) / (MAX - MIN)
    // For preference: higher = better.
    // We compute inline for all vehicles using full range G17:G26, H17:H26, etc.
    const normCostLow = (col, rng) =>
      `IF(MAX(${rng})-MIN(${rng})=0,0.5,1-(${col}${r}-MIN(${rng}))/(MAX(${rng})-MIN(${rng})))`;
    const normCostHigh = (col, rng) =>
      `IF(MAX(${rng})-MIN(${rng})=0,0.5,(${col}${r}-MIN(${rng}))/(MAX(${rng})-MIN(${rng})))`;

    data.push({ range: `${S}!U${r}`, values: [[
      `=IFERROR(` +
      `$B$6*${normCostLow('H','H17:H26')}+` +
      `$B$7*${normCostLow('I','I17:I26')}+` +
      `$B$8*${normCostLow('C','C17:C26')}+` +
      `$B$9*${normCostLow('M','M17:M26')}+` +
      `$B$10*${normCostLow('L','L17:L26')}+` +
      `$B$11*${normCostHigh('K','K17:K26')}+` +
      `$B$12*${normCostHigh('Q','Q17:Q26')}+` +
      `$B$13*${normCostHigh('S','S17:S26')},0)`
    ]] });

    // V: Recommendation
    data.push({ range: `${S}!V${r}`, values: [[
      `=IF(RANK(U${r},$U$17:$U$26,0)=1,"Best Overall",` +
      `IF(H${r}=MIN($H$17:$H$26),"Best Monthly Cost",` +
      `IF(I${r}=MIN($I$17:$I$26),"Best Long-Term Value",` +
      `IF(T${r}="Comfortable","Review Carefully",""))))`
    ]] });
  });

  // ── Result cards (rows 28-35) ──────────────────────────────────────────────
  const cardStart = 28;
  data.push({ range: `${S}!A${cardStart}`, values: [['RESULT CARDS']] });
  const cards = [
    ['Best Overall',         `=IFERROR(INDEX(B17:B26,MATCH(MAX(U17:U26),U17:U26,0)),"")`],
    ['Lowest Monthly Cost',  `=IFERROR(INDEX(B17:B26,MATCH(MIN(H17:H26),H17:H26,0)),"")`],
    ['Lowest 5-Year Cost',   `=IFERROR(INDEX(B17:B26,MATCH(MIN(I17:I26),I17:I26,0)),"")`],
    ['Lowest Cost per Mile', `=IFERROR(INDEX(B17:B26,MATCH(MIN(J17:J26),J17:J26,0)),"")`],
    ['Highest Resale Est.',  `=IFERROR(INDEX(B17:B26,MATCH(MAX(K17:K26),K17:K26,0)),"")`],
    ['Best Budget Fit',      `=IFERROR(INDEX(B17:B26,MATCH(MAX(IF(T17:T26="Comfortable",U17:U26,0)),IF(T17:T26="Comfortable",U17:U26,0),0)),"")`],
  ];
  cards.forEach(([label, formula], i) => {
    data.push({ range: `${S}!A${cardStart+1+i}:B${cardStart+1+i}`, values: [[label, formula]] });
  });

  await valuesBatchUpdate(id, data, 'vehicle-comparison-values');

  // ── Formatting ─────────────────────────────────────────────────────────────
  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(VC,0,500,0,22), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Title
  reqs.push({ mergeCells: { range: gridRange(VC,0,1,0,22), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(VC,0,1,0,22), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:15 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:VC, dimension:'ROWS', startIndex:0, endIndex:1 }, properties: { pixelSize:44 }, fields:'pixelSize' }});

  [1,2].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(VC,ri,ri+1,0,22), mergeType:'MERGE_ALL' }});
    reqs.push({ repeatCell: { range: gridRange(VC,ri,ri+1,0,22), cell: { userEnteredFormat: {
      backgroundColor: ri===1 ? hex(C.secondary) : hex(C.mutedBlue),
      textFormat: { italic:true, foregroundColor: hex(C.white), fontSize:9 },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});
  });

  // Weights section
  reqs.push({ repeatCell: { range: gridRange(VC,4,5,0,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});
  reqs.push({ repeatCell: { range: gridRange(VC,5,13,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.input), numberFormat: { type:'PERCENT', pattern:'0%' } } }, fields:'userEnteredFormat(backgroundColor,numberFormat)' }});

  // Weight total row — conditional via bold
  reqs.push({ repeatCell: { range: gridRange(VC,13,14,0,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold:true, foregroundColor: hex(C.white) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Column headers (row 16, index 15)
  reqs.push({ repeatCell: { range: gridRange(VC,15,16,0,22), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:9 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:VC, dimension:'ROWS', startIndex:15, endIndex:16 }, properties: { pixelSize:40 }, fields:'pixelSize' }});

  // Data rows alternating
  for (let i = 0; i < 10; i++) {
    const ri = 16 + i;
    reqs.push({ repeatCell: { range: gridRange(VC,ri,ri+1,0,22), cell: { userEnteredFormat: { backgroundColor: i%2===0 ? hex(C.white) : hex(C.altRow) } }, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Input cols P,Q,R,S (15-18) — pale yellow
  reqs.push({ repeatCell: { range: gridRange(VC,16,26,15,19), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Formula cols — pale blue
  reqs.push({ repeatCell: { range: gridRange(VC,16,26,1,15), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});
  reqs.push({ repeatCell: { range: gridRange(VC,16,26,19,22), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Currency formats C-N (2-13)
  const currFmt = { type:'CURRENCY', pattern:'$#,##0.00' };
  reqs.push({ repeatCell: { range: gridRange(VC,16,26,2,14), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});

  // Cost per mile J (9)
  reqs.push({ repeatCell: { range: gridRange(VC,16,26,9,10), cell: { userEnteredFormat: { numberFormat: { type:'NUMBER', pattern:'$0.000' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Weighted score U (20) as percent
  reqs.push({ repeatCell: { range: gridRange(VC,16,26,20,21), cell: { userEnteredFormat: { numberFormat: { type:'NUMBER', pattern:'0.00' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Result cards
  reqs.push({ mergeCells: { range: gridRange(VC,cardStart-1,cardStart,0,22), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(VC,cardStart-1,cardStart,0,22), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});
  reqs.push({ repeatCell: { range: gridRange(VC,cardStart,cardStart+cards.length,0,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.success), textFormat: { bold:true, foregroundColor: hex(C.mainText) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Column widths
  reqs.push({ updateDimensionProperties: { range: { sheetId:VC, dimension:'COLUMNS', startIndex:0, endIndex:1 }, properties: { pixelSize:70 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:VC, dimension:'COLUMNS', startIndex:1, endIndex:2 }, properties: { pixelSize:130 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:VC, dimension:'COLUMNS', startIndex:2, endIndex:15 }, properties: { pixelSize:110 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:VC, dimension:'COLUMNS', startIndex:15, endIndex:19 }, properties: { pixelSize:80 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:VC, dimension:'COLUMNS', startIndex:19, endIndex:22 }, properties: { pixelSize:120 }, fields:'pixelSize' }});

  reqs.push({ updateSheetProperties: { properties: { sheetId:VC, gridProperties: { frozenRowCount:6 } }, fields:'gridProperties.frozenRowCount' }});

  await batchUpdate(id, reqs, 'vehicle-comparison-format');
  console.log('✓ Vehicle Comparison');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
