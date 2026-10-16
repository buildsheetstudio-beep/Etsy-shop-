'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Dashboard'];
const S  = "'Dashboard'";
const VQ = "'Vehicle Details & Quotes'";
const BS = "'Buyer Setup'";
const MB = "'Monthly Budget Impact'";
const TCO= "'Total Cost of Ownership'";
const VC = "'Vehicle Comparison'";
const LC = "'Auto Loan Calculator'";

(async () => {
  const data = [];

  // ── Rows 1-2: Header ──────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['AUTO BUYING & BUDGET DASHBOARD']] });
  data.push({ range: `${S}!A2`, values: [['Affordability  •  Financing  •  Ownership Cost  •  Lease vs Buy']] });

  // ── Row 4-5: Controls ─────────────────────────────────────────────────────
  data.push({ range: `${S}!A4`, values: [['DASHBOARD CONTROLS']] });
  data.push({ range: `${S}!A5:B5`, values: [['Selected Vehicle ID', 'V-001']] });
  data.push({ range: `${S}!C5:D5`, values: [['Purchase Method', 'Finance']] });
  data.push({ range: `${S}!E5:F5`, values: [['Ownership Horizon (yrs)', 5]] });
  data.push({ range: `${S}!A6:B6`, values: [['Annual Miles', `=IFERROR('Buyer Setup'!B21,12000)`]] });
  data.push({ range: `${S}!C6:D6`, values: [['Fuel Price ($/gal)', 3.50]] });
  data.push({ range: `${S}!E6:F6`, values: [['Electricity Price ($/kWh)', 0.16]] });

  // ── Primary KPI Cards (rows 8-9: label, rows 10-11: value) ────────────────
  // 4 cards per row, 2 rows = 8 primary KPIs
  data.push({ range: `${S}!A8`, values: [['Out-the-Door Price']] });
  data.push({ range: `${S}!C8`, values: [['Monthly Payment']] });
  data.push({ range: `${S}!E8`, values: [['Monthly Vehicle Cost']] });
  data.push({ range: `${S}!G8`, values: [['Available Cash After Vehicle']] });
  data.push({ range: `${S}!A9`, values: [[
    `=IFERROR(INDEX(${VQ}!$W$6:$W$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),"")`
  ]] });
  data.push({ range: `${S}!C9`, values: [[
    `=IFERROR(IF(INDEX(${VQ}!$X$6:$X$105,MATCH(B5,${VQ}!$A$6:$A$105,0))=0,0,` +
    `-PMT(INDEX(${VQ}!$X$6:$X$105,MATCH(B5,${VQ}!$A$6:$A$105,0))/12,` +
    `INDEX(${VQ}!$Y$6:$Y$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),` +
    `MAX(0,A9-IFERROR('Buyer Setup'!B18,0)-IFERROR('Buyer Setup'!B19,0)+IFERROR('Buyer Setup'!B20,0)))),0)`
  ]] });
  data.push({ range: `${S}!E9`, values: [[
    `=IFERROR(C9+IFERROR(INDEX(${VQ}!$AA$6:$AA$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),0)+` +
    `IF(IFERROR(INDEX(${VQ}!$I$6:$I$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),"Gasoline")="Electric",` +
    `(B6/3.5)*D6/12,(B6/IFERROR(INDEX(${VQ}!$Z$6:$Z$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),28))*D6/12)+` +
    `IFERROR(INDEX(${VQ}!$AB$6:$AB$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),0)+25,0)`
  ]] });
  data.push({ range: `${S}!G9`, values: [[
    `=IFERROR('Buyer Setup'!B31-'Buyer Setup'!B32-'Buyer Setup'!B13-E9,0)`
  ]] });

  data.push({ range: `${S}!A11`, values: [['5-Year Ownership Cost']] });
  data.push({ range: `${S}!C11`, values: [['Cost per Mile']] });
  data.push({ range: `${S}!E11`, values: [['Emergency Fund Remaining']] });
  data.push({ range: `${S}!G11`, values: [['Affordability Status']] });

  // OTD for TCO calc
  const otd_db = `IFERROR(A9,0)`;
  const apr_db = `IFERROR(INDEX(${VQ}!$X$6:$X$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),0)`;
  const term_db= `IFERROR(INDEX(${VQ}!$Y$6:$X$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),60)`;
  const insM_db= `IFERROR(INDEX(${VQ}!$AA$6:$AA$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),0)`;
  const maintM_db=`IFERROR(INDEX(${VQ}!$AB$6:$AB$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),0)`;
  const cond_db = `IFERROR(INDEX(${VQ}!$G$6:$G$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),"Used")`;
  const resale5_db= `${otd_db}*IF(${cond_db}="New",0.88^5,0.90^5)`;
  const totInt5_db= `IF(${apr_db}=0,0,-PMT(${apr_db}/12,${term_db},MAX(0,${otd_db}-IFERROR('Buyer Setup'!B18,0)-IFERROR('Buyer Setup'!B19,0)+IFERROR('Buyer Setup'!B20,0)))*MIN(${term_db},60)-MAX(0,${otd_db}-IFERROR('Buyer Setup'!B18,0)))`;

  data.push({ range: `${S}!A12`, values: [[
    `=IFERROR(${otd_db}+${totInt5_db}+(${insM_db}*60)+E9/12*60*12+(${maintM_db}*60)+25*5-${resale5_db},0)`
  ]] });
  data.push({ range: `${S}!C12`, values: [[`=IFERROR(A12/(B6*5),0)`]] });
  data.push({ range: `${S}!E12`, values: [[`=IFERROR('Buyer Setup'!B14-'Buyer Setup'!B18,0)`]] });
  data.push({ range: `${S}!G12`, values: [[
    `=IF(E9="","Insufficient Data",` +
    `IF(G9<0,"Over Budget",` +
    `IF(E9/'Buyer Setup'!B31>0.20,"Tight",` +
    `IF(E9/'Buyer Setup'!B31>0.15,"Manageable","Comfortable"))))`
  ]] });

  // ── Secondary KPI Cards (rows 14-17) ─────────────────────────────────────
  data.push({ range: `${S}!A14`, values: [['Best Overall Vehicle']] });
  data.push({ range: `${S}!C14`, values: [['Lowest Monthly Cost']] });
  data.push({ range: `${S}!E14`, values: [['Lowest Long-Term Cost']] });
  data.push({ range: `${S}!G14`, values: [['Best Finance Option']] });
  data.push({ range: `${S}!A15`, values: [[`=IFERROR(INDEX('Vehicle Comparison'!B17:B26,MATCH(MAX('Vehicle Comparison'!U17:U26),'Vehicle Comparison'!U17:U26,0)),"")`]] });
  data.push({ range: `${S}!C15`, values: [[`=IFERROR(INDEX('Vehicle Comparison'!B17:B26,MATCH(MIN('Vehicle Comparison'!H17:H26),'Vehicle Comparison'!H17:H26,0)),"")`]] });
  data.push({ range: `${S}!E15`, values: [[`=IFERROR(INDEX('Vehicle Comparison'!B17:B26,MATCH(MIN('Vehicle Comparison'!I17:I26),'Vehicle Comparison'!I17:I26,0)),"")`]] });
  data.push({ range: `${S}!G15`, values: [[`=IFERROR(INDEX('Vehicle Comparison'!B17:B26,MATCH(MIN('Vehicle Comparison'!G17:G26),'Vehicle Comparison'!G17:G26,0)),"")`]] });

  data.push({ range: `${S}!A17`, values: [['Total Interest (Scenario A)']] });
  data.push({ range: `${S}!C17`, values: [['Trade-In Equity']] });
  data.push({ range: `${S}!E17`, values: [['Daily Cost']] });
  data.push({ range: `${S}!G17`, values: [['Lowest Cost per Mile']] });
  data.push({ range: `${S}!A18`, values: [[`=IFERROR('Auto Loan Calculator'!E16,0)`]] });
  data.push({ range: `${S}!C18`, values: [[`=IFERROR('Buyer Setup'!B35,0)`]] });
  data.push({ range: `${S}!E18`, values: [[`=IFERROR(E9*12/365,0)`]] });
  data.push({ range: `${S}!G18`, values: [[`=IFERROR(INDEX('Vehicle Comparison'!B17:B26,MATCH(MIN('Vehicle Comparison'!J17:J26),'Vehicle Comparison'!J17:J26,0)),"")`]] });

  // ── Decision Snapshot (row 20+) ───────────────────────────────────────────
  data.push({ range: `${S}!A20`, values: [['DECISION SNAPSHOT']] });
  const snapLabels = [
    'Selected Vehicle', 'Make / Model', 'Out-the-Door Price', 'Down Payment',
    'Trade-In Equity', 'APR / Term', 'Monthly Payment',
    'Insurance (Mo.)', 'Fuel / Electricity (Mo.)', 'Maintenance (Mo.)',
    'Remaining Monthly Cash', '5-Year Ownership Cost', 'Cost per Mile', 'Affordability Status',
  ];
  const snapVals = [
    `=B5`,
    `=IFERROR(INDEX(${VQ}!$C$6:$C$105,MATCH(B5,${VQ}!$A$6:$A$105,0))&" "&INDEX(${VQ}!$D$6:$D$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),"")`,
    `=IFERROR(A9,0)`,
    `=IFERROR('Buyer Setup'!B18,0)`,
    `=IFERROR('Buyer Setup'!B35,0)`,
    `=IFERROR(TEXT(INDEX(${VQ}!$X$6:$X$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),"0.00%")&" / "&INDEX(${VQ}!$Y$6:$Y$105,MATCH(B5,${VQ}!$A$6:$A$105,0))&" mo","")`,
    `=IFERROR(C9,0)`,
    `=IFERROR(INDEX(${VQ}!$AA$6:$AA$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),0)`,
    `=IFERROR(E9-C9-IFERROR(INDEX(${VQ}!$AA$6:$AA$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),0)-IFERROR(INDEX(${VQ}!$AB$6:$AB$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),0),0)`,
    `=IFERROR(INDEX(${VQ}!$AB$6:$AB$105,MATCH(B5,${VQ}!$A$6:$A$105,0)),0)`,
    `=IFERROR(G9,0)`,
    `=IFERROR(A12,0)`,
    `=IFERROR(C12,0)`,
    `=IFERROR(G12,"")`,
  ];
  snapLabels.forEach((label, i) => {
    data.push({ range: `${S}!A${i+21}:B${i+21}`, values: [[label, snapVals[i]]] });
  });

  // ── Disclaimer ─────────────────────────────────────────────────────────────
  const discRow = 37;
  data.push({ range: `${S}!A${discRow}`, values: [[
    'For educational comparison only. Not lending, financial, insurance, tax, legal, or purchasing advice. ' +
    'Replace estimates with actual dealer, lender, insurer, tax, registration, fuel, maintenance, and resale information before making a decision.'
  ]] });

  await valuesBatchUpdate(id, data, 'dashboard-values');

  // ── Formatting ─────────────────────────────────────────────────────────────
  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(DB,0,500,0,10), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Row 1: Main title
  reqs.push({ mergeCells: { range: gridRange(DB,0,1,0,9), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(DB,0,1,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:20 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:DB, dimension:'ROWS', startIndex:0, endIndex:1 }, properties: { pixelSize:56 }, fields:'pixelSize' }});

  // Row 2: Subtitle
  reqs.push({ mergeCells: { range: gridRange(DB,1,2,0,9), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(DB,1,2,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary),
    textFormat: { bold:false, foregroundColor: hex(C.white), fontSize:11, italic:true },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});

  // Controls section (row 4, index 3)
  reqs.push({ repeatCell: { range: gridRange(DB,3,4,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Control input cells (rows 5-6, B5,D5,F5,B6,D6,F6 — index 4-5)
  reqs.push({ repeatCell: { range: gridRange(DB,4,6,0,9), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Primary KPI cards — label rows (8,11 → index 7,10): bold, primary bg
  [7,10].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(DB,ri,ri+1,0,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:10 },
      wrapStrategy:'WRAP', horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }}, fields:'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ updateDimensionProperties: { range: { sheetId:DB, dimension:'ROWS', startIndex:ri, endIndex:ri+1 }, properties: { pixelSize:28 }, fields:'pixelSize' }});
  });

  // Value rows (9,12 → index 8,11): large formula text
  [8,11].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(DB,ri,ri+1,0,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { bold:true, foregroundColor: hex(C.primary), fontSize:14 },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
    reqs.push({ updateDimensionProperties: { range: { sheetId:DB, dimension:'ROWS', startIndex:ri, endIndex:ri+1 }, properties: { pixelSize:36 }, fields:'pixelSize' }});
  });

  // Currency format on primary KPIs
  const currFmt = { type:'CURRENCY', pattern:'$#,##0.00' };
  // A9,C9,E9,G9 = index col 0,2,4,6
  [0,2,4,6].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(DB,8,9,ci,ci+1), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});
  });
  // A12,C12 = OC + cost/mi
  reqs.push({ repeatCell: { range: gridRange(DB,11,12,0,1), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(DB,11,12,2,3), cell: { userEnteredFormat: { numberFormat: { type:'NUMBER', pattern:'$0.000' } } }, fields:'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(DB,11,12,4,5), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});

  // Secondary KPI section (rows 14-18, index 13-17)
  reqs.push({ repeatCell: { range: gridRange(DB,13,14,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:10 },
    wrapStrategy:'WRAP', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:DB, dimension:'ROWS', startIndex:13, endIndex:14 }, properties: { pixelSize:28 }, fields:'pixelSize' }});

  [14,16].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(DB,ri,ri+1,0,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:9 },
      wrapStrategy:'WRAP', verticalAlignment:'MIDDLE'
    }}, fields:'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' }});
  });
  [15,17].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(DB,ri,ri+1,0,9), cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel), textFormat: { bold:true, foregroundColor: hex(C.secondary), fontSize:12 },
      verticalAlignment:'MIDDLE'
    }}, fields:'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
    reqs.push({ updateDimensionProperties: { range: { sheetId:DB, dimension:'ROWS', startIndex:ri, endIndex:ri+1 }, properties: { pixelSize:32 }, fields:'pixelSize' }});
  });

  // Currency for secondary KPIs (A18, C18, E18)
  [0,2,4].forEach(ci => {
    reqs.push({ repeatCell: { range: gridRange(DB,17,18,ci,ci+1), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});
  });

  // Decision snapshot section (row 20, index 19)
  reqs.push({ mergeCells: { range: gridRange(DB,19,20,0,9), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(DB,19,20,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Snapshot rows (21-34, index 20-33)
  reqs.push({ repeatCell: { range: gridRange(DB,20,34,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.white), textFormat: { bold:false, foregroundColor: hex(C.mainText) }, wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)' }});
  reqs.push({ repeatCell: { range: gridRange(DB,20,34,1,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula), textFormat: { bold:true, foregroundColor: hex(C.primary) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Currency rows in snapshot: rows 23,24,25,27,28,29,30,31,32 (index 22-31 selectively)
  [22,23,24,26,27,28,29,30].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(DB,ri,ri+1,1,2), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});
  });
  // Cost per mile (index 31)
  reqs.push({ repeatCell: { range: gridRange(DB,31,32,1,2), cell: { userEnteredFormat: { numberFormat: { type:'NUMBER', pattern:'$0.000' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Disclaimer (row 37, index 36)
  reqs.push({ mergeCells: { range: gridRange(DB,36,39,0,9), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(DB,36,39,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex('#E8E6E1'),
    textFormat: { italic:true, foregroundColor: hex(C.secText), fontSize:9 },
    wrapStrategy:'WRAP', verticalAlignment:'MIDDLE',
    borders: {
      top:    { style:'SOLID_MEDIUM', color: hex(C.slate) },
      bottom: { style:'SOLID_MEDIUM', color: hex(C.slate) },
      left:   { style:'SOLID_MEDIUM', color: hex(C.slate) },
      right:  { style:'SOLID_MEDIUM', color: hex(C.slate) },
    }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment,borders)' }});

  // Column widths
  const dbColWidths = [[0,140],[1,140],[2,140],[3,140],[4,140],[5,140],[6,140],[7,140],[8,140]];
  dbColWidths.forEach(([ci,px]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId:DB, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties: { pixelSize:px }, fields:'pixelSize' }});
  });

  // Freeze rows 1:5 (index 0:5)
  reqs.push({ updateSheetProperties: { properties: { sheetId:DB, gridProperties: { frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' }});

  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('✓ Dashboard');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
