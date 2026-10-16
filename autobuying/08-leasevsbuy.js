'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const LVB = sheetMap['Lease vs Buy'];
const S   = "'Lease vs Buy'";
const VQ  = "'Vehicle Details & Quotes'";
const BS  = "'Buyer Setup'";

(async () => {
  const data = [];

  data.push({ range: `${S}!A1`, values: [['LEASE VS BUY ANALYSIS']] });
  data.push({ range: `${S}!A2`, values: [['Compare estimated total cost of leasing vs buying over the same period. Educational comparison only — confirm all terms with the dealer or lender.']] });
  data.push({ range: `${S}!A3`, values: [['Lease payment formula is an estimate. Money factor and residual must be confirmed with the dealer. Excess mileage, wear charges, and fees vary.']] });

  // ── Column headers ────────────────────────────────────────────────────────
  data.push({ range: `${S}!A5`, values: [['COMPARISON INPUT']] });
  data.push({ range: `${S}!C5:D5`, values: [['LEASE', 'BUY']] });

  // ── Lease inputs ──────────────────────────────────────────────────────────
  const leaseInputs = [
    ['Vehicle ID',               'V-005', ''],
    ['MSRP ($)',                  36500,   ''],
    ['Negotiated Cap Cost ($)',   34800,   ''],
    ['Cap Cost Reduction ($)',     2000,   ''],
    ['Lease Term (months)',          36,   ''],
    ['Money Factor',             0.00175,  ''],
    ['Residual Value (%)',          0.58,   ''],
    ['Annual Mileage Allowance', 12000,   ''],
    ['Acquisition Fee ($)',         650,   ''],
    ['Documentation Fee ($)',       350,   ''],
    ['Registration ($)',            285,   ''],
    ['Sales Tax Rate (%)',         0.081,   ''],
    ['Monthly Lease Payment',
      `=IFERROR((C9-C9*C13)+(C9+C9*C13)*C12*(1+C18),0)`,
      ''],
    ['Disposition Fee ($)',         395,   ''],
    ['Expected Excess Miles',      2000,   ''],
    ['Excess Mileage Rate ($/mi)', 0.25,   ''],
    ['Expected Wear Charges ($)',     0,   ''],
    ['Insurance (Monthly $)',       150,   ''],
    ['Maintenance (Monthly $)',      55,   ''],
  ];

  leaseInputs.forEach(([label, leaseVal, buyVal], i) => {
    data.push({ range: `${S}!A${i+6}`, values: [[label]] });
    data.push({ range: `${S}!C${i+6}`, values: [[leaseVal]] });
  });

  // ── Buy inputs (pulled from VQ + loan calc) ───────────────────────────────
  const buyInputRows = [
    ['Vehicle ID',            'V-001'],
    ['Out-the-Door Price',    `=IFERROR(INDEX(${VQ}!$W$6:$W$105,MATCH(D7,${VQ}!$A$6:$A$105,0)),0)`],
    ['Down Payment',          `=IFERROR('Buyer Setup'!B18,0)`],
    ['APR',                   `=IFERROR(INDEX(${VQ}!$X$6:$X$105,MATCH(D7,${VQ}!$A$6:$A$105,0)),0)`],
    ['Loan Term (months)',    `=IFERROR(INDEX(${VQ}!$Y$6:$Y$105,MATCH(D7,${VQ}!$A$6:$A$105,0)),60)`],
    ['Monthly Payment',       `=IFERROR(IF(D10=0,0,-PMT(D10/12,D11,MAX(0,D8-D9-IFERROR('Buyer Setup'!B19,0)+IFERROR('Buyer Setup'!B20,0)))),0)`],
    ['Total Interest',        `=IFERROR(D12*D11-MAX(0,D8-D9-IFERROR('Buyer Setup'!B19,0)+IFERROR('Buyer Setup'!B20,0)),0)`],
    ['Est. Resale Value',     `=IFERROR(D8*(1-0.12)^(C11/12),0)`],
    ['Insurance (Monthly $)', `=IFERROR(INDEX(${VQ}!$AA$6:$AA$105,MATCH(D7,${VQ}!$A$6:$A$105,0)),120)`],
    ['Maintenance (Monthly $)',`=IFERROR(INDEX(${VQ}!$AB$6:$AB$105,MATCH(D7,${VQ}!$A$6:$A$105,0)),60)`],
    ['Fuel/Elec (Monthly $)', `=IFERROR((IFERROR('Buyer Setup'!B21,12000)/28)*3.50/12,0)`],
  ];

  buyInputRows.forEach(([label, val], i) => {
    data.push({ range: `${S}!D${i+6}`, values: [[val]] });
  });

  // Note: buy inputs use column D starting at row 7 for vehicle ID
  data.push({ range: `${S}!D6`, values: [['V-001']] }); // Buy Vehicle ID

  // ── Comparison Metrics (rows 30+) ─────────────────────────────────────────
  const compStart = 30;
  data.push({ range: `${S}!A${compStart}`, values: [['COMPARISON METRICS']] });
  data.push({ range: `${S}!C${compStart}:D${compStart}`, values: [['LEASE','BUY']] });

  const leaseRefs = {
    upfront:   `=IFERROR(C10+C15+C18,0)`,                       // Cap Reduction + Acq + first mo
    monthly:   `=IFERROR(C19*(1+C18),0)`,                        // payment + tax
    totalPmts: `=IFERROR(C19*(1+C18)*C11,0)`,
    endFees:   `=IFERROR(C20+C21*C22+C23,0)`,                    // disposition + excess miles + wear
    equity:    0,
    cashOut:   `=IFERROR(C${compStart+1}+C${compStart+3}+C${compStart+4},0)`,
    netCost:   `=IFERROR(C${compStart+6}-C${compStart+5},0)`,
    costPmo:   `=IFERROR(C${compStart+7}/C11,0)`,
    costPmi:   `=IFERROR(C${compStart+7}/(C14*C11/12),0)`,
  };

  const buyRefs = {
    upfront:   `=IFERROR(D9,0)`,
    monthly:   `=IFERROR(D12,0)`,
    totalPmts: `=IFERROR(D12*D11,0)`,
    endFees:   0,
    equity:    `=IFERROR(D14,0)`,
    cashOut:   `=IFERROR(D${compStart+1}+D${compStart+3},0)`,
    netCost:   `=IFERROR(D${compStart+6}-D${compStart+5},0)`,
    costPmo:   `=IFERROR(D${compStart+7}/D11,0)`,
    costPmi:   `=IFERROR(D${compStart+7}/(IFERROR('Buyer Setup'!B21,12000)*D11/12),0)`,
  };

  const metrics = [
    ['Upfront Cash Required',   leaseRefs.upfront,   buyRefs.upfront],
    ['Monthly Payment',         leaseRefs.monthly,   buyRefs.monthly],
    ['Total Payments',          leaseRefs.totalPmts, buyRefs.totalPmts],
    ['End-of-Term Fees',        leaseRefs.endFees,   buyRefs.endFees],
    ['Equity / Asset Value',    leaseRefs.equity,    buyRefs.equity],
    ['Total Cash Outflow',      leaseRefs.cashOut,   buyRefs.cashOut],
    ['Estimated Net Cost',      leaseRefs.netCost,   buyRefs.netCost],
    ['Cost per Month',          leaseRefs.costPmo,   buyRefs.costPmo],
    ['Cost per Mile',           leaseRefs.costPmi,   buyRefs.costPmi],
  ];

  metrics.forEach(([label, lease, buy], i) => {
    data.push({ range: `${S}!A${compStart+1+i}:D${compStart+1+i}`, values: [[label,'',lease,buy]] });
  });

  // Recommendation
  const recRow = compStart + metrics.length + 2;
  data.push({ range: `${S}!A${recRow}`, values: [['RECOMMENDATION']] });
  data.push({ range: `${S}!C${recRow}`, values: [[
    `=IF(OR(C${compStart+7}="",D${compStart+7}=""),"Insufficient Data",` +
    `IF(ABS(C${compStart+7}-D${compStart+7})/MAX(C${compStart+7},D${compStart+7})<0.05,"Similar Estimated Cost",` +
    `IF(C${compStart+7}<D${compStart+7},"Lease Appears Lower Cost","Buy Appears Lower Cost")))`
  ]] });
  data.push({ range: `${S}!A${recRow+2}`, values: [[
    'This comparison reflects estimated total cost over the selected term only. Individual results vary based on actual fees, tax rates, excess mileage, wear-and-tear charges, and resale values. Do not assume one option is universally better.'
  ]] });

  await valuesBatchUpdate(id, data, 'lease-vs-buy-values');

  // ── Formatting ─────────────────────────────────────────────────────────────
  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(LVB,0,500,0,10), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Title
  reqs.push({ mergeCells: { range: gridRange(LVB,0,1,0,9), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(LVB,0,1,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:15 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:LVB, dimension:'ROWS', startIndex:0, endIndex:1 }, properties: { pixelSize:44 }, fields:'pixelSize' }});

  [1,2].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(LVB,ri,ri+1,0,9), mergeType:'MERGE_ALL' }});
    reqs.push({ repeatCell: { range: gridRange(LVB,ri,ri+1,0,9), cell: { userEnteredFormat: {
      backgroundColor: ri===1 ? hex(C.secondary) : hex(C.mutedBlue),
      textFormat: { italic:true, foregroundColor: hex(C.white), fontSize:9 },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});
  });

  // Section header (A5)
  reqs.push({ repeatCell: { range: gridRange(LVB,4,5,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Column headers for LEASE/BUY (C5:D5)
  reqs.push({ repeatCell: { range: gridRange(LVB,4,5,2,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold:true, foregroundColor: hex(C.white) }, horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  // Input cells (C6:C24 and D6:D16 that are editable) — pale yellow
  reqs.push({ repeatCell: { range: gridRange(LVB,5,24,2,3), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  reqs.push({ repeatCell: { range: gridRange(LVB,5,7,3,4), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Formula cells — pale blue
  reqs.push({ repeatCell: { range: gridRange(LVB,18,19,2,3), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});
  reqs.push({ repeatCell: { range: gridRange(LVB,7,16,3,4), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Comparison section
  reqs.push({ mergeCells: { range: gridRange(LVB,compStart-1,compStart,0,9), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(LVB,compStart-1,compStart,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  reqs.push({ repeatCell: { range: gridRange(LVB,compStart,compStart+1,2,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold:true, foregroundColor: hex(C.white) }, horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  // Metric rows
  reqs.push({ repeatCell: { range: gridRange(LVB,compStart+1,compStart+1+metrics.length,0,1), cell: { userEnteredFormat: { backgroundColor: hex(C.white), wrapStrategy:'WRAP' } }, fields:'userEnteredFormat(backgroundColor,wrapStrategy)' }});
  reqs.push({ repeatCell: { range: gridRange(LVB,compStart+1,compStart+1+metrics.length,2,4), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});

  const currFmt = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  reqs.push({ repeatCell: { range: gridRange(LVB,compStart+1,compStart+8,2,4), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(LVB,compStart+8,compStart+9,2,4), cell: { userEnteredFormat: { numberFormat: { type:'NUMBER', pattern:'$0.000' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Currency formats on lease inputs
  const currFmt2 = { type:'CURRENCY', pattern:'$#,##0.00' };
  [6,7,8,14,19,20,23,24].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(LVB,ri,ri+1,2,3), cell: { userEnteredFormat: { numberFormat: currFmt2 } }, fields:'userEnteredFormat.numberFormat' }});
  });
  // Percent inputs
  [11,12].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(LVB,ri,ri+1,2,3), cell: { userEnteredFormat: { numberFormat: { type:'PERCENT', pattern:'0.00%' } } }, fields:'userEnteredFormat.numberFormat' }});
  });

  // Recommendation row
  reqs.push({ mergeCells: { range: gridRange(LVB,recRow-1,recRow,0,9), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(LVB,recRow-1,recRow,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});
  reqs.push({ repeatCell: { range: gridRange(LVB,recRow,recRow+1,2,5), cell: { userEnteredFormat: {
    backgroundColor: hex(C.warning), textFormat: { bold:true, foregroundColor: hex(C.mainText) }, wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)' }});

  // Disclaimer note
  reqs.push({ mergeCells: { range: gridRange(LVB,recRow+1,recRow+3,0,9), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(LVB,recRow+1,recRow+3,0,9), cell: { userEnteredFormat: {
    backgroundColor: hex('#E8E6E1'), textFormat: { italic:true, foregroundColor: hex(C.secText), fontSize:9 },
    wrapStrategy:'WRAP', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)' }});

  // Column widths
  reqs.push({ updateDimensionProperties: { range: { sheetId:LVB, dimension:'COLUMNS', startIndex:0, endIndex:1 }, properties: { pixelSize:240 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:LVB, dimension:'COLUMNS', startIndex:2, endIndex:4 }, properties: { pixelSize:160 }, fields:'pixelSize' }});

  reqs.push({ updateSheetProperties: { properties: { sheetId:LVB, gridProperties: { frozenRowCount:6 } }, fields:'gridProperties.frozenRowCount' }});

  await batchUpdate(id, reqs, 'lease-vs-buy-format');
  console.log('✓ Lease vs Buy');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
