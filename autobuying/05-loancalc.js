'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const LC = sheetMap['Auto Loan Calculator'];
const S = "'Auto Loan Calculator'";
const VQ = "'Vehicle Details & Quotes'";
const BS = "'Buyer Setup'";

(async () => {
  const data = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['AUTO LOAN CALCULATOR']] });
  data.push({ range: `${S}!A2`, values: [['Compare two loan scenarios side-by-side. Yellow = your inputs. Blue = calculated.']] });
  data.push({ range: `${S}!A3`, values: [['Amount Financed = OTD Price − Down Payment − Trade-In Equity + Trade-In Loan Balance (negative equity adds to loan)']] });

  // ── Section A: Controls ───────────────────────────────────────────────────
  data.push({ range: `${S}!A5`, values: [['LOAN CONTROLS']] });
  data.push({ range: `${S}!A6:B6`, values: [['Selected Vehicle ID', 'V-001']] });
  data.push({ range: `${S}!A7:B7`, values: [['Selected Quote ID', `=IFERROR(INDEX('Vehicle Details & Quotes'!$B$6:$B$105,MATCH(B6,'Vehicle Details & Quotes'!$A$6:$A$105,0)),"")`]] });
  data.push({ range: `${S}!A8`, values: [['── Pulled Quote Data (read-only) ──']] });
  data.push({ range: `${S}!A9:B9`, values: [['Make / Model', `=IFERROR(INDEX(${VQ}!$C$6:$C$105,MATCH(B6,${VQ}!$A$6:$A$105,0))&" "&INDEX(${VQ}!$D$6:$D$105,MATCH(B6,${VQ}!$A$6:$A$105,0)),"")`]] });
  data.push({ range: `${S}!A10:B10`, values: [['Out-the-Door Price', `=IFERROR(INDEX(${VQ}!$W$6:$W$105,MATCH(B6,${VQ}!$A$6:$A$105,0)),0)`]] });
  data.push({ range: `${S}!A11:B11`, values: [['Quoted APR', `=IFERROR(INDEX(${VQ}!$X$6:$X$105,MATCH(B6,${VQ}!$A$6:$A$105,0)),0)`]] });
  data.push({ range: `${S}!A12:B12`, values: [['Quoted Term (months)', `=IFERROR(INDEX(${VQ}!$Y$6:$Y$105,MATCH(B6,${VQ}!$A$6:$A$105,0)),60)`]] });
  data.push({ range: `${S}!A13:B13`, values: [['Trade-In Value', `=IFERROR('Buyer Setup'!B19,0)`]] });
  data.push({ range: `${S}!A14:B14`, values: [['Trade-In Loan Balance', `=IFERROR('Buyer Setup'!B20,0)`]] });
  data.push({ range: `${S}!A15:B15`, values: [['Buyer Down Payment', `=IFERROR('Buyer Setup'!B18,0)`]] });
  data.push({ range: `${S}!A16:B16`, values: [['Total Monthly Income', `=IFERROR('Buyer Setup'!B31,0)`]] });

  // ── Section B: Scenario Comparison ───────────────────────────────────────
  data.push({ range: `${S}!D5`, values: [['SCENARIO COMPARISON']] });
  data.push({ range: `${S}!D6:F6`, values: [['', 'Scenario A', 'Scenario B']] });

  const scenarioInputs = [
    ['APR',              '=IFERROR(B11,0)',          '=IFERROR(B11*0.9,0)'],
    ['Term (months)',    '=IFERROR(B12,60)',          60],
    ['Down Payment',     '=IFERROR(B15,0)',           '=IFERROR(B15,0)'],
    ['Extra Monthly Pmt',0,                           0],
    ['First Payment Date','Sep 1, 2026',              'Sep 1, 2026'],
  ];

  scenarioInputs.forEach(([label, a, b], i) => {
    data.push({ range: `${S}!D${i+7}:F${i+7}`, values: [[label, a, b]] });
  });

  // Calculated outputs
  const calcRow = 13;
  data.push({ range: `${S}!D${calcRow}`, values: [['── Calculated Results ──']] });

  // Amount Financed = MAX(0, OTD - Down - TradeInValue + TradeInLoanBalance)
  // But we keep negative equity additive: OTD - DownPmt - TradeInValue + TradeInLoanBal
  // (if net trade-in is positive, it reduces financed; if negative, increases it)
  const amtFin_A = `=IFERROR(MAX(0,B10-E9-B13+B14),0)`;
  const amtFin_B = `=IFERROR(MAX(0,B10-F9-B13+B14),0)`;
  const pmt_A    = `=IFERROR(IF(E7=0,0,-PMT(E7/12,E8,${amtFin_A.replace('=','')})),0)`;
  const pmt_B    = `=IFERROR(IF(F7=0,0,-PMT(F7/12,F8,${amtFin_B.replace('=','')})),0)`;

  const calcRows = [
    ['Amount Financed',       amtFin_A,  amtFin_B],
    ['Monthly Payment',       pmt_A,     pmt_B],
    ['Total Interest',        `=IFERROR(E${calcRow+2}*E8-E${calcRow+1},0)`, `=IFERROR(F${calcRow+2}*F8-F${calcRow+1},0)`],
    ['Total Paid',            `=IFERROR(E${calcRow+2}*E8,0)`,               `=IFERROR(F${calcRow+2}*F8,0)`],
    ['Payoff Date',           `=IFERROR(EDATE(DATEVALUE(E11),E8-1),"")`,    `=IFERROR(EDATE(DATEVALUE(F11),F8-1),"")`],
    ['Payment-to-Income %',   `=IFERROR(E${calcRow+2}/B16,0)`,              `=IFERROR(F${calcRow+2}/B16,0)`],
    ['Interest Saved (Extra)', `=IFERROR(E${calcRow+3}-E${calcRow+3},0)`,   `=IFERROR(F${calcRow+3}-F${calcRow+3},0)`],
  ];

  calcRows.forEach(([label, a, b], i) => {
    data.push({ range: `${S}!D${calcRow+1+i}:F${calcRow+1+i}`, values: [[label, a, b]] });
  });

  // ── Amortization Schedule ─────────────────────────────────────────────────
  const amoStart = 25;
  data.push({ range: `${S}!A${amoStart}`, values: [['AMORTIZATION SCHEDULE — SCENARIO A']] });
  data.push({ range: `${S}!A${amoStart+1}:H${amoStart+1}`, values: [[
    'Payment #','Payment Date','Beginning Balance','Scheduled Payment','Extra Payment','Principal','Interest','Ending Balance'
  ]] });

  for (let p = 1; p <= 84; p++) {
    const r = amoStart + 1 + p;
    const prevEnd = p === 1 ? `IFERROR(MAX(0,$B$10-$E$9-$B$13+$B$14),0)` : `H${r-1}`;
    data.push({ range: `${S}!A${r}:H${r}`, values: [[
      `=IF(${p===1 ? `IFERROR(MAX(0,$B$10-$E$9-$B$13+$B$14),0)` : `H${r-1}`}>0,${p},"")`,
      `=IF(A${r}="","",IFERROR(EDATE(DATEVALUE($E$11),A${r}-1),""))`,
      `=IF(A${r}="","",${p===1 ? `IFERROR(MAX(0,$B$10-$E$9-$B$13+$B$14),0)` : `H${r-1}`})`,
      `=IF(A${r}="","",IFERROR($E${calcRow+2},0))`,
      `=IF(A${r}="","",IFERROR($E$10,0))`,
      `=IF(A${r}="","",IFERROR(-PPMT($E$7/12,A${r},$E$8,-MAX(0,$B$10-$E$9-$B$13+$B$14)),0))`,
      `=IF(A${r}="","",IFERROR(-IPMT($E$7/12,A${r},$E$8,-MAX(0,$B$10-$E$9-$B$13+$B$14)),0))`,
      `=IF(A${r}="","",MAX(0,C${r}-F${r}-E${r}))`
    ]] });
  }

  await valuesBatchUpdate(id, data, 'loan-calc-values');

  // ── Formatting ─────────────────────────────────────────────────────────────
  const reqs = [];

  reqs.push({ repeatCell: { range: gridRange(LC,0,500,0,12), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Title
  reqs.push({ mergeCells: { range: gridRange(LC,0,1,0,12), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(LC,0,1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:15 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:LC, dimension:'ROWS', startIndex:0, endIndex:1 }, properties: { pixelSize:44 }, fields:'pixelSize' }});

  // Rows 2-3 (index 1-2)
  [1,2].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(LC,ri,ri+1,0,12), mergeType:'MERGE_ALL' }});
    reqs.push({ repeatCell: { range: gridRange(LC,ri,ri+1,0,12), cell: { userEnteredFormat: {
      backgroundColor: ri===1 ? hex(C.secondary) : hex(C.mutedBlue),
      textFormat: { italic:true, foregroundColor: hex(C.white), fontSize:9 },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)' }});
  });

  // Section headers (A5, D5)
  reqs.push({ repeatCell: { range: gridRange(LC,4,5,0,2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});
  reqs.push({ mergeCells: { range: gridRange(LC,4,5,3,7), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(LC,4,5,3,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Scenario column headers (D6:F6)
  reqs.push({ repeatCell: { range: gridRange(LC,5,6,4,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold:true, foregroundColor: hex(C.white) },
    horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  // Input cells — pale yellow (B6, E7-F11)
  reqs.push({ repeatCell: { range: gridRange(LC,5,6,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  reqs.push({ repeatCell: { range: gridRange(LC,6,11,4,6), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Formula cells (B7-B16, E13-F20) — pale blue
  reqs.push({ repeatCell: { range: gridRange(LC,6,16,1,2), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});
  reqs.push({ repeatCell: { range: gridRange(LC,12,21,4,6), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Currency formats
  const currFmt = { type:'CURRENCY', pattern:'$#,##0.00' };
  // B10 (OTD), B13(trade-in), B14, B15 (index 9,12,13,14)
  [9,12,13,14].forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(LC,ri,ri+1,1,2), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});
  });
  // Scenario calc rows: amount financed, payment, interest, total paid
  [calcRow,calcRow+1,calcRow+2,calcRow+3].map(r=>r-1).forEach(ri => {
    reqs.push({ repeatCell: { range: gridRange(LC,ri,ri+1,4,6), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});
  });

  // APR format (E7:F7, B11 index 6, 10)
  reqs.push({ repeatCell: { range: gridRange(LC,6,7,4,6), cell: { userEnteredFormat: { numberFormat: { type:'PERCENT', pattern:'0.00%' } } }, fields:'userEnteredFormat.numberFormat' }});
  reqs.push({ repeatCell: { range: gridRange(LC,10,11,1,2), cell: { userEnteredFormat: { numberFormat: { type:'PERCENT', pattern:'0.00%' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Payment-to-Income %
  reqs.push({ repeatCell: { range: gridRange(LC,calcRow+5,calcRow+6,4,6), cell: { userEnteredFormat: { numberFormat: { type:'PERCENT', pattern:'0.0%' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Amortization header
  reqs.push({ mergeCells: { range: gridRange(LC,amoStart-1,amoStart,0,8), mergeType:'MERGE_ALL' }});
  reqs.push({ repeatCell: { range: gridRange(LC,amoStart-1,amoStart,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold:true, foregroundColor: hex(C.white), fontSize:11 }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  reqs.push({ repeatCell: { range: gridRange(LC,amoStart,amoStart+1,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.mutedBlue), textFormat: { bold:true, foregroundColor: hex(C.mainText) },
    horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' }});

  // Currency in amortization: C, D, E, F, G, H cols (2-7)
  reqs.push({ repeatCell: { range: gridRange(LC,amoStart+1,amoStart+86,2,8), cell: { userEnteredFormat: { numberFormat: currFmt } }, fields:'userEnteredFormat.numberFormat' }});

  // Alternate amortization rows
  for (let i = 0; i < 84; i++) {
    const ri = amoStart + 1 + i;
    const bg = i % 2 === 0 ? C.white : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(LC,ri,ri+1,0,8), cell: { userEnteredFormat: { backgroundColor: hex(bg) } }, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Column widths
  reqs.push({ updateDimensionProperties: { range: { sheetId:LC, dimension:'COLUMNS', startIndex:0, endIndex:1 }, properties: { pixelSize:220 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:LC, dimension:'COLUMNS', startIndex:1, endIndex:2 }, properties: { pixelSize:140 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:LC, dimension:'COLUMNS', startIndex:3, endIndex:4 }, properties: { pixelSize:220 }, fields:'pixelSize' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId:LC, dimension:'COLUMNS', startIndex:4, endIndex:6 }, properties: { pixelSize:140 }, fields:'pixelSize' }});

  // Freeze rows 1:6
  reqs.push({ updateSheetProperties: { properties: { sheetId:LC, gridProperties: { frozenRowCount:6 } }, fields:'gridProperties.frozenRowCount' }});

  await batchUpdate(id, reqs, 'loan-calc-format');
  console.log('✓ Auto Loan Calculator');

  // Store calcRow for reference by other scripts
  const amoStart_val = amoStart;
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });

const amoStart = 25;
