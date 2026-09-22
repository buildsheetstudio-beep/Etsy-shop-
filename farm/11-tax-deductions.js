'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, EXPENSE_CATS, TAX_GROUPS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TD = sheetMap['Tax Deduction Summary'];
const S = "'Tax Deduction Summary'";

// Unique tax groups in a stable display order
const TAX_GROUP_ORDER = [
  'Cost of Goods / Production Inputs',
  'Livestock & Animal Care',
  'Vehicle & Transportation',
  'Equipment & Repairs',
  'Land & Facilities',
  'Utilities',
  'Labor & Contract Services',
  'Insurance & Professional',
  'Selling & Administrative',
  'Other / Review Required',
];

(async () => {
  const data = [];
  const yr = '$B$2', qtr = '$B$3';

  // SUMIFS for deductible expenses filtered by year, quarter, and category
  const catAmt = (cat) =>
    `IFERROR(IF(${qtr}="Full Year",` +
      `SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$G$6:$G$505,"${cat}",'Expense Log'!$O$6:$O$505,"Yes"),` +
      `SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$E$6:$E$505,${qtr},'Expense Log'!$G$6:$G$505,"${cat}",'Expense Log'!$O$6:$O$505,"Yes")` +
    `),0)`;

  const groupAmt = (group) =>
    `IFERROR(IF(${qtr}="Full Year",` +
      `SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$P$6:$P$505,"${group}",'Expense Log'!$O$6:$O$505,"Yes"),` +
      `SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$E$6:$E$505,${qtr},'Expense Log'!$P$6:$P$505,"${group}",'Expense Log'!$O$6:$O$505,"Yes")` +
    `),0)`;

  // Header
  data.push({ range:`${S}!A1`, values:[['TAX DEDUCTION SUMMARY']] });
  data.push({ range:`${S}!A2:B2`, values:[['Reporting Year:', 2026]] });
  data.push({ range:`${S}!A3:B3`, values:[['Quarter:', 'Full Year']] });

  // Disclaimer (required)
  data.push({ range:`${S}!A5:G5`, values:[['IMPORTANT DISCLAIMER: This spreadsheet is for organizational and informational purposes only. It does not constitute tax, legal, or financial advice. Consult a qualified CPA or tax professional before filing. Deductibility depends on your specific situation and applicable tax law.']] });

  // ── Summary cards rows 8-10 ────────────────────────────────────────────────
  data.push({ range:`${S}!A8:F8`, values:[['TOTAL DEDUCTIBLE EXPENSES','','TOTAL ALL EXPENSES','','DEDUCTIBLE %','']] });
  data.push({ range:`${S}!A9`, values:[[`=IFERROR(IF(${qtr}="Full Year",SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$O$6:$O$505,"Yes"),SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$E$6:$E$505,${qtr},'Expense Log'!$O$6:$O$505,"Yes")),0)`]] });
  data.push({ range:`${S}!C9`, values:[[`=IFERROR(IF(${qtr}="Full Year",SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr}),SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$E$6:$E$505,${qtr})),0)`]] });
  data.push({ range:`${S}!E9`, values:[[`=IFERROR(A9/C9,0)`]] });

  // ── Tax Group Summary rows 13-23 ──────────────────────────────────────────
  data.push({ range:`${S}!A13:C13`, values:[['TAX GROUP','DEDUCTIBLE AMOUNT','% OF TOTAL']] });
  TAX_GROUP_ORDER.forEach((grp, i) => {
    const r = 14 + i;
    data.push({ range:`${S}!A${r}:C${r}`, values:[[
      grp,
      `=${groupAmt(grp)}`,
      `=IFERROR(B${r}/B24,0)`,
    ]] });
  });
  data.push({ range:`${S}!A24:C24`, values:[['TOTAL', `=SUM(B14:B23)`, `=1`]] });

  // ── Category Detail rows 27-57 ────────────────────────────────────────────
  data.push({ range:`${S}!A27:E27`, values:[['EXPENSE CATEGORY','TAX GROUP','DEDUCTIBLE AMOUNT','TRANSACTION COUNT','Deductible?']] });
  EXPENSE_CATS.forEach((cat, i) => {
    const r = 28 + i;
    const grp = TAX_GROUPS[cat] || 'Other / Review Required';
    const countFormula = `IFERROR(IF(${qtr}="Full Year",COUNTIFS('Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$G$6:$G$505,"${cat}",'Expense Log'!$O$6:$O$505,"Yes"),COUNTIFS('Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$E$6:$E$505,${qtr},'Expense Log'!$G$6:$G$505,"${cat}",'Expense Log'!$O$6:$O$505,"Yes")),0)`;
    data.push({ range:`${S}!A${r}:E${r}`, values:[[
      cat,
      grp,
      `=${catAmt(cat)}`,
      `=${countFormula}`,
      `=IFERROR(IF(C${r}>0,"Yes","—"),"—")`,
    ]] });
  });
  data.push({ range:`${S}!A58:E58`, values:[['TOTAL', '', `=SUM(C28:C57)`, `=SUM(D28:D57)`, '']] });

  await valuesBatchUpdate(id, data, 'tax-values');

  // Formatting
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(TD,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const currency = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  const pct      = { type:'PERCENT',  pattern:'0.0%' };
  const medium   = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin     = { style:'SOLID',        color:hex(C.border) };

  // Title
  reqs.push({ mergeCells:{ range:gridRange(TD,0,1,0,7), mergeType:'MERGE_ALL' } });
  fmt(0,1,0,7,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,4,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial' } } });
  fmt(1,4,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' } } });

  // Disclaimer row
  reqs.push({ mergeCells:{ range:gridRange(TD,4,6,0,7), mergeType:'MERGE_ALL' } });
  fmt(4,6,0,7,{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:8, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });

  // Summary cards
  const sCardCols = [[0,2],[2,4],[4,6]];
  sCardCols.forEach(([c1,c2]) => {
    reqs.push({ mergeCells:{ range:gridRange(TD,7,8,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(TD,8,10,c1,c2), mergeType:'MERGE_ALL' } });
  });
  fmt(7,8,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:8, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(8,10,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:16, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  reqs.push({ repeatCell:{ range:gridRange(TD,8,10,0,2), cell:{ userEnteredFormat:{ numberFormat:currency } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(TD,8,10,2,4), cell:{ userEnteredFormat:{ numberFormat:currency } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(TD,8,10,4,6), cell:{ userEnteredFormat:{ numberFormat:pct } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(TD,7,10,0,6), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Tax group table
  fmt(12,13,0,3,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 13; r < 24; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,3,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(23,24,0,3,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' } } });
  reqs.push({ repeatCell:{ range:gridRange(TD,13,24,1,2), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(TD,13,24,2,3), cell:{ userEnteredFormat:{ numberFormat:pct, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(TD,12,24,0,3), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Category detail table
  fmt(26,27,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 27; r < 58; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,5,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(57,58,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' } } });
  reqs.push({ repeatCell:{ range:gridRange(TD,27,58,2,3), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(TD,27,58,3,4), cell:{ userEnteredFormat:{ horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(TD,26,58,0,5), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Col widths
  [[0,220],[1,200],[2,130],[3,120],[4,100],[5,100],[6,100]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:TD, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:TD, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:TD, dimension:'ROWS', startIndex:4, endIndex:6 }, properties:{ pixelSize:36 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:TD, dimension:'ROWS', startIndex:1, endIndex:120 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Donut chart — deductible by tax group
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Deductible Expenses by Tax Group',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    pieChart:{
      legendPosition:'RIGHT_LEGEND',
      pieHole:0.4,
      domain:{ sourceRange:{ sources:[{ sheetId:TD, startRowIndex:12, endRowIndex:23, startColumnIndex:0, endColumnIndex:1 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:TD, startRowIndex:12, endRowIndex:23, startColumnIndex:1, endColumnIndex:2 }] } }
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:TD, rowIndex:60, columnIndex:0 }, widthPixels:460, heightPixels:320 } } } } });

  // Bar chart — deductible by category
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Deductible Expenses by Category',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'BAR',
      legendPosition:'NO_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Amount ($)' },{ position:'LEFT_AXIS', title:'Category' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:TD, startRowIndex:26, endRowIndex:58, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:TD, startRowIndex:26, endRowIndex:58, startColumnIndex:2, endColumnIndex:3 }] } }, targetAxis:'BOTTOM_AXIS' }]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:TD, rowIndex:60, columnIndex:4 }, widthPixels:420, heightPixels:320 } } } } });

  // freeze past the disclaimer/cards — rows 0-9 (10 rows) would split merged cards, freeze just title+controls
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:TD, gridProperties:{ frozenRowCount:4, frozenColumnCount:0 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'tax-format');
  console.log('Tax Deduction Summary complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
