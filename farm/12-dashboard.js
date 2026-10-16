'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Dashboard'];
const S = "'Dashboard'";

(async () => {
  const data = [];
  const yr = '$B$2', per = '$B$3', ent = '$B$4';

  // Helpers
  const iAmt = () =>
    `IFERROR(IF(${ent}="All Enterprises",IF(${per}="Full Year",SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr}),SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${per})),IF(${per}="Full Year",SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$F$6:$F$505,${ent}),SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${per},'Income Log'!$F$6:$F$505,${ent}))),0)`;

  const eAmt = () =>
    `IFERROR(IF(${ent}="All Enterprises",IF(${per}="Full Year",SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr}),SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,${per})),IF(${per}="Full Year",SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$F$6:$F$505,${ent}),SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,${per},'Expense Log'!$F$6:$F$505,${ent}))),0)`;

  // Title & controls
  data.push({ range:`${S}!A1`, values:[['FARM FINANCIAL DASHBOARD']] });
  data.push({ range:`${S}!A2:B2`, values:[['Reporting Year:', 2026]] });
  data.push({ range:`${S}!A3:B3`, values:[['Period:', 'Full Year']] });
  data.push({ range:`${S}!A4:B4`, values:[['Enterprise:', 'All Enterprises']] });

  // ── Primary KPI cards row 6 (labels) & 7 (values) ─────────────────────────
  data.push({ range:`${S}!A6:L6`, values:[['TOTAL INCOME','','TOTAL EXPENSES','','NET PROFIT','','PROFIT MARGIN','','YTD TRANSACTIONS','','ANNUAL GOAL','']] });
  data.push({ range:`${S}!A7`, values:[[`=${iAmt()}`]] });
  data.push({ range:`${S}!C7`, values:[[`=${eAmt()}`]] });
  data.push({ range:`${S}!E7`, values:[[`=IFERROR(A7-C7,0)`]] });
  data.push({ range:`${S}!G7`, values:[[`=IFERROR(E7/A7,0)`]] });
  // Transaction count
  data.push({ range:`${S}!I7`, values:[[
    `=IFERROR(IF(${ent}="All Enterprises",IF(${per}="Full Year",COUNTIF('Income Log'!$C$6:$C$505,${yr}),COUNTIFS('Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${per})),IF(${per}="Full Year",COUNTIFS('Income Log'!$C$6:$C$505,${yr},'Income Log'!$F$6:$F$505,${ent}),COUNTIFS('Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${per},'Income Log'!$F$6:$F$505,${ent}))),0)+IFERROR(IF(${ent}="All Enterprises",IF(${per}="Full Year",COUNTIF('Expense Log'!$C$6:$C$505,${yr}),COUNTIFS('Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,${per})),IF(${per}="Full Year",COUNTIFS('Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$F$6:$F$505,${ent}),COUNTIFS('Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,${per},'Expense Log'!$F$6:$F$505,${ent}))),0)`,
  ]] });
  data.push({ range:`${S}!K7`, values:[[`=IFERROR(SUM('Profit Goals'!$B$9:$B$20),0)`]] });

  // ── Secondary KPI row 9 (labels) & 10 (values) ─────────────────────────────
  data.push({ range:`${S}!A9:L9`, values:[['ACTUAL VS GOAL','','BUDGET VARIANCE (INC)','','BUDGET VARIANCE (EXP)','','TOP ENTERPRISE','','INCOME ENTRIES','','EXPENSE ENTRIES','']] });
  data.push({ range:`${S}!A10`, values:[[`=IFERROR(E7/K7,0)`]] });  // pct of annual goal
  data.push({ range:`${S}!C10`, values:[[`=IFERROR('Budget vs Actual'!D26,0)`]] }); // income variance
  data.push({ range:`${S}!E10`, values:[[`=IFERROR('Budget vs Actual'!D60,0)`]] }); // expense variance
  data.push({ range:`${S}!G10`, values:[[`=IFERROR(INDEX('Enterprise Profitability'!$A$9:$A$23,MATCH(MAX('Enterprise Profitability'!$D$9:$D$23),'Enterprise Profitability'!$D$9:$D$23,0)),"")`]] });
  data.push({ range:`${S}!I10`, values:[[`=IFERROR(COUNTIF('Income Log'!$C$6:$C$505,${yr}),0)`]] });
  data.push({ range:`${S}!K10`, values:[[`=IFERROR(COUNTIF('Expense Log'!$C$6:$C$505,${yr}),0)`]] });

  // ── Top 5 Enterprises rows 13-18 ──────────────────────────────────────────
  data.push({ range:`${S}!A13:E13`, values:[['TOP 5 ENTERPRISES BY NET PROFIT','','','','']] });
  data.push({ range:`${S}!A14:E14`, values:[['Enterprise','Income','Expenses','Net Profit','Margin']] });
  // Pull from Enterprise Profitability tab sorted by net profit (D col) — use LARGE/INDEX/MATCH pattern
  // Rows 9-23 in Enterprise Profitability
  for (let i = 0; i < 5; i++) {
    const r = 15 + i;
    // Find the row with the i+1 largest net profit
    const rankFormula = `IFERROR(INDEX('Enterprise Profitability'!$A$9:$A$23,MATCH(LARGE('Enterprise Profitability'!$D$9:$D$23,${i+1}),'Enterprise Profitability'!$D$9:$D$23,0)),"")`;
    data.push({ range:`${S}!A${r}:E${r}`, values:[[
      `=${rankFormula}`,
      `=IFERROR(VLOOKUP(A${r},'Enterprise Profitability'!$A$9:$J$23,2,FALSE),0)`,
      `=IFERROR(VLOOKUP(A${r},'Enterprise Profitability'!$A$9:$J$23,3,FALSE),0)`,
      `=IFERROR(VLOOKUP(A${r},'Enterprise Profitability'!$A$9:$J$23,4,FALSE),0)`,
      `=IFERROR(VLOOKUP(A${r},'Enterprise Profitability'!$A$9:$J$23,5,FALSE),0)`,
    ]] });
  }

  // ── Recent Transactions rows 21-31 ────────────────────────────────────────
  data.push({ range:`${S}!A21:F21`, values:[['5 MOST RECENT INCOME ENTRIES','','','','','']] });
  data.push({ range:`${S}!A22:F22`, values:[['Date','Enterprise','Category','Description','Amount','Status']] });
  // LARGE on date col to get most recent — use INDEX/MATCH on rank of date
  for (let i = 0; i < 5; i++) {
    const r = 23 + i;
    const dateRef = `LARGE('Income Log'!$B$6:$B$505,${i+1})`;
    const rowRef  = `MATCH(${dateRef},'Income Log'!$B$6:$B$505,0)`;
    data.push({ range:`${S}!A${r}:F${r}`, values:[[
      `=IFERROR(LARGE('Income Log'!$B$6:$B$505,${i+1}),"")`,
      `=IFERROR(INDEX('Income Log'!$F$6:$F$505,${rowRef}),"")`,
      `=IFERROR(INDEX('Income Log'!$G$6:$G$505,${rowRef}),"")`,
      `=IFERROR(INDEX('Income Log'!$I$6:$I$505,${rowRef}),"")`,
      `=IFERROR(INDEX('Income Log'!$L$6:$L$505,${rowRef}),0)`,
      `=IFERROR(INDEX('Income Log'!$M$6:$M$505,${rowRef}),"")`,
    ]] });
  }

  data.push({ range:`${S}!H21:M21`, values:[['5 MOST RECENT EXPENSE ENTRIES','','','','','']] });
  data.push({ range:`${S}!H22:M22`, values:[['Date','Enterprise','Category','Vendor','Amount','Status']] });
  for (let i = 0; i < 5; i++) {
    const r = 23 + i;
    const dateRef = `LARGE('Expense Log'!$B$6:$B$505,${i+1})`;
    const rowRef  = `MATCH(${dateRef},'Expense Log'!$B$6:$B$505,0)`;
    data.push({ range:`${S}!H${r}:M${r}`, values:[[
      `=IFERROR(LARGE('Expense Log'!$B$6:$B$505,${i+1}),"")`,
      `=IFERROR(INDEX('Expense Log'!$F$6:$F$505,${rowRef}),"")`,
      `=IFERROR(INDEX('Expense Log'!$G$6:$G$505,${rowRef}),"")`,
      `=IFERROR(INDEX('Expense Log'!$H$6:$H$505,${rowRef}),"")`,
      `=IFERROR(INDEX('Expense Log'!$M$6:$M$505,${rowRef}),0)`,
      `=IFERROR(INDEX('Expense Log'!$N$6:$N$505,${rowRef}),"")`,
    ]] });
  }

  // Disclaimer at bottom
  data.push({ range:`${S}!A32:M32`, values:[['DISCLAIMER: This dashboard is for informational purposes only. Consult a qualified CPA or tax professional for tax and financial advice. All figures depend on data entered in the log tabs.']] });

  await valuesBatchUpdate(id, data, 'dashboard-values');

  // Formatting
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(DB,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const currency = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  const pct      = { type:'PERCENT',  pattern:'0.0%' };
  const num      = { type:'NUMBER',   pattern:'#,##0' };
  const medium   = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin     = { style:'SOLID',        color:hex(C.border) };
  const dateF    = { type:'DATE', pattern:'MMM D, YYYY' };

  // Title row
  reqs.push({ mergeCells:{ range:gridRange(DB,0,1,0,13), mergeType:'MERGE_ALL' } });
  fmt(0,1,0,13,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:22, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,5,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial' } } });
  fmt(1,5,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' } } });

  // Primary KPI cards (rows 6-8, 0-indexed 5-7)
  const kpiCols = [[0,2],[2,4],[4,6],[6,8],[8,10],[10,12]];
  kpiCols.forEach(([c1,c2]) => {
    reqs.push({ mergeCells:{ range:gridRange(DB,5,6,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(DB,6,8,c1,c2), mergeType:'MERGE_ALL' } });
  });
  fmt(5,6,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:8, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(6,8,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:18, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  // Number formats for primary cards
  for (const ci of [0,2,4,10]) {
    reqs.push({ repeatCell:{ range:gridRange(DB,6,8,ci,ci+2), cell:{ userEnteredFormat:{ numberFormat:currency } }, fields:'userEnteredFormat.numberFormat' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(DB,6,8,6,8), cell:{ userEnteredFormat:{ numberFormat:pct } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(DB,6,8,8,10), cell:{ userEnteredFormat:{ numberFormat:num } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(DB,5,8,0,12), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Secondary KPI cards (rows 9-11, 0-indexed 8-10)
  kpiCols.forEach(([c1,c2]) => {
    reqs.push({ mergeCells:{ range:gridRange(DB,8,9,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(DB,9,11,c1,c2), mergeType:'MERGE_ALL' } });
  });
  fmt(8,9,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:8, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(9,11,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:14, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  reqs.push({ repeatCell:{ range:gridRange(DB,9,11,0,2), cell:{ userEnteredFormat:{ numberFormat:pct } }, fields:'userEnteredFormat.numberFormat' } });
  for (const ci of [2,4]) {
    reqs.push({ repeatCell:{ range:gridRange(DB,9,11,ci,ci+2), cell:{ userEnteredFormat:{ numberFormat:currency } }, fields:'userEnteredFormat.numberFormat' } });
  }
  for (const ci of [8,10]) {
    reqs.push({ repeatCell:{ range:gridRange(DB,9,11,ci,ci+2), cell:{ userEnteredFormat:{ numberFormat:num } }, fields:'userEnteredFormat.numberFormat' } });
  }
  reqs.push({ updateBorders:{ range:gridRange(DB,8,11,0,12), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Top enterprises table
  reqs.push({ mergeCells:{ range:gridRange(DB,12,13,0,5), mergeType:'MERGE_ALL' } });
  fmt(12,13,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(13,14,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 14; r < 20; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,5,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  for (const ci of [1,2,3]) reqs.push({ repeatCell:{ range:gridRange(DB,14,20,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(DB,14,20,4,5), cell:{ userEnteredFormat:{ numberFormat:pct, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(DB,12,20,0,5), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Recent Income table
  reqs.push({ mergeCells:{ range:gridRange(DB,20,21,0,6), mergeType:'MERGE_ALL' } });
  fmt(20,21,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.success), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(21,22,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 22; r < 28; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,6,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(DB,22,28,0,1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(DB,22,28,4,5), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(DB,20,28,0,6), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Recent Expense table
  reqs.push({ mergeCells:{ range:gridRange(DB,20,21,7,13), mergeType:'MERGE_ALL' } });
  fmt(20,21,7,13,{ userEnteredFormat:{ backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(21,22,7,13,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 22; r < 28; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,7,13,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(DB,22,28,7,8), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(DB,22,28,11,12), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(DB,20,28,7,13), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Disclaimer
  reqs.push({ mergeCells:{ range:gridRange(DB,31,33,0,13), mergeType:'MERGE_ALL' } });
  fmt(31,33,0,13,{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:8, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });

  // Row heights
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:48 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:1, endIndex:5 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:5, endIndex:8 }, properties:{ pixelSize:44 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:8, endIndex:11 }, properties:{ pixelSize:36 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:11, endIndex:100 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Col widths — 13 columns A:M
  [[0,120],[1,130],[2,130],[3,130],[4,80],[5,80],[6,80],[7,100],[8,130],[9,130],[10,140],[11,90],[12,90]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });

  // Charts: place below row 32 (0-indexed)
  // Chart 1: Monthly income vs expenses column
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Monthly Income vs Expenses',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'COLUMN',
      legendPosition:'BOTTOM_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Month' },{ position:'LEFT_AXIS', title:'Amount ($)' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:sheetMap['Annual Summary'], startRowIndex:7, endRowIndex:21, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:sheetMap['Annual Summary'], startRowIndex:7, endRowIndex:21, startColumnIndex:1, endColumnIndex:2 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:sheetMap['Annual Summary'], startRowIndex:7, endRowIndex:21, startColumnIndex:2, endColumnIndex:3 }] } }, targetAxis:'LEFT_AXIS' },
      ]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:34, columnIndex:0 }, widthPixels:430, heightPixels:280 } } } } });

  // Chart 2: Net profit by month line
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Monthly Net Profit',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'LINE',
      legendPosition:'NO_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Month' },{ position:'LEFT_AXIS', title:'Profit ($)' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:sheetMap['Annual Summary'], startRowIndex:7, endRowIndex:21, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:sheetMap['Annual Summary'], startRowIndex:7, endRowIndex:21, startColumnIndex:3, endColumnIndex:4 }] } }, targetAxis:'LEFT_AXIS' }]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:34, columnIndex:5 }, widthPixels:340, heightPixels:280 } } } } });

  // Chart 3: Enterprise net profit bar
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Net Profit by Enterprise',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'BAR',
      legendPosition:'NO_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Net Profit ($)' },{ position:'LEFT_AXIS', title:'Enterprise' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:sheetMap['Enterprise Profitability'], startRowIndex:7, endRowIndex:24, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:sheetMap['Enterprise Profitability'], startRowIndex:7, endRowIndex:24, startColumnIndex:3, endColumnIndex:4 }] } }, targetAxis:'BOTTOM_AXIS' }]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:34, columnIndex:9 }, widthPixels:370, heightPixels:280 } } } } });

  // Chart 4: Income donut
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Income by Category',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    pieChart:{
      legendPosition:'RIGHT_LEGEND',
      pieHole:0.4,
      domain:{ sourceRange:{ sources:[{ sheetId:sheetMap['Budget vs Actual'], startRowIndex:7, endRowIndex:26, startColumnIndex:0, endColumnIndex:1 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:sheetMap['Budget vs Actual'], startRowIndex:7, endRowIndex:26, startColumnIndex:2, endColumnIndex:3 }] } }
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:52, columnIndex:0 }, widthPixels:430, heightPixels:300 } } } } });

  // Chart 5: Expense donut
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Expense by Category',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    pieChart:{
      legendPosition:'RIGHT_LEGEND',
      pieHole:0.4,
      domain:{ sourceRange:{ sources:[{ sheetId:sheetMap['Budget vs Actual'], startRowIndex:28, endRowIndex:60, startColumnIndex:0, endColumnIndex:1 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:sheetMap['Budget vs Actual'], startRowIndex:28, endRowIndex:60, startColumnIndex:2, endColumnIndex:3 }] } }
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:52, columnIndex:7 }, widthPixels:430, heightPixels:300 } } } } });

  // Freeze top 4 rows (title + 3 controls), no col freeze
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:DB, gridProperties:{ frozenRowCount:5, frozenColumnCount:0 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
