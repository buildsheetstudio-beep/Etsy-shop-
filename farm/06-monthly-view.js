'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS, INCOME_CATS, EXPENSE_CATS, ENTERPRISES } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const MV = sheetMap['Monthly View'];
const S = "'Monthly View'";

// Controls: B2=Year, B3=Month, B4=Enterprise
// Helper: income SUMIFS with enterprise filter
const iAmt = (yrRef, moRef, entRef, extraCrit='') => {
  const base = `'Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yrRef},'Income Log'!$D$6:$D$505,${moRef}${extraCrit}`;
  return `IF(${entRef}="All Enterprises",SUMIFS(${base}),SUMIFS(${base},'Income Log'!$F$6:$F$505,${entRef}))`;
};
const eAmt = (yrRef, moRef, entRef, extraCrit='') => {
  const base = `'Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yrRef},'Expense Log'!$D$6:$D$505,${moRef}${extraCrit}`;
  return `IF(${entRef}="All Enterprises",SUMIFS(${base}),SUMIFS(${base},'Expense Log'!$F$6:$F$505,${entRef}))`;
};
const iCnt = (yrRef, moRef, entRef) => {
  const base = `'Income Log'!$C$6:$C$505,${yrRef},'Income Log'!$D$6:$D$505,${moRef}`;
  return `IF(${entRef}="All Enterprises",COUNTIFS(${base}),COUNTIFS(${base},'Income Log'!$F$6:$F$505,${entRef}))`;
};
const eCnt = (yrRef, moRef, entRef) => {
  const base = `'Expense Log'!$C$6:$C$505,${yrRef},'Expense Log'!$D$6:$D$505,${moRef}`;
  return `IF(${entRef}="All Enterprises",COUNTIFS(${base}),COUNTIFS(${base},'Expense Log'!$F$6:$F$505,${entRef}))`;
};

(async () => {
  const data = [];
  const yr = '$B$2', mo = '$B$3', ent = '$B$4';

  // ── Title & controls ───────────────────────────────────────────────────────
  data.push({ range:`${S}!A1`, values:[['MONTHLY VIEW']] });
  data.push({ range:`${S}!A2:B2`, values:[['Reporting Year:', 2026]] });
  data.push({ range:`${S}!A3:B3`, values:[['Month:', 'August']] });
  data.push({ range:`${S}!A4:B4`, values:[['Enterprise:', 'All Enterprises']] });

  // ── Summary cards row 7 labels / row 8 values ─────────────────────────────
  data.push({ range:`${S}!B7:M7`, values:[['TOTAL INCOME','','TOTAL EXPENSES','','NET PROFIT','','PROFIT MARGIN','','INCOME TXNS','','EXPENSE TXNS','']] });
  data.push({ range:`${S}!B8`, values:[[`=IFERROR(${iAmt(yr,mo,ent)},0)`]] });
  data.push({ range:`${S}!D8`, values:[[`=IFERROR(${eAmt(yr,mo,ent)},0)`]] });
  data.push({ range:`${S}!F8`, values:[[`=IFERROR(B8-D8,0)`]] });
  data.push({ range:`${S}!H8`, values:[[`=IFERROR(F8/B8,0)`]] });
  data.push({ range:`${S}!J8`, values:[[`=IFERROR(${iCnt(yr,mo,ent)},0)`]] });
  data.push({ range:`${S}!L8`, values:[[`=IFERROR(${eCnt(yr,mo,ent)},0)`]] });

  // ── Income by Category table rows 11-29 ───────────────────────────────────
  data.push({ range:`${S}!A11:C11`, values:[['INCOME BY CATEGORY','','Amount']] });
  INCOME_CATS.forEach((cat, i) => {
    const r = 12 + i;
    const crit = `,'Income Log'!$G$6:$G$505,"${cat}"`;
    data.push({ range:`${S}!A${r}`, values:[[cat]] });
    data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(${iAmt(yr,mo,ent,crit)},0)`]] });
  });
  data.push({ range:`${S}!A29:C29`, values:[['TOTAL','',`=SUM(C12:C28)`]] });

  // ── Expense by Category table rows 32-63 ──────────────────────────────────
  data.push({ range:`${S}!A32:C32`, values:[['EXPENSE BY CATEGORY','','Amount']] });
  EXPENSE_CATS.forEach((cat, i) => {
    const r = 33 + i;
    const crit = `,'Expense Log'!$G$6:$G$505,"${cat}"`;
    data.push({ range:`${S}!A${r}`, values:[[cat]] });
    data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(${eAmt(yr,mo,ent,crit)},0)`]] });
  });
  data.push({ range:`${S}!A63:C63`, values:[['TOTAL','',`=SUM(C33:C62)`]] });

  // ── Enterprise summary table rows 66-84 ───────────────────────────────────
  data.push({ range:`${S}!A66:F66`, values:[['ENTERPRISE SUMMARY','Income','Expenses','Net Profit','Margin','Count']] });
  ENTERPRISES.forEach((ent_name, i) => {
    const r = 67 + i;
    const iE = `IFERROR(IF("${ent_name}"="All Enterprises",SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${mo}),SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${mo},'Income Log'!$F$6:$F$505,"${ent_name}")),0)`;
    const eE = `IFERROR(SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,${mo},'Expense Log'!$F$6:$F$505,"${ent_name}"),0)`;
    data.push({ range:`${S}!A${r}:F${r}`, values:[[
      ent_name,
      `=IFERROR(SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${mo},'Income Log'!$F$6:$F$505,"${ent_name}"),0)`,
      `=IFERROR(SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,${mo},'Expense Log'!$F$6:$F$505,"${ent_name}"),0)`,
      `=IFERROR(B${r}-C${r},0)`,
      `=IFERROR(D${r}/B${r},0)`,
      `=IFERROR(COUNTIFS('Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${mo},'Income Log'!$F$6:$F$505,"${ent_name}"),0)`,
    ]] });
  });
  data.push({ range:`${S}!A82:F82`, values:[['TOTAL',`=SUM(B67:B81)`,`=SUM(C67:C81)`,`=IFERROR(B82-C82,0)`,`=IFERROR(D82/B82,0)`,`=SUM(F67:F81)`]] });

  await valuesBatchUpdate(id, data, 'monthly-values');

  // ── Formatting ─────────────────────────────────────────────────────────────
  const reqs = [];
  const merge = (r1,r2,c1,c2) => reqs.push({ mergeCells:{ range:gridRange(MV,r1,r2,c1,c2), mergeType:'MERGE_ALL' } });
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(MV,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const currency = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  const pct      = { type:'PERCENT',  pattern:'0.0%' };
  const medium   = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin     = { style:'SOLID',        color:hex(C.border) };

  // Title
  merge(0,1,0,13); merge(0,1,0,13);
  fmt(0,1,0,13,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Controls
  fmt(1,5,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,5,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Summary cards area row 7 (labels) + row 8 (values)
  const cardCols = [[1,3],[3,5],[5,7],[7,9],[9,11],[11,13]];
  cardCols.forEach(([c1,c2]) => { merge(6,7,c1,c2); merge(7,9,c1,c2); });
  fmt(6,7,1,13,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:8, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(7,9,1,13,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:18, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  // Currency for income, expenses, profit cards
  for (const c of [1,3,5]) {
    reqs.push({ repeatCell:{ range:gridRange(MV,7,9,c,c+2), cell:{ userEnteredFormat:{ numberFormat:currency } }, fields:'userEnteredFormat.numberFormat' } });
  }
  // Percent for margin card
  reqs.push({ repeatCell:{ range:gridRange(MV,7,9,7,9), cell:{ userEnteredFormat:{ numberFormat:pct } }, fields:'userEnteredFormat.numberFormat' } });

  // Table headers
  for (const r of [10,31,65]) {
    fmt(r,r+1,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  }

  // Income by category table
  for (let r = 11; r < 29; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,3,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(11,29,2,3,{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT', textFormat:{ fontSize:9, fontFamily:'Arial' } } });
  fmt(28,29,0,3,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'LEFT' } });
  reqs.push({ repeatCell:{ range:gridRange(MV,28,29,2,3), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT', textFormat:{ foregroundColor:hex(C.white), bold:true } } }, fields:'userEnteredFormat' } });

  // Expense by category table
  for (let r = 32; r < 62; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,3,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(32,63,2,3,{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT', textFormat:{ fontSize:9, fontFamily:'Arial' } } });
  fmt(62,63,0,3,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' } } });
  reqs.push({ repeatCell:{ range:gridRange(MV,62,63,2,3), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT', textFormat:{ foregroundColor:hex(C.white), bold:true } } }, fields:'userEnteredFormat' } });

  // Enterprise table
  for (let r = 66; r < 82; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,6,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  for (const ci of [1,2,3]) fmt(66,82,ci,ci+1,{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } });
  fmt(66,82,4,5,{ userEnteredFormat:{ numberFormat:pct, horizontalAlignment:'RIGHT' } });
  fmt(81,82,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' } } });
  for (const ci of [1,2,3]) reqs.push({ repeatCell:{ range:gridRange(MV,81,82,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currency, textFormat:{ foregroundColor:hex(C.white), bold:true } } }, fields:'userEnteredFormat' } });

  // Borders
  [[5,9,1,13],[10,29,0,3],[31,63,0,3],[65,82,0,6]].forEach(([r1,r2,c1,c2]) => {
    reqs.push({ updateBorders:{ range:gridRange(MV,r1,r2,c1,c2), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });
  });

  // Col widths
  [[0,160],[1,110],[2,20],[3,110],[4,20],[5,110],[6,20],[7,90],[8,20],[9,70],[10,20],[11,70],[12,20]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:MV, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:MV, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:MV, dimension:'ROWS', startIndex:1, endIndex:6 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:MV, dimension:'ROWS', startIndex:6, endIndex:7 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:MV, dimension:'ROWS', startIndex:7, endIndex:9 }, properties:{ pixelSize:38 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:MV, dimension:'ROWS', startIndex:9, endIndex:85 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Charts (donut income, donut expense, column enterprise)
  // Chart 1: Income by Category (C12:C28 values, A12:A28 labels)
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Income by Category',
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    backgroundColor:hex(C.panel),
    pieChart:{
      legendPosition:'RIGHT_LEGEND',
      pieHole:0.4,
      domain:{ sourceRange:{ sources:[{ sheetId:MV, startRowIndex:11, endRowIndex:28, startColumnIndex:0, endColumnIndex:1 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:MV, startRowIndex:11, endRowIndex:28, startColumnIndex:2, endColumnIndex:3 }] } }
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:MV, rowIndex:10, columnIndex:4 }, widthPixels:380, heightPixels:280 } } } } });

  // Chart 2: Expense by Category
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Expenses by Category',
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    backgroundColor:hex(C.panel),
    pieChart:{
      legendPosition:'RIGHT_LEGEND',
      pieHole:0.4,
      domain:{ sourceRange:{ sources:[{ sheetId:MV, startRowIndex:32, endRowIndex:62, startColumnIndex:0, endColumnIndex:1 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:MV, startRowIndex:32, endRowIndex:62, startColumnIndex:2, endColumnIndex:3 }] } }
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:MV, rowIndex:31, columnIndex:4 }, widthPixels:380, heightPixels:300 } } } } });

  // Chart 3: Income vs Expenses by Enterprise (bar)
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Income vs Expenses by Enterprise',
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    backgroundColor:hex(C.panel),
    basicChart:{
      chartType:'BAR',
      legendPosition:'BOTTOM_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Amount ($)' },{ position:'LEFT_AXIS', title:'Enterprise' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:MV, startRowIndex:65, endRowIndex:82, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:MV, startRowIndex:65, endRowIndex:82, startColumnIndex:1, endColumnIndex:2 }] } }, targetAxis:'BOTTOM_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:MV, startRowIndex:65, endRowIndex:82, startColumnIndex:2, endColumnIndex:3 }] } }, targetAxis:'BOTTOM_AXIS' },
      ]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:MV, rowIndex:65, columnIndex:7 }, widthPixels:420, heightPixels:320 } } } } });

  // Freeze rows 1-6
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:MV, gridProperties:{ frozenRowCount:6, frozenColumnCount:0 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'monthly-format');
  console.log('Monthly View complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
