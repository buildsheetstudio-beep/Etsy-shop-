'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, INCOME_CATS, EXPENSE_CATS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BA = sheetMap['Budget vs Actual'];
const S = "'Budget vs Actual'";

(async () => {
  const data = [];
  const yr = '$B$2', per = '$B$3';

  const iActual = (cat) => `IFERROR(IF(${per}="Full Year",SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$G$6:$G$505,"${cat}"),SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${per},'Income Log'!$G$6:$G$505,"${cat}")),0)`;
  const eActual = (cat) => `IFERROR(IF(${per}="Full Year",SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$G$6:$G$505,"${cat}"),SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,${per},'Expense Log'!$G$6:$G$505,"${cat}")),0)`;

  data.push({ range:`${S}!A1`, values:[['BUDGET VS ACTUAL']] });
  data.push({ range:`${S}!A2:B2`, values:[['Reporting Year:', 2026]] });
  data.push({ range:`${S}!A3:B3`, values:[['Period:', 'Full Year']] });

  // ── Income Budget Table rows 8-27 ──────────────────────────────────────────
  data.push({ range:`${S}!A8:F8`, values:[['INCOME BY CATEGORY','Annual Budget','Actual Income','Variance','Variance %','Status']] });
  INCOME_CATS.forEach((cat, i) => {
    const r = 9 + i;
    // Budget from Farm Setup col C, rows 34-50 (B34 = row index 33 in 0-based = row 34 in 1-based)
    const budgetRef = `IFERROR(VLOOKUP("${cat}",'Farm Setup'!$B$34:$C$50,2,FALSE),0)`;
    data.push({ range:`${S}!A${r}:F${r}`, values:[[
      cat,
      `=${budgetRef}`,
      `=${iActual(cat)}`,
      `=IFERROR(C${r}-B${r},0)`,
      `=IFERROR(D${r}/B${r},0)`,
      `=IFERROR(IF(C${r}>=B${r}*1.05,"Above Plan",IF(C${r}>=B${r}*0.9,"On Track","Below Plan")),"")`,
    ]] });
  });
  data.push({ range:`${S}!A26:F26`, values:[['TOTAL',`=SUM(B9:B25)`,`=SUM(C9:C25)`,`=IFERROR(C26-B26,0)`,`=IFERROR(D26/B26,0)`,'']] });

  // ── Expense Budget Table rows 29-60 ────────────────────────────────────────
  data.push({ range:`${S}!A29:F29`, values:[['EXPENSE BY CATEGORY','Annual Budget','Actual Expense','Variance','Variance %','Status']] });
  EXPENSE_CATS.forEach((cat, i) => {
    const r = 30 + i;
    const budgetRef = `IFERROR(VLOOKUP("${cat}",'Farm Setup'!$B$52:$C$81,2,FALSE),0)`;
    data.push({ range:`${S}!A${r}:F${r}`, values:[[
      cat,
      `=${budgetRef}`,
      `=${eActual(cat)}`,
      `=IFERROR(B${r}-C${r},0)`,
      `=IFERROR(D${r}/B${r},0)`,
      `=IFERROR(IF(C${r}<=B${r}*0.9,"Under Budget",IF(C${r}<=B${r}*1.05,"Near Budget","Over Budget")),"")`,
    ]] });
  });
  data.push({ range:`${S}!A60:F60`, values:[['TOTAL',`=SUM(B30:B59)`,`=SUM(C30:C59)`,`=IFERROR(B60-C60,0)`,`=IFERROR(D60/B60,0)`,'']] });

  // ── Summary cards rows 62-64 ───────────────────────────────────────────────
  data.push({ range:`${S}!A62:H62`, values:[['BUDGETED INCOME','','ACTUAL INCOME','','INCOME VARIANCE','','BUDGETED EXPENSES','ACTUAL EXPENSES']] });
  data.push({ range:`${S}!A63`, values:[[`=B26`]] });
  data.push({ range:`${S}!C63`, values:[[`=C26`]] });
  data.push({ range:`${S}!E63`, values:[[`=D26`]] });
  data.push({ range:`${S}!G63`, values:[[`=B60`]] });
  data.push({ range:`${S}!H63`, values:[[`=C60`]] });

  await valuesBatchUpdate(id, data, 'budget-values');

  // Formatting
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(BA,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const currency = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  const pct      = { type:'PERCENT',  pattern:'0.0%' };
  const medium   = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin     = { style:'SOLID',        color:hex(C.border) };

  reqs.push({ mergeCells:{ range:gridRange(BA,0,1,0,6), mergeType:'MERGE_ALL' } });
  fmt(0,1,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,4,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial' } } });
  fmt(1,4,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' } } });

  // Table headers
  for (const r of [7,28]) {
    fmt(r,r+1,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  }

  // Income table rows
  for (let r = 8; r < 26; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,6,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  for (const ci of [1,2,3]) reqs.push({ repeatCell:{ range:gridRange(BA,8,26,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(BA,8,26,4,5), cell:{ userEnteredFormat:{ numberFormat:pct, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  fmt(25,26,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' } } });

  // Expense table rows
  for (let r = 29; r < 60; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,6,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  for (const ci of [1,2,3]) reqs.push({ repeatCell:{ range:gridRange(BA,29,60,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(BA,29,60,4,5), cell:{ userEnteredFormat:{ numberFormat:pct, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  fmt(59,60,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' } } });

  // Summary cards
  const cardCols = [[0,1],[2,3],[4,5],[6,7],[7,8]];
  cardCols.forEach(([c1,c2]) => { reqs.push({ mergeCells:{ range:gridRange(BA,61,62,c1,c2), mergeType:'MERGE_ALL' } }); reqs.push({ mergeCells:{ range:gridRange(BA,62,64,c1,c2), mergeType:'MERGE_ALL' } }); });
  fmt(61,62,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:8, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(62,64,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:16, fontFamily:'Arial' }, numberFormat:currency, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Borders
  [[7,26,0,6],[28,60,0,6],[61,64,0,8]].forEach(([r1,r2,c1,c2]) => {
    reqs.push({ updateBorders:{ range:gridRange(BA,r1,r2,c1,c2), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });
  });

  [[0,160],[1,110],[2,110],[3,100],[4,90],[5,110]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:BA, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:BA, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:BA, dimension:'ROWS', startIndex:1, endIndex:200 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Charts
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Budget vs Actual — Income',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'BAR',
      legendPosition:'BOTTOM_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Amount ($)' },{ position:'LEFT_AXIS', title:'Category' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:BA, startRowIndex:7, endRowIndex:26, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:BA, startRowIndex:7, endRowIndex:26, startColumnIndex:1, endColumnIndex:2 }] } }, targetAxis:'BOTTOM_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:BA, startRowIndex:7, endRowIndex:26, startColumnIndex:2, endColumnIndex:3 }] } }, targetAxis:'BOTTOM_AXIS' },
      ]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:BA, rowIndex:65, columnIndex:0 }, widthPixels:480, heightPixels:300 } } } } });

  reqs.push({ addChart:{ chart:{ spec:{
    title:'Budget vs Actual — Expenses',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'BAR',
      legendPosition:'BOTTOM_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Amount ($)' },{ position:'LEFT_AXIS', title:'Category' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:BA, startRowIndex:28, endRowIndex:60, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:BA, startRowIndex:28, endRowIndex:60, startColumnIndex:1, endColumnIndex:2 }] } }, targetAxis:'BOTTOM_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:BA, startRowIndex:28, endRowIndex:60, startColumnIndex:2, endColumnIndex:3 }] } }, targetAxis:'BOTTOM_AXIS' },
      ]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:BA, rowIndex:65, columnIndex:6 }, widthPixels:480, heightPixels:300 } } } } });

  reqs.push({ updateSheetProperties:{ properties:{ sheetId:BA, gridProperties:{ frozenRowCount:6, frozenColumnCount:0 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'budget-format');
  console.log('Budget vs Actual complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
