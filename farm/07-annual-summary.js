'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const AS = sheetMap['Annual Summary'];
const S = "'Annual Summary'";

(async () => {
  const data = [];
  const yr = '$B$2', ent = '$B$3';
  const iAmt = (mo) => `IF(${ent}="All Enterprises",SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,"${mo}"),SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,"${mo}",'Income Log'!$F$6:$F$505,${ent}))`;
  const eAmt = (mo) => `IF(${ent}="All Enterprises",SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,"${mo}"),SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,"${mo}",'Expense Log'!$F$6:$F$505,${ent}))`;
  const iCnt = (mo) => `IF(${ent}="All Enterprises",COUNTIFS('Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,"${mo}"),COUNTIFS('Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,"${mo}",'Income Log'!$F$6:$F$505,${ent}))`;
  const eCnt = (mo) => `IF(${ent}="All Enterprises",COUNTIFS('Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,"${mo}"),COUNTIFS('Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,"${mo}",'Expense Log'!$F$6:$F$505,${ent}))`;

  // Title & controls
  data.push({ range:`${S}!A1`, values:[['ANNUAL SUMMARY']] });
  data.push({ range:`${S}!A2:B2`, values:[['Reporting Year:', 2026]] });
  data.push({ range:`${S}!A3:B3`, values:[['Enterprise:', 'All Enterprises']] });

  // Monthly table header row 8
  data.push({ range:`${S}!A8:H8`, values:[['Month','Income','Expenses','Net Profit','Profit Margin','Inc Txns','Exp Txns','Cumulative Profit']] });

  // Monthly rows 9-20
  MONTHS.forEach((mo, i) => {
    const r = 9 + i;
    data.push({ range:`${S}!A${r}:H${r}`, values:[[
      mo,
      `=IFERROR(${iAmt(mo)},0)`,
      `=IFERROR(${eAmt(mo)},0)`,
      `=IFERROR(B${r}-C${r},0)`,
      `=IFERROR(D${r}/B${r},0)`,
      `=IFERROR(${iCnt(mo)},0)`,
      `=IFERROR(${eCnt(mo)},0)`,
      i===0 ? `=D9` : `=IFERROR(H${r-1}+D${r},0)`,
    ]] });
  });

  // Annual Total row 21
  data.push({ range:`${S}!A21:H21`, values:[['ANNUAL TOTAL',`=SUM(B9:B20)`,`=SUM(C9:C20)`,`=IFERROR(B21-C21,0)`,`=IFERROR(D21/B21,0)`,`=SUM(F9:F20)`,`=SUM(G9:G20)`,`=D21`]] });

  // Summary cards rows 23-25
  data.push({ range:`${S}!B23:I23`, values:[['ANNUAL INCOME','','ANNUAL EXPENSES','','ANNUAL PROFIT','','ANNUAL MARGIN','BEST PROFIT MONTH']] });
  data.push({ range:`${S}!B24`, values:[[`=B21`]] });
  data.push({ range:`${S}!D24`, values:[[`=C21`]] });
  data.push({ range:`${S}!F24`, values:[[`=D21`]] });
  data.push({ range:`${S}!H24`, values:[[`=IFERROR(D21/B21,0)`]] });
  data.push({ range:`${S}!I24`, values:[[`=IFERROR(INDEX(A9:A20,MATCH(MAX(D9:D20),D9:D20,0)),"")`]] });

  await valuesBatchUpdate(id, data, 'annual-values');

  // Formatting
  const reqs = [];
  const merge = (r1,r2,c1,c2) => reqs.push({ mergeCells:{ range:gridRange(AS,r1,r2,c1,c2), mergeType:'MERGE_ALL' } });
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(AS,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const currency = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  const pct      = { type:'PERCENT',  pattern:'0.0%' };
  const medium   = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin     = { style:'SOLID',        color:hex(C.border) };

  merge(0,1,0,8);
  fmt(0,1,0,9,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,4,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,4,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Cards
  const cardCols = [[1,3],[3,5],[5,7],[7,8],[8,9]];
  cardCols.forEach(([c1,c2]) => { merge(22,23,c1,c2); merge(23,25,c1,c2); });
  fmt(22,23,1,9,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:8, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(23,25,1,9,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:16, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (const c of [1,3,5]) reqs.push({ repeatCell:{ range:gridRange(AS,23,25,c,c+2), cell:{ userEnteredFormat:{ numberFormat:currency } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(AS,23,25,7,8), cell:{ userEnteredFormat:{ numberFormat:pct } }, fields:'userEnteredFormat.numberFormat' } });

  // Monthly table
  fmt(7,8,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 8; r < 20; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,8,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(20,21,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' } } });
  for (const ci of [1,2,3,7]) reqs.push({ repeatCell:{ range:gridRange(AS,8,21,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(AS,8,21,4,5), cell:{ userEnteredFormat:{ numberFormat:pct, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });

  // Borders
  [[7,21,0,8],[22,25,1,9]].forEach(([r1,r2,c1,c2]) => {
    reqs.push({ updateBorders:{ range:gridRange(AS,r1,r2,c1,c2), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });
  });

  // Col widths
  [[0,110],[1,110],[2,110],[3,110],[4,90],[5,70],[6,70],[7,120],[8,120]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:AS, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:AS, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:AS, dimension:'ROWS', startIndex:1, endIndex:100 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Charts: Combo income/expense/profit, Line margin, Column quarterly profit
  // Chart 1: Combo monthly income vs expenses (columns) + profit (line)
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Monthly Income, Expenses & Profit',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'COMBO',
      legendPosition:'BOTTOM_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Month' },{ position:'LEFT_AXIS', title:'Amount ($)' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:AS, startRowIndex:7, endRowIndex:21, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:AS, startRowIndex:7, endRowIndex:21, startColumnIndex:1, endColumnIndex:2 }] } }, targetAxis:'LEFT_AXIS', type:'COLUMN' },
        { series:{ sourceRange:{ sources:[{ sheetId:AS, startRowIndex:7, endRowIndex:21, startColumnIndex:2, endColumnIndex:3 }] } }, targetAxis:'LEFT_AXIS', type:'COLUMN' },
        { series:{ sourceRange:{ sources:[{ sheetId:AS, startRowIndex:7, endRowIndex:21, startColumnIndex:3, endColumnIndex:4 }] } }, targetAxis:'LEFT_AXIS', type:'LINE' },
      ]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:AS, rowIndex:26, columnIndex:0 }, widthPixels:500, heightPixels:300 } } } } });

  // Chart 2: Monthly margin line
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Monthly Profit Margin',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'LINE',
      legendPosition:'BOTTOM_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Month' },{ position:'LEFT_AXIS', title:'Margin' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:AS, startRowIndex:7, endRowIndex:21, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:AS, startRowIndex:7, endRowIndex:21, startColumnIndex:4, endColumnIndex:5 }] } }, targetAxis:'LEFT_AXIS' }]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:AS, rowIndex:26, columnIndex:6 }, widthPixels:350, heightPixels:280 } } } } });

  // Freeze rows 1-6
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:AS, gridProperties:{ frozenRowCount:6, frozenColumnCount:0 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'annual-format');
  console.log('Annual Summary complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
