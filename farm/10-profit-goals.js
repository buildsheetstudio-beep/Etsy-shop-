'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PG = sheetMap['Profit Goals'];
const S = "'Profit Goals'";

(async () => {
  const data = [];
  const yr = '$B$2';

  // Monthly goal from Farm Setup A85:C96 via VLOOKUP on month name
  const goalAmt = (mo) => `IFERROR(VLOOKUP("${mo}",'Farm Setup'!$B$85:$C$96,2,FALSE),0)`;
  // Actual income - expenses for that month
  const incAmt  = (mo) => `IFERROR(SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,"${mo}"),0)`;
  const expAmt  = (mo) => `IFERROR(SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,"${mo}"),0)`;

  data.push({ range:`${S}!A1`, values:[['PROFIT GOALS']] });
  data.push({ range:`${S}!A2:B2`, values:[['Reporting Year:', 2026]] });

  // ── Summary cards rows 5-7 ─────────────────────────────────────────────────
  data.push({ range:`${S}!A5:H5`, values:[['ANNUAL GOAL','','ACTUAL PROFIT','','REMAINING','','REQUIRED MONTHLY AVG','CURRENT MONTHLY AVG']] });
  // Cards reference the monthly table (rows 9-20) that we build below
  data.push({ range:`${S}!A6`, values:[[`=IFERROR(SUM(B9:B20),0)`]] });
  data.push({ range:`${S}!C6`, values:[[`=IFERROR(SUM(D9:D20),0)`]] });
  data.push({ range:`${S}!E6`, values:[[`=IFERROR(MAX(0,A6-C6),0)`]] });
  // Required monthly avg = remaining / months left in year (months where actual=0 and not yet passed goal)
  data.push({ range:`${S}!G6`, values:[[`=IFERROR(E6/MAX(1,COUNTIF(D9:D20,0)),0)`]] });
  // Current avg = actual / months with any profit recorded
  data.push({ range:`${S}!H6`, values:[[`=IFERROR(C6/MAX(1,COUNTIF(D9:D20,">"&0)),0)`]] });

  // ── Monthly goal table rows 9-20 ───────────────────────────────────────────
  data.push({ range:`${S}!A8:G8`, values:[['Month','Profit Goal','Income','Expenses','Actual Profit','Variance','Status']] });

  MONTHS.forEach((mo, i) => {
    const r = 9 + i;
    data.push({ range:`${S}!A${r}:G${r}`, values:[[
      mo,
      `=${goalAmt(mo)}`,
      `=${incAmt(mo)}`,
      `=${expAmt(mo)}`,
      `=IFERROR(C${r}-D${r},0)`,
      `=IFERROR(E${r}-B${r},0)`,
      `=IFERROR(IF(E${r}>=B${r},"Goal Met",IF(E${r}>=B${r}*0.9,"In Progress","Below Goal")),"")`,
    ]] });
  });

  // Annual total row 21
  data.push({ range:`${S}!A21:G21`, values:[[
    'ANNUAL TOTAL',
    `=SUM(B9:B20)`,
    `=SUM(C9:C20)`,
    `=SUM(D9:D20)`,
    `=IFERROR(C21-D21,0)`,
    `=IFERROR(E21-B21,0)`,
    '',
  ]] });

  // ── Cumulative table rows 24-36 ────────────────────────────────────────────
  data.push({ range:`${S}!A23:E23`, values:[['CUMULATIVE TRACKING','','','','']] });
  data.push({ range:`${S}!A24:E24`, values:[['Month','Cumulative Goal','Cumulative Actual','Cumulative Variance','On Track?']] });

  MONTHS.forEach((mo, i) => {
    const r = 25 + i;
    const goalSum = `SUM(B9:B${9+i})`;
    const profitSum = `SUM(E9:E${9+i})`;
    data.push({ range:`${S}!A${r}:E${r}`, values:[[
      mo,
      `=IFERROR(${goalSum},0)`,
      `=IFERROR(${profitSum},0)`,
      `=IFERROR(C${r}-B${r},0)`,
      `=IFERROR(IF(C${r}>=B${r},"On Track","Behind"),"")`
    ]] });
  });

  await valuesBatchUpdate(id, data, 'goals-values');

  // Formatting
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(PG,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const currency = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  const medium   = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin     = { style:'SOLID',        color:hex(C.border) };

  reqs.push({ mergeCells:{ range:gridRange(PG,0,1,0,8), mergeType:'MERGE_ALL' } });
  fmt(0,1,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,3,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial' } } });
  fmt(1,3,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' } } });

  // Summary cards
  const cardCols = [[0,2],[2,4],[4,6],[6,7],[7,8]];
  cardCols.forEach(([c1,c2]) => {
    reqs.push({ mergeCells:{ range:gridRange(PG,4,5,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(PG,5,7,c1,c2), mergeType:'MERGE_ALL' } });
  });
  fmt(4,5,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:8, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(5,7,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:16, fontFamily:'Arial' }, numberFormat:currency, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  reqs.push({ updateBorders:{ range:gridRange(PG,4,7,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Monthly table header
  fmt(7,8,0,7,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 8; r < 21; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,7,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(20,21,0,7,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' } } });
  for (const ci of [1,2,3,4,5]) {
    reqs.push({ repeatCell:{ range:gridRange(PG,8,21,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  }
  reqs.push({ updateBorders:{ range:gridRange(PG,7,21,0,7), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Cumulative section header
  reqs.push({ mergeCells:{ range:gridRange(PG,22,23,0,5), mergeType:'MERGE_ALL' } });
  fmt(22,23,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(23,24,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 24; r < 36; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,5,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  for (const ci of [1,2,3]) {
    reqs.push({ repeatCell:{ range:gridRange(PG,24,36,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  }
  reqs.push({ updateBorders:{ range:gridRange(PG,23,36,0,5), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Col widths
  [[0,110],[1,110],[2,110],[3,110],[4,110],[5,110],[6,140],[7,140]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:PG, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PG, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PG, dimension:'ROWS', startIndex:1, endIndex:100 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Chart 1: Monthly Goal vs Actual Profit (column)
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Monthly Goal vs Actual Profit',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'COLUMN',
      legendPosition:'BOTTOM_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Month' },{ position:'LEFT_AXIS', title:'Amount ($)' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:PG, startRowIndex:7, endRowIndex:21, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:PG, startRowIndex:7, endRowIndex:21, startColumnIndex:1, endColumnIndex:2 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:PG, startRowIndex:7, endRowIndex:21, startColumnIndex:4, endColumnIndex:5 }] } }, targetAxis:'LEFT_AXIS' },
      ]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:PG, rowIndex:37, columnIndex:0 }, widthPixels:460, heightPixels:300 } } } } });

  // Chart 2: Cumulative Goal vs Actual (line)
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Cumulative Goal vs Actual',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'LINE',
      legendPosition:'BOTTOM_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Month' },{ position:'LEFT_AXIS', title:'Cumulative ($)' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:PG, startRowIndex:23, endRowIndex:36, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:PG, startRowIndex:23, endRowIndex:36, startColumnIndex:1, endColumnIndex:2 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:PG, startRowIndex:23, endRowIndex:36, startColumnIndex:2, endColumnIndex:3 }] } }, targetAxis:'LEFT_AXIS' },
      ]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:PG, rowIndex:37, columnIndex:5 }, widthPixels:460, heightPixels:300 } } } } });

  // freeze row 1 (title) + row 2 (year control) only — cards span rows 5-7 so we can't freeze into them
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:PG, gridProperties:{ frozenRowCount:3, frozenColumnCount:0 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'goals-format');
  console.log('Profit Goals complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
