'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C, ENTERPRISES } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const EP = sheetMap['Enterprise Profitability'];
const S = "'Enterprise Profitability'";

(async () => {
  const data = [];
  const yr = '$B$2', per = '$B$3';

  const iAmt = (ent) => `IFERROR(IF(${per}="Full Year",SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$F$6:$F$505,"${ent}"),SUMIFS('Income Log'!$L$6:$L$505,'Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${per},'Income Log'!$F$6:$F$505,"${ent}")),0)`;
  const eAmt = (ent) => `IFERROR(IF(${per}="Full Year",SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$F$6:$F$505,"${ent}"),SUMIFS('Expense Log'!$M$6:$M$505,'Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,${per},'Expense Log'!$F$6:$F$505,"${ent}")),0)`;

  data.push({ range:`${S}!A1`, values:[['ENTERPRISE PROFITABILITY']] });
  data.push({ range:`${S}!A2:B2`, values:[['Reporting Year:', 2026]] });
  data.push({ range:`${S}!A3:B3`, values:[['Period:', 'Full Year']] });

  // Table header row 8
  data.push({ range:`${S}!A8:J8`, values:[['Enterprise','Income','Expenses','Net Profit','Margin','% of Total Income','% of Total Profit','Inc Txns','Exp Txns','Performance Flag']] });

  // Enterprise rows 9-23
  ENTERPRISES.forEach((ent, i) => {
    const r = 9 + i;
    data.push({ range:`${S}!A${r}:J${r}`, values:[[
      ent,
      `=${iAmt(ent)}`,
      `=${eAmt(ent)}`,
      `=IFERROR(B${r}-C${r},0)`,
      `=IFERROR(D${r}/B${r},0)`,
      `=IFERROR(B${r}/B24,0)`,
      `=IFERROR(D${r}/D24,0)`,
      `=IFERROR(IF(${per}="Full Year",COUNTIFS('Income Log'!$C$6:$C$505,${yr},'Income Log'!$F$6:$F$505,"${ent}"),COUNTIFS('Income Log'!$C$6:$C$505,${yr},'Income Log'!$D$6:$D$505,${per},'Income Log'!$F$6:$F$505,"${ent}")),0)`,
      `=IFERROR(IF(${per}="Full Year",COUNTIFS('Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$F$6:$F$505,"${ent}"),COUNTIFS('Expense Log'!$C$6:$C$505,${yr},'Expense Log'!$D$6:$D$505,${per},'Expense Log'!$F$6:$F$505,"${ent}")),0)`,
      `=IFERROR(IF(E${r}>=0.25,"Strong Profit",IF(E${r}>0,"Profitable",IF(E${r}>=-0.05,"Near Break-Even","Loss — Review Required"))),"")`,
    ]] });
  });

  // Total row 24
  data.push({ range:`${S}!A24:J24`, values:[['TOTAL',`=SUM(B9:B23)`,`=SUM(C9:C23)`,`=IFERROR(B24-C24,0)`,`=IFERROR(D24/B24,0)`,`=1`,`=1`,`=SUM(H9:H23)`,`=SUM(I9:I23)`,'']] });

  await valuesBatchUpdate(id, data, 'enterprise-values');

  // Formatting
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(EP,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const currency = { type:'CURRENCY', pattern:'$#,##0.00;[Red]-$#,##0.00' };
  const pct      = { type:'PERCENT',  pattern:'0.0%' };
  const medium   = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin     = { style:'SOLID',        color:hex(C.border) };

  reqs.push({ mergeCells:{ range:gridRange(EP,0,1,0,10), mergeType:'MERGE_ALL' } });
  fmt(0,1,0,10,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,4,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial' } } });
  fmt(1,4,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' } } });

  fmt(7,8,0,10,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 8; r < 23; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,10,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  fmt(23,24,0,10,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' } } });

  for (const ci of [1,2,3]) reqs.push({ repeatCell:{ range:gridRange(EP,7,24,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:currency, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });
  for (const ci of [4,5,6]) reqs.push({ repeatCell:{ range:gridRange(EP,7,24,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:pct, horizontalAlignment:'RIGHT' } }, fields:'userEnteredFormat' } });

  reqs.push({ updateBorders:{ range:gridRange(EP,7,24,0,10), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  [[0,140],[1,110],[2,100],[3,100],[4,80],[5,110],[6,110],[7,70],[8,70],[9,140]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:EP, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:EP, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:EP, dimension:'ROWS', startIndex:1, endIndex:60 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Charts
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Net Profit by Enterprise',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'BAR',
      legendPosition:'NO_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Net Profit ($)' },{ position:'LEFT_AXIS', title:'Enterprise' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:EP, startRowIndex:7, endRowIndex:24, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:EP, startRowIndex:7, endRowIndex:24, startColumnIndex:3, endColumnIndex:4 }] } }, targetAxis:'BOTTOM_AXIS' }]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:EP, rowIndex:26, columnIndex:0 }, widthPixels:420, heightPixels:320 } } } } });

  reqs.push({ addChart:{ chart:{ spec:{
    title:'Income vs Expenses by Enterprise',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    basicChart:{
      chartType:'COLUMN',
      legendPosition:'BOTTOM_LEGEND',
      headerCount:1,
      axis:[{ position:'BOTTOM_AXIS', title:'Enterprise' },{ position:'LEFT_AXIS', title:'Amount ($)' }],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:EP, startRowIndex:7, endRowIndex:24, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:EP, startRowIndex:7, endRowIndex:24, startColumnIndex:1, endColumnIndex:2 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:EP, startRowIndex:7, endRowIndex:24, startColumnIndex:2, endColumnIndex:3 }] } }, targetAxis:'LEFT_AXIS' },
      ]
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:EP, rowIndex:26, columnIndex:6 }, widthPixels:420, heightPixels:320 } } } } });

  reqs.push({ addChart:{ chart:{ spec:{
    title:'Income Share by Enterprise',
    backgroundColor:hex(C.panel),
    titleTextFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:10 },
    pieChart:{
      legendPosition:'RIGHT_LEGEND',
      pieHole:0.4,
      domain:{ sourceRange:{ sources:[{ sheetId:EP, startRowIndex:7, endRowIndex:23, startColumnIndex:0, endColumnIndex:1 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:EP, startRowIndex:7, endRowIndex:23, startColumnIndex:1, endColumnIndex:2 }] } }
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:EP, rowIndex:48, columnIndex:0 }, widthPixels:420, heightPixels:300 } } } } });

  reqs.push({ updateSheetProperties:{ properties:{ sheetId:EP, gridProperties:{ frozenRowCount:6, frozenColumnCount:0 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });

  await batchUpdate(id, reqs, 'enterprise-format');
  console.log('Enterprise Profitability complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
