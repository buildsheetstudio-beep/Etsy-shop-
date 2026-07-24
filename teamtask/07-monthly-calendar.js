'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const MC = sheetMap['Monthly Calendar'];
const S = "'Monthly Calendar'";

const ML = "'Master Task Log'";
const mon = '$B$2', yr = '$B$3', mbr = '$B$4', proj = '$B$5';

// Calendar grid: 7 cols (Mon-Sun) x 6 rows = 42 cells
// We'll use a helper table in cols J:S to precompute day numbers for the grid
// Grid starts at row 9, columns A-G

// Date of first day of selected month
const firstDay = `DATE(${yr},MATCH(${mon},'Reference Data'!$P$2:$P$13,0),1)`;

// Day of week for first day (1=Sun, 2=Mon...7=Sat) — convert to Mon-first (0=Mon..6=Sun)
const firstDOW = `MOD(WEEKDAY(${firstDay},2)-1,7)`; // 0=Mon,1=Tue,...,6=Sun

// Day number in cell (row r 0..5, col c 0..6)
function cellDay(r, c) {
  const cellN = r * 7 + c + 1; // 1..42
  return `IFERROR(IF(AND(${cellN}>${firstDOW},${cellN}-${firstDOW}<=DAY(DATE(${yr},MATCH(${mon},'Reference Data'!$P$2:$P$13,0)+1,0))),${cellN}-${firstDOW},""),"")`;
}

// Date of a given cell
function cellDate(r, c) {
  const cellN = r * 7 + c + 1;
  return `IFERROR(IF(AND(${cellN}>${firstDOW},${cellN}-${firstDOW}<=DAY(DATE(${yr},MATCH(${mon},'Reference Data'!$P$2:$P$13,0)+1,0))),DATE(${yr},MATCH(${mon},'Reference Data'!$P$2:$P$13,0),${cellN}-${firstDOW}),""),"")`;
}

// Count tasks on a date expression (using DATE(yr,monthNum,day))
function taskCountOnDate(dateExpr) {
  return `IFERROR(SUMPRODUCT((${ML}!$I$6:$I$505=${dateExpr})*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(${ML}!$K$6:$K$505<>"Cancelled")),0)`;
}

// Get nth task name on a date
function taskNameOnDate(dateExpr, n) {
  return `IFERROR(INDEX(${ML}!$B$6:$B$505,SMALL(IF((${ML}!$I$6:$I$505=${dateExpr})*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(${ML}!$K$6:$K$505<>"Cancelled"),ROW(${ML}!$B$6:$B$505)-ROW(${ML}!$B$6)+1),${n})),"")`;
}

(async () => {
  const data = [];

  data.push({ range:`${S}!A1`, values:[['MONTHLY CALENDAR']] });
  data.push({ range:`${S}!A2:B2`, values:[['Month:', 'July']] });
  data.push({ range:`${S}!A3:B3`, values:[['Year:', 2026]] });
  data.push({ range:`${S}!A4:B4`, values:[['Team Member:', 'All Team Members']] });
  data.push({ range:`${S}!A5:B5`, values:[['Project:', 'All Projects']] });

  // Summary cards row 7-8
  data.push({ range:`${S}!A7:L7`, values:[['TASKS DUE','','COMPLETED','','OVERDUE','','COMPLETION %','','BUSIEST DAY','','BUSIEST MEMBER','']] });

  const monthNum = `MATCH(${mon},'Reference Data'!$P$2:$P$13,0)`;
  const inMonth = `(YEAR(${ML}!$I$6:$I$505)=${yr})*(MONTH(${ML}!$I$6:$I$505)=${monthNum})`;

  data.push({ range:`${S}!A8`, values:[[`=IFERROR(SUMPRODUCT(${inMonth}*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(${ML}!$K$6:$K$505<>"Cancelled")),0)`]] });
  data.push({ range:`${S}!C8`, values:[[`=IFERROR(SUMPRODUCT(${inMonth}*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(${ML}!$K$6:$K$505="Completed")),0)`]] });
  data.push({ range:`${S}!E8`, values:[[`=IFERROR(SUMPRODUCT(${inMonth}*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(${ML}!$S$6:$S$505="Yes")),0)`]] });
  data.push({ range:`${S}!G8`, values:[[`=IFERROR(C8/A8,0)`]] });
  data.push({ range:`${S}!I8`, values:[[`=IFERROR(TEXT(DATE(${yr},${monthNum},MATCH(MAX(COUNTIFS(${ML}!$I$6:$I$505,DATE(${yr},${monthNum},ROW(INDIRECT("1:"&DAY(DATE(${yr},${monthNum}+1,0))))),${ML}!$K$6:$K$505,"<>Cancelled")),COUNTIFS(${ML}!$I$6:$I$505,DATE(${yr},${monthNum},ROW(INDIRECT("1:"&DAY(DATE(${yr},${monthNum}+1,0))))),${ML}!$K$6:$K$505,"<>Cancelled"),0)),"MMM D"),"")`]] });
  data.push({ range:`${S}!K8`, values:[[`=IFERROR(INDEX(${ML}!$F$6:$F$505,MATCH(MAX(COUNTIF(${ML}!$F$6:$F$505,${ML}!$F$6:$F$505)*(YEAR(${ML}!$I$6:$I$505)=${yr})*(MONTH(${ML}!$I$6:$I$505)=${monthNum})),COUNTIF(${ML}!$F$6:$F$505,${ML}!$F$6:$F$505)*(YEAR(${ML}!$I$6:$I$505)=${yr})*(MONTH(${ML}!$I$6:$I$505)=${monthNum}),0)),"")`]] });

  // Calendar day-of-week header row 9
  data.push({ range:`${S}!A9:G9`, values:[['Mon','Tue','Wed','Thu','Fri','Sat','Sun']] });

  // Calendar grid rows 10-15 (6 rows), cols A-G (7 cols)
  const COLS = ['A','B','C','D','E','F','G'];
  for (let r = 0; r < 6; r++) {
    for (let c = 0; c < 7; c++) {
      const row = 10 + r * 4; // each day cell is 4 rows tall (day num + 3 task names)
      const col = COLS[c];

      // We'll use single rows per day for simplicity — 6 rows of calendar
      const actualRow = 10 + r;
      const cellN = r * 7 + c + 1;
      const dayNumExpr = cellDay(r, c);
      const dateExpr = cellDate(r, c);

      // Single row calendar cell: day number + count + up to 2 task names
      data.push({ range:`${S}!${col}${actualRow}`, values:[[
        `=IFERROR(IF(${dayNumExpr}="","",${dayNumExpr}&IF(${taskCountOnDate(dateExpr)}>0,CHAR(10)&${taskNameOnDate(dateExpr,1)}&IF(${taskCountOnDate(dateExpr)}>1,CHAR(10)&${taskNameOnDate(dateExpr,2)}&IF(${taskCountOnDate(dateExpr)}>2," + "&(${taskCountOnDate(dateExpr)}-2)&" more",""),""),"")),"")`
      ]] });
    }
  }

  // Monthly detail table row 18+
  data.push({ range:`${S}!A18:F18`, values:[['MONTHLY TASK DETAIL','','','','','']] });
  data.push({ range:`${S}!A19:F19`, values:[['Due Date','Task Name','Assigned To','Project','Priority','Status']] });

  for (let t = 0; t < 20; t++) {
    const r = 20 + t;
    const matchF = `SMALL(IF((YEAR(${ML}!$I$6:$I$505)=${yr})*(MONTH(${ML}!$I$6:$I$505)=${monthNum})*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(${ML}!$K$6:$K$505<>"Cancelled"),ROW(${ML}!$I$6:$I$505)-ROW(${ML}!$I$6)+1),${t+1})`;
    data.push({ range:`${S}!A${r}`, values:[[`=IFERROR(INDEX(${ML}!$I$6:$I$505,${matchF}),"")`]] });
    data.push({ range:`${S}!B${r}`, values:[[`=IFERROR(INDEX(${ML}!$B$6:$B$505,${matchF}),"")`]] });
    data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(INDEX(${ML}!$F$6:$F$505,${matchF}),"")`]] });
    data.push({ range:`${S}!D${r}`, values:[[`=IFERROR(INDEX(${ML}!$D$6:$D$505,${matchF}),"")`]] });
    data.push({ range:`${S}!E${r}`, values:[[`=IFERROR(INDEX(${ML}!$L$6:$L$505,${matchF}),"")`]] });
    data.push({ range:`${S}!F${r}`, values:[[`=IFERROR(INDEX(${ML}!$K$6:$K$505,${matchF}),"")`]] });
  }

  await valuesBatchUpdate(id, data, 'monthly-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(MC,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };
  const pctF   = { type:'PERCENT', pattern:'0%' };

  // Title
  reqs.push({ mergeCells:{ range:gridRange(MC,0,2,0,12), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:18, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,6,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial' } } });
  fmt(1,6,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' } } });

  // Summary cards
  const kpiCols = [[0,2],[2,4],[4,6],[6,8],[8,10],[10,12]];
  kpiCols.forEach(([c1,c2]) => {
    reqs.push({ mergeCells:{ range:gridRange(MC,6,7,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(MC,7,9,c1,c2), mergeType:'MERGE_ALL' } });
  });
  fmt(6,7,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:8, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(7,9,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:14, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  reqs.push({ repeatCell:{ range:gridRange(MC,7,9,6,8), cell:{ userEnteredFormat:{ numberFormat:pctF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(MC,6,9,0,12), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Calendar day-of-week header
  fmt(8,9,0,7,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Calendar cells
  for (let r = 9; r < 15; r++) {
    const bg = r%2===0 ? C.panel : '#F4F1F6'; // alternate lavender
    fmt(r,r+1,0,7,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:8, fontFamily:'Arial' }, verticalAlignment:'TOP', wrapStrategy:'WRAP' } });
  }
  // Weekend cols F,G (5,6) subtle tint
  for (const ci of [5,6]) {
    reqs.push({ repeatCell:{ range:gridRange(MC,9,15,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex('#F0EDE8') } }, fields:'userEnteredFormat.backgroundColor' } });
  }
  reqs.push({ updateBorders:{ range:gridRange(MC,8,15,0,7), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Detail table
  reqs.push({ mergeCells:{ range:gridRange(MC,17,18,0,6), mergeType:'MERGE_ALL' } });
  fmt(17,18,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(18,19,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 19; r < 40; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,6,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(MC,19,40,0,1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(MC,17,40,0,6), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Col widths
  [[0,100],[1,100],[2,100],[3,100],[4,100],[5,100],[6,100],[7,90],[8,90],[9,90],[10,100],[11,100]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:MC, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:MC, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:36 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:MC, dimension:'ROWS', startIndex:2, endIndex:6 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:MC, dimension:'ROWS', startIndex:9, endIndex:15 }, properties:{ pixelSize:72 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:MC, dimension:'ROWS', startIndex:7, endIndex:9 }, properties:{ pixelSize:36 }, fields:'pixelSize' } });

  reqs.push({ updateSheetProperties:{ properties:{ sheetId:MC, gridProperties:{ frozenRowCount:6 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'monthly-format');
  console.log('Monthly Calendar complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
