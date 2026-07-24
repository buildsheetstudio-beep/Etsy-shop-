'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const WC = sheetMap['Weekly Calendar'];
const S = "'Weekly Calendar'";

// Controls: B2 = Week Start Date, B3 = Team Member, B4 = Project, B5 = Department
// ML = Master Task Log abbreviation for formulas
const ML = "'Master Task Log'";
const wkStart = '$B$2', mbr = '$B$3', proj = '$B$4', dept = '$B$5';

// Helper: build a per-day count formula for summary cards
// dayOff = 0..6 (0=Mon..6=Sun relative to wkStart)
const dayDate = (n) => `(${wkStart}+${n})`;

// COUNTIFS with optional member/project/dept filter
function buildFilter(extraCriteria) {
  const base = [
    `${ML}!$I$6:$I$505`,  // due date
    'dueCol',             // placeholder replaced per use
    `${ML}!$K$6:$K$505`, // status
    '<>Cancelled',
  ];
  const parts = [
    [`${ML}!$F$6:$F$505`, mbr, `${mbr}="All Team Members"`],
    [`${ML}!$D$6:$D$505`, proj, `${proj}="All Projects"`],
    [`${ML}!$E$6:$E$505`, dept, `${dept}="All Departments"`],
  ];
  return parts; // return filter structure for use in formulas
}

// Count tasks due on a specific date with filters
function countDay(dayN, statusFilter) {
  const d = dayN === 0 ? wkStart : `${wkStart}+${dayN}`;
  const statusClause = statusFilter ? `,${ML}!$K$6:$K$505,"${statusFilter}"` : `,${ML}!$K$6:$K$505,"<>Cancelled"`;
  return `IFERROR(SUMPRODUCT((${ML}!$I$6:$I$505=${d})*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(IF(${dept}="All Departments",1,${ML}!$E$6:$E$505=${dept}))*(${ML}!$K$6:$K$505<>"Cancelled")${statusFilter?`*(${ML}!$K$6:$K$505="${statusFilter}")`:''}),0)`;
}

// Week-total summary
function weekTotal(col, criteria) {
  return `IFERROR(SUMPRODUCT((${ML}!$I$6:$I$505>=${wkStart})*(${ML}!$I$6:$I$505<=${wkStart}+6)*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(IF(${dept}="All Departments",1,${ML}!$E$6:$E$505=${dept}))${criteria}),0)`;
}

// For the detail table below: use SMALL/INDEX/MATCH approach for first 15 tasks
// This uses a helper column approach with COUNTIFS and SMALL to find task indices

(async () => {
  const data = [];

  // Title
  data.push({ range:`${S}!A1`, values:[['WEEKLY CALENDAR']] });
  data.push({ range:`${S}!A2:B2`, values:[['Week Start Date:', '2026-07-27']] });
  data.push({ range:`${S}!A3:B3`, values:[['Team Member:', 'All Team Members']] });
  data.push({ range:`${S}!A4:B4`, values:[['Project:', 'All Projects']] });
  data.push({ range:`${S}!A5:B5`, values:[['Department:', 'All Departments']] });

  // Summary cards row 7 (labels) & 8 (values)
  data.push({ range:`${S}!A7:L7`, values:[[
    'TASKS DUE THIS WEEK','','COMPLETED','','OVERDUE','','URGENT','','EST. HOURS','','COMPLETION %',''
  ]] });
  // Cards formulas
  const weekDue     = weekTotal('', '*(${ML}!$K$6:$K$505<>"Cancelled")');
  const weekComp    = weekTotal('', `*(${ML}!$K$6:$K$505="Completed")`);
  const weekOverdue = `IFERROR(SUMPRODUCT((${ML}!$I$6:$I$505>=${wkStart})*(${ML}!$I$6:$I$505<=${wkStart}+6)*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(IF(${dept}="All Departments",1,${ML}!$E$6:$E$505=${dept}))*(${ML}!$S$6:$S$505="Yes")),0)`;
  const weekUrgent  = weekTotal('', `*(${ML}!$L$6:$L$505="Urgent")`);
  const weekHours   = `IFERROR(SUMPRODUCT((${ML}!$I$6:$I$505>=${wkStart})*(${ML}!$I$6:$I$505<=${wkStart}+6)*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(IF(${dept}="All Departments",1,${ML}!$E$6:$E$505=${dept}))*(${ML}!$N$6:$N$505)),0)`;
  const weekPct     = `IFERROR(${weekComp}/${weekTotal('','*(${ML}!$K$6:$K$505<>"Cancelled")')},0)`;

  data.push({ range:`${S}!A8`, values:[[`=${weekTotal('','*(${ML}!$K$6:$K$505<>"Cancelled")')}`]] });
  data.push({ range:`${S}!C8`, values:[[`=${weekTotal('',`*(${ML}!$K$6:$K$505="Completed")`)}`]] });
  data.push({ range:`${S}!E8`, values:[[`=${weekOverdue}`]] });
  data.push({ range:`${S}!G8`, values:[[`=${weekUrgent}`]] });
  data.push({ range:`${S}!I8`, values:[[`=${weekHours}`]] });
  data.push({ range:`${S}!K8`, values:[[`=IFERROR(C8/A8,0)`]] });

  // Day columns header row 10
  const DAYNAMES = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
  data.push({ range:`${S}!A10:G10`, values:[DAYNAMES.map((d,i) => `${d} (=TEXT(${wkStart}+${i},"MMM D"))`)
    .map((f,i) => `=${i===0 ? `TEXT(${wkStart},"ddd MMM D")` : `TEXT(${wkStart}+${i},"ddd MMM D")`}`)
  ] });

  // Actually push day names as text with date formulas
  data.push({ range:`${S}!A10`, values:[[`=TEXT(${wkStart},"ddd MMM D")`]] });
  data.push({ range:`${S}!B10`, values:[[`=TEXT(${wkStart}+1,"ddd MMM D")`]] });
  data.push({ range:`${S}!C10`, values:[[`=TEXT(${wkStart}+2,"ddd MMM D")`]] });
  data.push({ range:`${S}!D10`, values:[[`=TEXT(${wkStart}+3,"ddd MMM D")`]] });
  data.push({ range:`${S}!E10`, values:[[`=TEXT(${wkStart}+4,"ddd MMM D")`]] });
  data.push({ range:`${S}!F10`, values:[[`=TEXT(${wkStart}+5,"ddd MMM D")`]] });
  data.push({ range:`${S}!G10`, values:[[`=TEXT(${wkStart}+6,"ddd MMM D")`]] });

  // Day task count rows 11 (tasks due that day)
  for (let d = 0; d < 7; d++) {
    const col = String.fromCharCode(65 + d);
    const dayD = d === 0 ? wkStart : `${wkStart}+${d}`;
    data.push({ range:`${S}!${col}11`, values:[[
      `=IFERROR(SUMPRODUCT((${ML}!$I$6:$I$505=${dayD})*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(IF(${dept}="All Departments",1,${ML}!$E$6:$E$505=${dept}))*(${ML}!$K$6:$K$505<>"Cancelled")),0)&" task(s)"`,
    ]] });
  }

  // Day task detail rows 12-17 (up to 5 task names per day using IFERROR+INDEX+SMALL)
  for (let d = 0; d < 7; d++) {
    const col = String.fromCharCode(65 + d);
    const dayD = d === 0 ? wkStart : `${wkStart}+${d}`;
    for (let t = 0; t < 5; t++) {
      const r = 12 + t;
      // SMALL-INDEX approach: find t+1-th row where due date matches and filters apply
      data.push({ range:`${S}!${col}${r}`, values:[[
        `=IFERROR(INDEX(${ML}!$B$6:$B$505,SMALL(IF((${ML}!$I$6:$I$505=${dayD})*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(IF(${dept}="All Departments",1,${ML}!$E$6:$E$505=${dept}))*(${ML}!$K$6:$K$505<>"Cancelled"),ROW(${ML}!$B$6:$B$505)-ROW(${ML}!$B$6)+1),${t+1})),"")`,
      ]] });
    }
  }

  // Detail table below row 20
  data.push({ range:`${S}!A20:H20`, values:[['WEEKLY TASK DETAIL','','','','','','','']] });
  data.push({ range:`${S}!A21:H21`, values:[['Due Date','Task ID','Task Name','Assigned To','Project','Priority','Status','Days Remaining']] });

  for (let t = 0; t < 15; t++) {
    const r = 22 + t;
    const rankN = t + 1;
    const matchFormula = `SMALL(IF((${ML}!$I$6:$I$505>=${wkStart})*(${ML}!$I$6:$I$505<=${wkStart}+6)*(IF(${mbr}="All Team Members",1,${ML}!$F$6:$F$505=${mbr}))*(IF(${proj}="All Projects",1,${ML}!$D$6:$D$505=${proj}))*(IF(${dept}="All Departments",1,${ML}!$E$6:$E$505=${dept}))*(${ML}!$K$6:$K$505<>"Cancelled"),ROW(${ML}!$I$6:$I$505)-ROW(${ML}!$I$6)+1),${rankN})`;
    data.push({ range:`${S}!A${r}`, values:[[`=IFERROR(INDEX(${ML}!$I$6:$I$505,${matchFormula}),"")`]] });
    data.push({ range:`${S}!B${r}`, values:[[`=IFERROR(INDEX(${ML}!$A$6:$A$505,${matchFormula}),"")`]] });
    data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(INDEX(${ML}!$B$6:$B$505,${matchFormula}),"")`]] });
    data.push({ range:`${S}!D${r}`, values:[[`=IFERROR(INDEX(${ML}!$F$6:$F$505,${matchFormula}),"")`]] });
    data.push({ range:`${S}!E${r}`, values:[[`=IFERROR(INDEX(${ML}!$D$6:$D$505,${matchFormula}),"")`]] });
    data.push({ range:`${S}!F${r}`, values:[[`=IFERROR(INDEX(${ML}!$L$6:$L$505,${matchFormula}),"")`]] });
    data.push({ range:`${S}!G${r}`, values:[[`=IFERROR(INDEX(${ML}!$K$6:$K$505,${matchFormula}),"")`]] });
    data.push({ range:`${S}!H${r}`, values:[[`=IFERROR(INDEX(${ML}!$R$6:$R$505,${matchFormula}),"")`]] });
  }

  await valuesBatchUpdate(id, data, 'weekly-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(WC,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };
  const numF   = { type:'NUMBER', pattern:'0' };
  const pctF   = { type:'PERCENT', pattern:'0%' };

  // Title
  reqs.push({ mergeCells:{ range:gridRange(WC,0,2,0,12), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:18, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(1,6,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial' } } });
  fmt(1,6,1,2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' } } });
  reqs.push({ repeatCell:{ range:gridRange(WC,1,2,1,2), cell:{ userEnteredFormat:{ numberFormat:dateF } }, fields:'userEnteredFormat.numberFormat' } });

  // Summary cards
  const kpiCols = [[0,2],[2,4],[4,6],[6,8],[8,10],[10,12]];
  kpiCols.forEach(([c1,c2]) => {
    reqs.push({ mergeCells:{ range:gridRange(WC,6,7,c1,c2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(WC,7,9,c1,c2), mergeType:'MERGE_ALL' } });
  });
  fmt(6,7,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:8, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(7,9,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:18, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  reqs.push({ repeatCell:{ range:gridRange(WC,7,9,10,12), cell:{ userEnteredFormat:{ numberFormat:pctF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(WC,6,9,0,12), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Day header row
  fmt(9,10,0,7,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  // Day task count row
  fmt(10,11,0,7,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  // Day task name rows 12-17
  for (let r = 11; r < 17; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,7,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });
  }
  // Weekend cols (F=5, G=6) — muted tint
  for (const ci of [5,6]) {
    reqs.push({ repeatCell:{ range:gridRange(WC,9,17,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex('#F0EDE8') } }, fields:'userEnteredFormat.backgroundColor' } });
  }
  reqs.push({ updateBorders:{ range:gridRange(WC,9,17,0,7), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Detail table
  reqs.push({ mergeCells:{ range:gridRange(WC,19,20,0,8), mergeType:'MERGE_ALL' } });
  fmt(19,20,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(20,21,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 21; r < 37; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,8,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(WC,21,37,0,1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  reqs.push({ repeatCell:{ range:gridRange(WC,21,37,7,8), cell:{ userEnteredFormat:{ numberFormat:{ type:'NUMBER', pattern:'0;[Red]-0' }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(WC,19,37,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Conditional formatting — overdue in detail table
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(WC,21,37,0,8)], booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=AND($G22<>"Completed",$G22<>"Cancelled",$G22<>"",$A22<TODAY())` }] }, format:{ backgroundColor:hex('#FEF0F0') } } }, index:0 } });
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(WC,21,37,0,8)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'Completed' }] }, format:{ backgroundColor:hex('#E8F5E9') } } }, index:0 } });
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(WC,21,37,5,6)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'Urgent' }] }, format:{ backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white) } } } }, index:0 } });

  // Col widths
  [[0,120],[1,120],[2,120],[3,120],[4,120],[5,120],[6,120],[7,90],[8,90],[9,90],[10,90],[11,90]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:WC, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:WC, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:36 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:WC, dimension:'ROWS', startIndex:2, endIndex:6 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:WC, dimension:'ROWS', startIndex:9, endIndex:17 }, properties:{ pixelSize:50 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:WC, dimension:'ROWS', startIndex:7, endIndex:9 }, properties:{ pixelSize:36 }, fields:'pixelSize' } });

  // Freeze rows 1-6 (0-indexed 0-5) - avoid merges that span into frozen area
  // Title merge is rows 0-1, controls are rows 1-5 = no full-width merges after row 1
  // Cards merges at rows 6-8 (0-indexed). Freeze only 6 rows = rows 0-5 = fine
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:WC, gridProperties:{ frozenRowCount:6 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'weekly-format');
  console.log('Weekly Calendar complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
