'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PP = sheetMap['Project Progress'];
const S = "'Project Progress'";

// 10 active projects matching Team Setup
const PROJECTS = [
  'Website Redesign','Product Launch','Client Onboarding','Monthly Reporting',
  'Social Media Campaign','Team Operations','Event Planning','Process Improvement',
  'Training Program','General Tasks'
];

// Row for each project in the summary table (rows 9-18, 0-indexed rows 8-17)
// Columns: A=Project ID, B=Project Name, C=Owner, D=Dept, E=Start, F=Target End,
// G=Status, H=Priority, I=Total Tasks, J=Completed, K=In Progress, L=Overdue,
// M=Completion%, N=Est Hours Total, O=Hrs Logged, P=Health, Q=Days Remaining

function totalTasksFormula(r) {
  return `=IFERROR(COUNTIF('Master Task Log'!$D$6:$D$505,B${r}),0)`;
}
function completedFormula(r) {
  return `=IFERROR(COUNTIFS('Master Task Log'!$D$6:$D$505,B${r},'Master Task Log'!$K$6:$K$505,"Completed"),0)`;
}
function inProgressFormula(r) {
  return `=IFERROR(COUNTIFS('Master Task Log'!$D$6:$D$505,B${r},'Master Task Log'!$K$6:$K$505,"In Progress"),0)`;
}
function overdueFormula(r) {
  return `=IFERROR(COUNTIFS('Master Task Log'!$D$6:$D$505,B${r},'Master Task Log'!$J$6:$J$505,"<"&TODAY(),'Master Task Log'!$K$6:$K$505,"<>Completed",'Master Task Log'!$K$6:$K$505,"<>Cancelled"),0)`;
}
function completionPctFormula(r) {
  return `=IFERROR(IF(I${r}=0,0,J${r}/I${r}),0)`;
}
function estHrsFormula(r) {
  return `=IFERROR(SUMIF('Master Task Log'!$D$6:$D$505,B${r},'Master Task Log'!$N$6:$N$505),0)`;
}
function hrsLoggedFormula(r) {
  return `=IFERROR(SUMIF('Master Task Log'!$D$6:$D$505,B${r},'Master Task Log'!$O$6:$O$505),0)`;
}
function healthFormula(r) {
  // Completed: 100%, Delayed: overdue>0 OR (past end date and <100%), At Risk: <50% and >50% of timeline elapsed, On Track otherwise
  return `=IFERROR(IF(M${r}=1,"Completed",IF(L${r}>0,"Delayed",IF(AND(F${r}<>"",F${r}<TODAY(),M${r}<1),"Delayed",IF(AND(F${r}<>"",E${r}<>"",TODAY()>E${r},(TODAY()-E${r})/(F${r}-E${r})>0.5,M${r}<0.5),"At Risk",IF(M${r}=0,"Not Started","On Track"))))),"—")`;
}
function daysRemainingFormula(r) {
  return `=IFERROR(IF(F${r}="","",MAX(0,F${r}-TODAY())),"")`;
}

// Owner lookup from Team Setup (col C = owner of project = B30:B44 names, col C of Team Setup = owner)
// Team Setup project table: col B=name, col C=owner, col D=dept, col E=start, col F=end, col G=status, col H=priority
function ownerFormula(r) {
  return `=IFERROR(INDEX('Team Setup'!$C$30:$C$44,MATCH(B${r},'Team Setup'!$B$30:$B$44,0)),"")`;
}
function deptFormula(r) {
  return `=IFERROR(INDEX('Team Setup'!$D$30:$D$44,MATCH(B${r},'Team Setup'!$B$30:$B$44,0)),"")`;
}
function startFormula(r) {
  return `=IFERROR(INDEX('Team Setup'!$E$30:$E$44,MATCH(B${r},'Team Setup'!$B$30:$B$44,0)),"")`;
}
function endFormula(r) {
  return `=IFERROR(INDEX('Team Setup'!$F$30:$F$44,MATCH(B${r},'Team Setup'!$B$30:$B$44,0)),"")`;
}
function statusFormula(r) {
  return `=IFERROR(INDEX('Team Setup'!$G$30:$G$44,MATCH(B${r},'Team Setup'!$B$30:$B$44,0)),"")`;
}
function priorityFormula(r) {
  return `=IFERROR(INDEX('Team Setup'!$H$30:$H$44,MATCH(B${r},'Team Setup'!$B$30:$B$44,0)),"")`;
}
function projectIdFormula(r) {
  return `=IFERROR(INDEX('Team Setup'!$A$30:$A$44,MATCH(B${r},'Team Setup'!$B$30:$B$44,0)),"")`;
}

// Selected project detail section (rows 22-27) driven by B22 = project picker
// Row 22: control label + picker
// Row 23: Owner | Dept | Start | End
// Row 24: Status | Priority | Total Tasks | Completed
// Row 25: Health | Completion % | Est Hours | Hours Logged
// Row 26: Overdue Tasks | Days Remaining | In Progress | Cancelled
// Row 27: Task breakdown by status (bar-style counts)

// Milestone/task detail table for selected project: rows 30-49 (20 tasks)
// Columns: A=Rank, B=Task ID, C=Task Name, D=Assigned To, E=Priority, F=Due Date,
// G=Est Hrs, H=Status, I=Overdue?

function selectedOwner() { return `=IFERROR(INDEX($C$9:$C$18,MATCH($B$21,$B$9:$B$18,0)),"—")`; }
function selectedDept() { return `=IFERROR(INDEX($D$9:$D$18,MATCH($B$21,$B$9:$B$18,0)),"—")`; }
function selectedStart() { return `=IFERROR(INDEX($E$9:$E$18,MATCH($B$21,$B$9:$B$18,0)),"—")`; }
function selectedEnd() { return `=IFERROR(INDEX($F$9:$F$18,MATCH($B$21,$B$9:$B$18,0)),"—")`; }
function selectedStatus() { return `=IFERROR(INDEX($G$9:$G$18,MATCH($B$21,$B$9:$B$18,0)),"—")`; }
function selectedPriority() { return `=IFERROR(INDEX($H$9:$H$18,MATCH($B$21,$B$9:$B$18,0)),"—")`; }
function selectedTotalTasks() { return `=IFERROR(INDEX($I$9:$I$18,MATCH($B$21,$B$9:$B$18,0)),0)`; }
function selectedCompleted() { return `=IFERROR(INDEX($J$9:$J$18,MATCH($B$21,$B$9:$B$18,0)),0)`; }
function selectedInProgress() { return `=IFERROR(INDEX($K$9:$K$18,MATCH($B$21,$B$9:$B$18,0)),0)`; }
function selectedOverdue() { return `=IFERROR(INDEX($L$9:$L$18,MATCH($B$21,$B$9:$B$18,0)),0)`; }
function selectedPct() { return `=IFERROR(INDEX($M$9:$M$18,MATCH($B$21,$B$9:$B$18,0)),0)`; }
function selectedEstHrs() { return `=IFERROR(INDEX($N$9:$N$18,MATCH($B$21,$B$9:$B$18,0)),0)`; }
function selectedHrsLogged() { return `=IFERROR(INDEX($O$9:$O$18,MATCH($B$21,$B$9:$B$18,0)),0)`; }
function selectedHealth() { return `=IFERROR(INDEX($P$9:$P$18,MATCH($B$21,$B$9:$B$18,0)),"—")`; }
function selectedDaysRem() { return `=IFERROR(INDEX($Q$9:$Q$18,MATCH($B$21,$B$9:$B$18,0)),"—")`; }
function selectedCancelled() { return `=IFERROR(COUNTIFS('Master Task Log'!$D$6:$D$505,$B$21,'Master Task Log'!$K$6:$K$505,"Cancelled"),0)`; }

// Task detail table for selected project
function detailRankFormula(r) { return `=IFERROR(IF($B$21="","",${r-29}),"")` ; }
function detailTaskId(r) {
  return `=IFERROR(IF($B$21="","",INDEX('Master Task Log'!$A$6:$A$505,MATCH(SMALL(IF('Master Task Log'!$D$6:$D$505=$B$21,ROW('Master Task Log'!$A$6:$A$505)-5,9999),ROW()-29),'Master Task Log'!$A$6:$A$505,0))),"")`;
}
// Non-array SMALL/INDEX/MATCH pattern:
function detailField(r, col) {
  // SMALL of row numbers where project matches; then INDEX that col
  return `=IFERROR(IF($B$21="","",INDEX('Master Task Log'!$${col}$6:$${col}$505,SMALL(IF('Master Task Log'!$D$6:$D$505=$B$21,ROW('Master Task Log'!$D$6:$D$505)-5),${r-29}))),"")`;
}

(async () => {
  const data = [];

  // Title & subtitle
  data.push({ range:`${S}!A1`, values:[['PROJECT PROGRESS TRACKER']] });
  data.push({ range:`${S}!A2`, values:[['Monitor project health, completion rates, and team workload by project. Select a project below for a detailed task breakdown.']] });

  // Controls row 4
  data.push({ range:`${S}!A4`, values:[['Filter by:']] });
  data.push({ range:`${S}!B4`, values:[['All Projects']] });

  // Section header for summary table
  data.push({ range:`${S}!A7`, values:[['PROJECT SUMMARY TABLE']] });

  // Summary table headers row 8
  data.push({ range:`${S}!A8:Q8`, values:[[
    'Project ID','Project Name','Owner','Department','Start Date','Target End',
    'Status','Priority','Total Tasks','Completed','In Progress','Overdue',
    'Completion %','Est. Hours','Hrs Logged','Health','Days Left'
  ]] });

  // Summary table rows 9-18
  for (let i = 0; i < PROJECTS.length; i++) {
    const r = 9 + i;
    data.push({ range:`${S}!B${r}`, values:[[PROJECTS[i]]] });
  }

  await valuesBatchUpdate(id, data, 'pp-values-1');

  const data2 = [];
  // Formula columns for summary table
  for (let i = 0; i < PROJECTS.length; i++) {
    const r = 9 + i;
    data2.push({ range:`${S}!A${r}`, values:[[projectIdFormula(r)]] });
    data2.push({ range:`${S}!C${r}`, values:[[ownerFormula(r)]] });
    data2.push({ range:`${S}!D${r}`, values:[[deptFormula(r)]] });
    data2.push({ range:`${S}!E${r}`, values:[[startFormula(r)]] });
    data2.push({ range:`${S}!F${r}`, values:[[endFormula(r)]] });
    data2.push({ range:`${S}!G${r}`, values:[[statusFormula(r)]] });
    data2.push({ range:`${S}!H${r}`, values:[[priorityFormula(r)]] });
    data2.push({ range:`${S}!I${r}`, values:[[totalTasksFormula(r)]] });
    data2.push({ range:`${S}!J${r}`, values:[[completedFormula(r)]] });
    data2.push({ range:`${S}!K${r}`, values:[[inProgressFormula(r)]] });
    data2.push({ range:`${S}!L${r}`, values:[[overdueFormula(r)]] });
    data2.push({ range:`${S}!M${r}`, values:[[completionPctFormula(r)]] });
    data2.push({ range:`${S}!N${r}`, values:[[estHrsFormula(r)]] });
    data2.push({ range:`${S}!O${r}`, values:[[hrsLoggedFormula(r)]] });
    data2.push({ range:`${S}!P${r}`, values:[[healthFormula(r)]] });
    data2.push({ range:`${S}!Q${r}`, values:[[daysRemainingFormula(r)]] });
  }

  // Selected project section header row 20
  data2.push({ range:`${S}!A20`, values:[['SELECTED PROJECT DETAIL']] });

  // Selected project picker row 21
  data2.push({ range:`${S}!A21`, values:[['Select Project:']] });
  data2.push({ range:`${S}!B21`, values:[['Website Redesign']] });

  // Selected project cards — 2 rows, 4 cols each (rows 22-23)
  data2.push({ range:`${S}!A22`, values:[['Owner']] });
  data2.push({ range:`${S}!B22`, values:[[selectedOwner()]] });
  data2.push({ range:`${S}!C22`, values:[['Department']] });
  data2.push({ range:`${S}!D22`, values:[[selectedDept()]] });
  data2.push({ range:`${S}!E22`, values:[['Start Date']] });
  data2.push({ range:`${S}!F22`, values:[[selectedStart()]] });
  data2.push({ range:`${S}!G22`, values:[['Target End']] });
  data2.push({ range:`${S}!H22`, values:[[selectedEnd()]] });

  data2.push({ range:`${S}!A23`, values:[['Status']] });
  data2.push({ range:`${S}!B23`, values:[[selectedStatus()]] });
  data2.push({ range:`${S}!C23`, values:[['Priority']] });
  data2.push({ range:`${S}!D23`, values:[[selectedPriority()]] });
  data2.push({ range:`${S}!E23`, values:[['Health']] });
  data2.push({ range:`${S}!F23`, values:[[selectedHealth()]] });
  data2.push({ range:`${S}!G23`, values:[['Days Left']] });
  data2.push({ range:`${S}!H23`, values:[[selectedDaysRem()]] });

  // KPI cards row 24-25
  data2.push({ range:`${S}!A24`, values:[['Total Tasks']] });
  data2.push({ range:`${S}!B24`, values:[[selectedTotalTasks()]] });
  data2.push({ range:`${S}!C24`, values:[['Completed']] });
  data2.push({ range:`${S}!D24`, values:[[selectedCompleted()]] });
  data2.push({ range:`${S}!E24`, values:[['In Progress']] });
  data2.push({ range:`${S}!F24`, values:[[selectedInProgress()]] });
  data2.push({ range:`${S}!G24`, values:[['Cancelled']] });
  data2.push({ range:`${S}!H24`, values:[[selectedCancelled()]] });

  data2.push({ range:`${S}!A25`, values:[['Overdue']] });
  data2.push({ range:`${S}!B25`, values:[[selectedOverdue()]] });
  data2.push({ range:`${S}!C25`, values:[['Completion %']] });
  data2.push({ range:`${S}!D25`, values:[[selectedPct()]] });
  data2.push({ range:`${S}!E25`, values:[['Est. Hours']] });
  data2.push({ range:`${S}!F25`, values:[[selectedEstHrs()]] });
  data2.push({ range:`${S}!G25`, values:[['Hours Logged']] });
  data2.push({ range:`${S}!H25`, values:[[selectedHrsLogged()]] });

  // Task detail table header row 27
  data2.push({ range:`${S}!A27`, values:[['TASK DETAIL — SELECTED PROJECT']] });
  data2.push({ range:`${S}!A28:I28`, values:[[
    '#','Task ID','Task Name','Assigned To','Priority','Due Date','Est. Hrs','Status','Overdue?'
  ]] });

  // Task detail rows 29-48 (20 rows)
  for (let i = 0; i < 20; i++) {
    const r = 29 + i;
    const n = i + 1;
    data2.push({ range:`${S}!A${r}`, values:[[n]] });
    data2.push({ range:`${S}!B${r}`, values:[[`=IFERROR(IF($B$21="","",INDEX('Master Task Log'!$A$6:$A$505,SMALL(IF('Master Task Log'!$D$6:$D$505=$B$21,ROW('Master Task Log'!$D$6:$D$505)-5),${n}))),"")` ]] });
    data2.push({ range:`${S}!C${r}`, values:[[`=IFERROR(IF($B$21="","",INDEX('Master Task Log'!$C$6:$C$505,SMALL(IF('Master Task Log'!$D$6:$D$505=$B$21,ROW('Master Task Log'!$D$6:$D$505)-5),${n}))),"")` ]] });
    data2.push({ range:`${S}!D${r}`, values:[[`=IFERROR(IF($B$21="","",INDEX('Master Task Log'!$F$6:$F$505,SMALL(IF('Master Task Log'!$D$6:$D$505=$B$21,ROW('Master Task Log'!$D$6:$D$505)-5),${n}))),"")` ]] });
    data2.push({ range:`${S}!E${r}`, values:[[`=IFERROR(IF($B$21="","",INDEX('Master Task Log'!$G$6:$G$505,SMALL(IF('Master Task Log'!$D$6:$D$505=$B$21,ROW('Master Task Log'!$D$6:$D$505)-5),${n}))),"")` ]] });
    data2.push({ range:`${S}!F${r}`, values:[[`=IFERROR(IF($B$21="","",INDEX('Master Task Log'!$J$6:$J$505,SMALL(IF('Master Task Log'!$D$6:$D$505=$B$21,ROW('Master Task Log'!$D$6:$D$505)-5),${n}))),"")` ]] });
    data2.push({ range:`${S}!G${r}`, values:[[`=IFERROR(IF($B$21="","",INDEX('Master Task Log'!$N$6:$N$505,SMALL(IF('Master Task Log'!$D$6:$D$505=$B$21,ROW('Master Task Log'!$D$6:$D$505)-5),${n}))),"")` ]] });
    data2.push({ range:`${S}!H${r}`, values:[[`=IFERROR(IF($B$21="","",INDEX('Master Task Log'!$K$6:$K$505,SMALL(IF('Master Task Log'!$D$6:$D$505=$B$21,ROW('Master Task Log'!$D$6:$D$505)-5),${n}))),"")` ]] });
    data2.push({ range:`${S}!I${r}`, values:[[`=IFERROR(IF(AND(F${r}<>"",F${r}<TODAY(),H${r}<>"Completed",H${r}<>"Cancelled"),"Yes","No"),"")` ]] });
  }

  await valuesBatchUpdate(id, data2, 'pp-values-2');

  // ── Formatting ──
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(PP,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };
  const pctF   = { type:'PERCENT', pattern:'0%' };

  // Title merge rows 0-1
  reqs.push({ mergeCells:{ range:gridRange(PP,0,2,0,17), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:18, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Subtitle rows 2-3
  reqs.push({ mergeCells:{ range:gridRange(PP,2,4,0,17), mergeType:'MERGE_ALL' } });
  fmt(2,4,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Controls row 4 (row index 3)
  fmt(3,4,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:10, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(3,4,1,4,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });

  // Section header: PROJECT SUMMARY TABLE row 6 (index 6)
  reqs.push({ mergeCells:{ range:gridRange(PP,6,7,0,17), mergeType:'MERGE_ALL' } });
  fmt(6,7,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Table header row 8 (index 7)
  fmt(7,8,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Data rows 9-18 (indices 8-17)
  for (let r = 8; r < 18; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,17,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols: A(0), C(2), D(3), E(4), F(5), G(6), H(7), I(8), J(9), K(10), L(11), M(12), N(13), O(14), P(15), Q(16)
  const formulaCols = [0,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16];
  for (const ci of formulaCols) {
    reqs.push({ repeatCell:{ range:gridRange(PP,8,18,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  // Input col B (1) — project name
  reqs.push({ repeatCell:{ range:gridRange(PP,8,18,1,2), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  // Date cols E(4), F(5)
  for (const ci of [4,5]) {
    reqs.push({ repeatCell:{ range:gridRange(PP,8,18,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Percentage col M (12)
  reqs.push({ repeatCell:{ range:gridRange(PP,8,18,12,13), cell:{ userEnteredFormat:{ numberFormat:pctF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  // Numbers centered
  for (const ci of [8,9,10,11,13,14,16]) {
    reqs.push({ repeatCell:{ range:gridRange(PP,8,18,ci,ci+1), cell:{ userEnteredFormat:{ horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat.horizontalAlignment' } });
  }
  reqs.push({ updateBorders:{ range:gridRange(PP,7,18,0,17), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Health conditional formatting
  const healthColors = [
    { val:'Completed', bg:C.success, fg:C.white },
    { val:'On Track',  bg:'#D4EDDA', fg:C.mainText },
    { val:'At Risk',   bg:C.warning, fg:C.mainText },
    { val:'Delayed',   bg:C.attention, fg:C.white },
    { val:'Not Started', bg:C.panel, fg:C.secText },
  ];
  healthColors.forEach(({ val, bg, fg }) => {
    reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(PP,8,18,15,16)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:val }] }, format:{ backgroundColor:hex(bg), textFormat:{ foregroundColor:hex(fg), bold:true } } } }, index:0 } });
  });

  // ── SELECTED PROJECT DETAIL section ──
  // Section header row 19 (index 19)
  reqs.push({ mergeCells:{ range:gridRange(PP,19,20,0,17), mergeType:'MERGE_ALL' } });
  fmt(19,20,0,17,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Picker row 20 (index 20)
  fmt(20,21,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ bold:true, fontSize:10, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  fmt(20,21,1,4,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });

  // Card rows 21-24 (indices 21-24) — label + value pairs
  const cardBg = C.panel;
  const labelFmt = { userEnteredFormat:{ backgroundColor:hex(C.lavender), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } };
  const valueFmt = { userEnteredFormat:{ backgroundColor:hex(cardBg), textFormat:{ fontSize:10, fontFamily:'Arial', foregroundColor:hex(C.primary), bold:true }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } };
  // 4 pairs per row, cols A-H (0-7)
  for (let rowIdx = 21; rowIdx <= 24; rowIdx++) {
    for (let pair = 0; pair < 4; pair++) {
      const c = pair * 2;
      reqs.push({ repeatCell:{ range:gridRange(PP,rowIdx,rowIdx+1,c,c+1), cell:labelFmt, fields:'userEnteredFormat' } });
      reqs.push({ repeatCell:{ range:gridRange(PP,rowIdx,rowIdx+1,c+1,c+2), cell:valueFmt, fields:'userEnteredFormat' } });
    }
  }
  // Date format for start/end in row 21 (D22=col3, F22=col5, H22=col7 → indices 21, cols 3,5,7)
  for (const ci of [5,7]) {
    reqs.push({ repeatCell:{ range:gridRange(PP,21,22,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:dateF } }, fields:'userEnteredFormat.numberFormat' } });
  }
  // Pct format for completion % in row 24 (D25 = col3)
  reqs.push({ repeatCell:{ range:gridRange(PP,24,25,3,4), cell:{ userEnteredFormat:{ numberFormat:pctF } }, fields:'userEnteredFormat.numberFormat' } });

  reqs.push({ updateBorders:{ range:gridRange(PP,21,25,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Task detail section header row 26 (index 26)
  reqs.push({ mergeCells:{ range:gridRange(PP,26,27,0,9), mergeType:'MERGE_ALL' } });
  fmt(26,27,0,9,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Task detail header row 27 (index 27)
  fmt(27,28,0,9,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Task detail data rows 28-47 (indices 28-47)
  for (let r = 28; r < 48; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,9,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols for detail table: B(1),C(2),D(3),E(4),F(5),G(6),H(7),I(8)
  for (const ci of [1,2,3,4,5,6,7,8]) {
    reqs.push({ repeatCell:{ range:gridRange(PP,28,48,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  // Date col F(5)
  reqs.push({ repeatCell:{ range:gridRange(PP,28,48,5,6), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  // Overdue col I(8) — conditional color
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(PP,28,48,8,9)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'Yes' }] }, format:{ backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true } } } }, index:0 } });

  reqs.push({ updateBorders:{ range:gridRange(PP,27,48,0,9), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Column widths
  [[0,80],[1,130],[2,120],[3,120],[4,100],[5,100],[6,90],[7,80],[8,80],[9,80],[10,80],[11,70],[12,85],[13,80],[14,80],[15,90],[16,75]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:PP, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  // Row heights
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PP, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:36 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PP, dimension:'ROWS', startIndex:2, endIndex:7 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PP, dimension:'ROWS', startIndex:7, endIndex:50 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Freeze 6 rows (title 0-1, subtitle 2-3, controls 3 row, section header 6 = index 6 is row 7)
  // Title merges: rows 0-1 and 2-3. Freeze after row 3 = frozenRowCount:4, but we also want header visible.
  // Keep it simple: freeze 8 rows to include through the table header
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:PP, gridProperties:{ frozenRowCount:8 } }, fields:'gridProperties.frozenRowCount' } });

  // Charts
  // Chart 1: Completion % by project — horizontal bar (col N=13 for chart data)
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Project Completion %',
    basicChart:{
      chartType:'BAR',
      legendPosition:'NO_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Completion %', format:{ bold:false } },
        { position:'LEFT_AXIS', title:'Project' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:1, endColumnIndex:2 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:12, endColumnIndex:13 }] } }, targetAxis:'BOTTOM_AXIS' }],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:PP, rowIndex:7, columnIndex:9 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:260 } } } } });

  // Chart 2: Tasks overview — grouped column (Total, Completed, In Progress, Overdue)
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Task Status by Project',
    basicChart:{
      chartType:'COLUMN',
      legendPosition:'BOTTOM_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Project' },
        { position:'LEFT_AXIS', title:'Task Count' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:1, endColumnIndex:2 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:8, endColumnIndex:9 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:9, endColumnIndex:10 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:10, endColumnIndex:11 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:11, endColumnIndex:12 }] } }, targetAxis:'LEFT_AXIS' },
      ],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:PP, rowIndex:7, columnIndex:13 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:260 } } } } });

  // Chart 3: Health distribution — pie chart
  // We'll build a mini summary in cols R-S (17-18) for health counts
  const healthVals = ['Completed','On Track','At Risk','Delayed','Not Started'];
  const data3 = [];
  data3.push({ range:`${S}!R8`, values:[['Health']] });
  data3.push({ range:`${S}!S8`, values:[['Count']] });
  healthVals.forEach((v, i) => {
    const r = 9 + i;
    data3.push({ range:`${S}!R${r}`, values:[[v]] });
    data3.push({ range:`${S}!S${r}`, values:[[`=COUNTIF($P$9:$P$18,"${v}")`]] });
  });
  await valuesBatchUpdate(id, data3, 'pp-values-3');

  reqs.push({ addChart:{ chart:{ spec:{
    title:'Project Health Distribution',
    pieChart:{
      legendPosition:'LABELED_LEGEND',
      domain:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:13, startColumnIndex:17, endColumnIndex:18 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:13, startColumnIndex:18, endColumnIndex:19 }] } },
      pieHole:0.4
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:PP, rowIndex:19, columnIndex:9 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:260 } } } } });

  // Chart 4: Est Hours vs Logged Hours — column chart
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Est. vs Logged Hours by Project',
    basicChart:{
      chartType:'COLUMN',
      legendPosition:'BOTTOM_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Project' },
        { position:'LEFT_AXIS', title:'Hours' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:1, endColumnIndex:2 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:13, endColumnIndex:14 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:PP, startRowIndex:7, endRowIndex:18, startColumnIndex:14, endColumnIndex:15 }] } }, targetAxis:'LEFT_AXIS' },
      ],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:PP, rowIndex:19, columnIndex:13 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:260 } } } } });

  await batchUpdate(id, reqs, 'pp-format');
  console.log('Project Progress complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
