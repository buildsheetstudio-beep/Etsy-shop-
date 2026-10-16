'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Team Dashboard'];
const S = "'Team Dashboard'";

// Controls: B2=Year, B3=Period (Month/Quarter/All Time), B4=Team Member, B5=Department, B6=Project
// KPI cards row 8-9 (6 primary), row 11-12 (6 secondary)
// Team workload table rows 15-25
// Upcoming deadlines rows 28-38
// Project snapshot rows 41-51
// Charts overlay

// ── KPI formulas ──
// All formulas filter by year (col B in Master Task Log? No - we use Due Date col J year)
// Controls: B2=Year, B3=Period ("All Time","Q1","Q2","Q3","Q4","January",...,"December")
// B4=Team Member ("All Team Members" or name)
// B5=Department ("All Departments" or dept)
// B6=Project ("All Projects" or project)

// Helper: SUMPRODUCT filter base
// MTL cols: A=TaskID, B=Task#, C=TaskName, D=Project, E=Department, F=AssignedTo,
//           G=Priority, H=TaskType, I=StartDate, J=DueDate, K=Status, L=CompletedDate,
//           M=DaysRemaining, N=EstHrs, O=HrsLogged, P=OverdueFlag, Q=RecID, R=Tags, ...V=Notes

function kpiTotalTasks() {
  return `=IFERROR(SUMPRODUCT(
(YEAR('Master Task Log'!$J$6:$J$505)=$B$4)*
IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*
IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*
IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*
('Master Task Log'!$A$6:$A$505<>"")),0)`.replace(/\n/g,'');
}

function kpiCompleted() {
  return `=IFERROR(SUMPRODUCT(
(YEAR('Master Task Log'!$J$6:$J$505)=$B$4)*
IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*
IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*
IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*
('Master Task Log'!$K$6:$K$505="Completed")),0)`.replace(/\n/g,'');
}

function kpiInProgress() {
  return `=IFERROR(SUMPRODUCT(
(YEAR('Master Task Log'!$J$6:$J$505)=$B$4)*
IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*
IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*
IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*
('Master Task Log'!$K$6:$K$505="In Progress")),0)`.replace(/\n/g,'');
}

function kpiOverdue() {
  return `=IFERROR(SUMPRODUCT(
('Master Task Log'!$J$6:$J$505<>"")*
('Master Task Log'!$J$6:$J$505<TODAY())*
IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*
IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*
IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*
('Master Task Log'!$K$6:$K$505<>"Completed")*
('Master Task Log'!$K$6:$K$505<>"Cancelled")),0)`.replace(/\n/g,'');
}

function kpiUrgent() {
  return `=IFERROR(SUMPRODUCT(
IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*
IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*
IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*
('Master Task Log'!$G$6:$G$505="Urgent")*
('Master Task Log'!$K$6:$K$505<>"Completed")*
('Master Task Log'!$K$6:$K$505<>"Cancelled")),0)`.replace(/\n/g,'');
}

function kpiCompletionRate() {
  return `=IFERROR(IF(A8=0,0,C8/A8),0)`;
}

function kpiBlocked() {
  return `=IFERROR(SUMPRODUCT(
IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*
IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*
IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*
('Master Task Log'!$K$6:$K$505="Blocked")),0)`.replace(/\n/g,'');
}

function kpiEstHours() {
  return `=IFERROR(SUMPRODUCT(
(YEAR('Master Task Log'!$J$6:$J$505)=$B$4)*
IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*
IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*
IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*
('Master Task Log'!$N$6:$N$505)),0)`.replace(/\n/g,'');
}

function kpiHrsLogged() {
  return `=IFERROR(SUMPRODUCT(
(YEAR('Master Task Log'!$J$6:$J$505)=$B$4)*
IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*
IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*
IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*
('Master Task Log'!$O$6:$O$505)),0)`.replace(/\n/g,'');
}

function kpiActiveProjects() {
  return `=IFERROR(COUNTIF('Team Setup'!$G$30:$G$44,"In Progress"),0)`;
}

function kpiActiveMembers() {
  return `=IFERROR(COUNTIF('Team Setup'!$G$20:$G$34,TRUE),0)`;
}

// Team workload table (rows 15-25, 10 members + 1 total)
// Members from Team Setup B20:B34
// Cols: A=Member, B=Total Assigned, C=Completed, D=In Progress, E=Overdue, F=Urgent, G=Est Hrs, H=Completion%, I=Capacity, J=Load Status
const MEMBERS = ['Alex Morgan','Jamie Lee','Taylor Brooks','Jordan Patel','Casey Rivera','Morgan Chen','Riley Davis','Sam Thompson','Avery Walker','Quinn Martinez'];

function memberTotal(name) {
  return `=IFERROR(COUNTIF('Master Task Log'!$F$6:$F$505,"${name}"),0)`;
}
function memberCompleted(name) {
  return `=IFERROR(COUNTIFS('Master Task Log'!$F$6:$F$505,"${name}",'Master Task Log'!$K$6:$K$505,"Completed"),0)`;
}
function memberInProgress(name) {
  return `=IFERROR(COUNTIFS('Master Task Log'!$F$6:$F$505,"${name}",'Master Task Log'!$K$6:$K$505,"In Progress"),0)`;
}
function memberOverdue(name) {
  return `=IFERROR(COUNTIFS('Master Task Log'!$F$6:$F$505,"${name}",'Master Task Log'!$J$6:$J$505,"<"&TODAY(),'Master Task Log'!$K$6:$K$505,"<>Completed",'Master Task Log'!$K$6:$K$505,"<>Cancelled"),0)`;
}
function memberUrgent(name) {
  return `=IFERROR(COUNTIFS('Master Task Log'!$F$6:$F$505,"${name}",'Master Task Log'!$G$6:$G$505,"Urgent",'Master Task Log'!$K$6:$K$505,"<>Completed",'Master Task Log'!$K$6:$K$505,"<>Cancelled"),0)`;
}
function memberEstHrs(name) {
  return `=IFERROR(SUMIF('Master Task Log'!$F$6:$F$505,"${name}",'Master Task Log'!$N$6:$N$505),0)`;
}
function memberCompPct(totalCell, completedCell) {
  return `=IFERROR(IF(${totalCell}=0,0,${completedCell}/${totalCell}),0)`;
}
function memberCapacity(name) {
  return `=IFERROR(INDEX('Team Setup'!$F$20:$F$34,MATCH("${name}",'Team Setup'!$B$20:$B$34,0)),20)`;
}
function memberLoadStatus(inProgCell, capCell) {
  // In Progress tasks vs capacity
  return `=IFERROR(IF(${capCell}=0,"—",IF(${inProgCell}/${capCell}>1,"Over Capacity",IF(${inProgCell}/${capCell}>=0.8,"Near Capacity",IF(${inProgCell}/${capCell}>=0.5,"Balanced","Available")))),"—")`;
}

// Upcoming deadlines (rows 28-37, 10 tasks due soonest from today, not completed/cancelled)
function deadlineField(rank, col) {
  return `=IFERROR(INDEX('Master Task Log'!$${col}$6:$${col}$505,MATCH(SMALL(IF(('Master Task Log'!$K$6:$K$505<>"Completed")*('Master Task Log'!$K$6:$K$505<>"Cancelled")*('Master Task Log'!$J$6:$J$505<>""),'Master Task Log'!$J$6:$J$505-DATE(2000,1,1)),${rank}),IF(('Master Task Log'!$K$6:$K$505<>"Completed")*('Master Task Log'!$K$6:$K$505<>"Cancelled")*('Master Task Log'!$J$6:$J$505<>""),'Master Task Log'!$J$6:$J$505-DATE(2000,1,1)),0)),"")`;
}

// Project snapshot (rows 41-50, 10 projects)
const PROJECTS_LIST = [
  'Website Redesign','Product Launch','Client Onboarding','Monthly Reporting',
  'Social Media Campaign','Team Operations','Event Planning','Process Improvement',
  'Training Program','General Tasks'
];
function snapTotal(proj) { return `=IFERROR(COUNTIF('Master Task Log'!$D$6:$D$505,"${proj}"),0)`; }
function snapCompleted(proj) { return `=IFERROR(COUNTIFS('Master Task Log'!$D$6:$D$505,"${proj}",'Master Task Log'!$K$6:$K$505,"Completed"),0)`; }
function snapOverdue(proj) { return `=IFERROR(COUNTIFS('Master Task Log'!$D$6:$D$505,"${proj}",'Master Task Log'!$J$6:$J$505,"<"&TODAY(),'Master Task Log'!$K$6:$K$505,"<>Completed",'Master Task Log'!$K$6:$K$505,"<>Cancelled"),0)`; }
function snapHealth(totalCell, completedCell, overdueCell, projRow) {
  return `=IFERROR(IF(${totalCell}=0,"—",IF(${completedCell}=${totalCell},"Completed",IF(${overdueCell}>0,"Delayed",IF(${completedCell}/${totalCell}<0.3,"At Risk","On Track")))),"—")`;
}
function snapPct(totalCell, completedCell) { return `=IFERROR(IF(${totalCell}=0,0,${completedCell}/${totalCell}),0)`; }

(async () => {
  // Pre-pass: read existing merges for this sheet and unmerge any in the top 6 rows
  const { google: goog } = require('googleapis');
  const fs2 = require('fs');
  const s2 = JSON.parse(fs2.readFileSync(__dirname + '/client_secret.json'));
  const auth2 = new goog.auth.OAuth2(s2.installed.client_id, s2.installed.client_secret, s2.installed.redirect_uris[0]);
  auth2.setCredentials(JSON.parse(fs2.readFileSync(__dirname + '/tokens.json')));
  const sheets2 = goog.sheets({ version: 'v4', auth: auth2 });
  const spreadsheetInfo = await sheets2.spreadsheets.get({ spreadsheetId:id, fields:'sheets(properties.sheetId,merges)' });
  const dbSheet = (spreadsheetInfo.data.sheets || []).find(s => s.properties && s.properties.sheetId === DB);
  const existingMerges = (dbSheet && dbSheet.merges) ? dbSheet.merges.filter(m => m.startRowIndex < 6) : [];
  if (existingMerges.length > 0) {
    await sheets2.spreadsheets.batchUpdate({
      spreadsheetId:id,
      requestBody:{ requests: existingMerges.map(range => ({ unmergeCells:{ range } })) }
    });
  }

  const data = [];

  // Title
  data.push({ range:`${S}!A1`, values:[['TEAM DASHBOARD']] });
  data.push({ range:`${S}!A2`, values:[['Your real-time command center for team task management. Use the controls below to filter by year, period, member, department, or project.']] });

  // Controls rows 4-6 (labels col A, values col B)
  data.push({ range:`${S}!A4`, values:[['Year:']] });
  data.push({ range:`${S}!B4`, values:[[2026]] });
  data.push({ range:`${S}!C4`, values:[['Period:']] });
  data.push({ range:`${S}!D4`, values:[['All Time']] });
  data.push({ range:`${S}!E4`, values:[['Team Member:']] });
  data.push({ range:`${S}!F4`, values:[['All Team Members']] });
  data.push({ range:`${S}!G4`, values:[['Department:']] });
  data.push({ range:`${S}!H4`, values:[['All Departments']] });
  data.push({ range:`${S}!I4`, values:[['Project:']] });
  data.push({ range:`${S}!J4`, values:[['All Projects']] });

  // Section header row 6 (KPIs)
  data.push({ range:`${S}!A6`, values:[['KEY PERFORMANCE INDICATORS']] });

  // KPI primary cards row 7-8 (6 cards, 2 cols each: label row 7, value row 8)
  const kpiPrimary = [
    ['Total Tasks',kpiTotalTasks()],
    ['Completed',kpiCompleted()],
    ['In Progress',kpiInProgress()],
    ['Overdue',kpiOverdue()],
    ['Urgent',kpiUrgent()],
    ['Completion Rate',kpiCompletionRate()],
  ];
  kpiPrimary.forEach(([label, formula], i) => {
    const col = String.fromCharCode(65 + i*2); // A,C,E,G,I,K
    const col2 = String.fromCharCode(65 + i*2 + 1);
    data.push({ range:`${S}!${col}7:${col2}7`, values:[[label,'']] });
    data.push({ range:`${S}!${col}8:${col2}8`, values:[[formula,'']] });
  });

  // KPI secondary cards row 10-11
  const kpiSecondary = [
    ['Blocked',kpiBlocked()],
    ['Est. Hours',kpiEstHours()],
    ['Hours Logged',kpiHrsLogged()],
    ['Active Projects',kpiActiveProjects()],
    ['Active Members',kpiActiveMembers()],
    ['Avg Tasks/Member',`=IFERROR(IF(I11=0,0,A8/I11),0)`],
  ];
  kpiSecondary.forEach(([label, formula], i) => {
    const col = String.fromCharCode(65 + i*2);
    const col2 = String.fromCharCode(65 + i*2 + 1);
    data.push({ range:`${S}!${col}10:${col2}10`, values:[[label,'']] });
    data.push({ range:`${S}!${col}11:${col2}11`, values:[[formula,'']] });
  });

  // Team workload section header row 13
  data.push({ range:`${S}!A13`, values:[['TEAM WORKLOAD']] });

  // Workload table header row 14
  data.push({ range:`${S}!A14:J14`, values:[[
    'Team Member','Total Tasks','Completed','In Progress','Overdue','Urgent','Est. Hours','Completion %','Capacity','Load Status'
  ]] });

  // Workload rows 15-24
  MEMBERS.forEach((name, i) => {
    const r = 15 + i;
    const totalCell = `B${r}`;
    const compCell = `C${r}`;
    const inProgCell = `D${r}`;
    const capCell = `I${r}`;
    data.push({ range:`${S}!A${r}`, values:[[name]] });
    data.push({ range:`${S}!B${r}`, values:[[memberTotal(name)]] });
    data.push({ range:`${S}!C${r}`, values:[[memberCompleted(name)]] });
    data.push({ range:`${S}!D${r}`, values:[[memberInProgress(name)]] });
    data.push({ range:`${S}!E${r}`, values:[[memberOverdue(name)]] });
    data.push({ range:`${S}!F${r}`, values:[[memberUrgent(name)]] });
    data.push({ range:`${S}!G${r}`, values:[[memberEstHrs(name)]] });
    data.push({ range:`${S}!H${r}`, values:[[memberCompPct(totalCell, compCell)]] });
    data.push({ range:`${S}!I${r}`, values:[[memberCapacity(name)]] });
    data.push({ range:`${S}!J${r}`, values:[[memberLoadStatus(inProgCell, capCell)]] });
  });

  // Upcoming deadlines section header row 26
  data.push({ range:`${S}!A26`, values:[['UPCOMING DEADLINES']] });

  // Deadlines header row 27
  data.push({ range:`${S}!A27:G27`, values:[[
    '#','Task ID','Task Name','Assigned To','Project','Due Date','Status'
  ]] });

  // Deadline rows 28-37
  for (let rank = 1; rank <= 10; rank++) {
    const r = 27 + rank;
    data.push({ range:`${S}!A${r}`, values:[[rank]] });
    data.push({ range:`${S}!B${r}`, values:[[deadlineField(rank, 'A')]] });
    data.push({ range:`${S}!C${r}`, values:[[deadlineField(rank, 'C')]] });
    data.push({ range:`${S}!D${r}`, values:[[deadlineField(rank, 'F')]] });
    data.push({ range:`${S}!E${r}`, values:[[deadlineField(rank, 'D')]] });
    data.push({ range:`${S}!F${r}`, values:[[deadlineField(rank, 'J')]] });
    data.push({ range:`${S}!G${r}`, values:[[deadlineField(rank, 'K')]] });
  }

  // Project snapshot section header row 39
  data.push({ range:`${S}!A39`, values:[['PROJECT SNAPSHOT']] });

  // Snapshot header row 40
  data.push({ range:`${S}!A40:F40`, values:[[
    'Project','Total Tasks','Completed','Overdue','Completion %','Health'
  ]] });

  // Snapshot rows 41-50
  PROJECTS_LIST.forEach((proj, i) => {
    const r = 41 + i;
    const totalCell = `B${r}`;
    const compCell = `C${r}`;
    const overdueCell = `D${r}`;
    data.push({ range:`${S}!A${r}`, values:[[proj]] });
    data.push({ range:`${S}!B${r}`, values:[[snapTotal(proj)]] });
    data.push({ range:`${S}!C${r}`, values:[[snapCompleted(proj)]] });
    data.push({ range:`${S}!D${r}`, values:[[snapOverdue(proj)]] });
    data.push({ range:`${S}!E${r}`, values:[[snapPct(totalCell, compCell)]] });
    data.push({ range:`${S}!F${r}`, values:[[snapHealth(totalCell, compCell, overdueCell, r)]] });
  });

  // Last updated row 52
  data.push({ range:`${S}!A52`, values:[['Dashboard refreshes automatically when the spreadsheet is opened or recalculated. Data reflects the Master Task Log.']] });

  await valuesBatchUpdate(id, data, 'dashboard-values');

  // ── Formatting ──
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(DB,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };
  const pctF   = { type:'PERCENT', pattern:'0%' };

  // Title row 0
  reqs.push({ mergeCells:{ range:gridRange(DB,0,1,0,12), mergeType:'MERGE_ALL' } });
  fmt(0,1,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:22, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Subtitle row 1
  reqs.push({ mergeCells:{ range:gridRange(DB,1,2,0,12), mergeType:'MERGE_ALL' } });
  fmt(1,2,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Controls row 3 (index 3)
  // 5 label+value pairs across A-J
  for (let i = 0; i < 5; i++) {
    const c = i * 2;
    fmt(3,4,c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'RIGHT', verticalAlignment:'MIDDLE' } });
    fmt(3,4,c+1,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.input), textFormat:{ fontSize:10, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }

  // KPI primary section header row 5 (index 5)
  reqs.push({ mergeCells:{ range:gridRange(DB,5,6,0,12), mergeType:'MERGE_ALL' } });
  fmt(5,6,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Primary KPI cards rows 6-7 (labels row 6, values row 7)
  for (let i = 0; i < 6; i++) {
    const c = i * 2;
    reqs.push({ mergeCells:{ range:gridRange(DB,6,7,c,c+2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(DB,7,8,c,c+2), mergeType:'MERGE_ALL' } });
    fmt(6,7,c,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.lavender), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(7,8,c,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:20, fontFamily:'Arial', foregroundColor:hex(C.primary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  }
  // Completion rate is a pct
  reqs.push({ repeatCell:{ range:gridRange(DB,7,8,10,12), cell:{ userEnteredFormat:{ numberFormat:pctF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(DB,6,8,0,12), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Secondary KPI cards rows 9-10 (labels row 9, values row 10)
  // blank row 8 = index 8 separator
  fmt(8,9,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.bg) } });
  for (let i = 0; i < 6; i++) {
    const c = i * 2;
    reqs.push({ mergeCells:{ range:gridRange(DB,9,10,c,c+2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(DB,10,11,c,c+2), mergeType:'MERGE_ALL' } });
    fmt(9,10,c,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.mutedBlue), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(10,11,c,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:16, fontFamily:'Arial', foregroundColor:hex(C.secondary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  }
  reqs.push({ updateBorders:{ range:gridRange(DB,9,11,0,12), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Team Workload section header row 12 (index 12)
  reqs.push({ mergeCells:{ range:gridRange(DB,12,13,0,10), mergeType:'MERGE_ALL' } });
  fmt(12,13,0,10,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Workload header row 13 (index 13)
  fmt(13,14,0,10,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Workload rows 14-23 (indices)
  for (let r = 14; r < 24; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,10,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols B-J (1-9) for workload
  for (const ci of [1,2,3,4,5,6,7,8,9]) {
    reqs.push({ repeatCell:{ range:gridRange(DB,14,24,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Pct col H (7)
  reqs.push({ repeatCell:{ range:gridRange(DB,14,24,7,8), cell:{ userEnteredFormat:{ numberFormat:pctF } }, fields:'userEnteredFormat.numberFormat' } });
  // Load status conditional formatting col J (9)
  const loadColors = [
    { val:'Over Capacity', bg:C.attention, fg:C.white },
    { val:'Near Capacity', bg:C.warning, fg:C.mainText },
    { val:'Balanced', bg:C.success, fg:C.white },
    { val:'Available', bg:'#D4EDDA', fg:C.mainText },
  ];
  loadColors.forEach(({ val, bg, fg }) => {
    reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(DB,14,24,9,10)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:val }] }, format:{ backgroundColor:hex(bg), textFormat:{ foregroundColor:hex(fg), bold:true } } } }, index:0 } });
  });
  reqs.push({ updateBorders:{ range:gridRange(DB,13,24,0,10), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Upcoming deadlines section header row 25 (index 25)
  reqs.push({ mergeCells:{ range:gridRange(DB,25,26,0,7), mergeType:'MERGE_ALL' } });
  fmt(25,26,0,7,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Deadlines header row 26 (index 26)
  fmt(26,27,0,7,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Deadline rows 27-36 (indices)
  for (let r = 27; r < 37; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,7,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols B-G (1-6) for deadlines
  for (const ci of [1,2,3,4,5,6]) {
    reqs.push({ repeatCell:{ range:gridRange(DB,27,37,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  // Due date col F (5)
  reqs.push({ repeatCell:{ range:gridRange(DB,27,37,5,6), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  // Overdue deadline highlight: due < today and not completed
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(DB,27,37,5,6)], booleanRule:{ condition:{ type:'DATE_BEFORE', values:[{ relativeDate:'TODAY' }] }, format:{ textFormat:{ foregroundColor:hex(C.attention), bold:true } } } }, index:0 } });
  reqs.push({ updateBorders:{ range:gridRange(DB,26,37,0,7), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Project snapshot section header row 38 (index 38)
  reqs.push({ mergeCells:{ range:gridRange(DB,38,39,0,6), mergeType:'MERGE_ALL' } });
  fmt(38,39,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Snapshot header row 39 (index 39)
  fmt(39,40,0,6,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Snapshot data rows 40-49 (indices)
  for (let r = 40; r < 50; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,6,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols for snapshot B-F (1-5)
  for (const ci of [1,2,3,4,5]) {
    reqs.push({ repeatCell:{ range:gridRange(DB,40,50,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Pct col E (4)
  reqs.push({ repeatCell:{ range:gridRange(DB,40,50,4,5), cell:{ userEnteredFormat:{ numberFormat:pctF } }, fields:'userEnteredFormat.numberFormat' } });
  // Health col F (5)
  const healthColors2 = [
    { val:'Completed', bg:C.success, fg:C.white },
    { val:'On Track',  bg:'#D4EDDA', fg:C.mainText },
    { val:'At Risk',   bg:C.warning, fg:C.mainText },
    { val:'Delayed',   bg:C.attention, fg:C.white },
  ];
  healthColors2.forEach(({ val, bg, fg }) => {
    reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(DB,40,50,5,6)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:val }] }, format:{ backgroundColor:hex(bg), textFormat:{ foregroundColor:hex(fg), bold:true } } } }, index:0 } });
  });
  reqs.push({ updateBorders:{ range:gridRange(DB,39,50,0,6), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Footer note row 51 (index 51)
  reqs.push({ mergeCells:{ range:gridRange(DB,51,52,0,12), mergeType:'MERGE_ALL' } });
  fmt(51,52,0,12,{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:8, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });

  // Column widths
  [[0,140],[1,80],[2,80],[3,80],[4,70],[5,70],[6,80],[7,80],[8,80],[9,110],[10,80],[11,80]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  // Row heights
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:44 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:2, endIndex:4 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:5, endIndex:6 }, properties:{ pixelSize:26 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:6, endIndex:8 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:7, endIndex:8 }, properties:{ pixelSize:48 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:8, endIndex:9 }, properties:{ pixelSize:10 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:9, endIndex:10 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:10, endIndex:11 }, properties:{ pixelSize:44 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:11, endIndex:55 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Freeze: title merges rows 0-3 (2 merges), controls row 3 = 4 rows total
  // But we want header sticky so freeze through row 5 (section header)
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:DB, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' } });

  // Charts
  // Chart 1: Completion % by team member — horizontal bar
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Completion % by Team Member',
    basicChart:{
      chartType:'BAR',
      legendPosition:'NO_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Completion %' },
        { position:'LEFT_AXIS', title:'Team Member' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:13, endRowIndex:24, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:13, endRowIndex:24, startColumnIndex:7, endColumnIndex:8 }] } }, targetAxis:'BOTTOM_AXIS' }],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:13, columnIndex:10 }, offsetXPixels:0, offsetYPixels:0, widthPixels:400, heightPixels:280 } } } } });

  // Chart 2: Task count by status — donut/pie
  // Build a mini pivot in cols M-N (12-13) rows 6-12
  const statusLabels = ['Not Started','In Progress','Waiting','Blocked','In Review','Completed','Cancelled'];
  const data2 = [];
  data2.push({ range:`${S}!M6`, values:[['Status']] });
  data2.push({ range:`${S}!N6`, values:[['Count']] });
  statusLabels.forEach((s, i) => {
    const r = 7 + i;
    data2.push({ range:`${S}!M${r}`, values:[[s]] });
    data2.push({ range:`${S}!N${r}`, values:[[`=IFERROR(COUNTIF('Master Task Log'!$K$6:$K$505,"${s}"),0)`]] });
  });
  await valuesBatchUpdate(id, data2, 'dashboard-pivot');

  reqs.push({ addChart:{ chart:{ spec:{
    title:'Tasks by Status',
    pieChart:{
      legendPosition:'LABELED_LEGEND',
      domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:5, endRowIndex:13, startColumnIndex:12, endColumnIndex:13 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:5, endRowIndex:13, startColumnIndex:13, endColumnIndex:14 }] } },
      pieHole:0.4
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:25, columnIndex:7 }, offsetXPixels:0, offsetYPixels:0, widthPixels:400, heightPixels:280 } } } } });

  // Chart 3: Overdue tasks by team member — column
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Overdue Tasks by Member',
    basicChart:{
      chartType:'COLUMN',
      legendPosition:'NO_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Team Member' },
        { position:'LEFT_AXIS', title:'Overdue Tasks' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:13, endRowIndex:24, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:13, endRowIndex:24, startColumnIndex:4, endColumnIndex:5 }] } }, targetAxis:'LEFT_AXIS' }],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:25, columnIndex:10 }, offsetXPixels:0, offsetYPixels:0, widthPixels:400, heightPixels:280 } } } } });

  // Chart 4: Project completion % — horizontal bar
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Project Completion %',
    basicChart:{
      chartType:'BAR',
      legendPosition:'NO_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Completion %' },
        { position:'LEFT_AXIS', title:'Project' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:39, endRowIndex:50, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:39, endRowIndex:50, startColumnIndex:4, endColumnIndex:5 }] } }, targetAxis:'BOTTOM_AXIS' }],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:39, columnIndex:7 }, offsetXPixels:0, offsetYPixels:0, widthPixels:400, heightPixels:280 } } } } });

  // Chart 5: Tasks by priority (stacked) — column
  // Mini pivot in cols O-R (14-17)
  const priorityLabels = ['Low','Medium','High','Urgent'];
  const data3 = [];
  data3.push({ range:`${S}!O6`, values:[['Priority']] });
  data3.push({ range:`${S}!P6`, values:[['Count']] });
  priorityLabels.forEach((p, i) => {
    const r = 7 + i;
    data3.push({ range:`${S}!O${r}`, values:[[p]] });
    data3.push({ range:`${S}!P${r}`, values:[[`=IFERROR(COUNTIFS('Master Task Log'!$G$6:$G$505,"${p}",'Master Task Log'!$K$6:$K$505,"<>Completed",'Master Task Log'!$K$6:$K$505,"<>Cancelled"),0)`]] });
  });
  await valuesBatchUpdate(id, data3, 'dashboard-pivot2');

  reqs.push({ addChart:{ chart:{ spec:{
    title:'Open Tasks by Priority',
    pieChart:{
      legendPosition:'LABELED_LEGEND',
      domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:5, endRowIndex:10, startColumnIndex:14, endColumnIndex:15 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:5, endRowIndex:10, startColumnIndex:15, endColumnIndex:16 }] } },
      pieHole:0
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:39, columnIndex:10 }, offsetXPixels:0, offsetYPixels:0, widthPixels:400, heightPixels:280 } } } } });

  // Chart 6: Est vs logged hours by member
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Est. vs Logged Hours by Member',
    basicChart:{
      chartType:'COLUMN',
      legendPosition:'BOTTOM_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Team Member' },
        { position:'LEFT_AXIS', title:'Hours' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:13, endRowIndex:24, startColumnIndex:0, endColumnIndex:1 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:13, endRowIndex:24, startColumnIndex:6, endColumnIndex:7 }] } }, targetAxis:'LEFT_AXIS' },
      ],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:13, columnIndex:7 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:260 } } } } });

  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('Team Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
