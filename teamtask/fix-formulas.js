'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ML = sheetMap['Master Task Log'];
const DB = sheetMap['Team Dashboard'];
const MC = sheetMap['Monthly Calendar'];

// ── Lists used in validations ──────────────────────────────────────────────
const TASK_TYPES  = ['Task','Milestone','Meeting Action','Follow-Up','Review','Approval','Administrative','Recurring Responsibility'];
const PRIORITIES  = ['Low','Medium','High','Urgent'];
const EFFORTS     = ['Quick','Small','Medium','Large','Major'];
const YES_NO      = ['Yes','No'];

function listRule(values) {
  return { condition:{ type:'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: v })) }, showCustomUi:true, strict:true };
}

// ── Dashboard KPI bare-expression builders (no leading =) ──────────────────
// Control cells: B4=Year, F4=Team Member, H4=Department, J4=Project
// MTL columns: I=Due Date, K=Status, L=Priority, N=Est Hours, O=Actual Hours,
//              D=Project, E=Department, F=Assigned To, A=Task ID

function baseFilter() {
  return `(YEAR('Master Task Log'!$I$6:$I$505)=$B$4)*IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)`;
}
const totalTasksExpr    = `IFERROR(SUMPRODUCT(${baseFilter()}*('Master Task Log'!$A$6:$A$505<>"")),0)`;
const completedExpr     = `IFERROR(SUMPRODUCT(${baseFilter()}*('Master Task Log'!$K$6:$K$505="Completed")),0)`;
const inProgressExpr    = `IFERROR(SUMPRODUCT(${baseFilter()}*('Master Task Log'!$K$6:$K$505="In Progress")),0)`;
const overdueExpr       = `IFERROR(SUMPRODUCT(('Master Task Log'!$I$6:$I$505<>"")*('Master Task Log'!$I$6:$I$505<TODAY())*IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*('Master Task Log'!$K$6:$K$505<>"Completed")*('Master Task Log'!$K$6:$K$505<>"Cancelled")),0)`;
const urgentExpr        = `IFERROR(SUMPRODUCT(IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*('Master Task Log'!$L$6:$L$505="Urgent")*('Master Task Log'!$K$6:$K$505<>"Completed")*('Master Task Log'!$K$6:$K$505<>"Cancelled")),0)`;
// K8: completionRate — embeds bare exprs to avoid leading-= embedding bug
const completionRateExpr = `IFERROR(IF(${totalTasksExpr}=0,0,${completedExpr}/${totalTasksExpr}),0)`;
const blockedExpr       = `IFERROR(SUMPRODUCT(IF($F$4="All Team Members",1,'Master Task Log'!$F$6:$F$505=$F$4)*IF($H$4="All Departments",1,'Master Task Log'!$E$6:$E$505=$H$4)*IF($J$4="All Projects",1,'Master Task Log'!$D$6:$D$505=$J$4)*('Master Task Log'!$K$6:$K$505="Blocked")),0)`;
const estHoursExpr      = `IFERROR(SUMPRODUCT(${baseFilter()}*('Master Task Log'!$N$6:$N$505)),0)`;
const hrsLoggedExpr     = `IFERROR(SUMPRODUCT(${baseFilter()}*('Master Task Log'!$O$6:$O$505)),0)`;
const activeMembersExpr = `IFERROR(COUNTIF('Team Setup'!$G$20:$G$34,TRUE),0)`;
// K11: avgTasks — embeds bare exprs to avoid leading-= embedding bug
const avgTasksExpr      = `IFERROR(IF(${activeMembersExpr}=0,0,${totalTasksExpr}/${activeMembersExpr}),0)`;

// ── Deadline formula (fixed: $J$→$I$ for Due Date sort key and display col) ─
function deadlineFieldFixed(rank, col) {
  return `=IFERROR(INDEX('Master Task Log'!$${col}$6:$${col}$505,MATCH(SMALL(IF(('Master Task Log'!$K$6:$K$505<>"Completed")*('Master Task Log'!$K$6:$K$505<>"Cancelled")*('Master Task Log'!$I$6:$I$505<>""),'Master Task Log'!$I$6:$I$505-DATE(2000,1,1)),${rank}),IF(('Master Task Log'!$K$6:$K$505<>"Completed")*('Master Task Log'!$K$6:$K$505<>"Cancelled")*('Master Task Log'!$I$6:$I$505<>""),'Master Task Log'!$I$6:$I$505-DATE(2000,1,1)),0)),"")`;
}

const MEMBERS = ['Alex Morgan','Jamie Lee','Taylor Brooks','Jordan Patel','Casey Rivera','Morgan Chen','Riley Davis','Sam Thompson','Avery Walker','Quinn Martinez'];
const PROJECTS_LIST = ['Website Redesign','Product Launch','Client Onboarding','Monthly Reporting','Social Media Campaign','Team Operations','Event Planning','Process Improvement','Training Program','General Tasks'];

(async () => {
  // ── 1. Team Setup B13 — fix active member COUNTIF range ───────────────────
  // Bug: COUNTIF(J19:J33,TRUE) — Active? is col G, member rows are 20-34
  await valuesBatchUpdate(id, [
    { range: `'Team Setup'!B13`, values: [['=IFERROR(COUNTIF(G20:G34,TRUE),0)']] },
  ], 'fix-ts-b13');
  console.log('1/5 Team Setup B13 fixed');

  // ── 2. Fix validations: year dropdowns + wrong MTL columns ────────────────
  const dvReqs = [];
  // Clear year validation from Dashboard B4 (number 2026 fails string ONE_OF_LIST)
  dvReqs.push({ setDataValidation: { range: gridRange(DB, 3, 4, 1, 2) } });
  // Clear year validation from Monthly Calendar B3 (same issue)
  dvReqs.push({ setDataValidation: { range: gridRange(MC, 2, 3, 1, 2) } });
  // Clear wrong MLT validations: G(6)=Requested By got PRIORITIES, H(7)=Start Date got TASK_TYPES,
  //   S(18)=Overdue? formula got EFFORTS, T(19)=Progress% formula got RECURRENCE
  dvReqs.push({ setDataValidation: { range: gridRange(ML, 5, 505, 6, 7) } });
  dvReqs.push({ setDataValidation: { range: gridRange(ML, 5, 505, 7, 8) } });
  dvReqs.push({ setDataValidation: { range: gridRange(ML, 5, 505, 18, 19) } });
  dvReqs.push({ setDataValidation: { range: gridRange(ML, 5, 505, 19, 20) } });
  // Add correct MTL validations: C(2)=Task Type, L(11)=Priority, M(12)=Effort, P(15)=Recurring?
  dvReqs.push({ setDataValidation: { range: gridRange(ML, 5, 505, 2, 3),  rule: listRule(TASK_TYPES) } });
  dvReqs.push({ setDataValidation: { range: gridRange(ML, 5, 505, 11, 12), rule: listRule(PRIORITIES) } });
  dvReqs.push({ setDataValidation: { range: gridRange(ML, 5, 505, 12, 13), rule: listRule(EFFORTS) } });
  dvReqs.push({ setDataValidation: { range: gridRange(ML, 5, 505, 15, 16), rule: listRule(YES_NO) } });
  await batchUpdate(id, dvReqs, 'fix-validations');
  console.log('2/5 Validations fixed');

  // ── 3. Dashboard — fix 10 KPI cells ───────────────────────────────────────
  // Bugs: wrong control refs ($B$2→$B$4, $B$4→$F$4, $B$5→$H$4, $B$6→$J$4),
  //       wrong MTL cols ($J$→$I$ for Due Date, $G$→$L$ for Priority),
  //       formula embedding bug in K8 (#ERROR) and K11 (#ERROR)
  const kpiData = [
    { range: `'Team Dashboard'!A8`,  values: [[`=${totalTasksExpr}`]] },
    { range: `'Team Dashboard'!C8`,  values: [[`=${completedExpr}`]] },
    { range: `'Team Dashboard'!E8`,  values: [[`=${inProgressExpr}`]] },
    { range: `'Team Dashboard'!G8`,  values: [[`=${overdueExpr}`]] },
    { range: `'Team Dashboard'!I8`,  values: [[`=${urgentExpr}`]] },
    { range: `'Team Dashboard'!K8`,  values: [[`=${completionRateExpr}`]] },
    { range: `'Team Dashboard'!A11`, values: [[`=${blockedExpr}`]] },
    { range: `'Team Dashboard'!C11`, values: [[`=${estHoursExpr}`]] },
    { range: `'Team Dashboard'!E11`, values: [[`=${hrsLoggedExpr}`]] },
    { range: `'Team Dashboard'!K11`, values: [[`=${avgTasksExpr}`]] },
  ];
  await valuesBatchUpdate(id, kpiData, 'fix-kpi-cells');
  console.log('3/5 Dashboard KPI cells fixed');

  // ── 4. Dashboard — fix workload, deadlines, snapshot ─────────────────────
  const dashData = [];

  // Workload rows 15-24: memberOverdue ($J$→$I$) and memberUrgent ($G$→$L$)
  MEMBERS.forEach((name, i) => {
    const r = 15 + i;
    dashData.push({ range: `'Team Dashboard'!E${r}`, values: [[
      `=IFERROR(COUNTIFS('Master Task Log'!$F$6:$F$505,"${name}",'Master Task Log'!$I$6:$I$505,"<"&TODAY(),'Master Task Log'!$K$6:$K$505,"<>Completed",'Master Task Log'!$K$6:$K$505,"<>Cancelled"),0)`,
    ]] });
    dashData.push({ range: `'Team Dashboard'!F${r}`, values: [[
      `=IFERROR(COUNTIFS('Master Task Log'!$F$6:$F$505,"${name}",'Master Task Log'!$L$6:$L$505,"Urgent",'Master Task Log'!$K$6:$K$505,"<>Completed",'Master Task Log'!$K$6:$K$505,"<>Cancelled"),0)`,
    ]] });
  });

  // Deadlines rows 28-37: fix sort key $J$→$I$, due date col J→I, task name col C→B
  for (let rank = 1; rank <= 10; rank++) {
    const r = 27 + rank;
    dashData.push({ range: `'Team Dashboard'!B${r}`, values: [[deadlineFieldFixed(rank, 'A')]] }); // Task ID
    dashData.push({ range: `'Team Dashboard'!C${r}`, values: [[deadlineFieldFixed(rank, 'B')]] }); // Task Name (was 'C'=Task Type, fixed to 'B'=Task Name)
    dashData.push({ range: `'Team Dashboard'!D${r}`, values: [[deadlineFieldFixed(rank, 'F')]] }); // Assigned To
    dashData.push({ range: `'Team Dashboard'!E${r}`, values: [[deadlineFieldFixed(rank, 'D')]] }); // Project
    dashData.push({ range: `'Team Dashboard'!F${r}`, values: [[deadlineFieldFixed(rank, 'I')]] }); // Due Date (was 'J'=Completed Date, fixed to 'I'=Due Date)
    dashData.push({ range: `'Team Dashboard'!G${r}`, values: [[deadlineFieldFixed(rank, 'K')]] }); // Status
  }

  // Snapshot rows 41-50: snapOverdue ($J$→$I$)
  PROJECTS_LIST.forEach((proj, i) => {
    const r = 41 + i;
    dashData.push({ range: `'Team Dashboard'!D${r}`, values: [[
      `=IFERROR(COUNTIFS('Master Task Log'!$D$6:$D$505,"${proj}",'Master Task Log'!$I$6:$I$505,"<"&TODAY(),'Master Task Log'!$K$6:$K$505,"<>Completed",'Master Task Log'!$K$6:$K$505,"<>Cancelled"),0)`,
    ]] });
  });

  await valuesBatchUpdate(id, dashData, 'fix-dashboard-workload-deadlines-snapshot');
  console.log('4/5 Dashboard workload, deadlines, snapshot fixed');

  console.log('5/5 All fixes complete');
  console.log('Spreadsheet:', `https://docs.google.com/spreadsheets/d/${id}/edit`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
