'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ML = sheetMap['Master Task Log'];
const RP = sheetMap['Recurring Task Planner'];
const WC = sheetMap['Weekly Calendar'];
const MC = sheetMap['Monthly Calendar'];
const PP = sheetMap['Project Progress'];
const DB = sheetMap['Team Dashboard'];

function listRule(values) {
  return { condition:{ type:'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue:v })) }, showCustomUi:true, strict:true };
}
function rangeRule(sheetName, col, startRow, endRow) {
  return { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`='${sheetName}'!$${col}$${startRow}:$${col}$${endRow}` }] }, showCustomUi:true, strict:false };
}

const STATUSES = ['Not Started','In Progress','Waiting','Blocked','In Review','Completed','Cancelled'];
const PRIORITIES = ['Low','Medium','High','Urgent'];
const EFFORTS = ['Quick','Small','Medium','Large','Major'];
const TASK_TYPES = ['Task','Milestone','Meeting Action','Follow-Up','Review','Approval','Administrative','Recurring Responsibility'];
const RECURRENCE = ['None','Daily','Weekdays','Weekly','Every 2 Weeks','Monthly','Quarterly','Yearly'];
const YES_NO = ['Yes','No'];
const PERIODS = ['All Time','January','February','March','April','May','June','July','August','September','October','November','December','Q1','Q2','Q3','Q4'];
const YEARS = ['2024','2025','2026','2027','2028'];
const HEALTH = ['Not Started','On Track','At Risk','Delayed','Completed'];
const PROJECTS = ['Website Redesign','Product Launch','Client Onboarding','Monthly Reporting','Social Media Campaign','Team Operations','Event Planning','Process Improvement','Training Program','General Tasks','All Projects'];
const DEPARTMENTS = ['Leadership','Operations','Marketing','Sales','Finance','Customer Service','Human Resources','Product','Creative','Administration','General Team','All Departments'];
const TEAM_MEMBERS = ['Alex Morgan','Jamie Lee','Taylor Brooks','Jordan Patel','Casey Rivera','Morgan Chen','Riley Davis','Sam Thompson','Avery Walker','Quinn Martinez','Member 11','Member 12','Member 13','Member 14','Member 15','All Team Members'];

(async () => {
  const reqs = [];

  const addValidation = (sheetId, r1, r2, c1, c2, rule) => {
    reqs.push({ setDataValidation:{ range:gridRange(sheetId,r1,r2,c1,c2), rule } });
  };

  // ── Master Task Log (cols 0-indexed, data rows 5-504) ──
  // C=Project(3), D=Dept(4), E=Dept? No — let's check column order:
  // A=TaskID(0), B=TaskNum(1), C=TaskName(2), D=Project(3), E=Department(4),
  // F=AssignedTo(5), G=Priority(6), H=TaskType(7), I=StartDate(8), J=DueDate(9),
  // K=Status(10), L=CompletedDate(11), M=DaysRemaining(12), N=EstHrs(13),
  // O=HrsLogged(14), P=OverdueFlag(15), Q=RecID(16), R=Tags(17), S=Notes(18)...

  // Project col D (index 3)
  addValidation(ML,5,505,3,4, listRule(PROJECTS.filter(p=>p!=='All Projects')));
  // Department col E (index 4)
  addValidation(ML,5,505,4,5, listRule(DEPARTMENTS.filter(d=>d!=='All Departments')));
  // Assigned To col F (index 5)
  addValidation(ML,5,505,5,6, listRule(TEAM_MEMBERS.filter(m=>m!=='All Team Members')));
  // Priority col G (index 6)
  addValidation(ML,5,505,6,7, listRule(PRIORITIES));
  // Task Type col H (index 7)
  addValidation(ML,5,505,7,8, listRule(TASK_TYPES));
  // Status col K (index 10)
  addValidation(ML,5,505,10,11, listRule(STATUSES));
  // Effort col (check what column effort is — from 04-mastertask headers:
  // A=Task ID, B=#, C=Task Name, D=Project, E=Dept, F=Assigned To, G=Priority,
  // H=Task Type, I=Start Date, J=Due Date, K=Status, L=Completed Date,
  // M=Days Remaining, N=Est. Hours, O=Hours Logged, P=Overdue?, Q=Rec ID,
  // R=Tags, S=Effort, T=Recurrence, U=Blocked By, V=Notes
  // Effort col S (index 18)
  addValidation(ML,5,505,18,19, listRule(EFFORTS));
  // Recurrence col T (index 19)
  addValidation(ML,5,505,19,20, listRule(RECURRENCE));

  // ── Recurring Task Planner ──
  // Priority col F (index 5), Effort col G (index 6), Recurrence col H (index 7)
  // Active col M (index 12) — already BOOLEAN checkbox, skip
  addValidation(RP,5,20,5,6, listRule(PRIORITIES));
  addValidation(RP,5,20,6,7, listRule(EFFORTS));
  addValidation(RP,5,20,7,8, listRule(RECURRENCE.filter(r=>r!=='None')));

  // Assigned To col E (index 4) in Recurring
  addValidation(RP,5,20,4,5, listRule(TEAM_MEMBERS.filter(m=>m!=='All Team Members')));
  // Project col C (index 2)
  addValidation(RP,5,20,2,3, listRule(PROJECTS.filter(p=>p!=='All Projects')));
  // Department col D (index 3)
  addValidation(RP,5,20,3,4, listRule(DEPARTMENTS.filter(d=>d!=='All Departments')));

  // ── Weekly Calendar controls ──
  // B3=Team Member (row 2, col 1), B4=Project (row 3, col 1), B5=Dept (row 4, col 1)
  addValidation(WC,2,3,1,2, listRule(TEAM_MEMBERS));
  addValidation(WC,3,4,1,2, listRule(PROJECTS));
  addValidation(WC,4,5,1,2, listRule(DEPARTMENTS));

  // ── Monthly Calendar controls ──
  // B2=Month (row 1, col 1), B3=Year (row 2, col 1), B4=Team Member (row 3, col 1), B5=Project (row 4, col 1)
  const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  addValidation(MC,1,2,1,2, listRule(MONTHS));
  addValidation(MC,2,3,1,2, listRule(YEARS));
  addValidation(MC,3,4,1,2, listRule(TEAM_MEMBERS));
  addValidation(MC,4,5,1,2, listRule(PROJECTS));

  // ── Project Progress controls ──
  // B4=Filter by (row 3, col 1) and B21=Selected Project (row 20, col 1)
  addValidation(PP,3,4,1,2, listRule([...PROJECTS.filter(p=>p!=='All Projects'), 'All Projects']));
  addValidation(PP,20,21,1,2, listRule(PROJECTS.filter(p=>p!=='All Projects')));

  // ── Team Dashboard controls ──
  // B4=Year (row 3 col 1), D4=Period (row 3 col 3), F4=Team Member (row 3 col 5)
  // H4=Department (row 3 col 7), J4=Project (row 3 col 9)
  addValidation(DB,3,4,1,2, listRule(YEARS));
  addValidation(DB,3,4,3,4, listRule(PERIODS));
  addValidation(DB,3,4,5,6, listRule(TEAM_MEMBERS));
  addValidation(DB,3,4,7,8, listRule(DEPARTMENTS));
  addValidation(DB,3,4,9,10, listRule(PROJECTS));

  // ── Team Setup ──
  const TS = sheetMap['Team Setup'];
  // Project status col G (index 6) rows 30-44 (indices 29-43)
  addValidation(TS,29,44,6,7, listRule(['Not Started','Planning','In Progress','On Hold','Completed','Cancelled']));
  // Project priority col H (index 7)
  addValidation(TS,29,44,7,8, listRule(PRIORITIES));
  // Department col C (index 2) for member table rows 20-34 (indices 19-33)
  addValidation(TS,19,34,2,3, listRule(DEPARTMENTS.filter(d=>d!=='All Departments')));

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
