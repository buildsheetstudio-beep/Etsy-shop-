'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TSK = sheetMap['✅ Task Manager'];

const HEADERS = ['Task','Life Category','Priority','Status','Due Date','Notes'];
const ROWS = [
  ['Read 12 books this year','Personal Growth','Medium','In Progress','2025-12-31','Currently on book 3'],
  ['Complete online Python course','Personal Growth','High','In Progress','2025-06-30','Lesson 14 of 40'],
  ['Start meditation practice — 10 min/day','Health & Wellness','High','In Progress','2025-02-28','Using Headspace app'],
  ['Schedule annual physical','Health & Wellness','High','Done','2025-01-15','Booked for Jan 20'],
  ['Run a 5K','Health & Wellness','Medium','In Progress','2025-05-15','Training 3x/week'],
  ['Update LinkedIn profile','Career & Business','Medium','Done','2025-01-10','Added new skills'],
  ['Ask for performance review','Career & Business','High','To Do','2025-03-01','Prep talking points first'],
  ['Research side business ideas','Career & Business','Low','In Progress','2025-04-30','Looking at Etsy and freelance'],
  ['Complete Q1 project deliverables','Career & Business','High','In Progress','2025-03-31','On track'],
  ['Deep clean apartment','Home & Household','Medium','Done','2025-01-20','Done — feels great'],
  ['Organize closet','Home & Household','Low','To Do','2025-02-28',''],
  ['Fix leaky bathroom faucet','Home & Household','High','To Do','2025-01-18','Need to call landlord'],
  ['Set up home office properly','Home & Household','Medium','In Progress','2025-03-15','Need monitor stand'],
  ['Plan monthly dinner party','Relationships & Social','Medium','To Do','2025-02-22','Invite 6 people'],
  ['Call parents every Sunday','Relationships & Social','High','In Progress','2025-12-31','Recurring reminder'],
  ['Reconnect with college friends','Relationships & Social','Low','To Do','2025-03-31','Plan a group hangout'],
  ['Plan sister\'s birthday trip','Relationships & Social','High','In Progress','2025-04-01','Looking at weekend options'],
  ['Build emergency fund — 3 months expenses','Finance','High','In Progress','2025-12-31','$4,200 goal, at $1,800'],
  ['Pay off credit card balance','Finance','High','In Progress','2025-06-30','$2,400 remaining'],
  ['Create monthly budget spreadsheet','Finance','Medium','Done','2025-01-05','Using this planner!'],
  ['Research Roth IRA options','Finance','Medium','To Do','2025-04-30','Compare 3 brokerages'],
  ['Automate savings transfers','Finance','Low','Done','2025-01-12','$300/month set up'],
  ['Write morning pages for 30 days','Personal Growth','Low','Done','2025-01-31','Completed the challenge!'],
  ['Take a digital detox weekend','Health & Wellness','Low','To Do','2025-03-29',''],
  ['Attend a networking event','Career & Business','Medium','To Do','2025-02-20','Look up local events'],
  ['Meal prep on Sundays','Home & Household','Medium','In Progress','2025-12-31','Recurring weekly'],
  ['Text a friend I haven\'t talked to','Relationships & Social','Low','Done','2025-01-14','Reached out to Marcus'],
  ['Track spending for 30 days','Finance','High','Done','2025-01-31','Eye-opening — $400 on food!'],
  ['Journal 3 things I\'m grateful for daily','Personal Growth','Medium','In Progress','2025-12-31','Using this planner'],
  ['Go to bed by 10:30pm consistently','Health & Wellness','High','To Do','2025-02-15','Sleep quality priority'],
];

(async () => {
  if (ROWS.length !== 30) throw new Error(`Expected 30 rows, got ${ROWS.length}`);
  const data = [
    { range: "'✅ Task Manager'!A1:F1", values: [HEADERS] },
    { range: "'✅ Task Manager'!A2:F31", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'tsk-values');

  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(TSK,0,1,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TSK, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  for (let r = 1; r <= 30; r++) {
    reqs.push({ repeatCell: { range: gridRange(TSK,r,r+1,0,6), cell: { userEnteredFormat: { backgroundColor: hex(r%2===1?C.altRow:C.white), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  }
  reqs.push({ repeatCell: { range: gridRange(TSK,1,31,0,1), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });
  reqs.push({ repeatCell: { range: gridRange(TSK,1,31,5,6), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });
  [[0,240],[1,150],[2,80],[3,90],[4,100],[5,200]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: TSK, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TSK, dimension: 'ROWS', startIndex: 1, endIndex: 31 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });
  await batchUpdate(id, reqs, 'tsk-format');
  console.log('Task Manager complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
