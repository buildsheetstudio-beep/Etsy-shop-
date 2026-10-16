'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RFL = sheetMap['📝 Weekly & Monthly Reflection'];

const HEADERS = ['Period','Period Type','Wins This Period','Challenges Faced','Lessons Learned','Goals for Next Period','Overall Rating (1-10)'];
const ROWS = [
  ['Week of Jan 6','Weekly','Started meditation daily — 7/7 days. Reconnected with Marcus.','Overspent on food — $150 above budget.','Tracking before spending makes a real difference.','Stick to $80/week food budget. Read every day.',8],
  ['Week of Jan 13','Weekly','Finished 3 Python lessons. Went to bed by 10:30 four nights.','Skipped gym twice. Felt guilty after.','Skipping the gym hurts sleep more than I thought.','Gym 3x. One digital-free evening.',7],
  ['January 2025','Monthly','Set up automated savings. Finished annual physical. Read one full book.','Emergency fund still below goal. Exercise inconsistent.','Automation removes the willpower requirement. Just set it up.','Add gym to calendar, not just to-do list. Read book 2.',7],
  ['Week of Jan 20','Weekly','Deep cleaned apartment — felt amazing. Meal prepped Sunday.','Work deadline stress spilled into personal time.','Batching chores (Sunday reset) saves mental energy all week.','Keep Sunday reset. Close work laptop at 6pm.',8],
  ['Week of Jan 27','Weekly','Stayed under budget for the first time this month. Called parents twice.','Python lessons slipping — only one done this week.','I need to block dedicated learning time, not just "fit it in."','Block 8–8:30am for Python. Keep budget streak.',7],
  ['Week of Feb 3','Weekly','Ran 3 miles without stopping — new personal best!','Missed meditation 4 out of 7 days.','Running is improving. Meditation needs a better trigger.','Meditate right after alarm. Track it in this planner.',8],
  ['February 2025','Monthly','Personal best run. Paid off $200 of credit card. Had a great dinner party.','Dental co-pay threw off budget. Motivation dipped mid-month.','Mid-month dips are normal — don\'t catastrophize, just reset.','Build $50 buffer into each month\'s budget. Start book 3.',7],
  ['Week of Feb 10','Weekly','Reached out to two friends I\'d lost touch with. Organized home office.','Overspent on Valentine\'s dinner — but no regrets.','Some "overspending" is intentional — know the difference.','Write out this week\'s top 3 priorities on Sunday night.',9],
];

(async () => {
  if (ROWS.length !== 8) throw new Error(`Expected 8 rows, got ${ROWS.length}`);
  const data = [
    { range: "'📝 Weekly & Monthly Reflection'!A1:G1", values: [HEADERS] },
    { range: "'📝 Weekly & Monthly Reflection'!A2:G9", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'rfl-values');

  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(RFL,0,1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: RFL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });
  for (let r = 1; r <= 8; r++) {
    reqs.push({ repeatCell: { range: gridRange(RFL,r,r+1,0,7), cell: { userEnteredFormat: { backgroundColor: hex(r%2===1?C.altRow:C.white), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 }, wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)' } });
  }
  reqs.push({ repeatCell: { range: gridRange(RFL,1,9,6,7), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true, fontSize: 12 } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });
  [[0,130],[1,80],[2,240],[3,200],[4,240],[5,200],[6,80]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: RFL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: RFL, dimension: 'ROWS', startIndex: 1, endIndex: 9 }, properties: { pixelSize: 72 }, fields: 'pixelSize' } });
  await batchUpdate(id, reqs, 'rfl-format');
  console.log('Weekly & Monthly Reflection complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
