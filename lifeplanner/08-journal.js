'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const JRN = sheetMap['📓 Journal & Gratitude Log'];

const HEADERS = ['Date','Gratitude 1','Gratitude 2','Gratitude 3','Journal Notes','Mood'];
const ROWS = [
  ['2025-01-06','My warm apartment on a cold morning','A job I actually don\'t hate','My good health','Feeling motivated going into this week. Set up the life planner. Excited to track everything.','Great'],
  ['2025-01-07','Good coffee before work','A friend texting to check in','Running water','Long work day but manageable. Texted Marcus — haven\'t talked in months. He\'s doing well.','Good'],
  ['2025-01-08','Rest after a tiring day','Groceries already prepped','My cozy bed','Tired tonight. Skipped gym. Didn\'t beat myself up — just went to sleep earlier.','Okay'],
  ['2025-01-09','A sunny January day','Making progress on Python','Evening walk outside','Python lesson 14 done. The sunset was genuinely beautiful tonight. Good reset.','Good'],
  ['2025-01-10','Finishing the work week','My playlist on the commute','Plans with friends this weekend','End-of-week tiredness but in a good way. Looking forward to the weekend.','Good'],
  ['2025-01-11','A long Saturday sleep-in','Cooking a real meal','Calling home','Made pasta from scratch. Called parents. Two hours of reading. This is what weekends are for.','Great'],
  ['2025-01-12','A reset Sunday','Gym session that felt strong','Prepping for the week','Meal prepped for the week. Planned my goals. Feel genuinely ready for Monday for once.','Great'],
  ['2025-01-13','My morning routine holding steady','Progress feeling real','A good conversation today','Hard day at work — unexpected conflict. Handled it okay. Journaling helps more than I expected.','Okay'],
  ['2025-01-14','Warm tea on a gray day','Small wins still count','My support system','Feeling a little flat today. Not sad, just… muted. Still did all my habits. That counts.','Low'],
  ['2025-01-15','Payday','Feeling caught up for once','Room cleaned and calm','Paid rent, transferred to savings, still have money left. Small win. Budget is working.','Good'],
  ['2025-01-16','Friends who check in without prompting','My couch and a good book','Music I forgot I loved','Nothing dramatic happened today. Just quiet and nice. Grateful for ordinary good days.','Good'],
  ['2025-01-17','A healthy lunch I actually enjoyed','Fresh air on a walk','Momentum building','Work felt heavy. Walk at lunch helped reset. Remembering to give myself breaks.','Okay'],
  ['2025-01-18','Progress on the emergency fund','Running farther than last month','Clarity about priorities','Ran 2.8 miles — almost 3! Proud of myself. Finance section really helps me see progress.','Great'],
  ['2025-01-19','Sleep','Really good sleep','Seriously, sleep','Rough day. Nothing specific, just heavy. Everything felt harder than it should. Going to bed early.','Rough'],
];

(async () => {
  if (ROWS.length !== 14) throw new Error(`Expected 14 rows, got ${ROWS.length}`);
  const data = [
    { range: "'📓 Journal & Gratitude Log'!A1:F1", values: [HEADERS] },
    { range: "'📓 Journal & Gratitude Log'!A2:F15", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'jrn-values');

  const reqs = [];
  reqs.push({ repeatCell: { range: gridRange(JRN,0,1,0,6), cell: { userEnteredFormat: { backgroundColor: hex(C.raspberry), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: JRN, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  for (let r = 1; r <= 14; r++) {
    reqs.push({ repeatCell: { range: gridRange(JRN,r,r+1,0,6), cell: { userEnteredFormat: { backgroundColor: hex(r%2===1?C.altRow:C.white), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 }, wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)' } });
  }
  [[0,100],[1,180],[2,180],[3,180],[4,300],[5,70]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: JRN, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: JRN, dimension: 'ROWS', startIndex: 1, endIndex: 15 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  await batchUpdate(id, reqs, 'jrn-format');
  console.log('Journal & Gratitude Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
