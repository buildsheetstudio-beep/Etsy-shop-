'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const THR = sheetMap['🗓️ Therapy & Appointment Log'];

// Columns: A=Date, B=Provider Type, C=Provider Name, D=Topics Discussed, E=Follow-Up Needed(checkbox), F=Next Appointment Date, G=Notes
const HEADERS = ['Date','Provider Type','Provider Name','Topics Discussed','Follow-Up Needed','Next Appointment Date','Notes'];

const ROWS = [
  ['2025-01-07', 'Psychiatrist', 'Dr. Ramirez', 'Medication review, depression management', false, '2025-02-04', 'Stable on current doses. Check lithium levels next visit.'],
  ['2025-01-09', 'Therapist',   'Dr. Chen',    'Coping strategies for low mood, sleep hygiene', false, '2025-01-23', 'CBT worksheet assigned.'],
  ['2025-01-23', 'Therapist',   'Dr. Chen',    'Worksheet review, behavioral activation', false, '2025-02-06', 'Started small daily activity goals.'],
  ['2025-02-04', 'Psychiatrist', 'Dr. Ramirez', 'Mood tracking review, lithium blood levels', false, '2025-03-04', 'Lithium levels therapeutic. No changes.'],
  ['2025-02-06', 'Therapist',   'Dr. Chen',    'Progress on behavioral activation, goal-setting', false, '2025-02-20', 'Mood lifting. Good progress.'],
  ['2025-02-20', 'Therapist',   'Dr. Chen',    'Relationship stress, communication skills', false, '2025-03-06', 'Partner may join next session.'],
  ['2025-03-04', 'Psychiatrist', 'Dr. Ramirez', 'Hypomanic symptoms emerging, sleep disruption', true, '2025-03-18', 'Watch sleep carefully. Call if escalates.'],
  ['2025-03-06', 'Therapist',   'Dr. Chen',    'Early warning signs, hypomanic thought patterns', false, '2025-03-20', 'Reviewed personal warning signs list.'],
  ['2025-03-18', 'Psychiatrist', 'Dr. Ramirez', 'Hypomania update — symptoms manageable', false, '2025-04-08', 'No medication change needed. Good self-monitoring.'],
  ['2025-04-08', 'Psychiatrist', 'Dr. Ramirez', 'Stable review, annual blood work ordered', false, '2025-05-06', 'Great stable period. Maintain routines.'],
  ['2025-06-04', 'Psychiatrist', 'Dr. Ramirez', 'Manic episode — urgent appointment', true, '2025-06-11', 'Dose of lithium increased. Dose of quetiapine doubled. Daily check-in plan.'],
  ['2025-06-11', 'Psychiatrist', 'Dr. Ramirez', 'Manic follow-up — symptoms improving', true, '2025-06-18', 'Continuing dose adjustment. Monitoring closely.'],
  ['2025-06-12', 'Therapist',   'Dr. Chen',    'Manic episode debriefing, safety planning review', true, '2025-06-26', 'Updated safety plan. Reviewed spending decisions.'],
  ['2025-07-10', 'Therapist',   'Dr. Chen',    'Post-manic depression, guilt processing', false, '2025-07-24', 'Working on self-compassion. Regret is normal, not a verdict.'],
  ['2025-08-19', 'Psychiatrist', 'Dr. Ramirez', 'Recovery check-in, medication consolidation', false, '2025-09-16', 'Slightly reducing quetiapine. Mood stable.'],
];

(async () => {
  if (ROWS.length !== 15) throw new Error(`Expected 15 rows, got ${ROWS.length}`);

  const data = [
    { range: "'🗓️ Therapy & Appointment Log'!A1:G1", values: [HEADERS] },
    { range: "'🗓️ Therapy & Appointment Log'!A2:G16", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'thr-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(THR, 0, 1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 15; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(THR, r, r+1, 0, 7),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  reqs.push({ repeatCell: { range: gridRange(THR, 1, 16, 3, 4), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });
  reqs.push({ repeatCell: { range: gridRange(THR, 1, 16, 6, 7), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });

  [[0,100],[1,100],[2,120],[3,220],[4,80],[5,130],[6,260]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: THR, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: THR, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: THR, dimension: 'ROWS', startIndex: 1, endIndex: 16 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'thr-format');
  console.log('Therapy & Appointment Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
