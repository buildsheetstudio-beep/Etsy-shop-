'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SLP = sheetMap['😴 Sleep Tracker'];

// Columns: A=Date, B=Hours Slept, C=Sleep Quality, D=Bedtime, E=Wake Time, F=Notes
const HEADERS = ['Date','Hours Slept','Sleep Quality','Bedtime','Wake Time','Notes'];

const ROWS = [
  ['2025-01-02', 10.5, 'Poor',  '23:00', '09:30', 'Oversleeping — depressive pattern.'],
  ['2025-01-05', 11,   'Poor',  '22:30', '09:30', 'Hard to get up.'],
  ['2025-01-08', 10,   'Poor',  '23:00', '09:00', 'Missed medication — slept excessively.'],
  ['2025-01-12',  9,   'Fair',  '23:00', '08:00', 'Slight improvement.'],
  ['2025-01-19', 10,   'Poor',  '22:00', '08:00', 'Still oversleeping.'],
  ['2025-01-26',  8.5, 'Fair',  '23:00', '07:30', 'A bit better.'],
  ['2025-01-29',  8,   'Fair',  '23:00', '07:00', 'Getting closer to normal.'],
  ['2025-02-05',  7.5, 'Good',  '22:30', '06:00', 'Sleep normalizing.'],
  ['2025-02-12',  7,   'Good',  '22:00', '05:00', 'Energy improving.'],
  ['2025-02-19',  7,   'Great', '22:00', '05:00', 'Best sleep in weeks.'],
  ['2025-03-02',  5.5, 'Poor',  '00:00', '05:30', 'Woke at 4am energized — hypomania sign.'],
  ['2025-03-09',  5,   'Poor',  '00:30', '05:30', 'Only 5 hrs. Didn\'t feel tired at all.'],
  ['2025-03-12',  4.5, 'Poor',  '01:00', '05:30', 'Missed meds. Very reduced need for sleep.'],
  ['2025-03-19',  6,   'Fair',  '23:00', '05:00', 'Slightly better.'],
  ['2025-03-26',  7,   'Good',  '22:30', '05:30', 'Coming back down. Sleep improving.'],
  ['2025-04-06',  7.5, 'Great', '22:00', '05:30', 'Stable sleep. Feeling well.'],
  ['2025-04-14',  8,   'Great', '22:00', '06:00', 'Hiking tired me out nicely.'],
  ['2025-04-23',  7.5, 'Good',  '22:00', '05:30', 'Consistent.'],
  ['2025-05-07',  6.5, 'Fair',  '23:00', '05:30', 'Slight energy uptick — mild hypo.'],
  ['2025-05-18',  7.5, 'Good',  '22:00', '05:30', 'Back to baseline.'],
  ['2025-06-01',  3,   'Poor',  '02:00', '05:00', 'Only 3 hrs. Full mania. Felt amazing — bad sign.'],
  ['2025-06-03',  3.5, 'Poor',  '01:30', '05:00', 'Still very little sleep needed.'],
  ['2025-06-08',  4,   'Poor',  '01:00', '05:00', 'Dose adjustment started.'],
  ['2025-06-13',  5.5, 'Poor',  '00:00', '05:30', 'Slightly more sleep. Coming down.'],
  ['2025-06-17',  6.5, 'Fair',  '23:00', '05:30', 'More sleep. Mood settling.'],
  ['2025-06-28',  7,   'Good',  '22:30', '05:30', 'Back closer to normal.'],
  ['2025-07-08', 10.5, 'Poor',  '22:00', '08:30', 'Post-manic crash. Sleeping excessively.'],
  ['2025-08-05',  9,   'Fair',  '22:30', '07:30', 'Depression lifting. Still long sleep.'],
  ['2025-08-20',  7.5, 'Good',  '22:00', '05:30', 'Back to stable range.'],
  ['2025-09-10',  7.5, 'Great', '22:00', '05:30', 'Consistent sleep. Feeling stable.'],
];

(async () => {
  if (ROWS.length !== 30) throw new Error(`Expected 30 rows, got ${ROWS.length}`);

  const data = [
    { range: "'😴 Sleep Tracker'!A1:F1", values: [HEADERS] },
    { range: "'😴 Sleep Tracker'!A2:F31", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'slp-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(SLP, 0, 1, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 30; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(SLP, r, r+1, 0, 6),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  // Hours slept — centered bold
  reqs.push({ repeatCell: { range: gridRange(SLP, 1, 31, 1, 2), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });
  // Notes col — wrap
  reqs.push({ repeatCell: { range: gridRange(SLP, 1, 31, 5, 6), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });

  [[0,100],[1,90],[2,100],[3,80],[4,80],[5,260]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SLP, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: SLP, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'slp-format');
  console.log('Sleep Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
