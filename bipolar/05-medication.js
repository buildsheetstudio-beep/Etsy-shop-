'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const MED = sheetMap['💊 Medication & Treatment Tracker'];

// Columns: A=Date, B=Medication Name, C=Dosage, D=Frequency, E=Taken(checkbox), F=Side Effects Noted, G=Notes
const HEADERS = ['Date','Medication Name','Dosage','Frequency','Taken','Side Effects Noted','Notes'];

const ROWS = [
  ['2025-01-02','Lithium','600mg','Twice Daily',true,'','Morning and evening doses.'],
  ['2025-01-02','Quetiapine','50mg','Once Daily',true,'Mild drowsiness','Taken at bedtime.'],
  ['2025-01-05','Lithium','600mg','Twice Daily',true,'',''],
  ['2025-01-05','Quetiapine','50mg','Once Daily',true,'',''],
  ['2025-01-08','Lithium','600mg','Twice Daily',false,'','Forgot morning dose. Felt worse today.'],
  ['2025-01-08','Quetiapine','50mg','Once Daily',true,'',''],
  ['2025-01-12','Lithium','600mg','Twice Daily',true,'','Back on track.'],
  ['2025-01-12','Quetiapine','50mg','Once Daily',true,'Mild drowsiness',''],
  ['2025-01-19','Lithium','600mg','Twice Daily',true,'',''],
  ['2025-01-19','Quetiapine','50mg','Once Daily',true,'',''],
  ['2025-02-05','Lithium','600mg','Twice Daily',true,'','Mood improving.'],
  ['2025-02-05','Quetiapine','50mg','Once Daily',true,'',''],
  ['2025-02-19','Lithium','600mg','Twice Daily',true,'',''],
  ['2025-02-19','Quetiapine','50mg','Once Daily',true,'',''],
  ['2025-03-09','Lithium','600mg','Twice Daily',true,'','Slept 5hrs — watching closely.'],
  ['2025-03-09','Quetiapine','50mg','Once Daily',true,'',''],
  ['2025-03-12','Lithium','600mg','Twice Daily',false,'','Forgot. Hypo symptoms elevated.'],
  ['2025-03-12','Quetiapine','50mg','Once Daily',true,'',''],
  ['2025-04-06','Lithium','600mg','Twice Daily',true,'','Consistent adherence this month.'],
  ['2025-04-06','Quetiapine','50mg','Once Daily',true,'',''],
  ['2025-06-01','Lithium','600mg','Twice Daily',true,'','Manic — psychiatrist aware.'],
  ['2025-06-01','Quetiapine','100mg','Once Daily',true,'Significant drowsiness','Dose doubled by psychiatrist.'],
  ['2025-06-08','Lithium','900mg','Twice Daily',true,'Nausea, hand tremor','Dose increased due to manic episode.'],
  ['2025-06-08','Quetiapine','100mg','Once Daily',true,'Drowsiness, dry mouth',''],
  ['2025-06-13','Lithium','900mg','Twice Daily',true,'Mild tremor','Side effects settling.'],
  ['2025-06-13','Quetiapine','100mg','Once Daily',true,'Mild drowsiness',''],
  ['2025-07-08','Lithium','900mg','Twice Daily',true,'','Post-manic depression. Maintaining dose.'],
  ['2025-07-08','Quetiapine','100mg','Once Daily',true,'',''],
  ['2025-08-20','Lithium','900mg','Twice Daily',true,'','Back to stable.'],
  ['2025-08-20','Quetiapine','75mg','Once Daily',true,'Mild drowsiness','Dose slightly reduced.'],
];

(async () => {
  if (ROWS.length !== 30) throw new Error(`Expected 30 rows, got ${ROWS.length}`);
  ROWS.forEach((r, i) => { if (r.length !== 7) throw new Error(`Row ${i+1} has ${r.length} cols`); });

  const data = [
    { range: "'💊 Medication & Treatment Tracker'!A1:G1", values: [HEADERS] },
    { range: "'💊 Medication & Treatment Tracker'!A2:G31", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'med-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(MED, 0, 1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 30; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(MED, r, r+1, 0, 7),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  // Notes + side effects — wrap
  reqs.push({ repeatCell: { range: gridRange(MED, 1, 31, 5, 7), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });

  [[0,100],[1,130],[2,80],[3,110],[4,70],[5,180],[6,220]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: MED, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: MED, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: MED, dimension: 'ROWS', startIndex: 1, endIndex: 31 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'med-format');
  console.log('Medication & Treatment Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
