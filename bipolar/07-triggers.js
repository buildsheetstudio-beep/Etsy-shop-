'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TRG = sheetMap['⚡ Trigger & Stressor Log'];

// Columns: A=Date, B=Trigger/Stressor, C=Category, D=Impact Severity, E=Coping Strategy Used, F=Notes
const HEADERS = ['Date','Trigger / Stressor','Category','Impact Severity','Coping Strategy Used','Notes'];

const ROWS = [
  ['2025-01-08', 'Missed medication dose', 'Health', 'High', 'Called pharmacist, took dose when remembered', 'Mood significantly worse on missed-dose days.'],
  ['2025-01-15', 'Argument with partner', 'Relationship', 'Medium', 'Took a walk, called therapist next day', 'Felt withdrawn afterward.'],
  ['2025-01-22', 'Work deadline stress', 'Work/School', 'Medium', 'Broke task into small steps, set alarms', ''],
  ['2025-01-29', 'Unexpected bill', 'Financial', 'Low', 'Reviewed budget, set up payment plan', 'Financial worry always hits mood.'],
  ['2025-02-09', 'Poor sleep three nights running', 'Sleep Disruption', 'High', 'Adjusted bedtime, took quetiapine on schedule', 'Sleep disruption always precedes mood shift.'],
  ['2025-03-02', 'Woke 4am unable to sleep', 'Sleep Disruption', 'High', 'Noted in log, texted psychiatrist', 'Flagged as early mania warning.'],
  ['2025-03-09', 'Very high energy — hard to sit still', 'Health', 'High', 'Grounding exercise, called psychiatrist', 'Hypomania escalating.'],
  ['2025-03-12', 'Forgot medication again', 'Health', 'Medium', 'Set phone alarm for both doses', 'Missed doses correlate with mood shifts.'],
  ['2025-03-19', 'Social event — overstimulated', 'Relationship', 'Medium', 'Left early, rested at home', 'Social events can trigger escalation.'],
  ['2025-04-10', 'Minor work conflict', 'Work/School', 'Low', 'Discussed with manager, resolved quickly', 'Handled it. Good sign.'],
  ['2025-05-01', 'Late-night screen time', 'Sleep Disruption', 'Low', 'Set phone down earlier, wore blue-light glasses', ''],
  ['2025-06-01', 'Grandiose thinking at work', 'Health', 'High', 'Called psychiatrist immediately', 'Manic warning sign. Dose adjustment discussed.'],
  ['2025-06-03', 'Impulsive online shopping', 'Financial', 'High', 'Handed credit card to partner, deleted shopping apps', 'Spent several hundred dollars. Regret.'],
  ['2025-06-05', 'Conflict with family member (phone)', 'Relationship', 'High', 'Ended call, noted mood state', 'Irritability is symptom, not personal failing.'],
  ['2025-06-08', 'Hospitalization discussion with psychiatrist', 'Health', 'High', 'Agreed to dose increase, daily check-in', 'Scary but necessary conversation.'],
  ['2025-07-04', 'Holiday gathering — exhausted', 'Relationship', 'Medium', 'Left after 2 hours, took long nap', 'Post-manic fatigue. Protecting rest.'],
  ['2025-07-13', 'Intrusive thoughts about June spending', 'Financial', 'High', 'Journaled, discussed in therapy', 'Working through regret in therapy.'],
  ['2025-07-18', 'Couldn\'t taste food — low mood', 'Health', 'Medium', 'Called therapist, ate anyway', 'Anhedonia. Symptom, not character flaw.'],
  ['2025-08-12', 'Stressful news cycle', 'Other', 'Low', 'Media break, 30-min walk', 'Noticed external news can spike anxiety.'],
  ['2025-09-10', 'Work presentation — performance anxiety', 'Work/School', 'Medium', 'Deep breathing, prepared thoroughly', 'Went well. Stress manageable when stable.'],
];

(async () => {
  if (ROWS.length !== 20) throw new Error(`Expected 20 rows, got ${ROWS.length}`);

  const data = [
    { range: "'⚡ Trigger & Stressor Log'!A1:F1", values: [HEADERS] },
    { range: "'⚡ Trigger & Stressor Log'!A2:F21", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'trg-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(TRG, 0, 1, 0, 6),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 20; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(TRG, r, r+1, 0, 6),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  reqs.push({ repeatCell: { range: gridRange(TRG, 1, 21, 1, 2), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });
  reqs.push({ repeatCell: { range: gridRange(TRG, 1, 21, 4, 6), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });

  [[0,100],[1,200],[2,120],[3,100],[4,200],[5,220]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: TRG, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: TRG, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TRG, dimension: 'ROWS', startIndex: 1, endIndex: 21 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'trg-format');
  console.log('Trigger & Stressor Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
