'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSL = sheetMap['📅 Daily Symptom Log'];

// Columns: A=Date, B=Overall Severity % (0-100), C=Episode Type, D=Manic Symptoms Present,
//          E=Depressive Symptoms Present, F=Medication Taken (checkbox), G=Notes
const HEADERS = ['Date','Overall Severity %','Episode Type','Manic Symptoms Present','Depressive Symptoms Present','Medication Taken','Notes'];

// 60 rows spanning Jan–Oct 2025, covering every episode type with realistic patterns:
// - Winter: depressive episodes (Jan-Feb)
// - Spring hypomania (Mar-Apr)
// - Stable stretch (May)
// - Full manic episode (Jun)
// - Post-manic depressive crash (Jul-Aug)
// - Stabilizing (Sep-Oct)
// F (Medication Taken) is TRUE/FALSE for checkbox seeding via USER_ENTERED

const ROWS = [
  // January — depressive
  ['2025-01-02', 65, 'Depressive', '', 'Persistent low mood, Fatigue or low energy', true, 'Struggled to get out of bed. Called therapist.'],
  ['2025-01-05', 70, 'Depressive', '', 'Persistent low mood, Loss of interest or pleasure, Fatigue or low energy', true, 'No motivation. Stayed home.'],
  ['2025-01-08', 72, 'Depressive', '', 'Persistent low mood, Social withdrawal, Feelings of worthlessness', false, 'Missed medication. Very low.'],
  ['2025-01-12', 60, 'Depressive', '', 'Fatigue or low energy, Difficulty concentrating', true, 'A bit better. Made it to work.'],
  ['2025-01-15', 55, 'Depressive', '', 'Persistent low mood, Changes in appetite or sleep', true, 'Sleeping too much. Appetite down.'],
  ['2025-01-19', 68, 'Depressive', '', 'Persistent low mood, Social withdrawal', true, 'Cancelled plans. Stayed in.'],
  ['2025-01-22', 58, 'Depressive', '', 'Fatigue or low energy, Difficulty concentrating', true, 'Work was hard but managed.'],
  ['2025-01-26', 50, 'Depressive', '', 'Persistent low mood, Fatigue or low energy', true, 'Slight improvement. Did laundry.'],
  ['2025-01-29', 45, 'Depressive', '', 'Persistent low mood', true, 'Better day. Went for a short walk.'],
  // February — still depressive, slowly lifting
  ['2025-02-02', 48, 'Depressive', '', 'Fatigue or low energy, Changes in appetite or sleep', true, 'Sleep still disrupted.'],
  ['2025-02-05', 40, 'Depressive', '', 'Persistent low mood', true, 'Gradual lift. Cooked a real meal.'],
  ['2025-02-09', 35, 'Hypomanic', 'Elevated or irritable mood', '', true, 'Energy picking up. Mood brightening.'],
  ['2025-02-12', 30, 'Stable/Euthymic', '', '', true, 'Good day. Felt like myself.'],
  ['2025-02-16', 25, 'Stable/Euthymic', '', '', true, 'Stable. Productive at work.'],
  ['2025-02-19', 20, 'Stable/Euthymic', '', '', true, 'Best day in weeks.'],
  ['2025-02-23', 22, 'Stable/Euthymic', '', '', true, 'Maintained.'],
  // March — hypomania building
  ['2025-03-02', 30, 'Hypomanic', 'Elevated or irritable mood, Decreased need for sleep', '', true, 'Woke at 4am feeling great. Productive.'],
  ['2025-03-05', 38, 'Hypomanic', 'Elevated or irritable mood, Increased goal-directed activity', '', true, 'Lots of energy. Started 3 new projects.'],
  ['2025-03-09', 42, 'Hypomanic', 'Racing thoughts, Decreased need for sleep', '', true, 'Only slept 5 hrs. Not tired.'],
  ['2025-03-12', 45, 'Hypomanic', 'Racing thoughts, Rapid or pressured speech', '', false, 'Forgot meds. Talked too much at meeting.'],
  ['2025-03-16', 48, 'Hypomanic', 'Elevated or irritable mood, Distractibility', '', true, 'Hard to focus. Irritable with partner.'],
  ['2025-03-19', 40, 'Hypomanic', 'Increased goal-directed activity', '', true, 'Still elevated but manageable.'],
  ['2025-03-23', 35, 'Stable/Euthymic', '', '', true, 'Coming down. Sleep better last night.'],
  ['2025-03-26', 28, 'Stable/Euthymic', '', '', true, 'Stable. Good week overall.'],
  // April — stable
  ['2025-04-02', 20, 'Stable/Euthymic', '', '', true, 'Stable stretch. Enjoying spring.'],
  ['2025-04-06', 18, 'Stable/Euthymic', '', '', true, 'Best stretch in months.'],
  ['2025-04-10', 22, 'Stable/Euthymic', '', '', true, 'Minor stress at work, handled it.'],
  ['2025-04-14', 25, 'Stable/Euthymic', '', '', true, 'Went hiking. Feel grounded.'],
  ['2025-04-18', 20, 'Stable/Euthymic', '', '', true, 'Good sleep. Good mood.'],
  ['2025-04-23', 18, 'Stable/Euthymic', '', '', true, 'Stable.'],
  // May — still stable with very brief hypo blip
  ['2025-05-01', 22, 'Stable/Euthymic', '', '', true, 'Stable. Therapy went well.'],
  ['2025-05-07', 28, 'Hypomanic', 'Elevated or irritable mood', '', true, 'Energy up slightly. Manageable.'],
  ['2025-05-11', 24, 'Stable/Euthymic', '', '', true, 'Back to baseline. Kept up routines.'],
  ['2025-05-18', 20, 'Stable/Euthymic', '', '', true, 'Good week.'],
  ['2025-05-25', 18, 'Stable/Euthymic', '', '', true, 'Memorial Day. Relaxed. Stable.'],
  // June — manic episode
  ['2025-06-01', 55, 'Manic', 'Elevated or irritable mood, Decreased need for sleep, Racing thoughts', '', true, 'Slept 3 hours. Feel invincible. Red flag.'],
  ['2025-06-03', 68, 'Manic', 'Racing thoughts, Rapid or pressured speech, Increased goal-directed activity', '', true, 'Rapid speech noted by partner. Called psychiatrist.'],
  ['2025-06-05', 75, 'Manic', 'Grandiosity or inflated self-esteem, Impulsive or risky decisions', '', true, 'Spent a lot of money online. Regret already.'],
  ['2025-06-08', 80, 'Manic', 'Elevated or irritable mood, Racing thoughts, Decreased need for sleep, Grandiosity or inflated self-esteem', '', true, 'Hospitalization discussed. Dose adjusted.'],
  ['2025-06-10', 78, 'Manic', 'Racing thoughts, Distractibility', '', true, 'Still high. Dose change taking effect slowly.'],
  ['2025-06-13', 68, 'Manic', 'Elevated or irritable mood, Rapid or pressured speech', '', true, 'Coming down slightly.'],
  ['2025-06-17', 55, 'Hypomanic', 'Elevated or irritable mood', '', true, 'More grounded. Relief.'],
  ['2025-06-22', 42, 'Hypomanic', 'Elevated or irritable mood', '', true, 'Still above baseline but manageable.'],
  ['2025-06-28', 35, 'Stable/Euthymic', '', '', true, 'Crash avoided. Psychiatrist pleased.'],
  // July — post-manic depressive dip
  ['2025-07-04', 50, 'Depressive', '', 'Persistent low mood, Fatigue or low energy', true, 'Post-manic fatigue. Expected.'],
  ['2025-07-08', 60, 'Depressive', '', 'Persistent low mood, Social withdrawal, Fatigue or low energy', true, 'Hard week. Therapy tomorrow.'],
  ['2025-07-13', 65, 'Depressive', '', 'Persistent low mood, Feelings of worthlessness, Difficulty concentrating', true, 'Struggling with regret over June spending.'],
  ['2025-07-18', 58, 'Depressive', '', 'Fatigue or low energy, Changes in appetite or sleep', true, 'Appetite low. Sleeping a lot.'],
  ['2025-07-24', 52, 'Depressive', '', 'Persistent low mood, Fatigue or low energy', true, 'Slight improvement.'],
  ['2025-07-30', 45, 'Depressive', '', 'Persistent low mood', true, 'Therapy helping. Slowly lifting.'],
  // August — lifting
  ['2025-08-05', 40, 'Depressive', '', 'Fatigue or low energy', true, 'Better. Got outside.'],
  ['2025-08-12', 32, 'Depressive', '', 'Persistent low mood', true, 'Mood lifting week by week.'],
  ['2025-08-20', 25, 'Stable/Euthymic', '', '', true, 'Back to stable. Major relief.'],
  ['2025-08-27', 20, 'Stable/Euthymic', '', '', true, 'Good week. Exercised daily.'],
  // September — stable
  ['2025-09-03', 18, 'Stable/Euthymic', '', '', true, 'Stable. Starting new routine.'],
  ['2025-09-10', 22, 'Stable/Euthymic', '', '', true, 'Work stress but handled.'],
  ['2025-09-18', 20, 'Stable/Euthymic', '', '', true, 'Consistent sleep helping.'],
  ['2025-09-25', 18, 'Stable/Euthymic', '', '', true, 'Best month in a while.'],
  // October — stable, one hypo blip
  ['2025-10-02', 20, 'Stable/Euthymic', '', '', true, 'Stable.'],
  ['2025-10-09', 28, 'Hypomanic', 'Elevated or irritable mood', '', true, 'Energy uptick — monitoring closely.'],
];

(async () => {
  if (ROWS.length !== 60) throw new Error(`Expected 60 rows, got ${ROWS.length}`);
  ROWS.forEach((r, i) => { if (r.length !== 7) throw new Error(`Row ${i+1} has ${r.length} cols`); });

  const data = [
    { range: "'📅 Daily Symptom Log'!A1:G1", values: [HEADERS] },
    { range: "'📅 Daily Symptom Log'!A2:G61", values: ROWS },
  ];
  await valuesBatchUpdate(id, data, 'dsl-values');

  const reqs = [];

  reqs.push({
    repeatCell: {
      range: gridRange(DSL, 0, 1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 1; r <= 60; r++) {
    const bg = r % 2 === 1 ? C.altRow : C.white;
    reqs.push({
      repeatCell: {
        range: gridRange(DSL, r, r+1, 0, 7),
        cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 9 }, verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // Severity col B — centered
  reqs.push({ repeatCell: { range: gridRange(DSL, 1, 61, 1, 2), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } });
  // Notes col — wrap
  reqs.push({ repeatCell: { range: gridRange(DSL, 1, 61, 6, 7), cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(wrapStrategy)' } });
  // Symptom cols — wrap
  reqs.push({ repeatCell: { range: gridRange(DSL, 1, 61, 3, 5), cell: { userEnteredFormat: { wrapStrategy: 'WRAP', textFormat: { fontSize: 8 } } }, fields: 'userEnteredFormat(wrapStrategy,textFormat)' } });

  [[0,100],[1,80],[2,110],[3,180],[4,180],[5,90],[6,260]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSL, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  reqs.push({ updateDimensionProperties: { range: { sheetId: DSL, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSL, dimension: 'ROWS', startIndex: 1, endIndex: 61 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'dsl-format');
  console.log('Daily Symptom Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
