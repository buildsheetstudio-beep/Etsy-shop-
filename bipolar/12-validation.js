'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DSL = sheetMap['📅 Daily Symptom Log'];
const MED = sheetMap['💊 Medication & Treatment Tracker'];
const SLP = sheetMap['😴 Sleep Tracker'];
const TRG = sheetMap['⚡ Trigger & Stressor Log'];
const THR = sheetMap['🗓️ Therapy & Appointment Log'];

const REF_NAME = "'📋 Reference Data'";

function dropdown(sheetId, r1, r2, c1, c2, refRange) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: `=${REF_NAME}!${refRange}` }] },
        strict: true,
        showCustomUi: true,
      },
    },
  };
}

function checkbox(sheetId, r1, r2, c1, c2) {
  return {
    setDataValidation: {
      range: gridRange(sheetId, r1, r2, c1, c2),
      rule: {
        condition: { type: 'BOOLEAN' },
        strict: true,
        showCustomUi: true,
      },
    },
  };
}

(async () => {
  const reqs = [];

  // Daily Symptom Log (rows 2-61 = idx 1-60)
  // Col C = Episode Type (Ref A2:A5)
  reqs.push(dropdown(DSL, 1, 61, 2, 3, '$A$2:$A$5'));
  // Col F = Medication Taken checkbox
  reqs.push(checkbox(DSL, 1, 61, 5, 6));

  // Medication & Treatment Tracker (rows 2-31 = idx 1-30)
  // Col D = Frequency (Ref C2:C6)
  reqs.push(dropdown(MED, 1, 31, 3, 4, '$C$2:$C$6'));
  // Col E = Taken checkbox
  reqs.push(checkbox(MED, 1, 31, 4, 5));

  // Sleep Tracker (rows 2-31 = idx 1-30)
  // Col C = Sleep Quality (Ref D2:D5)
  reqs.push(dropdown(SLP, 1, 31, 2, 3, '$D$2:$D$5'));

  // Trigger & Stressor Log (rows 2-21 = idx 1-20)
  // Col C = Trigger Category (Ref E2:E8)
  reqs.push(dropdown(TRG, 1, 21, 2, 3, '$E$2:$E$8'));
  // Col D = Impact Severity (Ref F2:F4)
  reqs.push(dropdown(TRG, 1, 21, 3, 4, '$F$2:$F$4'));

  // Therapy & Appointment Log (rows 2-16 = idx 1-15)
  // Col B = Provider Type (Ref G2:G5)
  reqs.push(dropdown(THR, 1, 16, 1, 2, '$G$2:$G$5'));
  // Col E = Follow-Up Needed checkbox
  reqs.push(checkbox(THR, 1, 16, 4, 5));

  await batchUpdate(id, reqs, 'validation');
  console.log('Validation complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
