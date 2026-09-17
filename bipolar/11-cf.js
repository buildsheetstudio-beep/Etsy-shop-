'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const MED = sheetMap['💊 Medication & Treatment Tracker'];
const THR = sheetMap['🗓️ Therapy & Appointment Log'];

(async () => {
  const reqs = [];

  // Medication: Taken (col E = idx 4) = TRUE → sage; FALSE/blank → neutral gray
  // Data rows 2-31 (idx 1-30)
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(MED, 1, 31, 0, 7)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2=TRUE' }] },
          format: { backgroundColor: hex(C.sage) },
        },
      },
      index: 0,
    },
  });
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(MED, 1, 31, 0, 7)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E2=FALSE' }] },
          format: { backgroundColor: hex(C.missedGray) },
        },
      },
      index: 1,
    },
  });

  // Therapy: Follow-Up Needed (col E = idx 4) = TRUE AND Next Appt Date (col F = idx 5) is blank → amber highlight
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(THR, 1, 16, 0, 7)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($E2=TRUE,$F2="")' }] },
          format: { backgroundColor: hex(C.lightAmber) },
        },
      },
      index: 0,
    },
  });

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('CF rules complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
