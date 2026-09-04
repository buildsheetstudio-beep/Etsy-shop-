'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Reference']; // 6

(async () => {
  const reqs = [];

  // Column widths
  [0,2,4,6,8,10].forEach(ci => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: 160 }, fields: 'pixelSize' } });
  });

  // Header row format
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 12),
      cell: { userEnteredFormat: { backgroundColor: hex(C.dustyLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  // Alt rows
  for (let r = 1; r <= 13; r += 2) {
    reqs.push({ repeatCell: { range: gridRange(SID, r, r+1, 0, 12), cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderWhite) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  await batchUpdate(id, reqs, 'ref-format');

  await valuesBatchUpdate(id, [
    { range: 'Reference!A1', values: [['Life Areas']] },
    { range: 'Reference!A2:A8', values: [['Health & Fitness'],['Career & Purpose'],['Finance & Wealth'],['Relationships'],['Personal Growth'],['Fun & Adventure'],['Other']] },
    { range: 'Reference!C1', values: [['Statuses']] },
    { range: 'Reference!C2:C7', values: [['Not Started'],['In Progress'],['Complete'],['On Hold'],['Abandoned'],['Deferred']] },
    { range: 'Reference!E1', values: [['Quarters']] },
    { range: 'Reference!E2:E5', values: [['Q1: Jan–Mar'],['Q2: Apr–Jun'],['Q3: Jul–Sep'],['Q4: Oct–Dec']] },
    { range: 'Reference!G1', values: [['Months']] },
    { range: 'Reference!G2:G13', values: [['January'],['February'],['March'],['April'],['May'],['June'],['July'],['August'],['September'],['October'],['November'],['December']] },
    { range: 'Reference!I1', values: [['Priorities']] },
    { range: 'Reference!I2:I4', values: [['High'],['Medium'],['Low']] },
    { range: 'Reference!K1', values: [['Frequency']] },
    { range: 'Reference!K2:K7', values: [['Daily'],['6x/week'],['5x/week'],['3x/week'],['Weekly'],['As Needed']] },
  ], 'ref-values');

  console.log('Reference tab complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
