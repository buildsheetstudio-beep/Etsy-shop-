'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

// GSheets goal header rows are 3,9,15,21,27,33,39,45,51,57 (with title+header row offset)
const GOAL_GSHEETS_ROWS = [3, 9, 15, 21, 27, 33, 39, 45, 51, 57];

const LISTS = {
  A: ['Life Area','Career & Work','Health & Fitness','Finance & Wealth','Relationships','Personal Growth','Education & Skills','Travel & Adventure','Creativity & Hobbies','Spirituality & Wellbeing','Home & Family'],
  B: ['Goal Status','Not Started','In Progress','On Track','At Risk','Achieved','Paused','Abandoned'],
  C: ['Priority','High','Medium','Low'],
  D: ['Milestone Status','Not Started','In Progress','Done','Skipped'],
  E: ['Action Step Status','Not Started','In Progress','Done','Blocked','Deferred'],
  F: ['Score (1–5)','1','2','3','4','5'],
  G: ['Quarter','Q1','Q2','Q3','Q4'],
};

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'📋 Reference Data'";

  // Column widths
  Object.keys(LISTS).forEach((col, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: 180 }, fields: 'pixelSize',
    }});
  });
  // Col I for goal title helpers (wide)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'COLUMNS', startIndex: 8, endIndex: 9 },
    properties: { pixelSize: 260 }, fields: 'pixelSize',
  }});

  // Title banner (full width)
  reqs.push({ mergeCells: { range: gridRange(REF,0,1,0,10), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(REF,0,1,0,10),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midnightBlue),
      textFormat: { foregroundColor: hex(C.electricGold), bold: true, fontSize: 14 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['📋 Reference Data']] });

  // Header row styling (row 1, 0-indexed)
  reqs.push({ repeatCell: {
    range: gridRange(REF,1,2,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepNavy),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});

  // Data styling
  reqs.push({ repeatCell: {
    range: gridRange(REF,2,60,0,9),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.paleGold),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});

  // Title of each list in row 1 (0-indexed row 1 = GSheets row 2)
  const listHeaders = Object.keys(LISTS).map(col => LISTS[col][0]);
  listHeaders.push('Goal Titles');
  vals.push({ range: `${TAB}!A2`, values: [listHeaders] });

  // List values (start at GSheets row 3 = 0-indexed row 2)
  Object.entries(LISTS).forEach(([col, items]) => {
    vals.push({ range: `${TAB}!${col}3`, values: items.slice(1).map(v => [v]) });
  });

  // Col I: Goal title references (formulas into Goals tab)
  vals.push({ range: `${TAB}!I2`, values: [['Goal Titles']] });
  GOAL_GSHEETS_ROWS.forEach((r, i) => {
    vals.push({ range: `${TAB}!I${3+i}`, values: [[`=IFERROR('🎯 Goals & Milestones'!A${r},"")`]] });
  });

  await batchUpdate(id, reqs, 'ref-format');
  await valuesBatchUpdate(id, vals, 'ref-values');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
