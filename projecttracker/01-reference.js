'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

// All lookup lists
const LISTS = {
  'Project Statuses':       ['Not Started','Planning','Active','On Hold','At Risk','Complete','Cancelled','Archived'],
  'Task Statuses':          ['Backlog','Not Started','In Progress','Waiting','Blocked','Review','Complete','Cancelled'],
  'Task Priorities':        ['Low','Medium','High','Urgent'],
  'Project Priorities':     ['Low','Medium','High','Critical'],
  'Project Categories':     ['Client Project','Business Operations','Marketing','Product Development','Finance','Administrative','Personal','Creative','Event','Education','Home','Other'],
  'Task Types':             ['Task','Subtask','Milestone Task','Meeting','Approval','Deliverable','Follow-Up','Recurring Task','Other'],
  'Dependency Types':       ['None','Finish-to-Start','Start-to-Start','Finish-to-Finish','Start-to-Finish'],
  'Importance':             ['Important','Not Important'],
  'Urgency':                ['Urgent','Not Urgent'],
  'Recurrence Frequencies': ['None','Daily','Weekdays','Weekly','Biweekly','Monthly','Quarterly','Semiannual','Annual','Custom Interval'],
  'Recurrence Units':       ['Days','Weeks','Months','Years'],
  'Recurrence Statuses':    ['Active','Paused','Complete','Archived'],
  'Milestone Statuses':     ['Not Started','In Progress','Achieved','Delayed','Cancelled'],
  'Project Health':         ['On Track','Monitor','At Risk','Delayed','Complete','Insufficient Data'],
  'Workload Status':        ['Light','Balanced','Busy','Review Capacity','No Tasks'],
  'Yes / No':               ['Yes','No'],
  'Days of Week':           ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
  'Assignees':              ['Sarah Chen','Marcus Webb','Olivia Park','Tyler Banks','Zara Osei','Self','TBD'],
  'Pct Complete':           ['0%','25%','50%','75%','100%'],
};

(async () => {
  const fmt = [];
  const vals = [];

  // Title banner
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  fmt.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 14, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  vals.push({ range: `${S}!A1`, values: [['REFERENCE DATA — DO NOT EDIT']] });

  let col = 0;
  const entries = Object.entries(LISTS);
  entries.forEach(([listName, items], li) => {
    // Each list gets 2 columns; wrap after 5 lists
    const c = (li % 5) * 2;
    const baseRow = Math.floor(li / 5) * (Math.max(...entries.slice(Math.floor(li/5)*5, Math.floor(li/5)*5+5).map(([,v]) => v.length)) + 3) + 2;

    // Header
    fmt.push({ mergeCells: { range: gridRange(SID, baseRow, baseRow+1, c, c+2), mergeType: 'MERGE_ALL' } });
    fmt.push({
      repeatCell: {
        range: gridRange(SID, baseRow, baseRow+1, c, c+2),
        cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}${baseRow+1}`, values: [[listName]] });

    // Items
    items.forEach((item, ii) => {
      const r = baseRow + 1 + ii;
      fmt.push({
        repeatCell: {
          range: gridRange(SID, r, r+1, c, c+2),
          cell: { userEnteredFormat: { backgroundColor: ii % 2 === 0 ? hex(C.panel) : hex(C.altRow), textFormat: { fontSize: 9, fontFamily: 'Arial' } } },
          fields: 'userEnteredFormat(backgroundColor,textFormat)',
        },
      });
      vals.push({ range: `${S}!${String.fromCharCode(65+c)}${r+1}`, values: [[item]] });
    });
  });

  // Column widths
  for (let i = 0; i < 10; i++) {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: 140 }, fields: 'pixelSize' } });
  }

  // Hide the tab
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, hidden: true }, fields: 'hidden' } });
  // Freeze row 1
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' } });

  await batchUpdate(id, fmt, 'reference-fmt');
  await valuesBatchUpdate(id, vals, 'reference-vals');
  console.log('✓ Reference Data complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
