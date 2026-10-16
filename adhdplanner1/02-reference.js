'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title row
  vals.push({ range: `${S}!A1`, values: [['REFERENCE DATA — ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,16), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Column headers row 2
  const HEADERS = ['Task Status','Focus Stage','Today Capacity','Priority','Effort Level',
    'Time Estimate','Task Type','Rescue Action','Day Type','Energy Match',
    'Yes / No','Week Start','Month Names','Day Names','',''];
  vals.push({ range: `${S}!A2`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  // Data columns
  const COLS = {
    A: ['Inbox','Ready','In Progress','Waiting','Scheduled','Completed','Rescheduled','Archived'],
    B: ['Now','Next','Later','Event','Revisit'],
    C: ['Low','Medium','High'],
    D: ['Must Do','Important','Useful','Optional'],
    E: ['Quick','Standard','Deep Focus'],
    F: ['5 min','10 min','15 min','30 min','45 min','60 min','90+ min','Unsure'],
    G: ['Task','Event','Appointment','Deadline','Reminder','Errand','Follow-Up','Personal','Work','School','Household','Other'],
    H: ['Keep','Reschedule','Break Down','Archive'],
    I: ['Normal Day','Low-Capacity Day','Busy Day','Reset Day','Custom'],
    J: ['Low Effort','Medium Effort','High Effort','Flexible'],
    K: ['Yes','No'],
    L: ['Monday','Sunday'],
    M: ['January','February','March','April','May','June','July','August','September','October','November','December'],
    N: ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
  };

  Object.entries(COLS).forEach(([col, items]) => {
    const ci = col.charCodeAt(0) - 65;
    items.forEach((v, ri) => {
      vals.push({ range: `${S}!${col}${ri+3}`, values: [[v]] });
    });
    fmt.push({ repeatCell: { range: gridRange(SID,2,2+items.length,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.white), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });

  // Column widths
  const WIDTHS = [110,100,110,100,100,90,120,110,130,110,70,90,110,110];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  await valuesBatchUpdate(id, vals, '02-reference values');
  await batchUpdate(id, fmt, '02-reference format');
  console.log('✅  Reference Data done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
