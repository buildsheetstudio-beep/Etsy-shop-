'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CKL = sheetMap['✅ Planning Checklist'];

// [done, timeframe, task, priority]
const TASKS = [
  // 3+ Months Out
  [true,  '3+ Months Out', 'Agree on reunion date, duration, and location',                       'High'],
  [true,  '3+ Months Out', 'Set up a family group chat or email list',                            'High'],
  [true,  '3+ Months Out', 'Research and book venue or accommodation',                            'High'],
  [true,  '3+ Months Out', 'Set a budget and decide how costs will be shared',                    'High'],
  [true,  '3+ Months Out', 'Send save-the-date to all family members',                            'High'],
  [true,  '3+ Months Out', 'Create a preliminary guest list',                                     'Medium'],
  [true,  '3+ Months Out', 'Research activities suitable for all ages',                           'Medium'],
  [true,  '3+ Months Out', 'Designate a co-organiser or planning committee',                      'Medium'],
  // 2 Months Out
  [true,  '2 Months Out',  'Send formal invitations with RSVP link/form',                         'High'],
  [true,  '2 Months Out',  'Collect dietary requirements from all guests',                        'High'],
  [true,  '2 Months Out',  'Plan the weekend/event itinerary',                                    'High'],
  [true,  '2 Months Out',  'Book any activities or experiences in advance',                       'Medium'],
  [true,  '2 Months Out',  'Arrange transport coordination for out-of-towners',                   'Medium'],
  [true,  '2 Months Out',  'Begin planning the menu and potluck assignments',                     'Medium'],
  [true,  '2 Months Out',  'Organise accommodation assignments',                                  'High'],
  // 6 Weeks Out
  [false, '6 Weeks Out',   'Chase RSVPs from non-responders',                                     'High'],
  [false, '6 Weeks Out',   'Confirm final headcount with venue/caterer',                          'High'],
  [false, '6 Weeks Out',   'Order any printed materials (name tags, banners, signs)',             'Medium'],
  [false, '6 Weeks Out',   'Plan and purchase reunion merchandise or keepsakes (optional)',       'Low'],
  [false, '6 Weeks Out',   'Arrange for event photographer or designate a family photographer',   'Medium'],
  [false, '6 Weeks Out',   'Purchase non-perishable supplies (plates, napkins, decorations)',     'Medium'],
  // 4 Weeks Out
  [false, '4 Weeks Out',   'Confirm potluck assignments with all family units',                   'High'],
  [false, '4 Weeks Out',   'Finalise and share the itinerary with all guests',                   'High'],
  [false, '4 Weeks Out',   'Purchase prizes and supplies for activities/games',                   'Medium'],
  [false, '4 Weeks Out',   'Confirm all venue bookings and arrangements',                         'High'],
  [false, '4 Weeks Out',   'Plan and share driving directions and parking information',           'Medium'],
  [false, '4 Weeks Out',   'Create a family trivia quiz with fun facts about the family',        'Medium'],
  // 2 Weeks Out
  [false, '2 Weeks Out',   'Prepare name tags (include family branch colour coding)',             'Medium'],
  [false, '2 Weeks Out',   'Assemble welcome bags or arrival packs',                             'Medium'],
  [false, '2 Weeks Out',   'Confirm all arrivals and departure times',                           'High'],
  [false, '2 Weeks Out',   'Brief any family members helping with setup',                        'Medium'],
  [false, '2 Weeks Out',   'Download and print any maps or local guides',                        'Low'],
  // 1 Week Out
  [false, '1 Week Out',    'Confirm final attendance numbers one last time',                      'High'],
  [false, '1 Week Out',    'Prepare all decorations and signage',                                 'Medium'],
  [false, '1 Week Out',    'Create a family photo slideshow or memory board',                    'Medium'],
  [false, '1 Week Out',    'Purchase fresh food and perishables',                                 'High'],
  // Day Before
  [false, 'Day Before',    'Set up venue, decorations, and activity areas',                       'High'],
  [false, 'Day Before',    'Prepare food that can be made in advance',                           'High'],
  [false, 'Day Before',    'Set out name tags, welcome materials, and signage',                  'Medium'],
  [false, 'Day Before',    'Charge all cameras and devices',                                      'Low'],
  // Day Of
  [false, 'Day Of',        'Greet and welcome all arriving families',                             'High'],
  [false, 'Day Of',        'Run the itinerary and activities as planned',                         'High'],
  [false, 'Day Of',        'Capture photos, videos, and memories throughout',                    'Medium'],
  // After Reunion
  [false, 'After Reunion', 'Send a thank-you message and share the photo album',                 'High'],
  [false, 'After Reunion', 'Send cost-sharing summary to all contributing families',             'Medium'],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'✅ Planning Checklist'";
  const NCOLS = 8; // A-H

  // Column widths
  [40,150,320,90,140,110,120,220].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: CKL, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner (full width, no col A standalone since no col freeze needed)
  reqs.push({ mergeCells: { range: gridRange(CKL,0,1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(CKL,0,1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CKL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  ✅  Planning Checklist — Henderson Family Reunion 2025']] });

  // Row 1: summary header
  reqs.push({ mergeCells: { range: gridRange(CKL,1,2,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(CKL,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CKL, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Summary formula — single merged cell
  const sumFormula = `="  ✔ Done: "&COUNTIF(A4:A200,TRUE)&"  |  Total: "&COUNTA(C4:C200)&"  |  % Complete: "&TEXT(IFERROR(COUNTIF(A4:A200,TRUE)/COUNTA(C4:C200),0),"0%")&"  |  Overdue: "&COUNTIFS(A4:A200,FALSE,F4:F200,"<"&TODAY())`;
  vals.push({ range: `${TAB}!A2`, values: [[sumFormula]] });

  // Row 2: column headers (0-indexed row 2 = GSheets row 3)
  const hdrs = ['✓','TIMEFRAME','TASK','PRIORITY','ASSIGNED TO','DUE DATE','STATUS','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(CKL,2,3,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CKL, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [hdrs] });

  // Data rows (0-indexed 3+ = GSheets row 4+)
  reqs.push({ repeatCell: {
    range: gridRange(CKL,3,3+TASKS.length,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.white),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  // Alternating rows
  for (let i = 0; i < TASKS.length; i++) {
    if (i % 2 === 1) {
      reqs.push({ repeatCell: {
        range: gridRange(CKL, 3+i, 4+i, 0, NCOLS),
        cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) }},
        fields: 'userEnteredFormat.backgroundColor',
      }});
    }
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CKL, dimension: 'ROWS', startIndex: 3, endIndex: 3+TASKS.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Date format col F
  reqs.push({ repeatCell: {
    range: gridRange(CKL,3,3+TASKS.length,5,6),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' }}},
    fields: 'userEnteredFormat.numberFormat',
  }});
  // Center col A
  reqs.push({ repeatCell: {
    range: gridRange(CKL,3,3+TASKS.length,0,1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // Add checkbox validation for col A
  reqs.push({ setDataValidation: {
    range: gridRange(CKL,3,3+TASKS.length,0,1),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true, strict: true },
  }});

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(CKL,2,3+TASKS.length,0,NCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Write task data
  for (let i = 0; i < TASKS.length; i++) {
    const [done, tf, task, pri] = TASKS[i];
    const row = 4 + i;
    vals.push({ range: `${TAB}!A${row}`, values: [[done ? 'TRUE' : 'FALSE', tf, task, pri]] });
    // Set status for completed tasks
    if (done) {
      vals.push({ range: `${TAB}!G${row}`, values: [['Done']] });
    } else {
      vals.push({ range: `${TAB}!G${row}`, values: [['Not Started']] });
    }
  }

  await batchUpdate(id, reqs, 'checklist-format');
  await valuesBatchUpdate(id, vals, 'checklist-values');
  console.log('Planning Checklist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
