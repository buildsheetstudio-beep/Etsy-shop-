'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const CK = sheetMap['✅ Party Checklist'];

// Layout: Row1=Title, Row2=Summary, Row3=Headers, Rows4-48=45 tasks
// Cols: A=checkbox, B=timeframe, C=task, D=priority, E=due date, F=assigned to, G=status, H=notes
const NCOLS = 8;

const TASKS = [
  // [timeframe, task, priority, status, tick]
  // 8+ Weeks Out (8 tasks) — pre-ticked
  ['8+ Weeks Out','Set the date and time for the shower','High','Done',true],
  ['8+ Weeks Out','Choose a venue (home, restaurant, garden, hired hall)','High','Done',true],
  ['8+ Weeks Out','Agree on the theme and colour palette','High','Done',true],
  ['8+ Weeks Out','Create a budget and assign responsibilities','High','Done',true],
  ['8+ Weeks Out','Compile the initial guest list with mum-to-be','High','Done',true],
  ['8+ Weeks Out','Confirm mum-to-be\'s registry details and share with guests','Medium','Done',true],
  ['8+ Weeks Out','Book the venue if hiring an external space','High','Done',true],
  ['8+ Weeks Out','Plan who will co-host and divide responsibilities','Medium','Done',true],
  // 6 Weeks Out (7 tasks) — pre-ticked
  ['6 Weeks Out','Design and send invitations (physical or digital)','High','Done',true],
  ['6 Weeks Out','Plan the menu — food and drinks','High','Done',true],
  ['6 Weeks Out','Book caterer or assign food contributions','High','Done',true],
  ['6 Weeks Out','Order or plan the shower cake','High','Done',true],
  ['6 Weeks Out','Plan games and activities','Medium','Done',true],
  ['6 Weeks Out','Begin sourcing décor items','Medium','Done',true],
  ['6 Weeks Out','Arrange any special entertainment or photographers','Low','Done',true],
  // 4 Weeks Out (8 tasks) — not ticked
  ['4 Weeks Out','Chase RSVPs — follow up with non-responders','High','Not Started',false],
  ['4 Weeks Out','Finalise headcount and dietary requirements','High','Not Started',false],
  ['4 Weeks Out','Order flowers and florals','Medium','Not Started',false],
  ['4 Weeks Out','Purchase party favours and gift bags','Medium','Not Started',false],
  ['4 Weeks Out','Order or make decorations','Medium','Not Started',false],
  ['4 Weeks Out','Plan the party itinerary and run of show','High','Not Started',false],
  ['4 Weeks Out','Confirm all vendor bookings','High','Not Started',false],
  ['4 Weeks Out','Plan seating and layout','Low','Not Started',false],
  // 2 Weeks Out (6 tasks)
  ['2 Weeks Out','Confirm final guest numbers with caterer','High','Not Started',false],
  ['2 Weeks Out','Purchase non-perishable food and drinks','Medium','Not Started',false],
  ['2 Weeks Out','Prepare party games supplies and prizes','Medium','Not Started',false],
  ['2 Weeks Out','Create printed materials (menus, name cards, signage)','Medium','Not Started',false],
  ['2 Weeks Out','Brief any helpers on their roles','High','Not Started',false],
  ['2 Weeks Out','Confirm parking and transport for guests','Low','Not Started',false],
  // 1 Week Out (5 tasks)
  ['1 Week Out','Purchase fresh flowers (if not using a florist)','Medium','Not Started',false],
  ['1 Week Out','Prepare any DIY décor items','Medium','Not Started',false],
  ['1 Week Out','Prepare gift bags and favours','Medium','Not Started',false],
  ['1 Week Out','Confirm RSVP count one final time','High','Not Started',false],
  ['1 Week Out','Charge all cameras and devices','Low','Not Started',false],
  // Day Before (5 tasks)
  ['Day Before','Set up venue décor and layout','High','Not Started',false],
  ['Day Before','Prepare all food that can be made in advance','High','Not Started',false],
  ['Day Before','Chill drinks and prepare bar area','Medium','Not Started',false],
  ['Day Before','Prepare the welcome table and signage','Medium','Not Started',false],
  ['Day Before','Do a final run-through of the itinerary','High','Not Started',false],
  // Day Of (4 tasks)
  ['Day Of','Final venue setup and fresh flower placement','High','Not Started',false],
  ['Day Of','Set out food and drinks','High','Not Started',false],
  ['Day Of','Welcome guests and run the party!','High','Not Started',false],
  ['Day Of','Capture photos and video throughout','Medium','Not Started',false],
  // After Party (2 tasks)
  ['After Party','Send thank-you messages or cards within 1 week','High','Not Started',false],
  ['After Party','Share photos with guests and mum-to-be','Medium','Not Started',false],
];

(async () => {
  const reqs = [];

  // Column widths
  [40, 150, 300, 90, 110, 140, 120, 220].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: CK, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CK, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CK, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: CK, dimension: 'ROWS', startIndex: 3, endIndex: 48 },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Title row
  reqs.push({ mergeCells: { range: gridRange(CK, 0, 1, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(CK, 0, 1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary row 2
  reqs.push({ repeatCell: {
    range: gridRange(CK, 1, 2, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  // Merge pairs for summary stats
  [[0,2],[2,3],[3,5],[5,6],[6,8]].forEach(([c1,c2]) => {
    reqs.push({ mergeCells: { range: gridRange(CK, 1, 2, c1, c2), mergeType: 'MERGE_ALL' }});
  });

  // Column headers row 3
  reqs.push({ repeatCell: {
    range: gridRange(CK, 2, 3, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows — alternating
  for (let r = 3; r < 48; r++) {
    const bg = (r % 2 === 1) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(CK, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Checkbox col A — ivory + center
  reqs.push({ repeatCell: {
    range: gridRange(CK, 3, 48, 0, 1),
    cell: { userEnteredFormat: { backgroundColor: hex(C.ivory), horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
  }});

  // Task col C — slightly taller line height, wrap
  reqs.push({ repeatCell: {
    range: gridRange(CK, 3, 48, 2, 3),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' }},
    fields: 'userEnteredFormat.wrapStrategy',
  }});

  // Due date col E — center
  reqs.push({ repeatCell: {
    range: gridRange(CK, 3, 48, 4, 5),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'TEXT' } }},
    fields: 'userEnteredFormat(horizontalAlignment,numberFormat)',
  }});

  // Formula cells (B col for timeframe shown as dropdown): ivory bg
  reqs.push({ repeatCell: {
    range: gridRange(CK, 3, 48, 1, 2),
    cell: { userEnteredFormat: { backgroundColor: hex(C.ivory) }},
    fields: 'userEnteredFormat.backgroundColor',
  }});

  // Borders on data table
  reqs.push({ updateBorders: {
    range: gridRange(CK, 2, 48, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Checkbox data validation
  reqs.push({ setDataValidation: {
    range: gridRange(CK, 3, 48, 0, 1),
    rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
  }});

  await batchUpdate(id, reqs, 'checklist-format');

  // Values
  const vals = [];
  vals.push({ range: "'✅ Party Checklist'!A1", values: [['  ✅  Party Checklist — Baby Shower Planning Tasks']] });

  // Summary row
  vals.push({ range: "'✅ Party Checklist'!A2", values: [['TASKS SUMMARY']] });
  vals.push({ range: "'✅ Party Checklist'!C2", values: [['=COUNTA(C4:C200)&" total tasks"']] });
  vals.push({ range: "'✅ Party Checklist'!D2", values: [['=COUNTIF(A4:A200,TRUE)&" done"']] });
  vals.push({ range: "'✅ Party Checklist'!F2", values: [['=IFERROR(TEXT(COUNTIF(A4:A200,TRUE)/COUNTA(C4:C200),"0%")&" complete","—")']] });
  vals.push({ range: "'✅ Party Checklist'!G2", values: [['=COUNTIFS(A4:A200,FALSE,E4:E200,"<"&TODAY())&" overdue"']] });

  // Column headers
  vals.push({ range: "'✅ Party Checklist'!A3", values: [['✓','TIMEFRAME','TASK','PRIORITY','DUE DATE','ASSIGNED TO','STATUS','NOTES']] });

  // Task data — split checkbox col from rest
  const taskRows = TASKS.map(([tf, task, pri, status, tick]) => [tick ? 'TRUE' : 'FALSE', tf, task, pri, '', 'Emma Clarke', status, '']);
  vals.push({ range: "'✅ Party Checklist'!A4", values: taskRows });

  await valuesBatchUpdate(id, vals, 'checklist-values');

  console.log('Party Checklist complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
