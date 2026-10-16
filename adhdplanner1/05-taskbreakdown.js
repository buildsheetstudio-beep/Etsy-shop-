'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Task Breakdown'];
const S = "'Task Breakdown'";

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,600,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title row
  vals.push({ range: `${S}!A1`, values: [['TASK BREAKDOWN — ADHD FOCUS FLOW PLANNER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,11), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Instructions row 2
  vals.push({ range: `${S}!A2`, values: [['Break overwhelming tasks into manageable micro-steps. Link each parent task from Master Tasks. Check off each step as you go.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,11), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.lavTint), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE', padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Summary cards row 3
  const LBL3 = ['Parent Tasks','Total Steps','Completed Steps','In Progress','Blocked','Completion %','','','','',''];
  vals.push({ range: `${S}!A3`, values: [LBL3] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.headerMid), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});

  const F3 = [
    `=SUMPRODUCT((LEN(A25:A600)>0)*1)`,
    `=COUNTA(D25:D600)`,
    `=COUNTIF(I25:I600,TRUE)`,
    `=COUNTIF(H25:H600,"In Progress")`,
    `=COUNTIF(H25:H600,"Blocked")`,
    `=IFERROR(TEXT(COUNTIF(I25:I600,TRUE)/COUNTA(D25:D600),"0%"),"0%")`,
    '','','','','',
  ];
  vals.push({ range: `${S}!A4`, values: [F3] });
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.lavTint), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.lavender), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 4 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Spacer row 5
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,11), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // Parent task section header rows — each parent gets a 2-row header:
  // Row A: merged parent task header (purple background)
  // Row B: sub-step column headers
  // Then step rows (up to ~10 per parent)
  // We'll define 20 parent tasks with sample steps, then add blank rows for more

  // Parent tasks header row 6-7
  const PARENT_HDRS = ['Parent Task ID','Parent Task Name','','','','','','','','',''];
  vals.push({ range: `${S}!A6`, values: [PARENT_HDRS] });
  fmt.push({ mergeCells: { range: gridRange(SID,5,6,2,11), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,5,6,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.lavender), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Step column headers row 7
  const STEP_HDRS = ['Parent ID','Parent Task','Step #','Step Description','Time','Effort','Order','Status','Done?','Notes',''];
  vals.push({ range: `${S}!A7`, values: [STEP_HDRS] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Parent task header rows (rows 8-24, one per major parent task used as visual separators)
  const PARENTS = [
    { id: 'TASK-00001', name: 'Finish Q3 Project Report' },
    { id: 'TASK-00005', name: 'Prepare Presentation Slides' },
    { id: 'TASK-00013', name: 'Organize Digital Photos' },
    { id: 'TASK-00017', name: 'Reorganize Home Office' },
    { id: 'TASK-00020', name: 'Plan Family Vacation for Winter' },
    { id: 'TASK-00043', name: 'Plan Team Bonding Activity' },
    { id: 'TASK-00085', name: 'Plan Home Office Reorganization' },
    { id: 'TASK-00010', name: 'Write Performance Self-Review' },
  ];

  // We'll write them all into the step rows area starting at row 25 (index 24)
  // Format rows 8-24 as "parent section indicator" area
  for (let r = 7; r < 24; r++) {
    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,11), cell: { userEnteredFormat: {
      backgroundColor: hex(C.lavTint), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
    }}, fields: 'userEnteredFormat' }});
  }

  // Parent labels in rows 8-15
  const parentRows = PARENTS.map((p,i) => [p.id, p.name, '', '', '', '', '', '', '', '', '']);
  vals.push({ range: `${S}!A8`, values: parentRows });
  fmt.push({ repeatCell: { range: gridRange(SID,7,7+PARENTS.length,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.lavTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.lavender) },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 7+PARENTS.length }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Spacer row 24
  fmt.push({ repeatCell: { range: gridRange(SID,23,24,0,11), cell: { userEnteredFormat: { backgroundColor: hex(C.softGray) }}, fields: 'userEnteredFormat(backgroundColor)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 23, endIndex: 24 }, properties: { pixelSize: 6 }, fields: 'pixelSize' }});

  // Step data header row 25 repeated
  const SHDRS2 = ['Parent ID','Parent Task','Step #','Step Description','Time Est.','Effort','Order','Status','Done?','Notes',''];
  vals.push({ range: `${S}!A25`, values: [SHDRS2] });
  fmt.push({ repeatCell: { range: gridRange(SID,24,25,0,11), cell: { userEnteredFormat: {
    backgroundColor: hex(C.slate), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 24, endIndex: 25 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Step data rows (alternating background from row 26 = index 25)
  for (let r = 25; r < 600; r++) {
    const bg = r % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,r,r+1,0,11), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
      verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 25, endIndex: 600 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Sample step data
  // [ParentID, ParentTask, StepNum, StepDescription, TimeEst, Effort, Order, Status, Done, Notes]
  const STEPS = [
    // Q3 Report
    ['TASK-00001','Finish Q3 Project Report',1,'Review last quarter\'s data and pull key metrics','30 min','Standard',1,'Completed',true,'Use the BI dashboard'],
    ['TASK-00001','Finish Q3 Project Report',2,'Email Sarah for final sales numbers','5 min','Quick',2,'Completed',true,'She replied — numbers in email'],
    ['TASK-00001','Finish Q3 Project Report',3,'Outline report sections and key talking points','20 min','Standard',3,'Completed',true,''],
    ['TASK-00001','Finish Q3 Project Report',4,'Write executive summary paragraph','30 min','Deep Focus',4,'In Progress',false,'Do this first thing in the morning'],
    ['TASK-00001','Finish Q3 Project Report',5,'Build charts and visualizations','45 min','Standard',5,'Not Started',false,''],
    ['TASK-00001','Finish Q3 Project Report',6,'Draft full report body','90 min','Deep Focus',6,'Not Started',false,'Need 2-hour focus block'],
    ['TASK-00001','Finish Q3 Project Report',7,'Proofread and edit','30 min','Standard',7,'Not Started',false,''],
    ['TASK-00001','Finish Q3 Project Report',8,'Send to manager for review','5 min','Quick',8,'Not Started',false,''],
    // Presentation Slides
    ['TASK-00005','Prepare Presentation Slides',1,'Review meeting agenda and goals','10 min','Quick',1,'Completed',true,''],
    ['TASK-00005','Prepare Presentation Slides',2,'Gather data and supporting materials','30 min','Standard',2,'Completed',true,''],
    ['TASK-00005','Prepare Presentation Slides',3,'Create slide structure / outline','20 min','Standard',3,'In Progress',false,'7 slides planned'],
    ['TASK-00005','Prepare Presentation Slides',4,'Design title and section divider slides','30 min','Standard',4,'Not Started',false,''],
    ['TASK-00005','Prepare Presentation Slides',5,'Build content slides','60 min','Deep Focus',5,'Not Started',false,''],
    ['TASK-00005','Prepare Presentation Slides',6,'Add charts, tables, and visuals','45 min','Standard',6,'Not Started',false,''],
    ['TASK-00005','Prepare Presentation Slides',7,'Practice run-through (timed)','30 min','Standard',7,'Not Started',false,'Aim for under 20 min'],
    ['TASK-00005','Prepare Presentation Slides',8,'Final review and polish','20 min','Quick',8,'Not Started',false,''],
    // Organize Digital Photos
    ['TASK-00013','Organize Digital Photos',1,'Download all photos from phone and camera','15 min','Quick',1,'Not Started',false,''],
    ['TASK-00013','Organize Digital Photos',2,'Sort by year into folders','45 min','Standard',2,'Not Started',false,'Start with 2022-2026'],
    ['TASK-00013','Organize Digital Photos',3,'Delete obvious duplicates and blurry photos','60 min','Standard',3,'Not Started',false,''],
    ['TASK-00013','Organize Digital Photos',4,'Create event albums for major occasions','90 min','Deep Focus',4,'Not Started',false,''],
    ['TASK-00013','Organize Digital Photos',5,'Back up to cloud storage','30 min','Standard',5,'Not Started',false,'Google Photos or iCloud'],
    ['TASK-00013','Organize Digital Photos',6,'Share family albums with family members','15 min','Quick',6,'Not Started',false,''],
    // Reorganize Home Office
    ['TASK-00017','Reorganize Home Office',1,'Take "before" photos','5 min','Quick',1,'Not Started',false,''],
    ['TASK-00017','Reorganize Home Office',2,'Remove everything from desk','15 min','Quick',2,'Not Started',false,''],
    ['TASK-00017','Reorganize Home Office',3,'Sort items into keep/donate/trash','45 min','Standard',3,'Not Started',false,''],
    ['TASK-00017','Reorganize Home Office',4,'Donate or discard unwanted items','20 min','Quick',4,'Not Started',false,''],
    ['TASK-00017','Reorganize Home Office',5,'Deep clean desk surface and shelves','20 min','Quick',5,'Not Started',false,''],
    ['TASK-00017','Reorganize Home Office',6,'Organize cables and tech accessories','30 min','Standard',6,'Not Started',false,'Cable ties in junk drawer'],
    ['TASK-00017','Reorganize Home Office',7,'Set up ergonomic workspace layout','30 min','Standard',7,'Not Started',false,''],
    ['TASK-00017','Reorganize Home Office',8,'Label drawers and storage','20 min','Quick',8,'Not Started',false,''],
    ['TASK-00017','Reorganize Home Office',9,'Take "after" photos','5 min','Quick',9,'Not Started',false,''],
    // Family Vacation
    ['TASK-00020','Plan Family Vacation for Winter',1,'Agree on destination options with family','30 min','Standard',1,'Not Started',false,''],
    ['TASK-00020','Plan Family Vacation for Winter',2,'Set budget for trip','20 min','Standard',2,'Not Started',false,''],
    ['TASK-00020','Plan Family Vacation for Winter',3,'Research flight options and prices','45 min','Standard',3,'Not Started',false,''],
    ['TASK-00020','Plan Family Vacation for Winter',4,'Book flights','15 min','Quick',4,'Not Started',false,''],
    ['TASK-00020','Plan Family Vacation for Winter',5,'Research and book accommodations','60 min','Standard',5,'Not Started',false,''],
    ['TASK-00020','Plan Family Vacation for Winter',6,'Research activities and attractions','45 min','Standard',6,'Not Started',false,''],
    ['TASK-00020','Plan Family Vacation for Winter',7,'Create day-by-day itinerary','60 min','Deep Focus',7,'Not Started',false,''],
    ['TASK-00020','Plan Family Vacation for Winter',8,'Pack list and preparation','30 min','Standard',8,'Not Started',false,''],
    // Team Bonding
    ['TASK-00043','Plan Team Bonding Activity',1,'Survey team on preferred activity types','10 min','Quick',1,'Not Started',false,'Use Google Form'],
    ['TASK-00043','Plan Team Bonding Activity',2,'Review survey results','10 min','Quick',2,'Not Started',false,''],
    ['TASK-00043','Plan Team Bonding Activity',3,'Research 3 activity options within budget','45 min','Standard',3,'Not Started',false,'Budget: $500'],
    ['TASK-00043','Plan Team Bonding Activity',4,'Get manager approval for budget','10 min','Quick',4,'Not Started',false,''],
    ['TASK-00043','Plan Team Bonding Activity',5,'Book activity and send calendar invite','15 min','Quick',5,'Not Started',false,''],
    ['TASK-00043','Plan Team Bonding Activity',6,'Send details and any required info to team','10 min','Quick',6,'Not Started',false,''],
    // Performance Self-Review
    ['TASK-00010','Write Performance Self-Review',1,'Review goals set at start of year','20 min','Standard',1,'Not Started',false,'Check email from January'],
    ['TASK-00010','Write Performance Self-Review',2,'List major accomplishments and wins','30 min','Standard',2,'Not Started',false,''],
    ['TASK-00010','Write Performance Self-Review',3,'List areas for development','20 min','Standard',3,'Not Started',false,'Be honest but constructive'],
    ['TASK-00010','Write Performance Self-Review',4,'Gather feedback from peers and manager','20 min','Standard',4,'Not Started',false,''],
    ['TASK-00010','Write Performance Self-Review',5,'Write first draft of self-review','60 min','Deep Focus',5,'Not Started',false,''],
    ['TASK-00010','Write Performance Self-Review',6,'Review and polish draft','30 min','Standard',6,'Not Started',false,''],
    ['TASK-00010','Write Performance Self-Review',7,'Submit review','5 min','Quick',7,'Not Started',false,''],
    // Home Office Reorganization (Task-00085)
    ['TASK-00085','Plan Home Office Reorganization',1,'Measure office dimensions','15 min','Quick',1,'Not Started',false,''],
    ['TASK-00085','Plan Home Office Reorganization',2,'Sketch new layout options','30 min','Standard',2,'Not Started',false,''],
    ['TASK-00085','Plan Home Office Reorganization',3,'Research furniture options online','45 min','Standard',3,'Not Started',false,''],
    ['TASK-00085','Plan Home Office Reorganization',4,'Order any new furniture or accessories','20 min','Quick',4,'Not Started',false,''],
    ['TASK-00085','Plan Home Office Reorganization',5,'Wait for delivery and clear the room','10 min','Quick',5,'Not Started',false,''],
    ['TASK-00085','Plan Home Office Reorganization',6,'Set up new layout','90 min','Deep Focus',6,'Not Started',false,''],
    ['TASK-00085','Plan Home Office Reorganization',7,'Test ergonomics and adjust','20 min','Standard',7,'Not Started',false,''],
  ];

  vals.push({ range: `${S}!A26`, values: STEPS });

  // Done checkbox column I formatting note — set BooleanCondition via validation
  // (handled in validation script)

  // Column widths
  const WIDTHS = [90,180,60,280,80,90,60,100,60,180,60];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze rows only (title row merges across all cols)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 25 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '05-taskbreakdown values');
  await batchUpdate(id, fmt, '05-taskbreakdown format');
  console.log('✅  Task Breakdown done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
