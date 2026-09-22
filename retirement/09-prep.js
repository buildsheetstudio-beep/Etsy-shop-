'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PP = sheetMap['Prep & Party Day'];
const S = "'Prep & Party Day'";
const R = "'Reference Data'";

const HEADERS = ['Task ID','Phase','Task','Category','Assigned To','Start Date','Due Date','Priority','Status','Completed?','Completed Date','Days Remaining','Overdue?','Party-Day Time','Location / Area','Notes'];

// 40 tasks: [phase, task, category, assignedTo, startDate, dueDate, priority, status, completed, completedDate, partyTime, location, notes]
const TASKS = [
  // 8+ Weeks Before
  ['8+ Weeks Before','Book the venue','Venue','Michael Torres','Apr 1, 2026','Apr 15, 2026','Urgent','Complete',true,'Apr 8, 2026','','','Lakeside Event Hall confirmed'],
  ['8+ Weeks Before','Set party budget and planning committee','Planning','Michael Torres','Apr 1, 2026','Apr 15, 2026','High','Complete',true,'Apr 10, 2026','','','$4,500 budget approved'],
  ['8+ Weeks Before','Compile guest list','Guest Management','Angela Torres','Apr 5, 2026','Apr 20, 2026','High','Complete',true,'Apr 18, 2026','','','37 guests total'],
  ['8+ Weeks Before','Select retirement theme and colors','Decor','Rachel Kim','Apr 5, 2026','Apr 20, 2026','Medium','Complete',true,'Apr 16, 2026','','','Celebrating 32 Years of Service — Indigo + Gold'],
  ['8+ Weeks Before','Hire photographer','Photography','Michael Torres','Apr 10, 2026','Apr 30, 2026','High','Complete',true,'Apr 22, 2026','','','Riverside Photos booked'],
  ['8+ Weeks Before','Book DJ / entertainment','Entertainment','Michael Torres','Apr 10, 2026','Apr 30, 2026','High','Complete',true,'Apr 28, 2026','','','DJ Smooth and string quartet booked'],
  // 6 Weeks Before
  ['6 Weeks Before','Design and order invitations','Invitations','Rachel Kim','May 1, 2026','May 10, 2026','High','Complete',true,'May 5, 2026','','','PrintSmart order placed'],
  ['6 Weeks Before','Collect career photos from HR and family','Entertainment','Angela Torres','May 1, 2026','May 15, 2026','High','Complete',true,'May 12, 2026','','','32 years of photos compiled'],
  ['6 Weeks Before','Confirm speakers and assign speech times','Entertainment','Michael Torres','May 5, 2026','May 15, 2026','High','Complete',true,'May 14, 2026','','','10 speakers confirmed. Schedule distributed.'],
  ['6 Weeks Before','Order awards and personalized gifts','Gifts','Michael Torres','May 5, 2026','May 15, 2026','High','Complete',true,'May 9, 2026','','','Crystal award engraved with name and years'],
  // 4 Weeks Before
  ['4 Weeks Before','Send invitations','Invitations','Angela Torres','May 15, 2026','May 20, 2026','High','Complete',true,'May 18, 2026','','','All 37 invitations sent'],
  ['4 Weeks Before','Build career highlight slideshow','Entertainment','Rachel Kim','May 15, 2026','May 30, 2026','High','In Progress',false,'','','','Draft ready; photos being added'],
  ['4 Weeks Before','Confirm catering menu and dietary needs','Food','Michael Torres','May 15, 2026','May 25, 2026','High','Complete',true,'May 22, 2026','','','Lakeside Catering menu confirmed'],
  ['4 Weeks Before','Place decorations order','Decor','Rachel Kim','May 15, 2026','May 25, 2026','Medium','Complete',true,'May 20, 2026','','','Balloon arch, centerpieces, banner ordered'],
  ['4 Weeks Before','Build music playlist and brief DJ','Entertainment','Angela Torres','May 15, 2026','May 28, 2026','High','Complete',true,'May 25, 2026','','','Playlist shared with DJ Smooth'],
  // 3 Weeks Before
  ['3 Weeks Before','Follow up on outstanding RSVPs','Guest Management','Angela Torres','Jun 1, 2026','Jun 7, 2026','High','In Progress',false,'','','','6 non-responses; calls scheduled'],
  ['3 Weeks Before','Confirm final headcount with venue','Venue','Michael Torres','Jun 1, 2026','Jun 7, 2026','High','In Progress',false,'','','','Awaiting RSVP confirmations'],
  ['3 Weeks Before','Assign seating and table groups','Guest Management','Angela Torres','Jun 1, 2026','Jun 10, 2026','Medium','Not Started',false,'','','','Need final headcount first'],
  ['3 Weeks Before','Brief speakers on schedule and timing','Entertainment','Michael Torres','Jun 1, 2026','Jun 10, 2026','High','In Progress',false,'','','','Speaker call scheduled for Jun 6'],
  ['3 Weeks Before','Create event-day run of show','Planning','Michael Torres','Jun 1, 2026','Jun 10, 2026','High','Not Started',false,'','','',''],
  // 2 Weeks Before
  ['2 Weeks Before','Create seating chart and name tags','Guest Management','Angela Torres','Jun 10, 2026','Jun 17, 2026','Medium','Not Started',false,'','','','Print name tags after final count'],
  ['2 Weeks Before','Prepare memory box and guest book','Decor','Rachel Kim','Jun 10, 2026','Jun 17, 2026','Medium','Not Started',false,'','',''],
  ['2 Weeks Before','Order cake from bakery','Food','Michael Torres','Jun 10, 2026','Jun 14, 2026','High','In Progress',false,'','','','Sweet Endings Bakery — balance due Jun 20'],
  ['2 Weeks Before','Finalize slideshow and tribute video','Entertainment','Rachel Kim','Jun 10, 2026','Jun 17, 2026','High','Not Started',false,'','',''],
  ['2 Weeks Before','Confirm AV equipment with rental co.','Entertainment','Michael Torres','Jun 10, 2026','Jun 17, 2026','High','Not Started',false,'','','Riverside AV — Projector, PA confirmed'],
  // 1 Week Before
  ['1 Week Before','Confirm all vendors and final timing','Planning','Michael Torres','Jun 19, 2026','Jun 21, 2026','Urgent','Not Started',false,'','','','DJ, catering, AV, photographer'],
  ['1 Week Before','Print menus, programs, signage','Decor','Rachel Kim','Jun 19, 2026','Jun 22, 2026','High','Not Started',false,'','',''],
  ['1 Week Before','Rehearse slideshow and AV setup','Entertainment','Rachel Kim','Jun 19, 2026','Jun 22, 2026','High','Not Started',false,'','',''],
  ['1 Week Before','Prepare speaker gift bags and awards box','Gifts','Angela Torres','Jun 19, 2026','Jun 23, 2026','Medium','Not Started',false,'','',''],
  ['1 Week Before','Confirm strings quartet arrival time','Entertainment','Michael Torres','Jun 19, 2026','Jun 23, 2026','Medium','Not Started',false,'','','5:00 PM arrival'],
  // 3 Days Before
  ['3 Days Before','Confirm final guest count with caterer','Food','Michael Torres','Jun 24, 2026','Jun 24, 2026','Urgent','Not Started',false,'','','','Submit final count for plated service'],
  ['3 Days Before','Pick up decorations and supplies','Decor','Rachel Kim','Jun 24, 2026','Jun 25, 2026','High','Not Started',false,'','',''],
  ['3 Days Before','Create DJ briefing sheet','Entertainment','Michael Torres','Jun 24, 2026','Jun 25, 2026','High','Not Started',false,'','','Timeline, song list, cues'],
  // 1 Day Before
  ['1 Day Before','Set up venue (early access)','Setup','Event Team','Jun 26, 2026','Jun 26, 2026','Urgent','Not Started',false,'','3:00 PM','Lakeside Event Hall','2-hour setup window'],
  ['1 Day Before','Test all AV equipment','Entertainment','Michael Torres','Jun 26, 2026','Jun 26, 2026','Urgent','Not Started',false,'','4:30 PM','Lakeside Event Hall','Full run-through with DJ'],
  ['1 Day Before','Place centerpieces and name cards','Decor','Rachel Kim','Jun 26, 2026','Jun 26, 2026','High','Not Started',false,'','4:00 PM','Lakeside Event Hall',''],
  // Party Day
  ['Party Day','Welcome guests at entrance','Guest Management','Angela Torres','','','Urgent','Not Started',false,'','5:00 PM','Entrance','With name tags and programs'],
  ['Party Day','Manage gift and memory table','Gifts','Angela Torres','','','High','Not Started',false,'','5:00 PM','Gift Table','Keep organized; note contributors'],
  ['Party Day','Cue speakers per run of show','Entertainment','Michael Torres','','','Urgent','Not Started',false,'','5:30 PM','Stage Area',''],
  // After Party
  ['After Party','Send thank-you notes to vendors','Planning','Michael Torres','Jun 28, 2026','Jul 5, 2026','Medium','Not Started',false,'','','','Venue, catering, AV, photographer'],
  ['After Party','Begin thank-you card writing','Thank-You Cards','Patricia Collins','Jun 28, 2026','Jul 15, 2026','High','Not Started',false,'','','',''],
  ['After Party','Return AV rental equipment','Entertainment','Michael Torres','Jun 29, 2026','Jun 29, 2026','High','Not Started',false,'','','','Riverside AV pickup scheduled'],
];

(async () => {
  const reqs = [];
  const data = [];

  // Background + freeze
  reqs.push({ repeatCell: {
    range: gridRange(PP,0,300,0,16),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:PP, gridProperties:{ frozenRowCount:5 } },
    fields:'gridProperties.frozenRowCount'
  }});

  // Title
  data.push({ range:`${S}!A1`, values:[['PREP & PARTY DAY']] });
  reqs.push({ mergeCells: { range: gridRange(PP,0,1,0,16), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(PP,0,1,0,16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:16, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:PP, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties: { pixelSize:44 }, fields:'pixelSize'
  }});

  // Subtitle
  data.push({ range:`${S}!A2`, values:[['Plan and track every task from 8 weeks out to post-party follow-up']] });
  reqs.push({ mergeCells: { range: gridRange(PP,1,2,0,16), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(PP,1,2,0,16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Info rows 2-3
  data.push({ range:`${S}!A3`, values:[['Task ID, Days Remaining, and Overdue auto-calculate. Mark Completed? checkbox to track progress. Completed rows gray out.']] });
  data.push({ range:`${S}!A4`, values:[['Phases: 8+ Weeks Before → 6 Wks → 4 Wks → 3 Wks → 2 Wks → 1 Wk → 3 Days → 1 Day → Party Day → After Party']] });
  [2,3].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(PP,ri,ri+1,0,16), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(PP,ri,ri+1,0,16),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.secText), fontFamily:'Arial' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:PP, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:22 }, fields:'pixelSize'
    }});
  });

  // ── 6 Summary cards (row index 4) ─────────────────────────────────────────
  const summaryCards = [
    { label:'Days Until Party', formula:`=IFERROR(MAX(0,DATEVALUE("Jun 27, 2026")-TODAY()),"—")`, cols:[0,2,2,5] },
    { label:'Total Tasks',       formula:`=IFERROR(COUNTA(C8:C207),"—")`, cols:[5,7,7,9] },
    { label:'Completed',         formula:`=IFERROR(COUNTIF(J8:J207,TRUE),"—")`, cols:[9,11,11,12] },
    { label:'Remaining',         formula:`=IFERROR(COUNTA(C8:C207)-COUNTIF(J8:J207,TRUE),"—")`, cols:[12,13,13,14] },
    { label:'Overdue',           formula:`=IFERROR(COUNTIF(M8:M207,"Yes"),"—")`, cols:[14,15,15,16] },
  ];

  summaryCards.forEach(card => {
    const [lc1, lc2, vc1, vc2] = card.cols;
    const lr = String.fromCharCode(65+lc1) + '5';
    const vr = String.fromCharCode(65+vc1) + '5';
    data.push({ range:`${S}!${lr}`, values:[[card.label]] });
    data.push({ range:`${S}!${vr}`, values:[[card.formula]] });

    if (lc1 < lc2) {
      reqs.push({ mergeCells: { range: gridRange(PP,4,5,lc1,lc2), mergeType:'MERGE_ALL' } });
    }
    reqs.push({ repeatCell: {
      range: gridRange(PP,4,5,lc1,lc2 > lc1 ? lc2 : lc1+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    if (vc1 < vc2) {
      reqs.push({ mergeCells: { range: gridRange(PP,4,5,vc1,vc2), mergeType:'MERGE_ALL' } });
    }
    reqs.push({ repeatCell: {
      range: gridRange(PP,4,5,vc1,vc2 > vc1 ? vc2 : vc1+1),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold:true, fontSize:13, foregroundColor:hex(C.secondary), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:PP, dimension:'ROWS', startIndex:4, endIndex:5 },
    properties: { pixelSize:38 }, fields:'pixelSize'
  }});

  // ── Completion % row (index 5) ─────────────────────────────────────────────
  data.push({ range:`${S}!A6`, values:[['Completion %:']] });
  data.push({ range:`${S}!C6`, values:[[`=IFERROR(IF(COUNTA(C8:C207)=0,"—",COUNTIF(J8:J207,TRUE)/COUNTA(C8:C207)),"—")`]] });
  reqs.push({ repeatCell: {
    range: gridRange(PP,5,6,0,2),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'RIGHT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ repeatCell: {
    range: gridRange(PP,5,6,2,4),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.panel),
      textFormat: { bold:true, fontSize:13, foregroundColor:hex(C.secondary), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE',
      numberFormat:{ type:'PERCENT', pattern:'0%' }
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:PP, dimension:'ROWS', startIndex:5, endIndex:6 },
    properties: { pixelSize:28 }, fields:'pixelSize'
  }});

  // ── Header row (index 6) ───────────────────────────────────────────────────
  data.push({ range:`${S}!A7`, values:[HEADERS] });
  reqs.push({ repeatCell: {
    range: gridRange(PP,6,7,0,16),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:PP, dimension:'ROWS', startIndex:6, endIndex:7 },
    properties: { pixelSize:40 }, fields:'pixelSize'
  }});

  // ── Sample data rows 8-49 (index 7-48) ────────────────────────────────────
  TASKS.forEach(([phase, task, category, assignedTo, startDate, dueDate, priority, status, completed, completedDate, partyTime, location, notes], i) => {
    const row = 8 + i;
    const ri = row - 1;
    data.push({ range:`${S}!A${row}`, values:[[`=IF(C${row}="","","TASK-"&TEXT(ROW()-7,"000"))`]] });
    data.push({ range:`${S}!B${row}`, values:[[phase]] });
    data.push({ range:`${S}!C${row}`, values:[[task]] });
    data.push({ range:`${S}!D${row}`, values:[[category]] });
    data.push({ range:`${S}!E${row}`, values:[[assignedTo]] });
    if (startDate) data.push({ range:`${S}!F${row}`, values:[[startDate]] });
    if (dueDate) data.push({ range:`${S}!G${row}`, values:[[dueDate]] });
    data.push({ range:`${S}!H${row}`, values:[[priority]] });
    data.push({ range:`${S}!I${row}`, values:[[status]] });
    data.push({ range:`${S}!J${row}`, values:[[completed]] });
    if (completedDate) data.push({ range:`${S}!K${row}`, values:[[completedDate]] });
    // L: Days Remaining formula
    data.push({ range:`${S}!L${row}`, values:[[`=IF(G${row}="","",IF(J${row}=TRUE,0,G${row}-TODAY()))`]] });
    // M: Overdue formula
    data.push({ range:`${S}!M${row}`, values:[[`=IF(C${row}="","",IF(AND(J${row}<>TRUE,G${row}<>"",G${row}<TODAY()),"Yes","No"))`]] });
    if (partyTime) data.push({ range:`${S}!N${row}`, values:[[partyTime]] });
    if (location) data.push({ range:`${S}!O${row}`, values:[[location]] });
    if (notes) data.push({ range:`${S}!P${row}`, values:[[notes]] });
  });

  // Formula rows for remaining 50-207 (index 49-206)
  for (let row = 50; row <= 207; row++) {
    data.push({ range:`${S}!A${row}`, values:[[`=IF(C${row}="","","TASK-"&TEXT(ROW()-7,"000"))`]] });
    data.push({ range:`${S}!L${row}`, values:[[`=IF(G${row}="","",IF(J${row}=TRUE,0,G${row}-TODAY()))`]] });
    data.push({ range:`${S}!M${row}`, values:[[`=IF(C${row}="","",IF(AND(J${row}<>TRUE,G${row}<>"",G${row}<TODAY()),"Yes","No"))`]] });
  }

  // Alternating rows
  for (let ri = 7; ri < 207; ri += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(PP,ri,ri+1,0,16),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) }},
      fields:'userEnteredFormat(backgroundColor)'
    }});
  }

  // Formula col styling A,L,M
  [0,11,12].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(PP,7,207,ci,ci+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat:{fontSize:9,fontFamily:'Arial'} }},
      fields:'userEnteredFormat(backgroundColor,textFormat)'
    }});
  });

  // Dropdowns
  const dvReqs = [];
  // Phase dropdown col B (index 1)
  dvReqs.push({ setDataValidation: {
    range: gridRange(PP,7,207,1,2),
    rule: { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$Q$2:$Q$11` }] }, showCustomUi:true, strict:false }
  }});
  // Priority dropdown col H (index 7)
  dvReqs.push({ setDataValidation: {
    range: gridRange(PP,7,207,7,8),
    rule: { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$U$2:$U$5` }] }, showCustomUi:true, strict:false }
  }});
  // Status dropdown col I (index 8)
  dvReqs.push({ setDataValidation: {
    range: gridRange(PP,7,207,8,9),
    rule: { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$S$2:$S$6` }] }, showCustomUi:true, strict:false }
  }});
  // Completed checkbox col J (index 9)
  dvReqs.push({ setDataValidation: {
    range: gridRange(PP,7,207,9,10),
    rule: { condition:{ type:'BOOLEAN' }, showCustomUi:true }
  }});

  // Status CF col I (index 8)
  const statusCF = [
    { val:'Complete',    bg:C.success,   text:C.mainText },
    { val:'In Progress', bg:C.mutedBlue, text:C.mainText },
    { val:'Waiting',     bg:C.warning,   text:C.mainText },
    { val:'Not Started', bg:C.altRow,    text:C.mainText },
    { val:'Cancelled',   bg:C.border,    text:C.secText },
  ];
  statusCF.forEach((r, idx) => {
    reqs.push({ addConditionalFormatRule: { index:idx, rule: {
      ranges:[gridRange(PP,7,207,8,9)],
      booleanRule: {
        condition:{ type:'TEXT_EQ', values:[{userEnteredValue:r.val}] },
        format:{ backgroundColor:hex(r.bg), textFormat:{foregroundColor:hex(r.text)} }
      }
    }}});
  });

  // Priority CF col H (index 7)
  reqs.push({ addConditionalFormatRule: { index:5, rule: {
    ranges:[gridRange(PP,7,207,7,8)],
    booleanRule: {
      condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Urgent'}] },
      format:{ backgroundColor:hex(C.attention), textFormat:{foregroundColor:hex(C.white),bold:true} }
    }
  }}});
  reqs.push({ addConditionalFormatRule: { index:6, rule: {
    ranges:[gridRange(PP,7,207,7,8)],
    booleanRule: {
      condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'High'}] },
      format:{ backgroundColor:hex(C.warning) }
    }
  }}});

  // Full-row: completed = sage tint
  reqs.push({ addConditionalFormatRule: { index:7, rule: {
    ranges:[gridRange(PP,7,207,0,16)],
    booleanRule: {
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=$J8=TRUE`}] },
      format:{ backgroundColor:hex('#E8F0E8'), textFormat:{foregroundColor:hex(C.secText)} }
    }
  }}});

  // Full-row: overdue = rust tint
  reqs.push({ addConditionalFormatRule: { index:8, rule: {
    ranges:[gridRange(PP,7,207,0,16)],
    booleanRule: {
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=AND($M8="Yes",$J8=FALSE)`}] },
      format:{ backgroundColor:hex('#F5E8E8') }
    }
  }}});

  // Column widths
  const colW = [80,120,200,110,120,90,90,80,100,70,100,80,70,80,130,200];
  colW.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:PP, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
      properties: { pixelSize:w }, fields:'pixelSize'
    }});
  });

  // Tab color
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:PP, tabColorStyle:{ rgbColor:hex(C.primary) } },
    fields:'tabColorStyle'
  }});

  await batchUpdate(id, reqs);
  await batchUpdate(id, dvReqs);
  await valuesBatchUpdate(id, data);
  console.log('Prep & Party Day complete');
})();
