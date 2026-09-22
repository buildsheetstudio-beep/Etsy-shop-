'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const PP = sheetMap['Prep, Setup & Cleanup'];
const S = "'Prep, Setup & Cleanup'";

// 38 tasks: [Phase, Task, Category, AssignedTo, StartDate, DueDate, Priority, Status, Completed, CompletedDate, PartyDayTime, Location, Notes]
const TASKS = [
  ['8+ Weeks Before','Book venue','Venue','Sarah Hendricks','Jun 1, 2026','Jun 15, 2026','Urgent','Complete',true,'Jun 10, 2026','','','Confirmed — Riverside Community Room'],
  ['8+ Weeks Before','Set total budget','Planning','Sarah Hendricks','Jun 1, 2026','Jun 15, 2026','High','Complete',true,'Jun 5, 2026','','','$1,200 set'],
  ['8+ Weeks Before','Draft guest list','Guest Management','Sarah Hendricks','Jun 1, 2026','Jun 20, 2026','High','Complete',true,'Jun 18, 2026','','','27 guests finalised'],
  ['6 Weeks Before','Finalise guest list','Guest Management','Sarah Hendricks','Jun 20, 2026','Jul 1, 2026','High','Complete',true,'Jun 28, 2026','','',''],
  ['6 Weeks Before','Order/print invitations','Invitations','Sarah Hendricks','Jun 20, 2026','Jul 3, 2026','High','Complete',true,'Jul 3, 2026','','','Printed x30'],
  ['6 Weeks Before','Send invitations','Invitations','Sarah Hendricks','Jul 3, 2026','Jul 8, 2026','High','Complete',true,'Jul 7, 2026','','','All delivered by hand and email'],
  ['6 Weeks Before','Book cake','Cake','Sarah Hendricks','Jun 20, 2026','Jul 5, 2026','Urgent','Complete',true,'Jun 25, 2026','','','Sweet Bloom Bakery — deposit paid'],
  ['6 Weeks Before','Book entertainment','Entertainment','Sarah Hendricks','Jun 20, 2026','Jul 5, 2026','High','Complete',true,'Jul 1, 2026','','','Face painter booked'],
  ['4 Weeks Before','Plan menu','Food','Sarah Hendricks','Jul 1, 2026','Jul 15, 2026','High','Complete',true,'Jul 12, 2026','','','Full menu on Food, Drinks & Cake tab'],
  ['4 Weeks Before','Purchase decorations (online)','Decorations','Sarah Hendricks','Jul 10, 2026','Jul 20, 2026','Medium','Complete',true,'Jul 15, 2026','','','Balloon arch kit received'],
  ['4 Weeks Before','Purchase craft activity kits','Activities','Sarah Hendricks','Jul 20, 2026','Jul 28, 2026','Medium','Complete',true,'Jul 28, 2026','','','Fairy garden kits x20 — packed'],
  ['3 Weeks Before','Chase RSVPs','Guest Management','Sarah Hendricks','Jul 20, 2026','Jul 25, 2026','High','Complete',true,'Jul 22, 2026','','','3 still no response — following up'],
  ['3 Weeks Before','Confirm headcount with venue','Venue','Sarah Hendricks','Jul 20, 2026','Jul 28, 2026','High','In Progress',false,'','','','Final headcount needed by Aug 5'],
  ['3 Weeks Before','Confirm dietary needs','Food','Sarah Hendricks','Jul 20, 2026','Jul 28, 2026','High','In Progress',false,'','','','Follow up with Patel, Robinson families'],
  ['2 Weeks Before','Order food items (non-perishable)','Food','Sarah Hendricks','Aug 1, 2026','Aug 5, 2026','High','Not Started',false,'','','','Place Costco order'],
  ['2 Weeks Before','Set up party favor bags','Party Favors','Sarah Hendricks','Aug 1, 2026','Aug 8, 2026','Medium','Not Started',false,'','','','Fill bags with stickers, figurines'],
  ['2 Weeks Before','Confirm face painter','Entertainment','Sarah Hendricks','Aug 1, 2026','Aug 5, 2026','High','Not Started',false,'','','','Call to confirm arrival time 2PM'],
  ['2 Weeks Before','Purchase tableware','Supplies','Sarah Hendricks','Aug 1, 2026','Aug 8, 2026','Medium','Not Started',false,'','','','Already done in July — check stock'],
  ['1 Week Before','Confirm final guest numbers','Guest Management','Sarah Hendricks','Aug 8, 2026','Aug 10, 2026','High','Not Started',false,'','','','Final RSVP cutoff Aug 1'],
  ['1 Week Before','Confirm food orders','Food','Sarah Hendricks','Aug 8, 2026','Aug 10, 2026','Urgent','Not Started',false,'','','','Call Local Deli and Costco'],
  ['1 Week Before','Pick up non-perishable groceries','Shopping','Sarah Hendricks','Aug 8, 2026','Aug 12, 2026','High','Not Started',false,'','','',''],
  ['1 Week Before','Prepare pass-the-parcel','Activities','Sarah Hendricks','Aug 8, 2026','Aug 12, 2026','Medium','Not Started',false,'','','','Wrap 7 layers with prizes'],
  ['3 Days Before','Confirm with bakery','Cake','Sarah Hendricks','Aug 12, 2026','Aug 12, 2026','Urgent','Not Started',false,'','','','Confirm pickup for Aug 14'],
  ['3 Days Before','Prepare activity tables setup plan','Setup','Sarah Hendricks','Aug 12, 2026','Aug 12, 2026','Medium','Not Started',false,'','','','Sketch layout for venue'],
  ['1 Day Before','Pick up birthday cake','Cake','Sarah Hendricks','Aug 14, 2026','Aug 14, 2026','Urgent','Not Started',false,'','','','Sweet Bloom Bakery 10AM'],
  ['1 Day Before','Prepare homemade food items','Food','Sarah Hendricks','Aug 14, 2026','Aug 14, 2026','High','Not Started',false,'','','','Mini quiche, lemonade, fruit skewers'],
  ['1 Day Before','Set up party favor bags (final)','Party Favors','Sarah Hendricks','Aug 14, 2026','Aug 14, 2026','Medium','Not Started',false,'','','','Label each bag with child name'],
  ['1 Day Before','Charge camera / confirm photos plan','Photography','Sarah Hendricks','Aug 14, 2026','Aug 14, 2026','Low','Not Started',false,'','','',''],
  ['1 Day Before','Pack supply bag','Supplies','Sarah Hendricks','Aug 14, 2026','Aug 14, 2026','High','Not Started',false,'','','','Candles, lighter, knife, serving gear'],
  ['Party Day','Arrive at venue early','Setup','Sarah Hendricks','','Aug 15, 2026','Urgent','Not Started',false,'','12:30 PM','Riverside Room','Venue open from 12:30 PM'],
  ['Party Day','Set up decorations','Decorations','Sarah Hendricks','','Aug 15, 2026','High','Not Started',false,'','12:30 PM','Main Hall','Balloon arch, centrepieces, bunting'],
  ['Party Day','Set up food & drinks table','Food','Sarah Hendricks','','Aug 15, 2026','High','Not Started',false,'','1:00 PM','Dining Area',''],
  ['Party Day','Set up activity and gift table','Activities','Sarah Hendricks','','Aug 15, 2026','Medium','Not Started',false,'','1:15 PM','Activity Corner',''],
  ['Party Day','Welcome guests and manage arrivals','Guest Management','Sarah Hendricks','','Aug 15, 2026','High','Not Started',false,'','2:00 PM','Entrance',''],
  ['Party Day','Serve food during party','Food','Sarah Hendricks','','Aug 15, 2026','High','Not Started',false,'','2:30 PM','Dining Area',''],
  ['Party Day','Run activities and games','Activities','Sarah Hendricks','','Aug 15, 2026','High','Not Started',false,'','3:00 PM','Activity Corner',''],
  ['Party Day','Sing Happy Birthday and cut cake','Cake','Sarah Hendricks','','Aug 15, 2026','Urgent','Not Started',false,'','4:00 PM','Main Hall',''],
  ['After Party','Pack and return favor bags','Party Favors','Sarah Hendricks','Aug 15, 2026','Aug 15, 2026','Medium','Not Started',false,'','5:00 PM','','Give to each family on departure'],
  ['After Party','Clean venue','Cleanup','Sarah Hendricks','Aug 15, 2026','Aug 15, 2026','High','Not Started',false,'','5:00 PM','All Areas','Leave no later than 6PM'],
  ['After Party','Send thank-you messages','Guest Management','Sarah Hendricks','Aug 16, 2026','Aug 22, 2026','Medium','Not Started',false,'','','','Handwritten where possible'],
];

(async () => {
  const data = [];

  data.push({ range:`${S}!A1`, values:[['PARTY PREP, SETUP & CLEANUP']] });
  data.push({ range:`${S}!A3`, values:[['Plan and track every preparation task from booking to cleanup. Use phases and due dates to stay on schedule.']] });

  // Summary cards rows 5-6
  data.push({ range:`${S}!A5`, values:[['Days Until Party']] });
  data.push({ range:`${S}!B5`, values:[[`=IFERROR(MAX(0,DATEVALUE("August 15, 2026")-TODAY()),"—")`]] });
  data.push({ range:`${S}!C5`, values:[['Total Tasks']] });
  data.push({ range:`${S}!D5`, values:[[`=IFERROR(COUNTA($C$8:$C$207),0)`]] });
  data.push({ range:`${S}!E5`, values:[['Completed']] });
  data.push({ range:`${S}!F5`, values:[[`=IFERROR(COUNTIF($J$8:$J$207,TRUE),0)`]] });
  data.push({ range:`${S}!G5`, values:[['Remaining']] });
  data.push({ range:`${S}!H5`, values:[[`=IFERROR(D5-F5,0)`]] });
  data.push({ range:`${S}!A6`, values:[['Overdue Tasks']] });
  data.push({ range:`${S}!B6`, values:[[`=IFERROR(COUNTIF($M$8:$M$207,"Yes"),0)`]] });
  data.push({ range:`${S}!C6`, values:[['Completion %']] });
  data.push({ range:`${S}!D6`, values:[[`=IFERROR(IF(D5=0,0,F5/D5),0)`]] });
  data.push({ range:`${S}!E6`, values:[['In Progress']] });
  data.push({ range:`${S}!F6`, values:[[`=IFERROR(COUNTIF($I$8:$I$207,"In Progress"),0)`]] });
  data.push({ range:`${S}!G6`, values:[['Urgent Tasks Left']] });
  data.push({ range:`${S}!H6`, values:[[`=IFERROR(COUNTIFS($H$8:$H$207,"Urgent",$J$8:$J$207,FALSE),0)`]] });

  // Table section header row 7 (data row 8 in Sheets, index 7)
  data.push({ range:`${S}!A8`, values:[['PREPARATION TASK LIST']] });
  data.push({ range:`${S}!A9:P9`, values:[[
    'Task ID','Phase','Task','Category','Assigned To','Start Date','Due Date',
    'Priority','Status','Completed?','Completed Date','Days Remaining',
    'Overdue?','Party-Day Time','Location / Area','Notes'
  ]] });

  // Task ID, Days Remaining, Overdue formulas rows 10-209
  for (let r = 10; r <= 209; r++) {
    data.push({ range:`${S}!A${r}`, values:[[`=IF(C${r}="","","TASK-"&TEXT(ROW()-9,"000"))`]] });
    data.push({ range:`${S}!L${r}`, values:[[`=IF(G${r}="","",IF(J${r}=TRUE,0,G${r}-TODAY()))`]] });
    data.push({ range:`${S}!M${r}`, values:[[`=IF(C${r}="","",IF(AND(J${r}<>TRUE,G${r}<>"",(G${r}<TODAY())),"Yes","No"))`]] });
  }

  // Sample tasks
  TASKS.forEach(([phase,task,cat,assigned,startDate,dueDate,priority,status,completed,completedDate,partyTime,location,notes], i) => {
    const r = 10 + i;
    data.push({ range:`${S}!B${r}`, values:[[phase]] });
    data.push({ range:`${S}!C${r}`, values:[[task]] });
    data.push({ range:`${S}!D${r}`, values:[[cat]] });
    data.push({ range:`${S}!E${r}`, values:[[assigned]] });
    if (startDate) data.push({ range:`${S}!F${r}`, values:[[startDate]] });
    if (dueDate) data.push({ range:`${S}!G${r}`, values:[[dueDate]] });
    data.push({ range:`${S}!H${r}`, values:[[priority]] });
    data.push({ range:`${S}!I${r}`, values:[[status]] });
    data.push({ range:`${S}!J${r}`, values:[[completed]] });
    if (completedDate) data.push({ range:`${S}!K${r}`, values:[[completedDate]] });
    if (partyTime) data.push({ range:`${S}!N${r}`, values:[[partyTime]] });
    if (location) data.push({ range:`${S}!O${r}`, values:[[location]] });
    if (notes) data.push({ range:`${S}!P${r}`, values:[[notes]] });
  });

  await valuesBatchUpdate(id, data, 'prep-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(PP,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };
  const pctF   = { type:'PERCENT', pattern:'0%' };

  // Title rows 0-1
  reqs.push({ mergeCells:{ range:gridRange(PP,0,2,0,16), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Guidance rows 2-3
  reqs.push({ mergeCells:{ range:gridRange(PP,2,4,0,16), mergeType:'MERGE_ALL' } });
  fmt(2,4,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Cards rows 4-5
  for (let i = 0; i < 4; i++) {
    const c = i * 2;
    reqs.push({ mergeCells:{ range:gridRange(PP,4,5,c,c+1), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(PP,4,5,c+1,c+2), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(PP,5,6,c,c+1), mergeType:'MERGE_ALL' } });
    reqs.push({ mergeCells:{ range:gridRange(PP,5,6,c+1,c+2), mergeType:'MERGE_ALL' } });
    fmt(4,5,c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.lavender), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(4,5,c+1,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:14, fontFamily:'Arial', foregroundColor:hex(C.primary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(5,6,c,c+1,{ userEnteredFormat:{ backgroundColor:hex(C.mutedBlue), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    fmt(5,6,c+1,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:14, fontFamily:'Arial', foregroundColor:hex(C.secondary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  }
  // Pct format for D6 (index row 5, col 3)
  reqs.push({ repeatCell:{ range:gridRange(PP,5,6,3,4), cell:{ userEnteredFormat:{ numberFormat:pctF } }, fields:'userEnteredFormat.numberFormat' } });
  // Overdue tasks card B6 (index 5, col 1) — red if >0
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(PP,5,6,1,2)], booleanRule:{ condition:{ type:'NUMBER_GREATER', values:[{ userEnteredValue:'0' }] }, format:{ textFormat:{ foregroundColor:hex(C.attention), bold:true } } } }, index:0 } });
  reqs.push({ updateBorders:{ range:gridRange(PP,4,6,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Section header row 7 (index 7), table header row 8 (index 8)
  reqs.push({ mergeCells:{ range:gridRange(PP,7,8,0,16), mergeType:'MERGE_ALL' } });
  fmt(7,8,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(8,9,0,16,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });

  // Data rows 9-208 (indices)
  for (let r = 9; r < 209; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,16,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Formula cols A(0),L(11),M(12)
  for (const ci of [0,11,12]) {
    reqs.push({ repeatCell:{ range:gridRange(PP,9,209,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Input cols B-K(1-10),N-P(13-15)
  for (const ci of [1,2,3,4,5,6,7,8,9,10,13,14,15]) {
    reqs.push({ repeatCell:{ range:gridRange(PP,9,209,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  }
  // Date cols F(5),G(6),K(10)
  for (const ci of [5,6,10]) {
    reqs.push({ repeatCell:{ range:gridRange(PP,9,209,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Checkbox J(9) Completed
  for (let r = 9; r < 209; r++) {
    reqs.push({ setDataValidation:{ range:gridRange(PP,r,r+1,9,10), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
  }
  // Dropdowns: Phase B(1), Priority H(7), Status I(8)
  reqs.push({ setDataValidation:{ range:gridRange(PP,9,209,1,2), rule:{ condition:{ type:'ONE_OF_LIST', values:['8+ Weeks Before','6 Weeks Before','4 Weeks Before','3 Weeks Before','2 Weeks Before','1 Week Before','3 Days Before','1 Day Before','Party Day','After Party'].map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true } } });
  reqs.push({ setDataValidation:{ range:gridRange(PP,9,209,7,8), rule:{ condition:{ type:'ONE_OF_LIST', values:['Low','Medium','High','Urgent'].map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true } } });
  reqs.push({ setDataValidation:{ range:gridRange(PP,9,209,8,9), rule:{ condition:{ type:'ONE_OF_LIST', values:['Not Started','In Progress','Waiting','Complete','Cancelled'].map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true } } });

  // Status CF col I(8)
  const statusCF = [
    { val:'Complete',    bg:C.success,   fg:C.white },
    { val:'In Progress', bg:C.mutedBlue, fg:C.mainText },
    { val:'Waiting',     bg:C.warning,   fg:C.mainText },
    { val:'Not Started', bg:C.altRow,    fg:C.secText },
    { val:'Cancelled',   bg:C.border,    fg:C.secText },
  ];
  statusCF.forEach(({ val, bg, fg }) => {
    reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(PP,9,209,8,9)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:val }] }, format:{ backgroundColor:hex(bg), textFormat:{ foregroundColor:hex(fg), bold:true } } } }, index:0 } });
  });

  // Priority CF col H(7)
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(PP,9,209,7,8)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'Urgent' }] }, format:{ backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true } } } }, index:0 } });
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(PP,9,209,7,8)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'High' }] }, format:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true } } } }, index:0 } });

  // Full-row CF: Complete → sage tint
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(PP,9,209,0,16)], booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=$J10=TRUE` }] }, format:{ backgroundColor:{ red:0.568, green:0.706, blue:0.604, alpha:0.2 } } } }, index:0 } });
  // Overdue → muted rust tint
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(PP,9,209,0,16)], booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=AND($M10="Yes",$J10=FALSE)` }] }, format:{ backgroundColor:{ red:0.761, green:0.467, blue:0.447, alpha:0.15 } } } }, index:0 } });

  reqs.push({ updateBorders:{ range:gridRange(PP,8,209,0,16), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  [[0,80],[1,110],[2,200],[3,100],[4,120],[5,100],[6,100],[7,75],[8,100],[9,80],[10,100],[11,90],[12,75],[13,90],[14,120],[15,170]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:PP, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PP, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PP, dimension:'ROWS', startIndex:2, endIndex:9 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PP, dimension:'ROWS', startIndex:4, endIndex:6 }, properties:{ pixelSize:42 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:PP, dimension:'ROWS', startIndex:9, endIndex:209 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  reqs.push({ updateSheetProperties:{ properties:{ sheetId:PP, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'prep-format');
  console.log('Prep, Setup & Cleanup complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
