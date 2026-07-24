'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const ML = sheetMap['Master Task Log'];
const S = "'Master Task Log'";

// 40 tasks: [taskName, taskType, project, dept, assignedTo, requestedBy, startDate, dueDate, completedDate, status, priority, effort, estHrs, actualHrs, recurring, recurrenceId, progress, blocker, notes]
const TASKS = [
  ['Redesign homepage hero section','Task','Website Redesign','Marketing','Taylor Brooks','Alex Morgan','2026-01-15','2026-02-28','2026-02-20','Completed','High','Large',12,11,'No','','100','','Approved by stakeholders'],
  ['Write Q1 blog content','Task','Social Media Campaign','Marketing','Taylor Brooks','Taylor Brooks','2026-01-20','2026-03-10','','In Progress','Medium','Medium',8,3,'No','','40','','3 of 5 posts done'],
  ['Conduct team onboarding sessions','Task','Training Program','Human Resources','Avery Walker','Alex Morgan','2026-03-01','2026-04-15','','Not Started','High','Large',16,0,'No','','0','Awaiting curriculum sign-off',''],
  ['Monthly financial close — March','Recurring Responsibility','Monthly Reporting','Finance','Morgan Chen','Morgan Chen','2026-03-01','2026-03-31','2026-03-31','Completed','High','Medium',6,6,'Yes','REC-0004','100','',''],
  ['Develop product feature spec','Milestone','Product Launch','Product','Jordan Patel','Alex Morgan','2026-02-01','2026-03-15','2026-03-14','Completed','Urgent','Large',20,18,'No','','100','','Ahead of schedule'],
  ['Client onboarding call — TechNova','Meeting Action','Client Onboarding','Customer Service','Sam Thompson','Sam Thompson','2026-03-10','2026-03-20','2026-03-20','Completed','Medium','Small',2,2,'No','','100','','Client happy with process'],
  ['Review Q1 marketing analytics','Review','Social Media Campaign','Marketing','Taylor Brooks','Alex Morgan','2026-04-01','2026-04-10','','In Review','Medium','Small',3,2,'No','','90','','Awaiting final review sign-off'],
  ['Update employee handbook','Administrative','Team Operations','Human Resources','Avery Walker','Avery Walker','2026-02-15','2026-04-01','','In Progress','Low','Medium',10,4,'No','','40','','Policy review pending'],
  ['Set up CRM pipeline stages','Task','Client Onboarding','Sales','Casey Rivera','Jordan Patel','2026-01-20','2026-02-15','2026-02-14','Completed','High','Small',4,4,'No','','100','',''],
  ['Blocked: API integration for launch','Task','Product Launch','Product','Jordan Patel','Jordan Patel','2026-03-01','2026-04-30','','Blocked','Urgent','Major',40,10,'No','','25','Waiting for third-party API credentials','Priority escalated'],
  ['Weekly status report distribution','Recurring Responsibility','Team Operations','Operations','Jamie Lee','Jamie Lee','2026-01-05','2026-07-26','','In Progress','Medium','Quick',1,0,'Yes','REC-0001','0','','Due this week'],
  ['Design new product logo concepts','Task','Product Launch','Creative','Riley Davis','Jordan Patel','2026-02-15','2026-03-31','2026-03-28','Completed','High','Medium',8,7,'No','','100','','Selected v3 design'],
  ['Q2 budget forecast','Task','Monthly Reporting','Finance','Morgan Chen','Alex Morgan','2026-03-15','2026-04-15','','In Progress','High','Medium',6,2,'No','','35','','Needs dept input by Apr 5'],
  ['Social media scheduling — April','Task','Social Media Campaign','Marketing','Riley Davis','Taylor Brooks','2026-03-25','2026-04-05','2026-04-03','Completed','Medium','Small',3,3,'No','','100','','Posted on schedule'],
  ['Process improvement workshop prep','Task','Process Improvement','Operations','Jamie Lee','Jamie Lee','2026-02-10','2026-03-01','2026-02-28','Completed','Medium','Medium',8,9,'No','','100','','Workshop held successfully'],
  ['Approval: new vendor contract','Approval','Team Operations','Finance','Morgan Chen','Alex Morgan','2026-04-01','2026-04-20','','In Progress','High','Quick',2,1,'No','','50','','Under legal review'],
  ['Hire customer support associate','Task','Team Operations','Human Resources','Avery Walker','Alex Morgan','2026-03-01','2026-05-01','','In Progress','High','Large',20,8,'No','','40','','3 candidates in final round'],
  ['Write product launch press release','Task','Product Launch','Marketing','Taylor Brooks','Jordan Patel','2026-04-01','2026-05-15','','Not Started','Medium','Medium',6,0,'No','','0','',''],
  ['Audit event vendor quotes','Task','Event Planning','Human Resources','Avery Walker','Avery Walker','2026-04-10','2026-05-01','','Not Started','Medium','Small',4,0,'No','','0','','3 venues to compare'],
  ['Quarterly process audit — Q1','Recurring Responsibility','Process Improvement','Operations','Jamie Lee','Alex Morgan','2026-01-01','2026-03-31','2026-04-02','Completed','High','Large',12,13,'Yes','REC-0005','100','',''],
  ['Respond to overdue client follow-ups','Follow-Up','Client Onboarding','Customer Service','Sam Thompson','Sam Thompson','2026-03-10','2026-03-25','','Blocked','High','Small',3,1,'No','','33','Waiting on client response for 2 accounts',''],
  ['Monthly payroll review — March','Recurring Responsibility','Monthly Reporting','Finance','Morgan Chen','Morgan Chen','2026-03-01','2026-03-28','2026-03-28','Completed','High','Small',2,2,'Yes','REC-0002','100','',''],
  ['Review website accessibility issues','Task','Website Redesign','Creative','Riley Davis','Taylor Brooks','2026-03-15','2026-04-30','','In Progress','Medium','Medium',6,1,'No','','15','','WCAG checklist in progress'],
  ['Sales pipeline report — Q1','Task','Monthly Reporting','Sales','Casey Rivera','Alex Morgan','2026-04-01','2026-04-15','','In Progress','Medium','Small',3,1,'No','','30','','CRM export ready'],
  ['Write training material drafts','Task','Training Program','Human Resources','Avery Walker','Avery Walker','2026-03-15','2026-05-01','','In Progress','Low','Large',14,4,'No','','30','','Modules 1-3 drafted'],
  ['Follow up with overdue invoice','Follow-Up','Monthly Reporting','Finance','Morgan Chen','Morgan Chen','2026-03-20','2026-04-01','','In Progress','High','Quick',1,0,'No','','50','','Invoice #2047 — 30 days overdue'],
  ['Redesign mobile checkout flow','Task','Website Redesign','Product','Jordan Patel','Taylor Brooks','2026-03-01','2026-05-31','','Not Started','High','Major',30,0,'No','','0','Waiting for wireframe approval',''],
  ['Weekly team meeting agenda','Recurring Responsibility','Team Operations','Operations','Jamie Lee','Alex Morgan','2026-01-06','2026-07-28','','Not Started','Medium','Quick',1,0,'Yes','REC-0001','0','',''],
  ['Onboard new sales team member','Task','Team Operations','Sales','Casey Rivera','Avery Walker','2026-04-15','2026-05-15','','Not Started','Medium','Medium',8,0,'No','','0','','Starts May 1'],
  ['Create event sponsorship deck','Task','Event Planning','Marketing','Riley Davis','Avery Walker','2026-04-20','2026-05-20','','Not Started','Low','Medium',6,0,'No','','0','',''],
  ['Finalize product pricing model','Task','Product Launch','Finance','Morgan Chen','Jordan Patel','2026-03-20','2026-04-30','','In Progress','Urgent','Medium',8,3,'No','','35','','Competitor pricing analysis complete'],
  ['Update CRM contact records','Administrative','Client Onboarding','Customer Service','Sam Thompson','Casey Rivera','2026-03-01','2026-04-10','','In Progress','Low','Small',4,2,'No','','50','','400 records to update'],
  ['Beta test user recruitment','Task','Product Launch','Marketing','Casey Rivera','Jordan Patel','2026-04-01','2026-05-10','','Not Started','High','Medium',10,0,'No','','0','','Target: 50 beta testers'],
  ['Q2 social media content calendar','Task','Social Media Campaign','Marketing','Taylor Brooks','Taylor Brooks','2026-04-05','2026-04-30','','In Progress','Medium','Medium',8,4,'No','','50','','Draft complete, needs approval'],
  ['Internal audit: expense categories','Task','Process Improvement','Finance','Morgan Chen','Jamie Lee','2026-03-10','2026-04-20','','In Progress','Medium','Medium',5,2,'No','','40','',''],
  ['Set up project tracking templates','Administrative','Team Operations','Operations','Jamie Lee','Alex Morgan','2026-02-01','2026-03-01','2026-02-28','Completed','Low','Small',3,3,'No','','100','','This tracker!'],
  ['Design event marketing assets','Task','Event Planning','Creative','Riley Davis','Taylor Brooks','2026-05-01','2026-06-01','','Not Started','Medium','Medium',8,0,'No','','0','',''],
  ['Conduct performance review cycle','Review','Team Operations','Human Resources','Avery Walker','Alex Morgan','2026-04-15','2026-05-31','','Not Started','High','Large',20,0,'No','','0','','H1 reviews'],
  ['Overdue: finalize Q1 social report','Task','Social Media Campaign','Marketing','Taylor Brooks','Alex Morgan','2026-03-01','2026-03-31','','Blocked','High','Small',4,1,'No','','25','Analytics tool access issue','Past due — tool access requested'],
  ['Platform upgrade planning','Task','Process Improvement','Product','Jordan Patel','Jamie Lee','2026-04-10','2026-06-30','','Not Started','Medium','Major',24,0,'No','','0','','Vendor evaluation in progress'],
];

(async () => {
  const data = [];

  // Title
  data.push({ range:`${S}!A1`, values:[['MASTER TEAM TASK LOG']] });
  data.push({ range:`${S}!A2`, values:[['Enter each task once. Weekly, monthly, project, and dashboard views update automatically from this log.']] });

  // Column headers row 5
  data.push({ range:`${S}!A5:V5`, values:[[
    'Task ID','Task Name','Task Type','Project','Department','Assigned To','Requested By',
    'Start Date','Due Date','Completed Date','Status','Priority','Effort',
    'Est. Hours','Actual Hours','Recurring?','Recurrence ID',
    'Days Remaining','Overdue?','Progress %','Blocker / Dependency','Notes'
  ]] });

  // Auto-formulas rows 6-505
  for (let r = 6; r <= 505; r++) {
    data.push({ range:`${S}!A${r}`, values:[[`=IF(B${r}="","","TSK-"&TEXT(ROW()-5,"0000"))`]] });
    data.push({ range:`${S}!R${r}`, values:[[`=IF(I${r}="","",IF(K${r}="Completed",0,I${r}-TODAY()))`]] });
    data.push({ range:`${S}!S${r}`, values:[[`=IF(B${r}="","",IF(AND(K${r}<>"Completed",K${r}<>"Cancelled",I${r}<TODAY()),"Yes","No"))`]] });
  }

  // Sample tasks
  TASKS.forEach(([name,type,proj,dept,assigned,reqby,start,due,compDate,status,pri,effort,est,actual,rec,recId,prog,blocker,notes], i) => {
    const r = 6 + i;
    data.push({ range:`${S}!B${r}:T${r}`, values:[[
      name,type,proj,dept,assigned,reqby,start,due,compDate,status,pri,effort,est||'',actual||'',rec,recId,null,null,prog/100
    ]] });
    if (notes || blocker) {
      data.push({ range:`${S}!U${r}:V${r}`, values:[[blocker,notes]] });
    }
  });

  await valuesBatchUpdate(id, data, 'mastertask-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(ML,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF    = { type:'DATE', pattern:'MMM D, YYYY' };
  const pctF     = { type:'PERCENT', pattern:'0%' };
  const numF     = { type:'NUMBER', pattern:'0.0' };

  // Title rows
  reqs.push({ mergeCells:{ range:gridRange(ML,0,2,0,22), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,22,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:18, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  reqs.push({ mergeCells:{ range:gridRange(ML,2,4,0,22), mergeType:'MERGE_ALL' } });
  fmt(2,4,0,22,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Column header row 5
  fmt(4,5,0,22,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Data rows 6-505
  // Formula cols: A(0), R(17), S(18)
  fmt(5,505,0,1,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  fmt(5,505,17,19,{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });

  // Input cols
  fmt(5,505,1,17,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  fmt(5,505,19,22,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });

  // Date columns H,I,J (7,8,9)
  for (const ci of [7,8,9]) {
    reqs.push({ repeatCell:{ range:gridRange(ML,5,505,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER', backgroundColor:hex(C.input) } }, fields:'userEnteredFormat' } });
  }
  // Input yellow for text input cols
  for (const ci of [1,2,3,4,5,6,10,11,12,15,16,19,20,21]) {
    reqs.push({ repeatCell:{ range:gridRange(ML,5,505,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  }
  // Hours cols N,O (13,14)
  for (const ci of [13,14]) {
    reqs.push({ repeatCell:{ range:gridRange(ML,5,505,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:numF, horizontalAlignment:'RIGHT', backgroundColor:hex(C.input) } }, fields:'userEnteredFormat' } });
  }
  // Progress col T (19)
  reqs.push({ repeatCell:{ range:gridRange(ML,5,505,19,20), cell:{ userEnteredFormat:{ numberFormat:pctF, horizontalAlignment:'CENTER', backgroundColor:hex(C.input) } }, fields:'userEnteredFormat' } });
  // Days remaining: number with sign
  reqs.push({ repeatCell:{ range:gridRange(ML,5,505,17,18), cell:{ userEnteredFormat:{ numberFormat:{ type:'NUMBER', pattern:'0;[Red]-0' }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });

  // Alternating rows
  for (let r = 5; r < 505; r+=2) {
    fmt(r,r+1,0,22,{ userEnteredFormat:{ backgroundColor:hex(C.altRow) } });
  }

  // Borders
  reqs.push({ updateBorders:{ range:gridRange(ML,4,505,0,22), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Col widths
  [[0,85],[1,200],[2,120],[3,140],[4,120],[5,120],[6,120],[7,100],[8,100],[9,100],
   [10,95],[11,80],[12,80],[13,80],[14,80],[15,80],[16,100],[17,90],[18,75],[19,80],[20,160],[21,150]]
    .forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:ML, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:ML, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:36 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:ML, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:ML, dimension:'ROWS', startIndex:5, endIndex:505 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Unmerge title rows before applying column+row freeze to avoid conflict
  // (title merges go across cols A-V so col freeze A:B would fail without unmerge)
  reqs.push({ unmergeCells:{ range:gridRange(ML,0,2,0,22) } });
  reqs.push({ unmergeCells:{ range:gridRange(ML,2,4,0,22) } });
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:ML, gridProperties:{ frozenRowCount:5, frozenColumnCount:2 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });
  // Re-apply title merges after freeze (outside A:B)
  reqs.push({ mergeCells:{ range:gridRange(ML,0,2,2,22), mergeType:'MERGE_ALL' } });
  reqs.push({ mergeCells:{ range:gridRange(ML,2,4,2,22), mergeType:'MERGE_ALL' } });

  await batchUpdate(id, reqs, 'mastertask-format');
  console.log('Master Task Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
