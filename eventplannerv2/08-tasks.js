'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 6;

// 19 columns A:S
const HEADERS = [
  'Event ID','Phase','Task Name','Owner','Due Date','Status','Priority',
  'Days Until Due','Depends On','Vendor/Partner','Completed Date',
  'Est. Hours','Actual Hours','Budget Impact?','Cost Estimate',
  'Notes','Urgency Flag','Overdue?','Vendor Task?',
];

const COL_WIDTHS = [80,130,220,130,100,110,90,80,130,140,100,80,80,90,90,200,100,80,80];

const PHASES = ['Pre-Planning','Venue & Vendors','Marketing & Comms','Logistics & Ops',
  'Registration','Day-Of Setup','Day-Of Execution','Post-Event','Finance','Admin & Legal'];

// 100 tasks
const TASKS = [
  // Pre-Planning (EVT-001)
  ['EVT-001','Pre-Planning','Define event goals and KPIs','Diana Reyes','01/15/2026','Completed','High','','','','01/15/2026',4,4,'No',0,'Approved by board','','','No'],
  ['EVT-001','Pre-Planning','Set event dates and venue options','Diana Reyes','01/20/2026','Completed','High','','','','01/20/2026',3,3,'No',0,'','','','No'],
  ['EVT-001','Pre-Planning','Create event budget draft','Elena Park','01/25/2026','Completed','Critical','','','','01/25/2026',5,6,'Yes',75000,'Board-approved budget','','','No'],
  ['EVT-001','Pre-Planning','Form event planning committee','Diana Reyes','01/28/2026','Completed','High','','','','01/28/2026',2,2,'No',0,'8 members','','','No'],
  ['EVT-001','Pre-Planning','Submit permits and city filings','Marcus Cole','02/05/2026','Completed','Critical','','','','02/05/2026',4,4,'Yes',500,'Permits approved','','','No'],
  // Venue & Vendors (EVT-001)
  ['EVT-001','Venue & Vendors','Sign venue contract — ACC','Marcus Cole','02/01/2026','Completed','Critical','','','VEN-001','02/01/2026',2,2,'Yes',30000,'Deposit $5,000 paid','','','Yes'],
  ['EVT-001','Venue & Vendors','Source and contract catering','Elena Park','02/10/2026','Completed','Critical','','','VDR-001','02/10/2026',5,6,'Yes',18000,'Summit Catering signed','','','Yes'],
  ['EVT-001','Venue & Vendors','Contract AV provider','Marcus Cole','02/15/2026','Completed','Critical','','','VDR-003','02/15/2026',3,3,'Yes',35000,'ProAV signed','','','Yes'],
  ['EVT-001','Venue & Vendors','Book photography & video','James Liu','02/20/2026','Completed','High','','','VDR-007','02/20/2026',2,2,'Yes',14000,'','','','Yes'],
  ['EVT-001','Venue & Vendors','Arrange shuttle transportation','Marcus Cole','03/01/2026','Completed','Medium','','','VDR-012','03/01/2026',2,2,'Yes',4500,'City Shuttle signed','','','Yes'],
  // Marketing & Comms (EVT-001)
  ['EVT-001','Marketing & Comms','Design event branding and logo','James Liu','02/15/2026','Completed','High','','','','02/15/2026',8,10,'Yes',2500,'Logo approved','','','No'],
  ['EVT-001','Marketing & Comms','Launch event website','James Liu','02/28/2026','Completed','High','','','','02/28/2026',12,12,'Yes',1500,'Online registration live','','','No'],
  ['EVT-001','Marketing & Comms','Send save-the-dates','Elena Park','03/01/2026','Completed','High','','','','03/01/2026',2,2,'No',800,'Email + print','','','No'],
  ['EVT-001','Marketing & Comms','Launch social media campaign','James Liu','03/10/2026','Completed','Medium','','','','03/10/2026',6,7,'Yes',3000,'4 platforms active','','','No'],
  ['EVT-001','Marketing & Comms','Press release distribution','Elena Park','03/20/2026','Completed','Medium','','','','03/20/2026',4,4,'Yes',500,'10 media outlets','','','No'],
  // Logistics & Ops (EVT-001)
  ['EVT-001','Logistics & Ops','Create run-of-show document','Marcus Cole','04/01/2026','Completed','Critical','','','','04/01/2026',8,9,'No',0,'Distributed to all leads','','','No'],
  ['EVT-001','Logistics & Ops','Order signage and printed materials','James Liu','03/25/2026','Completed','High','','','VDR-009','03/25/2026',3,3,'Yes',3200,'Final files submitted','','','Yes'],
  ['EVT-001','Logistics & Ops','Confirm room layouts and diagrams','Marcus Cole','04/05/2026','Completed','High','','','','04/05/2026',3,4,'No',0,'Floor plans approved','','','No'],
  ['EVT-001','Logistics & Ops','Finalize catering menu and headcount','Elena Park','04/08/2026','Completed','Critical','','','VDR-001','04/08/2026',2,2,'No',0,'Final count submitted','','','Yes'],
  ['EVT-001','Logistics & Ops','Staff briefing and walkthroughs','Diana Reyes','04/13/2026','Completed','High','','','VDR-014','04/13/2026',4,4,'No',0,'All 20 staff confirmed','','','Yes'],
  // Registration (EVT-001)
  ['EVT-001','Registration','Set up online registration system','James Liu','02/28/2026','Completed','Critical','','','','02/28/2026',8,9,'Yes',500,'480 registered','','','No'],
  ['EVT-001','Registration','Process VIP invitations','Diana Reyes','03/05/2026','Completed','High','','','','03/05/2026',4,4,'No',0,'40 VIPs confirmed','','','No'],
  ['EVT-001','Registration','Design and print name badges','James Liu','04/10/2026','Completed','High','','','VDR-009','04/10/2026',3,3,'Yes',800,'480 badges printed','','','Yes'],
  ['EVT-001','Registration','Set up check-in station','Marcus Cole','04/14/2026','Completed','High','','','','04/14/2026',2,3,'No',0,'QR scanner ready','','','No'],
  // Day-Of Setup (EVT-001)
  ['EVT-001','Day-Of Setup','Venue walkthrough with ACC staff','Diana Reyes','04/14/2026','Completed','Critical','','','VEN-001','04/14/2026',2,2,'No',0,'Room configurations confirmed','','','Yes'],
  ['EVT-001','Day-Of Setup','AV setup and testing','Marcus Cole','04/15/2026','Completed','Critical','','','VDR-003','04/15/2026',4,5,'No',0,'All mics/screens tested','','','Yes'],
  ['EVT-001','Day-Of Setup','Registration table setup','Elena Park','04/15/2026','Completed','High','','','','04/15/2026',2,2,'No',0,'','','','No'],
  ['EVT-001','Day-Of Setup','Catering setup review','Marcus Cole','04/15/2026','Completed','High','','','VDR-001','04/15/2026',1,1,'No',0,'','','','Yes'],
  // Post-Event (EVT-001)
  ['EVT-001','Post-Event','Send attendee thank-you emails','Elena Park','04/20/2026','In Progress','High','','','','',3,0,'No',0,'Draft in progress','','','No'],
  ['EVT-001','Post-Event','Collect and analyze feedback surveys','James Liu','04/22/2026','Not Started','High','','','','',4,0,'No',0,'Survey link sent','','','No'],
  // EVT-002 tasks
  ['EVT-002','Pre-Planning','Finalize gala theme and color scheme','Diana Reyes','02/01/2026','Completed','High','','','','02/01/2026',3,3,'No',0,'Gold & blush theme','','','No'],
  ['EVT-002','Venue & Vendors','Sign Driskill Hotel contract','Marcus Cole','02/10/2026','Completed','Critical','','','VEN-002','02/10/2026',2,2,'Yes',16000,'Deposit paid','','','Yes'],
  ['EVT-002','Venue & Vendors','Contract catering — Gala Fine Dining','Elena Park','02/15/2026','Completed','Critical','','','VDR-002','02/15/2026',3,3,'Yes',22000,'Menu finalized','','','Yes'],
  ['EVT-002','Venue & Vendors','Book entertainment — The Jazzmakers','Marcus Cole','03/01/2026','Completed','High','','','VDR-010','03/01/2026',1,1,'Yes',3500,'Deposit paid','','','Yes'],
  ['EVT-002','Marketing & Comms','Design and mail gala invitations','James Liu','03/15/2026','Completed','High','','','','03/15/2026',5,6,'Yes',2000,'280 mailed','','','No'],
  ['EVT-002','Marketing & Comms','Solicit auction items','Elena Park','03/20/2026','In Progress','Critical','','','','',6,3,'No',0,'40 items so far — need 60','','','No'],
  ['EVT-002','Logistics & Ops','Create auction bid sheets','James Liu','04/20/2026','Not Started','High','','','','',4,0,'Yes',500,'','','','No'],
  ['EVT-002','Logistics & Ops','Coordinate valet logistics','Marcus Cole','04/25/2026','Not Started','Medium','','','VDR-015','',2,0,'No',0,'','','','Yes'],
  ['EVT-002','Day-Of Setup','Décor and floral setup','Diana Reyes','05/02/2026','Not Started','High','','','VDR-006','',3,0,'Yes',8500,'Setup starts at noon','','','Yes'],
  ['EVT-002','Finance','Close auction and process donations','Elena Park','05/04/2026','Not Started','Critical','','','','',4,0,'Yes',0,'Target: $80,000','','','No'],
  // EVT-003 tasks
  ['EVT-003','Pre-Planning','Secure park permit from City','Sam Torres','03/01/2026','Completed','Critical','','','','03/01/2026',3,3,'Yes',250,'Permit confirmed','','','No'],
  ['EVT-003','Venue & Vendors','Contract PA system rental','Sam Torres','04/01/2026','Completed','High','','','VDR-004','04/01/2026',2,2,'Yes',4500,'Park Sounds booked','','','Yes'],
  ['EVT-003','Venue & Vendors','Recruit health vendor partners','Priya Nair','04/15/2026','In Progress','Critical','','','VDR-016','',5,2,'No',0,'8 of 10 confirmed','','','No'],
  ['EVT-003','Marketing & Comms','Flyers and social media outreach','Sam Torres','04/20/2026','In Progress','High','','','','',6,3,'Yes',800,'','','','No'],
  ['EVT-003','Logistics & Ops','Map venue layout and vendor spots','Priya Nair','05/01/2026','Not Started','High','','','','',4,0,'No',0,'','','','No'],
  ['EVT-003','Day-Of Setup','Morning setup — tents and tables','Sam Torres','05/15/2026','Not Started','Critical','','','VDR-017','',3,0,'Yes',6000,'Delivery at 6 AM','','','Yes'],
  // EVT-006 tasks
  ['EVT-006','Pre-Planning','Register with city sports league','Sam Torres','03/15/2026','Completed','High','','','','03/15/2026',2,2,'Yes',100,'','','','No'],
  ['EVT-006','Venue & Vendors','Reserve park fields','Sam Torres','03/20/2026','Completed','Critical','','','VEN-007','03/20/2026',1,1,'Yes',600,'4 fields confirmed','','','Yes'],
  ['EVT-006','Registration','Open online team registration','Priya Nair','04/01/2026','Completed','High','','','','04/01/2026',4,4,'No',0,'16 teams registered','','','No'],
  ['EVT-006','Logistics & Ops','Recruit referees','Sam Torres','04/15/2026','In Progress','Critical','','','','',4,2,'Yes',600,'Need 8 referees','','','No'],
  ['EVT-006','Logistics & Ops','Order trophies and medals','Priya Nair','05/01/2026','In Progress','High','','','VDR-018','',2,1,'Yes',2200,'Files submitted','','','Yes'],
  ['EVT-006','Marketing & Comms','Announce tournament brackets','Sam Torres','05/20/2026','Not Started','Medium','','','','',3,0,'No',0,'','','','No'],
  ['EVT-006','Day-Of Setup','Field markings and equipment setup','Sam Torres','06/05/2026','Not Started','Critical','','','','',3,0,'No',0,'Setup 6 AM','','','No'],
  // EVT-007 tasks
  ['EVT-007','Venue & Vendors','Set up virtual event platform','James Liu','05/01/2026','Completed','Critical','','','VDR-019','05/01/2026',4,4,'Yes',1500,'Zoom Events licensed','','','Yes'],
  ['EVT-007','Marketing & Comms','Email campaign to job seekers','James Liu','05/10/2026','Completed','High','','','','05/10/2026',6,7,'Yes',500,'20,000 emails sent','','','No'],
  ['EVT-007','Marketing & Comms','Employer showcase invitations','Elena Park','05/15/2026','Completed','High','','','','05/15/2026',3,3,'No',0,'22 employers confirmed','','','No'],
  ['EVT-007','Registration','Open attendee registration portal','James Liu','05/01/2026','Completed','Critical','','','','05/01/2026',4,4,'No',0,'380 registered','','','No'],
  ['EVT-007','Logistics & Ops','Prep virtual breakout room schedule','Elena Park','06/01/2026','In Progress','High','','','','',3,2,'No',0,'','','','No'],
  ['EVT-007','Day-Of Execution','Moderate sessions and manage chat','James Liu','06/10/2026','Not Started','Critical','','','','',5,0,'No',0,'','','','No'],
  // EVT-013 tasks
  ['EVT-013','Pre-Planning','Confirm hybrid streaming infrastructure','Marcus Cole','04/01/2026','In Progress','Critical','','','VDR-003','',6,3,'Yes',35000,'AV contract pending','','','Yes'],
  ['EVT-013','Venue & Vendors','Finalize Palmer Events Center contract','Marcus Cole','04/15/2026','In Progress','Critical','','','VEN-003','',2,1,'Yes',24000,'Awaiting counter-sig','','','Yes'],
  ['EVT-013','Marketing & Comms','Launch conference website and reg','James Liu','04/20/2026','Not Started','Critical','','','','',12,0,'Yes',4000,'','','','No'],
  ['EVT-013','Marketing & Comms','Confirm keynote speaker lineup','Diana Reyes','04/30/2026','Not Started','Critical','','','','',6,0,'No',0,'3 keynotes pending','','','No'],
  ['EVT-013','Registration','Build multi-track registration system','James Liu','05/15/2026','Not Started','Critical','','','','',20,0,'Yes',2000,'','','','No'],
  ['EVT-013','Venue & Vendors','Contract exhibition hall vendors','Marcus Cole','05/01/2026','Not Started','High','','','','',8,0,'Yes',15000,'','','','No'],
  ['EVT-013','Logistics & Ops','Develop conference app','James Liu','05/15/2026','Not Started','High','','','','',40,0,'Yes',8000,'','','','No'],
  ['EVT-013','Logistics & Ops','Print conference materials','James Liu','07/01/2026','Not Started','High','','','VDR-009','',4,0,'Yes',3200,'','','','Yes'],
  // EVT-019 tasks
  ['EVT-019','Pre-Planning','City permit — Republic Square','Sam Torres','05/01/2026','Completed','Critical','','','','05/01/2026',3,3,'Yes',500,'Approved','','','No'],
  ['EVT-019','Venue & Vendors','Contract staging and tents','Sam Torres','05/15/2026','In Progress','Critical','','','VDR-024','',4,2,'Yes',8500,'Heritage Event Design','','','Yes'],
  ['EVT-019','Marketing & Comms','Partner with cultural organizations','Elena Park','05/20/2026','In Progress','High','','','','',10,5,'No',0,'12 orgs confirmed','','','No'],
  ['EVT-019','Marketing & Comms','Promote via radio and community flyers','Sam Torres','06/01/2026','Not Started','High','','','','',8,0,'Yes',2500,'','','','No'],
  ['EVT-019','Logistics & Ops','Cultural booth assignments','Sam Torres','07/01/2026','Not Started','High','','','','',6,0,'No',0,'30 booths','','','No'],
  ['EVT-019','Logistics & Ops','Volunteer recruitment and training','Priya Nair','07/15/2026','Not Started','High','','','','',8,0,'No',0,'Need 80 volunteers','','','No'],
  ['EVT-019','Day-Of Setup','Early morning setup — 5 AM','Sam Torres','08/15/2026','Not Started','Critical','','','VDR-024','',6,0,'No',0,'Crew of 20','','','Yes'],
  // EVT-025 tasks
  ['EVT-025','Pre-Planning','Partner with food bank','Sam Torres','08/01/2026','Not Started','Critical','','','','',4,0,'No',0,'Central Texas Food Bank','','','No'],
  ['EVT-025','Venue & Vendors','Secure Palmer Events Center','Sam Torres','08/15/2026','Not Started','Critical','','','VEN-003','',2,0,'Yes',12000,'','','','Yes'],
  ['EVT-025','Venue & Vendors','Catering — community meal logistics','Elena Park','09/01/2026','Not Started','Critical','','','','',6,0,'Yes',15000,'800 guests','','','No'],
  ['EVT-025','Logistics & Ops','Recruit and train 100 volunteers','Priya Nair','10/01/2026','Not Started','High','','','','',8,0,'No',0,'','','','No'],
  ['EVT-025','Marketing & Comms','Community outreach — flyers + social','Sam Torres','10/15/2026','Not Started','High','','','','',6,0,'Yes',1000,'','','','No'],
  ['EVT-025','Day-Of Setup','Table and meal station setup','Sam Torres','11/21/2026','Not Started','Critical','','','','',4,0,'No',0,'Setup starts 7 AM','','','No'],
  // Admin tasks
  ['EVT-001','Admin & Legal','Obtain event liability insurance','Marcus Cole','02/15/2026','Completed','Critical','','','VDR-023','02/15/2026',2,2,'Yes',4500,'Certificates on file','','','Yes'],
  ['EVT-001','Finance','Reconcile all vendor invoices','Elena Park','04/25/2026','Not Started','High','','','','',5,0,'Yes',0,'','','','No'],
  ['EVT-002','Finance','Reconcile gala revenue and costs','Elena Park','05/10/2026','Not Started','Critical','','','','',4,0,'Yes',0,'','','','No'],
  ['EVT-001','Post-Event','Upload event photos to shared drive','James Liu','04/20/2026','In Progress','Medium','','','VDR-007','',2,1,'No',0,'Photos received','','','Yes'],
  ['EVT-001','Post-Event','Write event debrief report','Diana Reyes','04/28/2026','Not Started','High','','','','',4,0,'No',0,'Board presentation 05/15','','','No'],
  ['EVT-002','Post-Event','Send donor acknowledgment letters','Diana Reyes','05/08/2026','Not Started','Critical','','','','',3,0,'No',0,'IRS requirement','','','No'],
  ['EVT-007','Post-Event','Send employer intro connections','James Liu','06/15/2026','Not Started','Medium','','','','',3,0,'No',0,'','','','No'],
  ['EVT-013','Admin & Legal','Secure conference insurance','Marcus Cole','05/15/2026','Not Started','Critical','','','VDR-023','',2,0,'Yes',4500,'','','','Yes'],
  ['EVT-013','Post-Event','Produce conference highlight video','James Liu','08/01/2026','Not Started','Medium','','','VDR-008','',8,0,'Yes',8000,'','','','Yes'],
  ['EVT-019','Post-Event','Document and archive festival outcomes','Sam Torres','09/01/2026','Not Started','Medium','','','','',4,0,'No',0,'For grant reporting','','','No'],
  ['EVT-022','Pre-Planning','Secure golf course deposit','Marcus Cole','06/01/2026','In Progress','Critical','','','VEN-009','',2,1,'Yes',9000,'','','','Yes'],
  ['EVT-022','Venue & Vendors','Recruit hole sponsors','Marcus Cole','06/15/2026','Not Started','High','','','','',8,0,'Yes',28000,'18 holes × $500','','','No'],
  ['EVT-022','Logistics & Ops','Order golf gifts and prizes','Marcus Cole','07/15/2026','Not Started','High','','','','',3,0,'Yes',3500,'','','','No'],
  ['EVT-026','Venue & Vendors','Contract DJ for year-end party','Diana Reyes','10/01/2026','Not Started','High','','','VDR-022','',1,0,'Yes',2500,'','','','Yes'],
  ['EVT-026','Logistics & Ops','Plan recognition awards program','Diana Reyes','11/01/2026','Not Started','High','','','','',4,0,'Yes',1500,'','','','No'],
  ['EVT-027','Logistics & Ops','Set up toy drop-off sites','Priya Nair','10/15/2026','Not Started','High','','','','',4,0,'No',0,'5 locations','','','No'],
  ['EVT-027','Marketing & Comms','Holiday toy drive media campaign','Priya Nair','10/20/2026','Not Started','Medium','','','','',6,0,'Yes',800,'','','','No'],
];

(async () => {
  const values = [];
  const fmt = [];

  // Title
  values.push({ range:"'Event Tasks'!A1", values:[['EVENT TASKS']] });
  values.push({ range:"'Event Tasks'!A2", values:[['Track all tasks by event and phase. Urgency and Overdue flags are auto-calculated.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,19), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,19), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,19), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,19), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:44}, fields:'pixelSize' }});

  // Stats row 3
  values.push({ range:"'Event Tasks'!A3", values:[[
    `=IFERROR("Total Tasks: "&COUNTA(C6:C500)&" | Completed: "&COUNTIF(F6:F500,"Completed")&" | In Progress: "&COUNTIF(F6:F500,"In Progress")&" | Overdue: "&COUNTIF(R6:R500,"YES"),"Loading...")`
  ]]});
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,19), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,19), cell:{userEnteredFormat:{
    backgroundColor:hex('#E9DBA7'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#5C4A00')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:2,endIndex:3},
    properties:{pixelSize:28}, fields:'pixelSize' }});

  // Info row 4
  values.push({ range:"'Event Tasks'!A4", values:[['ℹ Days Until Due (H), Urgency Flag (Q), and Overdue (R) are auto-calculated. Vendor Task (S) is a checkbox.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,3,4,0,19), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,3,4,0,19), cell:{userEnteredFormat:{
    backgroundColor:hex('#EEF4F8'), textFormat:{italic:true,fontSize:9,foregroundColor:hex('#4A5056')},
    verticalAlignment:'MIDDLE', padding:{left:10},
  }}, fields:'userEnteredFormat' }});

  // Header row 5
  values.push({ range:"'Event Tasks'!A5", values:[HEADERS] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,19), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5},
    properties:{pixelSize:32}, fields:'pixelSize' }});

  COL_WIDTHS.forEach((w,i)=>{
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1},
      properties:{pixelSize:w}, fields:'pixelSize' }});
  });

  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:5}},
    fields:'gridProperties.frozenRowCount' }});

  // Data rows with formulas
  const dataRows = TASKS.map((task,i)=>{
    const r = i+6;
    const row = [...task];
    // H=7: Days Until Due
    row[7] = `=IFERROR(IF(OR(F${r}="Completed",E${r}=""),"",INT(E${r}-TODAY())),"")`;
    // Q=16: Urgency Flag
    row[16] = `=IFERROR(IF(F${r}="Completed","✓ Done",IF(E${r}="","",IF(TODAY()>E${r},"OVERDUE",IF(E${r}-TODAY()<=7,"THIS WEEK",IF(E${r}-TODAY()<=30,"THIS MONTH","On Track"))))),"")`;
    // R=17: Overdue?
    row[17] = `=IFERROR(IF(AND(F${r}<>"Completed",E${r}<>"",TODAY()>E${r}),"YES",""),"")`;
    // S=18: Vendor Task checkbox — convert from string
    row[18] = row[18] === 'Yes';
    return row;
  });
  values.push({ range:"'Event Tasks'!A6", values: dataRows });

  // Alt row styling
  TASKS.forEach((_,i)=>{
    const ri = i+5;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,19), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#FDFCF9':'#F2F1EE'),
      textFormat:{foregroundColor:hex('#2F3336'),fontSize:9}, verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Formula col styling: H=7, Q=16, R=17
  TASKS.forEach((_,i)=>{
    const ri = i+5;
    [7,16,17].forEach(c=>{
      fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,c,c+1), cell:{userEnteredFormat:{
        backgroundColor:hex('#EEF4F8'), textFormat:{foregroundColor:hex('#3D6374'),fontSize:9},
        horizontalAlignment:'CENTER',
      }}, fields:'userEnteredFormat' }});
    });
  });

  // Checkbox: S=18
  fmt.push({ repeatCell:{ range: gridRange(SID,5,500,18,19), cell:{
    userEnteredFormat:{ horizontalAlignment:'CENTER' },
    dataValidation:{ condition:{ type:'BOOLEAN' }, showCustomUi:true },
  }, fields:'userEnteredFormat.horizontalAlignment,dataValidation' }});

  // Validations
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,1,2), rule:{
    condition:{ type:'ONE_OF_LIST', values: PHASES.map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:false,
  }}});
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,5,6), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Not Started','In Progress','Completed','Blocked','Deferred'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,6,7), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Critical','High','Medium','Low'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});

  // CF: Status (F=col5)
  const statusCF = [
    { v:'Completed',   bg:'#CBDCC5', text:'#2A5C35' },
    { v:'In Progress', bg:'#E9DBA7', text:'#5C4A00' },
    { v:'Not Started', bg:'#F2F1EE', text:'#4A5056' },
    { v:'Blocked',     bg:'#E8C0B8', text:'#5C1A1A' },
    { v:'Deferred',    bg:'#DDD5ED', text:'#3A1A5C' },
  ];
  statusCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,5,500,5,6)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // CF: Priority (G=col6)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,6,7)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Critical'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{foregroundColor:hex('#5C1A1A'),bold:true} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,6,7)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'High'}] },
      format:{ backgroundColor:hex('#F2D2B6'), textFormat:{foregroundColor:hex('#5C3A00')} } }
  }, index:0 }});

  // CF: Urgency (Q=col16)
  const urgencyCF = [
    { v:'OVERDUE',    bg:'#E8C0B8', text:'#5C1A1A' },
    { v:'THIS WEEK',  bg:'#F2D2B6', text:'#5C3A00' },
    { v:'THIS MONTH', bg:'#E9DBA7', text:'#5C4A00' },
    { v:'✓ Done',     bg:'#CBDCC5', text:'#2A5C35' },
    { v:'On Track',   bg:'#D5E4F2', text:'#1E3A5F' },
  ];
  urgencyCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,5,500,16,17)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // CF: Overdue row highlight (R=col17 = YES → whole row amber)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,0,19)],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:'=$R6="YES"'}] },
      format:{ backgroundColor:hex('#F2D2B6') } }
  }, index:0 }});

  // Currency format: O=14 Cost Estimate
  fmt.push({ repeatCell:{ range: gridRange(SID,5,500,14,15), cell:{userEnteredFormat:{
    numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
  }}, fields:'userEnteredFormat.numberFormat' }});

  // Filter
  fmt.push({ setBasicFilter:{ filter:{ range: gridRange(SID,4,500,0,19) }}});

  await valuesBatchUpdate(id, values, '08-tasks values');
  await batchUpdate(id, fmt, '08-tasks format');
  console.log('✅ Event Tasks done.');
})().catch(e => { console.error(e.errors||e.message||e); process.exit(1); });
