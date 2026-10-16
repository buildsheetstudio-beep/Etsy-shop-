'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 8;

const BUDGET_CATS = [
  ['BC-01','Venue & Facilities',75000,0,''],
  ['BC-02','Catering & Food',55000,0,''],
  ['BC-03','AV & Technology',48000,0,''],
  ['BC-04','Marketing & Advertising',22000,0,''],
  ['BC-05','Printing & Signage',8500,0,''],
  ['BC-06','Photography & Video',14000,0,''],
  ['BC-07','Entertainment',10000,0,''],
  ['BC-08','Security',5000,0,''],
  ['BC-09','Transportation',6500,0,''],
  ['BC-10','Décor & Flowers',12000,0,''],
  ['BC-11','Staffing',9000,0,''],
  ['BC-12','Insurance',5000,0,''],
  ['BC-13','Permits & Licenses',3000,0,''],
  ['BC-14','Speaker Fees',8000,0,''],
  ['BC-15','Awards & Recognition',4000,0,''],
  ['BC-16','Contingency Fund',15000,0,''],
  ['BC-17','Administrative',3000,0,''],
];

// 60 expenses
const EXPENSES = [
  ['EXP-001','EVT-001','Venue & Facilities','Austin Convention Center — Day 1-3','VEN-001','03/01/2026','Check',25000,25000,'Paid in Full','contract_acc.pdf','','',''],
  ['EXP-002','EVT-001','Catering & Food','Summit Catering Co — Day 1 Meals','VDR-001','03/10/2026','Check',6500,5000,'Deposit Paid','invoice_vdr001.pdf','','','Deposit 03/10'],
  ['EXP-003','EVT-001','AV & Technology','ProAV Solutions — Summit Full AV','VDR-003','03/15/2026','ACH',10000,10000,'Deposit Paid','invoice_vdr003.pdf','','',''],
  ['EXP-004','EVT-001','Marketing & Advertising','Social media ad campaign Q1','VDR-013','03/20/2026','Credit',3000,3000,'Paid in Full','receipt_mktg001.pdf','','',''],
  ['EXP-005','EVT-001','Photography & Video','Pixel Perfect Photography deposit','VDR-007','03/25/2026','Check',2000,2000,'Deposit Paid','invoice_vdr007.pdf','','',''],
  ['EXP-006','EVT-001','Photography & Video','Lens & Motion Video deposit','VDR-008','03/25/2026','Check',2500,2500,'Deposit Paid','invoice_vdr008.pdf','','',''],
  ['EXP-007','EVT-001','Printing & Signage','TexPrint — programs, badges, signage','VDR-009','04/01/2026','Check',3200,1000,'Deposit Paid','invoice_vdr009.pdf','','',''],
  ['EXP-008','EVT-001','Security','Austin Event Security — 2 days','VDR-011','04/01/2026','Check',2400,800,'Deposit Paid','invoice_vdr011.pdf','','',''],
  ['EXP-009','EVT-001','Transportation','City Shuttle Co — 3 day shuttle','VDR-012','04/01/2026','Check',4500,1500,'Deposit Paid','invoice_vdr012.pdf','','',''],
  ['EXP-010','EVT-001','Insurance','Event liability — 3 events','VDR-023','02/20/2026','Check',4500,4500,'Paid in Full','certificate_vdr023.pdf','','',''],
  ['EXP-011','EVT-001','Staffing','Field Events Staffing — 2 days','VDR-014','04/05/2026','ACH',9000,3000,'Deposit Paid','invoice_vdr014.pdf','','',''],
  ['EXP-012','EVT-001','Venue & Facilities','ACC Audio Breakout Rooms supplement','VEN-001','04/01/2026','Check',5000,5000,'Paid in Full','invoice_acc2.pdf','','','Room upgrade'],
  ['EXP-013','EVT-001','Catering & Food','Summit Catering Co — Day 2 Meals','VDR-001','04/01/2026','Check',6000,0,'Not Paid','','','','Due on delivery'],
  ['EXP-014','EVT-001','Catering & Food','Summit Catering Co — Day 3 Meals','VDR-001','04/01/2026','Check',5500,0,'Not Paid','','','','Due on delivery'],
  ['EXP-015','EVT-001','Speaker Fees','Dr. Sandra Okafor keynote fee','','04/01/2026','Check',3000,3000,'Paid in Full','receipt_speak001.pdf','','',''],
  ['EXP-016','EVT-001','Speaker Fees','Panel speakers honorariums (x4)','','04/01/2026','Check',2000,2000,'Paid in Full','receipt_speak002.pdf','','','$500 each'],
  ['EXP-017','EVT-001','AV & Technology','ProAV — Live streaming balance','VDR-003','04/15/2026','ACH',25000,0,'Not Paid','','','','Due post-event'],
  ['EXP-018','EVT-001','Décor & Flowers','Stage décor and greenery','','03/15/2026','Check',2000,2000,'Paid in Full','receipt_decor001.pdf','','',''],
  ['EXP-019','EVT-001','Administrative','Event coordination software annual','','02/01/2026','Credit',500,500,'Paid in Full','receipt_admin001.pdf','','',''],
  ['EXP-020','EVT-001','Permits & Licenses','Convention center permit fee','','02/10/2026','Check',300,300,'Paid in Full','receipt_permit001.pdf','','',''],
  // EVT-002
  ['EXP-021','EVT-002','Venue & Facilities','The Driskill Hotel — Grand Ballroom','VEN-002','02/15/2026','Check',16000,3000,'Deposit Paid','invoice_ven002.pdf','','',''],
  ['EXP-022','EVT-002','Catering & Food','Gala Fine Dining — 300 guests','VDR-002','02/20/2026','Check',22000,7000,'Deposit Paid','invoice_vdr002.pdf','','',''],
  ['EXP-023','EVT-002','Entertainment','The Jazzmakers — live trio','VDR-010','03/05/2026','Check',3500,1000,'Deposit Paid','invoice_vdr010.pdf','','',''],
  ['EXP-024','EVT-002','Décor & Flowers','Bloom Event Décor — gala florals','VDR-006','03/10/2026','Check',8500,2500,'Deposit Paid','invoice_vdr006.pdf','','',''],
  ['EXP-025','EVT-002','Photography & Video','Gala photography (included in EVT-001 pkg)','VDR-007','03/25/2026','Check',0,0,'Paid in Full','','','','Bundled with EVT-001'],
  ['EXP-026','EVT-002','Transportation','Quick Park Valet — 80 cars','VDR-015','04/01/2026','Check',1800,500,'Deposit Paid','invoice_vdr015.pdf','','',''],
  ['EXP-027','EVT-002','Marketing & Advertising','Gala invitation design & print','','03/01/2026','Credit',2000,2000,'Paid in Full','receipt_mktg002.pdf','','',''],
  ['EXP-028','EVT-002','Administrative','Auction management software license','','01/20/2026','Credit',200,200,'Paid in Full','receipt_admin002.pdf','','',''],
  // EVT-003
  ['EXP-029','EVT-003','Venue & Facilities','Zilker Park permit','VEN-005','03/05/2026','Check',500,500,'Paid in Full','receipt_ven005.pdf','','',''],
  ['EXP-030','EVT-003','AV & Technology','Park Sounds — outdoor PA','VDR-004','04/01/2026','Check',4500,1500,'Deposit Paid','invoice_vdr004.pdf','','',''],
  ['EXP-031','EVT-003','Décor & Flowers','GreenLawn Supply — tents, tables','VDR-017','04/05/2026','Check',6000,2000,'Deposit Paid','invoice_vdr017.pdf','','',''],
  ['EXP-032','EVT-003','Printing & Signage','Health fair flyers and banners','','04/10/2026','Credit',800,800,'Paid in Full','receipt_print003.pdf','','',''],
  ['EXP-033','EVT-003','Marketing & Advertising','Community outreach radio spots','','04/15/2026','Credit',500,500,'Paid in Full','receipt_mktg003.pdf','','',''],
  // EVT-004
  ['EXP-034','EVT-004','Venue & Facilities','WeWork Downtown — full day room','VEN-004','04/01/2026','Credit',800,800,'Paid in Full','invoice_ven004.pdf','','',''],
  ['EXP-035','EVT-004','Catering & Food','Fresh Catering — box lunches x40','VDR-020','05/10/2026','Check',800,0,'Not Paid','','','','Invoice pending'],
  ['EXP-036','EVT-004','Printing & Signage','Workshop workbooks x45','VDR-009','04/10/2026','Check',225,225,'Paid in Full','receipt_print004.pdf','','',''],
  ['EXP-037','EVT-004','AV & Technology','Projector rental (WeWork provided)','VEN-004','','Credit',0,0,'Paid in Full','','','','Included in room'],
  ['EXP-038','EVT-004','Marketing & Advertising','Workshop email campaign','','04/01/2026','Credit',200,200,'Paid in Full','receipt_mktg004.pdf','','',''],
  // EVT-006
  ['EXP-039','EVT-006','Venue & Facilities','Walnut Creek Park — field reservation','VEN-007','03/25/2026','Check',600,600,'Paid in Full','receipt_ven007.pdf','','',''],
  ['EXP-040','EVT-006','Awards & Recognition','Trophy House — trophies & medals','VDR-018','04/10/2026','Check',2200,700,'Deposit Paid','invoice_vdr018.pdf','','',''],
  ['EXP-041','EVT-006','Permits & Licenses','Youth sports event permit','','03/20/2026','Check',100,100,'Paid in Full','receipt_permit006.pdf','','',''],
  ['EXP-042','EVT-006','Marketing & Advertising','Tournament flyers and social media','','04/15/2026','Credit',400,400,'Paid in Full','receipt_mktg006.pdf','','',''],
  ['EXP-043','EVT-006','Catering & Food','Concession supplies — 2 days','','05/01/2026','Check',1200,0,'Not Paid','','','','Vendor confirmed'],
  // EVT-007
  ['EXP-044','EVT-007','AV & Technology','Zoom Events platform — annual','VDR-019','05/01/2026','Credit',1500,1500,'Paid in Full','receipt_vdr019.pdf','','',''],
  ['EXP-045','EVT-007','AV & Technology','Clarity Digital — virtual tech support','VDR-005','05/05/2026','Credit',3000,1000,'Deposit Paid','invoice_vdr005.pdf','','',''],
  ['EXP-046','EVT-007','Marketing & Advertising','Career fair email campaign','','05/10/2026','Credit',500,500,'Paid in Full','receipt_mktg007.pdf','','',''],
  // EVT-013
  ['EXP-047','EVT-013','Venue & Facilities','Palmer Events Center deposit','VEN-003','05/01/2026','Check',4000,4000,'Deposit Paid','invoice_ven003.pdf','','',''],
  ['EXP-048','EVT-013','AV & Technology','ProAV — hybrid conference AV','VDR-003','05/15/2026','ACH',10000,10000,'Deposit Paid','invoice_vdr003b.pdf','','',''],
  ['EXP-049','EVT-013','Marketing & Advertising','Bold Brand Marketing campaign','VDR-013','06/01/2026','Check',12000,4000,'Deposit Paid','invoice_vdr013.pdf','','',''],
  ['EXP-050','EVT-013','Insurance','Conference liability insurance','VDR-023','05/20/2026','Check',4500,4500,'Paid in Full','certificate_vdr023b.pdf','','',''],
  // EVT-019
  ['EXP-051','EVT-019','Venue & Facilities','Republic Square permit','','05/05/2026','Check',500,500,'Paid in Full','receipt_ven011.pdf','','',''],
  ['EXP-052','EVT-019','Décor & Flowers','Heritage Event Design — staging','VDR-024','05/20/2026','Check',8500,2500,'Deposit Paid','invoice_vdr024.pdf','','',''],
  ['EXP-053','EVT-019','Marketing & Advertising','Festival marketing campaign','','06/10/2026','Credit',2500,0,'Not Paid','','','',''],
  ['EXP-054','EVT-019','Entertainment','Cultural performers fees','','07/01/2026','Check',5000,0,'Not Paid','','','',''],
  ['EXP-055','EVT-019','Catering & Food','Food vendor coordination fee','','07/01/2026','Check',1500,0,'Not Paid','','','',''],
  // Revenue placeholder entries
];

// 30 revenue records
const REVENUE = [
  ['REV-001','EVT-001','Ticket Sales','General Admission tickets — 440 attendees','','03/31/2026','Bank Transfer',66000,66000,'Received','reg_system_001',''],
  ['REV-002','EVT-001','Ticket Sales','VIP tickets — 40 attendees','','03/31/2026','Bank Transfer',6000,6000,'Received','reg_system_002',''],
  ['REV-003','EVT-001','Sponsorship','Platinum Sponsor — Foundation Hub','WRN Hughes','04/01/2026','Check',15000,15000,'Received','receipt_sp001',''],
  ['REV-004','EVT-001','Sponsorship','Gold Sponsor — TechBridge','N Brooks','04/01/2026','Check',7500,7500,'Received','receipt_sp002',''],
  ['REV-005','EVT-001','Sponsorship','Silver Sponsors (x5)','','04/01/2026','Check',12500,12500,'Received','receipt_sp003',''],
  ['REV-006','EVT-001','Registration Fee','Late registration fees (x20)','','04/10/2026','Online',2000,2000,'Received','reg_system_003',''],
  ['REV-007','EVT-001','Merchandise','Logo merchandise sales at summit','','04/17/2026','Cash/Card',3500,3500,'Received','pos_receipt_001',''],
  ['REV-008','EVT-002','Ticket Sales','General Admission — 260 tickets','','04/15/2026','Online',65000,65000,'Received','gala_reg_001',''],
  ['REV-009','EVT-002','Ticket Sales','VIP tables — 8 tables x$5,000','','04/10/2026','Check',40000,40000,'Received','gala_reg_002',''],
  ['REV-010','EVT-002','Donation','Fund-a-Need live auction (estimated)','','05/02/2026','Various',25000,0,'Pending','',''],
  ['REV-011','EVT-002','Sponsorship','Event sponsor — Novak Enterprises','C Novak','03/15/2026','Check',10000,10000,'Received','receipt_sp004',''],
  ['REV-012','EVT-002','Sponsorship','Corporate table sponsors (x4)','','03/15/2026','Check',8000,8000,'Received','receipt_sp005',''],
  ['REV-013','EVT-003','Donation','Health fair booth donations (estimated)','','05/15/2026','Cash',500,0,'Pending','',''],
  ['REV-014','EVT-003','Sponsorship','City health dept co-sponsor','','03/01/2026','Grant',3000,3000,'Received','grant_ev003',''],
  ['REV-015','EVT-004','Ticket Sales','Workshop registrations — 35 paid','','05/10/2026','Online',2625,2625,'Received','workshop_reg_001',''],
  ['REV-016','EVT-004','Ticket Sales','On-site registrations — 5','','05/20/2026','Card',375,375,'Received','pos_receipt_002',''],
  ['REV-017','EVT-006','Ticket Sales','Team registration — 16 teams x$250','','05/25/2026','Online',4000,4000,'Received','soccer_reg_001',''],
  ['REV-018','EVT-006','Sponsorship','Tournament field sponsor','','04/15/2026','Check',1500,1500,'Received','receipt_sp006',''],
  ['REV-019','EVT-006','Ticket Sales','Spectator admission (estimated)','','06/06/2026','Cash',800,0,'Pending','',''],
  ['REV-020','EVT-007','Registration Fee','Virtual career fair — 500 free tickets','','','',0,0,'N/A','','Free event'],
  ['REV-021','EVT-007','Sponsorship','Employer showcase sponsor (x3)','','05/01/2026','Invoice',6000,6000,'Received','receipt_sp007',''],
  ['REV-022','EVT-013','Ticket Sales','General conference tickets (est. 600x$200)','','07/01/2026','Online',120000,0,'Pending','','Reg opens 05/01'],
  ['REV-023','EVT-013','Ticket Sales','VIP passes (est. 100x$500)','','07/01/2026','Online',50000,0,'Pending','',''],
  ['REV-024','EVT-013','Sponsorship','Title Sponsor','','05/01/2026','Check',25000,0,'Pending','','Contract out'],
  ['REV-025','EVT-013','Sponsorship','Gold sponsors (x4 x$5,000)','','05/15/2026','Check',20000,0,'Pending','',''],
  ['REV-026','EVT-019','Donation','Festival donations (estimated)','','08/16/2026','Various',5000,0,'Pending','',''],
  ['REV-027','EVT-019','Sponsorship','Cultural org. sponsors','','06/01/2026','Various',12000,0,'Pending','',''],
  ['REV-028','EVT-022','Ticket Sales','Golf tournament entry — 144 x$175','','09/01/2026','Online',25200,0,'Pending','',''],
  ['REV-029','EVT-022','Sponsorship','Hole sponsors (est. 12 holes)','','08/01/2026','Check',12000,0,'Pending','',''],
  ['REV-030','EVT-022','Merchandise','Golf merchandise sales','','09/09/2026','Cash',2000,0,'Pending','',''],
];

(async () => {
  const values = [];
  const fmt = [];

  // Title
  values.push({ range:"'Budget & Expenses'!A1", values:[['BUDGET & EXPENSES']] });
  values.push({ range:"'Budget & Expenses'!A2", values:[['Three sections: (1) Category Budget Summary — rows 5-25, (2) Expense Log — rows 35-1034, (3) Revenue Log — rows 1045-1344.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,13), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,13), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:44}, fields:'pixelSize' }});

  // KPI summary bar row 3
  values.push({ range:"'Budget & Expenses'!A3", values:[[
    `=IFERROR("Total Budget: $"&TEXT(SUM(C6:C22),"#,##0")&" | Total Spent: $"&TEXT(SUM(G41:G1040),"#,##0")&" | Total Revenue: $"&TEXT(SUM(F1050:F1350),"#,##0")&" | Net: $"&TEXT(SUM(F1050:F1350)-SUM(G41:G1040),"#,##0"),"Loading...")`
  ]]});
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,13), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:2,endIndex:3},
    properties:{pixelSize:28}, fields:'pixelSize' }});

  // ── SECTION 1: Category Budget (rows 5-25, 0-indexed 4-24) ────────
  values.push({ range:"'Budget & Expenses'!A5", values:[['BUDGET SUMMARY BY CATEGORY']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,4,5,0,13), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:11,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  const budgetHeaders = ['Cat ID','Category','Budget Allocated','Actual Spent','Variance','% Used','Status','Notes'];
  values.push({ range:"'Budget & Expenses'!A6", values:[budgetHeaders] });
  fmt.push({ repeatCell:{ range: gridRange(SID,5,6,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  const budgetData = BUDGET_CATS.map((cat,i)=>{
    const r = i+7;
    return [
      cat[0], cat[1], cat[2],
      `=IFERROR(SUMPRODUCT(('Budget & Expenses'!C$41:C$1040=B${r})*('Budget & Expenses'!G$41:G$1040)),0)`,
      `=IFERROR(C${r}-D${r},0)`,
      `=IFERROR(IF(C${r}=0,"",D${r}/C${r}),"")`,
      `=IFERROR(IF(F${r}="","",IF(F${r}>=1,"Over Budget",IF(F${r}>=0.8,"Near Limit","On Track"))),"")`,
      cat[4],
    ];
  });
  values.push({ range:"'Budget & Expenses'!A7", values: budgetData });

  // Budget totals row
  const totalRow = BUDGET_CATS.length + 7; // row after data
  values.push({ range:`'Budget & Expenses'!A${totalRow}`, values:[['','TOTAL',
    `=IFERROR(SUM(C7:C${totalRow-1}),0)`,
    `=IFERROR(SUM(D7:D${totalRow-1}),0)`,
    `=IFERROR(C${totalRow}-D${totalRow},0)`,
    `=IFERROR(D${totalRow}/C${totalRow},"")`,
    '','',
  ]]});
  fmt.push({ repeatCell:{ range: gridRange(SID,totalRow-1,totalRow,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,foregroundColor:hex('#FFFFFF')},
  }}, fields:'userEnteredFormat' }});

  // Budget category alt rows
  BUDGET_CATS.forEach((_,i)=>{
    const ri = i+6;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,8), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#EEF4F8':'#F7F6F3'),
      textFormat:{foregroundColor:hex('#2F3336'),fontSize:9}, verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Formula cols in budget (D=3,E=4,F=5,G=6)
  [3,4,5,6].forEach(c=>{
    BUDGET_CATS.forEach((_,i)=>{
      fmt.push({ repeatCell:{ range: gridRange(SID,i+6,i+7,c,c+1), cell:{userEnteredFormat:{
        backgroundColor:hex('#EEF4F8'), textFormat:{foregroundColor:hex('#3D6374'),fontSize:9},
        horizontalAlignment:'RIGHT',
      }}, fields:'userEnteredFormat' }});
    });
  });

  // Currency: C=2,D=3,E=4
  [2,3,4].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,5,30,c,c+1), cell:{userEnteredFormat:{
      numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
    }}, fields:'userEnteredFormat.numberFormat' }});
  });
  // Percent: F=5
  fmt.push({ repeatCell:{ range: gridRange(SID,5,30,5,6), cell:{userEnteredFormat:{
    numberFormat:{ type:'PERCENT', pattern:'0%' }
  }}, fields:'userEnteredFormat.numberFormat' }});

  // Budget CF: % Used >= 100% → red, >= 80% → amber
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,6,24,5,6)],
    booleanRule:{ condition:{ type:'NUMBER_GREATER_THAN_EQ', values:[{userEnteredValue:'1'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{foregroundColor:hex('#5C1A1A'),bold:true} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,6,24,5,6)],
    booleanRule:{ condition:{ type:'NUMBER_GREATER_THAN_EQ', values:[{userEnteredValue:'0.8'}] },
      format:{ backgroundColor:hex('#F2D2B6'), textFormat:{foregroundColor:hex('#5C3A00')} } }
  }, index:0 }});

  // Status CF (G=6 in budget)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,6,24,6,7)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Over Budget'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{foregroundColor:hex('#5C1A1A'),bold:true} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,6,24,6,7)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Near Limit'}] },
      format:{ backgroundColor:hex('#F2D2B6'), textFormat:{foregroundColor:hex('#5C3A00')} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,6,24,6,7)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'On Track'}] },
      format:{ backgroundColor:hex('#CBDCC5'), textFormat:{foregroundColor:hex('#2A5C35')} } }
  }, index:0 }});

  // ── SECTION 2: Expense Log (rows 35-1034, 0-indexed 34-1033) ──────
  // 0-indexed 34
  values.push({ range:"'Budget & Expenses'!A35", values:[['EXPENSE LOG']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,34,35,0,13), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,34,35,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:11,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:34,endIndex:35},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  const expHeaders = ['Exp ID','Event ID','Category','Description','Vendor/Payee','Date','Method',
    'Amount','Amount Paid','Payment Status','Receipt File','Verified?','Notes'];
  // 0-indexed 35
  values.push({ range:"'Budget & Expenses'!A36", values:[expHeaders] });
  fmt.push({ repeatCell:{ range: gridRange(SID,35,36,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  // Expense data rows (0-indexed 36 = row 37)
  const expData = EXPENSES.map((exp,i)=>{
    const row = [...exp];
    // Balance Due = Amount - Amount Paid (col index 7=H, 8=I)
    row[11] = false; // Verified checkbox
    return row;
  });
  values.push({ range:"'Budget & Expenses'!A37", values: expData });

  // Expense alt rows
  EXPENSES.forEach((_,i)=>{
    const ri = i+36;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,13), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#FDFCF9':'#F2F1EE'),
      textFormat:{foregroundColor:hex('#2F3336'),fontSize:9}, verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Checkbox for Verified (L=11)
  fmt.push({ repeatCell:{ range: gridRange(SID,36,1040,11,12), cell:{
    userEnteredFormat:{ horizontalAlignment:'CENTER' },
    dataValidation:{ condition:{ type:'BOOLEAN' }, showCustomUi:true },
  }, fields:'userEnteredFormat.horizontalAlignment,dataValidation' }});

  // Currency: H=7, I=8
  [7,8].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,36,1040,c,c+1), cell:{userEnteredFormat:{
      numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0.00' }
    }}, fields:'userEnteredFormat.numberFormat' }});
  });

  // Payment status CF (J=9)
  const paymentCF = [
    { v:'Paid in Full', bg:'#CBDCC5', text:'#2A5C35' },
    { v:'Deposit Paid', bg:'#D5E4F2', text:'#1E3A5F' },
    { v:'Partial',      bg:'#E9DBA7', text:'#5C4A00' },
    { v:'Not Paid',     bg:'#E8C0B8', text:'#5C1A1A' },
    { v:'N/A',          bg:'#F2F1EE', text:'#4A5056' },
  ];
  paymentCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,36,1040,9,10)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // Category validation for expenses
  fmt.push({ setDataValidation:{ range: gridRange(SID,36,1040,2,3), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Venue & Facilities','Catering & Food','AV & Technology',
      'Marketing & Advertising','Printing & Signage','Photography & Video','Entertainment',
      'Security','Transportation','Décor & Flowers','Staffing','Insurance','Permits & Licenses',
      'Speaker Fees','Awards & Recognition','Administrative','Other'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:false,
  }}});
  fmt.push({ setDataValidation:{ range: gridRange(SID,36,1040,9,10), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Not Paid','Deposit Paid','Partial','Paid in Full','N/A'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});

  // ── SECTION 3: Revenue Log (rows 1045-1344, 0-indexed 1044-1343) ─
  // 0-indexed 1044
  values.push({ range:"'Budget & Expenses'!A1045", values:[['REVENUE LOG']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,1044,1045,0,13), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1044,1045,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:11,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1044,endIndex:1045},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  const revHeaders = ['Rev ID','Event ID','Revenue Type','Description','Source/Payer','Date','Method',
    'Amount','Amount Received','Payment Status','Receipt/Ref','Notes'];
  values.push({ range:"'Budget & Expenses'!A1046", values:[revHeaders] });
  fmt.push({ repeatCell:{ range: gridRange(SID,1045,1046,0,12), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  values.push({ range:"'Budget & Expenses'!A1047", values: REVENUE });

  // Revenue alt rows
  REVENUE.forEach((_,i)=>{
    const ri = i+1046;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,12), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#F2F8F2':'#FDFCF9'),
      textFormat:{foregroundColor:hex('#2F3336'),fontSize:9}, verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Revenue CF: Received = green, Pending = amber
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,1046,1345,9,10)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Received'}] },
      format:{ backgroundColor:hex('#CBDCC5'), textFormat:{foregroundColor:hex('#2A5C35'),bold:true} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,1046,1345,9,10)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Pending'}] },
      format:{ backgroundColor:hex('#E9DBA7'), textFormat:{foregroundColor:hex('#5C4A00')} } }
  }, index:0 }});

  // Currency: H=7, I=8 for revenue
  [7,8].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,1046,1345,c,c+1), cell:{userEnteredFormat:{
      numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0.00' }
    }}, fields:'userEnteredFormat.numberFormat' }});
  });

  // Revenue type validation (C=2)
  fmt.push({ setDataValidation:{ range: gridRange(SID,1046,1345,2,3), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Ticket Sales','Sponsorship','Donation','Registration Fee',
      'Merchandise','Grant','Other'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:false,
  }}});

  // Column widths
  const colWidths = [75,80,160,220,160,100,90,100,100,110,160,80,180];
  colWidths.forEach((w,i)=>{
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1},
      properties:{pixelSize:w}, fields:'pixelSize' }});
  });

  // Freeze row 6 (budget headers)
  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:6}},
    fields:'gridProperties.frozenRowCount' }});

  // Filters on expense log
  fmt.push({ setBasicFilter:{ filter:{ range: gridRange(SID,35,1040,0,13) }}});

  await valuesBatchUpdate(id, values, '10-budget values');
  await batchUpdate(id, fmt, '10-budget format');
  console.log('✅ Budget & Expenses done.');
})().catch(e => { console.error(e.errors||e.message||e); process.exit(1); });
