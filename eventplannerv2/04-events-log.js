'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 2;

const CATEGORIES = ['Business','Conference','Workshop','Networking','Community','Nonprofit',
  'School','Fundraiser','Social','Birthday','Wedding','Family','Holiday','Sports','Virtual','Other'];
const CAT_COLORS = {
  Business:'#C9DCEF', Conference:'#D8CDE8', Workshop:'#F2D2B6', Networking:'#C5E0DD',
  Community:'#CBDCC5', Nonprofit:'#E9DBA7', School:'#D5E4F2',  Fundraiser:'#E8C0B8',
  Social:'#E9CCD8',   Birthday:'#F3D0BB',   Wedding:'#DDD0DF',  Family:'#EEE2C8',
  Holiday:'#D0E5D8',  Sports:'#C6D9E8',     Virtual:'#DDD5ED',  Other:'#E3E1DE',
};

// 29 columns A:AC
const HEADERS = [
  'Event ID','Event Name','Event Type','Start Date','End Date','Status',
  'Category','Venue / Location','City','State','Virtual Link',
  'Organizer','Co-Organizer','Target Attendance','Confirmed Attendance',
  'Ticket Price','Total Revenue','Total Cost','Net Profit/Loss',
  'Budget Allocated','Budget Spent','Budget Variance',
  'Recurrence','Next Occurrence','Tasks Complete','Tasks Total','% Complete',
  'Priority','Notes',
];
// A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12 N=13 O=14
// P=15 Q=16 R=17 S=18 T=19 U=20 V=21 W=22 X=23 Y=24 Z=25 AA=26 AB=27 AC=28

function colLetter(i) {
  if (i < 26) return String.fromCharCode(65+i);
  return 'A' + String.fromCharCode(65+i-26);
}

const COL_WIDTHS = [
  80,180,110,100,100,110,
  110,200,110,70,160,
  140,140,80,80,
  90,100,100,100,
  100,100,100,
  100,100,80,80,80,
  80,200,
];

// 30 sample events
const EVENTS = [
  ['EVT-001','Annual Leadership Summit','Conference','04/15/2026','04/17/2026','Confirmed','Conference','Austin Convention Center','Austin','TX','','Diana Reyes','Marcus Cole',500,480,150,'','','','75000','','','Annual','04/15/2027','','','','High','Keynote speakers confirmed'],
  ['EVT-002','Spring Fundraising Gala','Fundraiser','05/02/2026','05/02/2026','Confirmed','Fundraiser','The Driskill Hotel','Austin','TX','','Diana Reyes','','300',280,250,'','','','45000','','','None','','','','','Critical','Auction items needed'],
  ['EVT-003','Community Health Fair','Community','05/15/2026','05/15/2026','Planning','Community','Zilker Park','Austin','TX','','Elena Park','Sam Torres',1000,'',0,'','','','8000','','','Annual','05/15/2027','','','','Medium','Vendor booths needed'],
  ['EVT-004','Digital Marketing Workshop','Workshop','05/20/2026','05/20/2026','Confirmed','Workshop','WeWork Downtown','Austin','TX','','James Liu','',40,'',75,'','','','3200','','','None','','','','','High','Material kits needed'],
  ['EVT-005','Nonprofit Board Meeting','Business','05/27/2026','05/27/2026','Confirmed','Business','Foundation HQ','Austin','TX','','Diana Reyes','',25,'',0,'','','','500','','','Monthly','06/27/2026','','','','Medium','Quarterly reports due'],
  ['EVT-006','Youth Soccer Tournament','Sports','06/05/2026','06/06/2026','Planning','Sports','Walnut Creek Park','Austin','TX','','Sam Torres','Priya Nair',200,'',25,'','','','6000','','','Annual','06/05/2027','','','','High','Referee scheduling pending'],
  ['EVT-007','Virtual Career Fair','Virtual','06/10/2026','06/10/2026','Confirmed','Virtual','Online','','','https://zoom.us/ev001','James Liu','',500,'',0,'','','','2000','','','Annual','06/10/2027','','','','Medium','Platform link active'],
  ['EVT-008','Summer Networking Mixer','Networking','06/17/2026','06/17/2026','Planning','Networking','The Domain Rooftop','Austin','TX','','Elena Park','',120,'',35,'','','','4500','','','None','','','','','Medium','Catering selections needed'],
  ['EVT-009','School Science Fair','School','06/19/2026','06/19/2026','Confirmed','School','Anderson High School','Austin','TX','','Priya Nair','Sam Torres',300,'',0,'','','','1800','','','Annual','06/19/2027','','','','High','Judge registration open'],
  ['EVT-010','Birthday Bash — Thompson Family','Birthday','06/25/2026','06/25/2026','Planning','Birthday','Barton Springs Pavilion','Austin','TX','','Diana Reyes','',80,'',0,'','','','3500','','','Annual','06/25/2027','','','','Low','Surprise party'],
  ['EVT-011','Independence Day Block Party','Holiday','07/04/2026','07/04/2026','Planning','Holiday','Neighborhood Park','Austin','TX','','Sam Torres','',400,'',0,'','','','5000','','','Annual','07/04/2027','','','','Medium','Permits filed'],
  ['EVT-012','Grant Writing Seminar','Nonprofit','07/08/2026','07/08/2026','Confirmed','Nonprofit','Austin Public Library','Austin','TX','','Elena Park','',60,'',0,'','','','1200','','','None','','','','','Medium','Workbooks printing'],
  ['EVT-013','Hybrid Tech Conference','Conference','07/15/2026','07/17/2026','Planning','Conference','Palmer Events Center','Austin','TX','https://stream.tech/htc26','James Liu','Marcus Cole',800,'',200,'','','','120000','','','Annual','07/15/2027','','','','Critical','Hybrid AV contract pending'],
  ['EVT-014','Family Reunion Picnic','Family','07/18/2026','07/19/2026','Planning','Family','McKinney Falls State Park','Austin','TX','','Diana Reyes','',150,'',0,'','','','2500','','','Annual','07/18/2027','','','','Low','Potluck sign-up needed'],
  ['EVT-015','Volunteer Appreciation Dinner','Nonprofit','07/22/2026','07/22/2026','Confirmed','Nonprofit','The LINE Hotel','Austin','TX','','Elena Park','',100,95,0,'','','','7500','','','Annual','07/22/2027','','','','High','Menu approved'],
  ['EVT-016','Back-to-School Resource Fair','Community','08/01/2026','08/01/2026','Planning','Community','Pflugerville Community Center','Pflugerville','TX','','Priya Nair','Sam Torres',500,'',0,'','','','4000','','','Annual','08/01/2027','','','','High','Sponsor outreach needed'],
  ['EVT-017','Wellness & Yoga Retreat','Social','08/07/2026','08/09/2026','Planning','Social','Camp Longhorn','Burnet','TX','','Elena Park','',50,'',350,'','','','18000','','','Annual','08/07/2027','','','','Medium','Waitlist open'],
  ['EVT-018','Business Plan Competition','Business','08/12/2026','08/12/2026','Confirmed','Business','UT McCombs School','Austin','TX','','James Liu','Marcus Cole',200,'',0,'','','','9500','','','Annual','08/12/2027','','','','High','Judge panel confirmed'],
  ['EVT-019','Cultural Heritage Festival','Community','08/15/2026','08/16/2026','Planning','Community','Republic Square','Austin','TX','','Sam Torres','Elena Park',2000,'',0,'','','','25000','','','Annual','08/15/2027','','','','Critical','City permit approved'],
  ['EVT-020','Team Leadership Workshop','Workshop','08/19/2026','08/19/2026','Planning','Workshop','Foundation HQ','Austin','TX','','Diana Reyes','',30,'',0,'','','','1500','','','None','','','','','Medium','Facilitator booked'],
  ['EVT-021','Online Fundraising Webinar','Virtual','08/25/2026','08/25/2026','Confirmed','Virtual','Online','','','https://zoom.us/ev021','Elena Park','',200,'',0,'','','','800','','','None','','','','','Medium','Recording available after'],
  ['EVT-022','Fall Golf Tournament','Sports','09/09/2026','09/09/2026','Planning','Sports','Barton Creek Resort','Austin','TX','','Marcus Cole','',144,'',175,'','','','28000','','','Annual','09/09/2027','','','','High','Sponsorship spots open'],
  ['EVT-023','Kids Halloween Party','Birthday','10/17/2026','10/17/2026','Planning','Birthday','Brush Square Park','Austin','TX','','Priya Nair','',200,'',0,'','','','3000','','','Annual','10/17/2027','','','','Medium','Costume contest planned'],
  ['EVT-024','Board Development Retreat','Business','10/21/2026','10/23/2026','Planning','Business','Lost Pines Resort','Bastrop','TX','','Diana Reyes','',20,'',0,'','','','12000','','','Annual','10/21/2027','','','','High','Agenda drafting'],
  ['EVT-025','Community Thanksgiving Feast','Community','11/21/2026','11/21/2026','Planning','Community','Palmer Events Center','Austin','TX','','Sam Torres','Elena Park',800,'',0,'','','','15000','','','Annual','11/21/2027','','','','Critical','Food bank partner confirmed'],
  ['EVT-026','Year-End Staff Celebration','Social','12/11/2026','12/11/2026','Planning','Social','The Driskill Hotel','Austin','TX','','Diana Reyes','Marcus Cole',75,'',0,'','','','8500','','','Annual','12/11/2027','','','','High','AV and DJ needed'],
  ['EVT-027','Holiday Toy Drive','Fundraiser','12/01/2026','12/18/2026','Planning','Fundraiser','Multiple Locations','Austin','TX','','Priya Nair','',0,'',0,'','','','2000','','','Annual','12/01/2027','','','','High','Drop-off sites confirmed'],
  ['EVT-028','New Year Networking Event','Networking','12/30/2026','12/30/2026','Planning','Networking','W Austin Hotel','Austin','TX','','James Liu','',200,'',50,'','','','9000','','','Annual','12/30/2027','','','','Medium','Venue deposit due'],
  ['EVT-029','Alumni Reunion Weekend','Social','01/16/2027','01/17/2027','Planning','Social','Marriott Downtown','Austin','TX','','Marcus Cole','Elena Park',250,'',125,'','','','30000','','','Annual','01/16/2028','','','','High','Save-the-dates sent'],
  ['EVT-030','Spring Conference Planning','Conference','01/28/2027','01/28/2027','Planning','Conference','Foundation HQ','Austin','TX','','Diana Reyes','',15,'',0,'','','','500','','','None','','','','','Medium','Pre-planning kickoff'],
];

(async () => {
  const values = [];
  const fmt = [];

  // Title row
  values.push({ range: "'Master Event Log'!A1", values: [['MASTER EVENT LOG']] });
  values.push({ range: "'Master Event Log'!A2", values: [['Track all events — single-day, multi-day, virtual, hybrid, and recurring. Formulas auto-calculate revenue, costs, and completion.']] });

  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,29), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,29), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,29), cell:{ userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,29), cell:{ userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:44}, fields:'pixelSize' }});

  // Instruction row 3
  values.push({ range: "'Master Event Log'!A3", values:[['ℹ Formula columns (Q-V, Y-AA) are calculated automatically — do not edit them directly.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,0,29), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,29), cell:{ userEnteredFormat:{
    backgroundColor:hex('#EEF4F8'), textFormat:{italic:true,fontSize:9,foregroundColor:hex('#4A5056')},
    verticalAlignment:'MIDDLE', padding:{left:10},
  }}, fields:'userEnteredFormat' }});

  // Filter/freeze row 4 (blank gap) + header row 5
  values.push({ range: "'Master Event Log'!A5", values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,29), cell:{ userEnteredFormat:{
    backgroundColor:hex('#3D6374'),
    textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    wrapStrategy:'WRAP',
    borders:{ bottom:{ style:'SOLID', color:hex('#FFFFFF'), width:2 } },
  }}, fields:'userEnteredFormat' }});

  // Column widths
  COL_WIDTHS.forEach((w,i)=>{
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1},
      properties:{pixelSize:w}, fields:'pixelSize' }});
  });

  // Row heights: title rows
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:2},
    properties:{pixelSize:28}, fields:'pixelSize' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5},
    properties:{pixelSize:32}, fields:'pixelSize' }});

  // Freeze row 5 only (no frozenColumnCount — merged title rows span all columns)
  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:5}},
    fields:'gridProperties.frozenRowCount' }});

  // Data rows 6-35 (0-indexed 5-34)
  // Formula columns (0-indexed): Q=16 Total Revenue, R=17 Total Cost, S=18 Net P/L,
  //   V=21 Budget Variance, Y=24 Tasks %, Z=25 Tasks Total, AA=26 % Complete
  const dataValues = [];
  EVENTS.forEach((ev, i) => {
    const r = i + 6; // 1-indexed row
    const ri = i + 5; // 0-indexed

    // Columns: A=EventID, B=Name, C=Type, D=Start, E=End, F=Status,
    // G=Category, H=Venue, I=City, J=State, K=VirtualLink,
    // L=Organizer, M=Co-Org, N=TargetAtt, O=ConfirmedAtt,
    // P=TicketPrice, Q=TotalRevenue, R=TotalCost, S=NetPL,
    // T=BudgetAllocated, U=BudgetSpent, V=BudgetVariance,
    // W=Recurrence, X=NextOccurrence, Y=TasksComplete, Z=TasksTotal,
    // AA=PctComplete, AB=Priority, AC=Notes

    const row = [...ev]; // copy
    // Insert formulas for calculated cols
    // Q=16: Total Revenue = SUMPRODUCT from Guests tab
    row[16] = `=IFERROR(SUMPRODUCT(('Guests & Attendees'!B$6:B$500=B${r})*('Guests & Attendees'!L$6:L$500)),0)`;
    // R=17: Total Cost = Budget Spent
    row[17] = `=IFERROR(U${r},0)`;
    // S=18: Net P/L
    row[18] = `=IFERROR(Q${r}-R${r},0)`;
    // U=20: Budget Spent from Budget sheet
    row[20] = `=IFERROR(SUMPRODUCT(('Budget & Expenses'!B$41:B$1040=A${r})*('Budget & Expenses'!G$41:G$1040)),0)`;
    // V=21: Budget Variance
    row[21] = `=IFERROR(T${r}-U${r},0)`;
    // Y=24: Tasks Complete
    row[24] = `=IFERROR(SUMPRODUCT(('Event Tasks'!B$6:B$500=A${r})*('Event Tasks'!F$6:F$500="Completed")),0)`;
    // Z=25: Tasks Total
    row[25] = `=IFERROR(SUMPRODUCT(('Event Tasks'!B$6:B$500=A${r})*('Event Tasks'!C$6:C$500<>"")),0)`;
    // AA=26: % Complete
    row[26] = `=IFERROR(IF(Z${r}=0,"",Y${r}/Z${r}),"")`;

    dataValues.push(row);
  });

  values.push({ range: "'Master Event Log'!A6", values: dataValues });

  // Alt-row backgrounds (rows 6-35)
  EVENTS.forEach((ev, i) => {
    const ri = i + 5;
    const bg = i % 2 === 0 ? '#FDFCF9' : '#F2F1EE';
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,29), cell:{userEnteredFormat:{
      backgroundColor:hex(bg), textFormat:{foregroundColor:hex('#2F3336'),fontSize:9},
      verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Freeze formula cols styling (blue-gray)
  const formulaCols = [16,17,18,20,21,24,25,26];
  EVENTS.forEach((ev, i) => {
    const ri = i + 5;
    formulaCols.forEach(c => {
      fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,c,c+1), cell:{userEnteredFormat:{
        backgroundColor: hex('#EEF4F8'), textFormat:{foregroundColor:hex('#3D6374'),fontSize:9},
        horizontalAlignment:'RIGHT',
      }}, fields:'userEnteredFormat' }});
    });
  });

  // Data validations
  // C=2: Event Type
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,2,3), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Single-Day','Multi-Day','Virtual','Hybrid','Recurring'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  // F=5: Status
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,5,6), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Planning','Confirmed','In Progress','Completed','Cancelled','Postponed'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  // G=6: Category
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,6,7), rule:{
    condition:{ type:'ONE_OF_LIST', values: CATEGORIES.map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  // W=22: Recurrence
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,22,23), rule:{
    condition:{ type:'ONE_OF_LIST', values:['None','Daily','Weekly','Bi-Weekly','Monthly','Annual'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  // AB=27: Priority
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,27,28), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Critical','High','Medium','Low'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});

  // Number formats
  // P=15: currency (Ticket Price)
  fmt.push({ repeatCell:{ range: gridRange(SID,5,500,15,16), cell:{userEnteredFormat:{
    numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0.00' }
  }}, fields:'userEnteredFormat.numberFormat' }});
  // Q=16,R=17,S=18,T=19,U=20,V=21: currency
  [16,17,18,19,20,21].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,5,500,c,c+1), cell:{userEnteredFormat:{
      numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
    }}, fields:'userEnteredFormat.numberFormat' }});
  });
  // AA=26: percent
  fmt.push({ repeatCell:{ range: gridRange(SID,5,500,26,27), cell:{userEnteredFormat:{
    numberFormat:{ type:'PERCENT', pattern:'0%' }
  }}, fields:'userEnteredFormat.numberFormat' }});

  // Conditional formatting: Status (F=col 5)
  const statusCF = [
    { v:'Confirmed', bg:'#CBDCC5', text:'#2A5C35' },
    { v:'Completed', bg:'#92AE95', text:'#1A3A22' },
    { v:'Planning',  bg:'#D5E4F2', text:'#1E3A5F' },
    { v:'In Progress',bg:'#E9DBA7',text:'#5C4A00' },
    { v:'Cancelled', bg:'#E8C0B8', text:'#5C1A1A' },
    { v:'Postponed', bg:'#DDD5ED', text:'#3A1A5C' },
  ];
  statusCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,5,500,5,6)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // CF: Category color (G=col 6)
  CATEGORIES.forEach(cat=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,5,500,6,7)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:cat}] },
        format:{ backgroundColor:hex(CAT_COLORS[cat]) } }
    }, index:0 }});
  });

  // CF: Net P/L (S=col 18): negative = red, positive = green
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,18,19)],
    booleanRule:{ condition:{ type:'NUMBER_LESS', values:[{userEnteredValue:'0'}] },
      format:{ textFormat:{foregroundColor:hex('#BC7770'),bold:true} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,18,19)],
    booleanRule:{ condition:{ type:'NUMBER_GREATER', values:[{userEnteredValue:'0'}] },
      format:{ textFormat:{foregroundColor:hex('#3A6B3A'),bold:true} } }
  }, index:0 }});

  // CF: Budget Variance (V=col 21): negative = red
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,21,22)],
    booleanRule:{ condition:{ type:'NUMBER_LESS', values:[{userEnteredValue:'0'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{foregroundColor:hex('#5C1A1A'),bold:true} } }
  }, index:0 }});

  // CF: Priority (AB=col 27)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,27,28)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Critical'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{foregroundColor:hex('#5C1A1A'),bold:true} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,27,28)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'High'}] },
      format:{ backgroundColor:hex('#F2D2B6'), textFormat:{foregroundColor:hex('#5C3A00')} } }
  }, index:0 }});

  // Add filters on row 5
  fmt.push({ setBasicFilter:{ filter:{
    range: gridRange(SID,4,500,0,29),
  }}});

  await valuesBatchUpdate(id, values, '04-events-log values');
  await batchUpdate(id, fmt, '04-events-log format');
  console.log('✅ Master Event Log done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
