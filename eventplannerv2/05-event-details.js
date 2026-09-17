'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 3;

// 21 columns A:U
const HEADERS = [
  'Event ID','Session/Agenda Item','Session Type','Date','Start Time','End Time',
  'Duration','Room / Location','Speaker / Facilitator','Max Capacity','Attendance Count',
  'AV Required','Materials Needed','Live Stream?','Recording?',
  'Day Number','Multi-Day Conflict','Session Status','Attendee Feedback Score','Notes','Action Required',
];
// A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12 N=13 O=14 P=15 Q=16 R=17 S=18 T=19 U=20

const COL_WIDTHS = [80,200,120,100,90,90,80,160,160,80,90,80,160,80,80,70,110,110,100,200,110];

// 50 session records
const SESSIONS = [
  ['EVT-001','Registration & Welcome Reception','Networking','04/15/2026','08:00 AM','09:30 AM','','Atrium','Staff',500,'','Yes','Name badges','No','No','1','','Confirmed','','Badge printing confirmed',''],
  ['EVT-001','Opening Keynote: Leading Through Change','Keynote','04/15/2026','09:30 AM','11:00 AM','','Main Ballroom','Dr. Sandra Okafor',500,'','Yes','Slide deck','Yes','Yes','1','','Confirmed','','Mic test 8:00 AM',''],
  ['EVT-001','Breakout A: Strategy & Vision','Workshop','04/15/2026','11:15 AM','12:30 PM','','Room 101','Marcus Cole',80,'','No','Workbooks','No','No','1','','Confirmed','','Print 100 copies',''],
  ['EVT-001','Breakout B: Financial Leadership','Workshop','04/15/2026','11:15 AM','12:30 PM','','Room 102','Elena Park',60,'','No','Case studies','No','No','1','','Confirmed','','PDF distributed',''],
  ['EVT-001','Networking Lunch','Networking','04/15/2026','12:30 PM','01:30 PM','','Dining Hall','Catering Staff',500,'','No','Menu cards','No','No','1','','Confirmed','','Dietary labels needed',''],
  ['EVT-001','Panel: Nonprofit Governance','Panel','04/15/2026','01:30 PM','03:00 PM','','Main Ballroom','Diana Reyes',500,'','Yes','Bios','Yes','Yes','1','','Confirmed','','Panelists confirmed',''],
  ['EVT-001','Afternoon Breakout: HR & Culture','Workshop','04/15/2026','03:15 PM','04:30 PM','','Room 103','James Liu',70,'','No','Slides','No','No','1','','Planning','','Slides due 04/10',''],
  ['EVT-001','Day 1 Wrap-Up & Networking','Networking','04/15/2026','04:30 PM','06:00 PM','','Patio','Diana Reyes',500,'','No','','No','No','1','','Confirmed','','Sponsor tables setup',''],
  ['EVT-001','Morning Yoga Session (Optional)','Recreation','04/16/2026','06:30 AM','07:30 AM','','Pool Deck','Elena Park',40,'','No','Mats','No','No','2','','Confirmed','','Mats arriving 04/14',''],
  ['EVT-001','Keynote Day 2: Innovation & Tech','Keynote','04/16/2026','09:00 AM','10:30 AM','','Main Ballroom','Tech Panel',500,'','Yes','Slide deck','Yes','Yes','2','','Confirmed','','Tech check 8:00 AM',''],
  ['EVT-001','Workshop: Grant Writing Essentials','Workshop','04/16/2026','10:45 AM','12:00 PM','','Room 201','Priya Nair',60,'','No','Templates','No','No','2','','Confirmed','','Workbooks printing',''],
  ['EVT-001','Awards Luncheon','Ceremony','04/16/2026','12:00 PM','02:00 PM','','Ballroom C','Diana Reyes',500,'','Yes','Award plaques','Yes','Yes','2','','Confirmed','','Engraving ordered',''],
  ['EVT-001','Afternoon Keynote: Equity & Access','Keynote','04/16/2026','02:15 PM','03:30 PM','','Main Ballroom','Community Panel',500,'','Yes','Slide deck','Yes','Yes','2','','Planning','','Panelists being confirmed',''],
  ['EVT-001','Closing Reception Day 2','Networking','04/16/2026','05:00 PM','07:00 PM','','Rooftop Terrace','Staff',300,'','No','','No','No','2','','Confirmed','','Catering order placed',''],
  ['EVT-001','Closing Keynote & Call to Action','Keynote','04/17/2026','09:00 AM','10:30 AM','','Main Ballroom','Dr. Sandra Okafor',500,'','Yes','Slide deck','Yes','Yes','3','','Confirmed','','Final session',''],
  ['EVT-001','Workshop: Building Your Action Plan','Workshop','04/17/2026','10:45 AM','12:00 PM','','Breakout Rooms','Marcus Cole',80,'','No','Workbook pg 40','No','No','3','','Confirmed','','Group activity',''],
  ['EVT-001','Farewell Lunch & Networking','Networking','04/17/2026','12:00 PM','01:30 PM','','Dining Hall','Staff',500,'','No','','No','No','3','','Confirmed','','Boxed lunches ordered',''],
  ['EVT-002','VIP Cocktail Reception','Networking','05/02/2026','05:30 PM','07:00 PM','','Foyer','Diana Reyes',80,'','No','Name tags','No','No','1','','Confirmed','','VIP list 40 confirmed',''],
  ['EVT-002','Dinner Service','Reception','05/02/2026','07:00 PM','08:30 PM','','Grand Ballroom','Catering Staff',300,'','No','Menus','No','No','1','','Confirmed','','Seating chart done',''],
  ['EVT-002','Fund-a-Need Auction Segment','Fundraiser','05/02/2026','08:30 PM','09:30 PM','','Grand Ballroom','MC Jason Webb',300,'','Yes','Bid paddles','Yes','No','1','','Planning','','Auction items needed',''],
  ['EVT-002','Live Entertainment: The Jazzmakers','Entertainment','05/02/2026','09:30 PM','11:00 PM','','Grand Ballroom','The Jazzmakers',300,'','Yes','Stage setup','No','No','1','','Confirmed','','Sound check 4 PM',''],
  ['EVT-003','Information Booth Setup','Setup','05/15/2026','07:00 AM','09:00 AM','','Park Entry','Staff',10,'','No','Signage','No','No','1','','Confirmed','','Signage printing',''],
  ['EVT-003','Keynote: Community Health Today','Keynote','05/15/2026','09:30 AM','10:30 AM','','Main Stage','Dr. Mia Torres',1000,'','Yes','None','No','Yes','1','','Confirmed','','PA system rented',''],
  ['EVT-003','Health Screenings (All Day)','Health','05/15/2026','10:00 AM','04:00 PM','','Vendor Area','Health Partners',400,'','No','Supplies','No','No','1','','Planning','','Nurse scheduling needed',''],
  ['EVT-003','Kids Activity Zone','Recreation','05/15/2026','10:00 AM','04:00 PM','','Kids Area','Volunteers',200,'','No','Activity kits','No','No','1','','Planning','','Volunteer training 05/10',''],
  ['EVT-004','Morning Check-In','Registration','05/20/2026','08:30 AM','09:00 AM','','Lobby','Staff',40,'','No','Name tags','No','No','1','','Confirmed','','Badge system ready',''],
  ['EVT-004','Workshop: SEO Fundamentals','Workshop','05/20/2026','09:00 AM','11:00 AM','','Training Room','James Liu',40,'','Yes','Handouts','No','Yes','1','','Confirmed','','Laptop required',''],
  ['EVT-004','Working Lunch & Q&A','Networking','05/20/2026','11:00 AM','12:00 PM','','Breakout Area','James Liu',40,'','No','Lunch','No','No','1','','Confirmed','','Catering ordered',''],
  ['EVT-004','Workshop: Social Media Strategy','Workshop','05/20/2026','12:00 PM','02:00 PM','','Training Room','Elena Park',40,'','Yes','Workbook','No','Yes','1','','Confirmed','','Platform demos',''],
  ['EVT-004','Wrap-Up & Certification','Ceremony','05/20/2026','02:00 PM','03:00 PM','','Training Room','James Liu',40,'','No','Certificates','No','No','1','','Planning','','Certificates printing',''],
  ['EVT-006','Check-In & Team Registration','Registration','06/05/2026','07:00 AM','08:30 AM','','Main Entrance','Staff',200,'','No','Roster','No','No','1','','Confirmed','','Uniform check',''],
  ['EVT-006','Opening Ceremony','Ceremony','06/05/2026','08:30 AM','09:00 AM','','Main Field','Sam Torres',200,'','Yes','Flags','No','No','1','','Confirmed','','Anthem singer booked',''],
  ['EVT-006','Morning Bracket Matches','Sports','06/05/2026','09:00 AM','12:00 PM','','Fields 1-4','Referees',200,'','No','','No','No','1','','Confirmed','','Referee schedule set',''],
  ['EVT-006','Halftime & Food Vendors','Break','06/05/2026','12:00 PM','01:00 PM','','Concessions','Food Vendors',200,'','No','','No','No','1','','Confirmed','','4 vendors confirmed',''],
  ['EVT-006','Afternoon Matches & Semi-Finals','Sports','06/05/2026','01:00 PM','05:00 PM','','Fields 1-4','Referees',200,'','No','','No','No','1','','Confirmed','','',''],
  ['EVT-006','Day 1 Closing & Standings','Ceremony','06/05/2026','05:15 PM','05:45 PM','','Main Field','Sam Torres',200,'','Yes','','No','No','1','','Confirmed','','Scoreboard update',''],
  ['EVT-006','Championship Final','Sports','06/06/2026','10:00 AM','12:00 PM','','Main Field','Referees',200,'','Yes','Trophies','No','Yes','2','','Confirmed','','Trophy engraving',''],
  ['EVT-006','Awards Ceremony','Ceremony','06/06/2026','12:15 PM','01:00 PM','','Main Field','Sam Torres',200,'','Yes','Trophies','No','Yes','2','','Confirmed','','',''],
  ['EVT-007','Platform Login & Testing','Setup','06/10/2026','08:00 AM','09:00 AM','','Virtual','James Liu',500,'','No','Tech guide','No','No','1','','Confirmed','','Zoom backup ready',''],
  ['EVT-007','Opening Welcome & Agenda Review','Presentation','06/10/2026','09:00 AM','09:30 AM','','Virtual','Elena Park',500,'','Yes','Slides','Yes','Yes','1','','Confirmed','','Slide deck finalized',''],
  ['EVT-007','Employer Showcase Presentations','Presentation','06/10/2026','09:30 AM','12:00 PM','','Virtual','HR Teams',500,'','Yes','Company profiles','Yes','Yes','1','','Confirmed','','20 employers registered',''],
  ['EVT-007','Virtual Networking Rooms','Networking','06/10/2026','12:00 PM','01:00 PM','','Virtual Breakout','Facilitators',500,'','Yes','','Yes','No','1','','Confirmed','','Rooms configured',''],
  ['EVT-007','Afternoon Employer Sessions','Presentation','06/10/2026','01:00 PM','03:30 PM','','Virtual','HR Teams',500,'','Yes','','Yes','Yes','1','','Confirmed','','',''],
  ['EVT-007','Resume Review Workshops','Workshop','06/10/2026','03:30 PM','04:30 PM','','Virtual Breakout','Career Coaches',100,'','Yes','Resume template','Yes','No','1','','Planning','','3 coaches confirmed',''],
  ['EVT-007','Closing & Next Steps','Presentation','06/10/2026','04:30 PM','05:00 PM','','Virtual','Elena Park',500,'','Yes','Survey link','Yes','No','1','','Confirmed','','Post-event survey ready',''],
  ['EVT-013','Pre-Registration & Badge Pickup','Registration','07/15/2026','07:00 AM','09:00 AM','','Convention Foyer','Staff',800,'','No','Badges','No','No','1','','Planning','','Badge system ordered',''],
  ['EVT-013','Opening Keynote: Future of Work','Keynote','07/15/2026','09:00 AM','10:30 AM','','Main Hall','Industry Leaders',800,'','Yes','Slides','Yes','Yes','1','','Planning','','Speaker lineup pending',''],
  ['EVT-013','Track A: AI & Automation','Panel','07/15/2026','11:00 AM','12:30 PM','','Hall A',  'Tech Panelists',400,'','Yes','Slides','Yes','Yes','1','','Planning','','Panelist bios needed',''],
  ['EVT-013','Track B: Remote & Hybrid Ops','Workshop','07/15/2026','11:00 AM','12:30 PM','','Hall B','HR Experts',300,'','Yes','Workbook','Yes','Yes','1','','Planning','','Workbook in draft',''],
  ['EVT-013','Tech Expo Networking Lunch','Networking','07/15/2026','12:30 PM','02:00 PM','','Expo Hall','Staff',800,'','No','','Yes','No','1','','Planning','','Exhibitors finalizing',''],
];

(async () => {
  const values = [];
  const fmt = [];

  // Title
  values.push({ range:"'Event Details'!A1", values:[['EVENT DETAILS & AGENDA']] });
  values.push({ range:"'Event Details'!A2", values:[['Track sessions, agenda items, speakers, rooms, and AV needs for every event. Link by Event ID.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,21), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,21), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,21), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,21), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:44}, fields:'pixelSize' }});

  // Info row 3
  values.push({ range:"'Event Details'!A3", values:[['ℹ Duration (G), Day Number (P), and Multi-Day Conflict (Q) are auto-calculated.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,21), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,21), cell:{userEnteredFormat:{
    backgroundColor:hex('#EEF4F8'), textFormat:{italic:true,fontSize:9,foregroundColor:hex('#4A5056')},
    verticalAlignment:'MIDDLE', padding:{left:10},
  }}, fields:'userEnteredFormat' }});

  // Header row 5
  values.push({ range:"'Event Details'!A5", values: [HEADERS] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,21), cell:{userEnteredFormat:{
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

  // Build data rows with formulas
  const dataValues = [];
  SESSIONS.forEach((sess, i) => {
    const r = i + 6;
    const row = [...sess];
    // G=6: Duration formula
    row[6] = `=IFERROR(IF(AND(E${r}<>"",F${r}<>""),TEXT(TIMEVALUE(F${r})-TIMEVALUE(E${r}),"h:mm"),""),"")`;
    // P=15: Day Number formula (days since start date of event)
    row[15] = `=IFERROR(IF(AND(D${r}<>"",A${r}<>""),DAYS(D${r},IFERROR(INDEX('Master Event Log'!D$6:D$500,MATCH(A${r},'Master Event Log'!A$6:A$500,0)),D${r}))+1,""),"")`;
    // Q=16: Multi-Day Conflict — session date differs from event start date
    row[16] = `=IFERROR(IF(AND(A${r}<>"",D${r}<>"",P${r}>1),"Multi-Day",""),"")`;
    dataValues.push(row);
  });

  values.push({ range:"'Event Details'!A6", values: dataValues });

  // Alt rows
  SESSIONS.forEach((_, i) => {
    const ri = i + 5;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,21), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#FDFCF9':'#F2F1EE'),
      textFormat:{foregroundColor:hex('#2F3336'),fontSize:9},
      verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Formula col styling
  [6,15,16].forEach(c=>{
    SESSIONS.forEach((_,i)=>{
      const ri = i+5;
      fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,c,c+1), cell:{userEnteredFormat:{
        backgroundColor:hex('#EEF4F8'), textFormat:{foregroundColor:hex('#3D6374'),fontSize:9},
        horizontalAlignment:'CENTER',
      }}, fields:'userEnteredFormat' }});
    });
  });

  // Validations
  // C=2: Session Type
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,2,3), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Keynote','Workshop','Panel','Networking','Registration',
      'Ceremony','Reception','Breakout','Presentation','Health','Recreation','Entertainment',
      'Setup','Break','Sports','Fundraiser','Other'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:false,
  }}});
  // L=11,N=13,O=14: Yes/No
  [11,13,14].forEach(c=>{
    fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,c,c+1), rule:{
      condition:{ type:'ONE_OF_LIST', values:[{userEnteredValue:'Yes'},{userEnteredValue:'No'}] },
      showCustomUi:true, strict:true,
    }}});
  });
  // R=17: Session Status
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,17,18), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Planning','Confirmed','Completed','Cancelled'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  // S=18: Feedback score number 1-5
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,18,19), rule:{
    condition:{ type:'NUMBER_BETWEEN', values:[{userEnteredValue:'1'},{userEnteredValue:'5'}] },
    showCustomUi:false, strict:false,
  }}});

  // CF: Multi-Day Conflict (Q=col16)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,16,17)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Multi-Day'}] },
      format:{ backgroundColor:hex('#DDD5ED'), textFormat:{bold:true,foregroundColor:hex('#3A1A5C')} } }
  }, index:0 }});

  // CF: Session Status
  const ssCF = [
    { v:'Confirmed', bg:'#CBDCC5', text:'#2A5C35' },
    { v:'Completed', bg:'#92AE95', text:'#1A3A22' },
    { v:'Planning',  bg:'#D5E4F2', text:'#1E3A5F' },
    { v:'Cancelled', bg:'#E8C0B8', text:'#5C1A1A' },
  ];
  ssCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,5,500,17,18)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // CF: AV Required = Yes (L=col11)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,11,12)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Yes'}] },
      format:{ backgroundColor:hex('#F2D2B6'), textFormat:{foregroundColor:hex('#5C3A00'),bold:true} } }
  }, index:0 }});

  // U=20: Action Required CF
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,20,21)],
    booleanRule:{ condition:{ type:'TEXT_NOT_CONTAINS', values:[{userEnteredValue:''}] },
      format:{ backgroundColor:hex('#E9DBA7') } }
  }, index:0 }});

  // Filter on row 5
  fmt.push({ setBasicFilter:{ filter:{ range: gridRange(SID,4,500,0,21) }}});

  await valuesBatchUpdate(id, values, '05-event-details values');
  await batchUpdate(id, fmt, '05-event-details format');
  console.log('✅ Event Details done.');
})().catch(e => { console.error(e.errors||e.message||e); process.exit(1); });
