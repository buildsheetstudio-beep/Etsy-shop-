'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 5;

function cl(i) { if (i<26) return String.fromCharCode(65+i); return 'A'+String.fromCharCode(65+i-26); }

// Venue table headers (12 venues, rows 6:105, 0-indexed 5:104)
const VENUE_HEADERS = [
  'Venue ID','Venue Name','Venue Type','Address','City','State','Zip','Contact Name',
  'Contact Email','Contact Phone','Capacity','Rate Per Day','Deposit Paid','Total Cost',
  'Balance Due','Contract Status','Notes',
];
// A=0 through Q=16

// Vendor table headers (25 vendors, rows 115:414, 0-indexed 114:413)
const VENDOR_HEADERS = [
  'Vendor ID','Vendor Name','Category','Contact Name','Email','Phone','Website',
  'Event ID(s)','Service Description','Quote','Deposit','Balance Due',
  'Contract Status','Payment Status','Insurance on File?','Contract File',
  'Performance Rating','Notes',
];
// A=0 through R=17

const VENUES = [
  ['VEN-001','Austin Convention Center','Conference Center','500 E Cesar Chavez St','Austin','TX','78701','Patricia Lam','p.lam@austincc.com','(512)404-4000',6000,15000,5000,'','','Signed','Main summit venue'],
  ['VEN-002','The Driskill Hotel','Hotel','604 Brazos St','Austin','TX','78701','Carlos Medina','c.medina@driskill.com','(512)439-1234',300,8000,3000,'','','Signed','Gala & year-end'],
  ['VEN-003','Palmer Events Center','Conference Center','900 Barton Springs Rd','Austin','TX','78704','Sarah Wong','s.wong@palmerevents.com','(512)404-4500',5000,12000,4000,'','','Signed','Cultural festival'],
  ['VEN-004','WeWork Downtown Austin','Office','600 Congress Ave #200','Austin','TX','78701','Jessica Park','j.park@wework.com','(512)555-4001',40,800,200,'','','Signed','Workshops'],
  ['VEN-005','Zilker Park','Outdoor','2100 Barton Springs Rd','Austin','TX','78704','City Parks Dept','parks@austintexas.gov','(512)974-6700',5000,500,0,'','','Signed','Health fair'],
  ['VEN-006','Pflugerville Community Center','Community Hall','1460 Town Center Dr','Pflugerville','TX','78660','Mark Ellis','m.ellis@pflugerville.net','(512)990-6000',600,1200,400,'','','Signed','Resource fair'],
  ['VEN-007','Walnut Creek Park','Outdoor','12138 N Lamar Blvd','Austin','TX','78753','City Parks Dept','parks@austintexas.gov','(512)974-6700',500,300,0,'','','Signed','Youth soccer'],
  ['VEN-008','UT McCombs School','School','2110 Speedway','Austin','TX','78712','Events Office','events@mccombs.edu','(512)471-3611',400,0,0,'','','Signed','Biz plan competition'],
  ['VEN-009','Barton Creek Resort & Spa','Hotel','8212 Barton Club Dr','Austin','TX','78735','Events Manager','events@bartoncreek.com','(512)329-4000',500,9000,3000,'','','Sent','Golf tournament'],
  ['VEN-010','Lost Pines Resort','Hotel','575 Hyatt Lost Pines Rd','Bastrop','TX','78602','Corporate Events','events@lostpines.com','(512)308-1234',300,6000,2000,'','','Sent','Board retreat'],
  ['VEN-011','Republic Square','Outdoor','422 W 2nd St','Austin','TX','78701','City Events','events@austintexas.gov','(512)974-6700',3000,2000,500,'','','Signed','Heritage festival'],
  ['VEN-012','W Austin Hotel','Hotel','200 Lavaca St','Austin','TX','78701','Events Team','events@waustin.com','(512)542-3600',400,7500,2500,'','','Sent','New year event'],
];

const VENDORS = [
  ['VDR-001','Summit Catering Co','Catering','Ray Espinoza','ray@summitcatering.com','(512)555-1001','summitcatering.com','EVT-001','Full-service catering Day 1-3',18000,5000,'','Signed','Deposit Paid','Yes','contract_vdr001.pdf','','Menus finalized'],
  ['VDR-002','Gala Fine Dining','Catering','Isabelle Rocha','i.rocha@galafd.com','(512)555-1002','galafd.com','EVT-002','Plated dinner + VIP cocktails',22000,7000,'','Signed','Deposit Paid','Yes','contract_vdr002.pdf','','3-course menu confirmed'],
  ['VDR-003','ProAV Solutions','AV/Tech','Marcus Webb','m.webb@proav.com','(512)555-1003','proavsolutions.com','EVT-001,EVT-013','Full AV: lighting, sound, screens',35000,10000,'','Signed','Deposit Paid','Yes','contract_vdr003.pdf','','Live stream feed set'],
  ['VDR-004','Park Sounds','AV/Tech','Linda Yu','l.yu@parksounds.com','(512)555-1004','parksounds.com','EVT-003,EVT-011','Outdoor PA and generator',4500,1500,'','Signed','Deposit Paid','Yes','contract_vdr004.pdf','','Generator permit filed'],
  ['VDR-005','Clarity Digital','AV/Tech','Chris Fong','c.fong@claritydigital.com','(512)555-1005','claritydigital.com','EVT-007','Virtual platform + tech support',3000,1000,'','Signed','Paid in Full','Yes','contract_vdr005.pdf','','Zoom Pro licenses active'],
  ['VDR-006','Bloom Event Décor','Decoration','Maria Santos','m.santos@bloomeventdecor.com','(512)555-1006','bloomeventdecor.com','EVT-002,EVT-015','Florals, centerpieces, draping',8500,2500,'','Signed','Deposit Paid','Yes','contract_vdr006.pdf','','Color scheme: gold & blush'],
  ['VDR-007','Pixel Perfect Photography','Photography','Dana Lee','d.lee@pixelperfectphoto.com','(512)555-1007','pixelperfect.photo','EVT-001,EVT-002','Event photography both days',6000,2000,'','Signed','Deposit Paid','Yes','contract_vdr007.pdf','5','Photos in 72hrs'],
  ['VDR-008','Lens & Motion Video','Videography','Sam Park','s.park@lensmotion.com','(512)555-1008','lensmotion.com','EVT-001,EVT-013','Full event video + highlight reel',8000,2500,'','Signed','Deposit Paid','Yes','contract_vdr008.pdf','','Drone waiver filed'],
  ['VDR-009','TexPrint Solutions','Printing','Robin Hall','r.hall@texprinttx.com','(512)555-1009','texprintolutions.com','EVT-001,EVT-004','Programs, banners, name badges',3200,1000,'','Signed','Deposit Paid','Yes','contract_vdr009.pdf','','Final files due 04/10'],
  ['VDR-010','The Jazzmakers','Entertainment','Jason Webb','j.webb@jazzmakers.com','(512)555-1010','jazzmakers.com','EVT-002','Live jazz trio 9:30-11:00 PM',3500,1000,'','Signed','Deposit Paid','Yes','contract_vdr010.pdf','','Sound check 4:00 PM'],
  ['VDR-011','Austin Event Security','Security','Tom Riley','t.riley@austinsecurity.com','(512)555-1011','austinsecurity.com','EVT-001,EVT-002','4 officers each event',2400,800,'','Signed','Deposit Paid','Yes','contract_vdr011.pdf','','Uniforms provided'],
  ['VDR-012','City Shuttle Co','Transportation','Paul Gomez','p.gomez@cityshuttle.com','(512)555-1012','cityshuttle.com','EVT-001','Hotel to venue shuttle x3 days',4500,1500,'','Signed','Deposit Paid','Yes','contract_vdr012.pdf','','Schedule confirmed'],
  ['VDR-013','Bold Brand Marketing','Marketing','Stacy Liu','s.liu@boldbrand.com','(512)555-1013','boldbrand.com','EVT-013','Digital campaign + social ads',12000,4000,'','Signed','Deposit Paid','Yes','contract_vdr013.pdf','','Campaign live 06/01'],
  ['VDR-014','Field Events Staffing','Staffing','Kevin Moore','k.moore@fieldevents.com','(512)555-1014','fieldevents.com','EVT-001,EVT-019','20 event staff per day',9000,3000,'','Signed','Deposit Paid','Yes','contract_vdr014.pdf','','Briefing 04/14'],
  ['VDR-015','Quick Park Valet','Transportation','Leo Castillo','l.castillo@quickpark.com','(512)555-1015','quickpark.com','EVT-002','Valet service 80 cars',1800,500,'','Signed','Deposit Paid','Yes','contract_vdr015.pdf','','Setup at 5 PM'],
  ['VDR-016','City Health Partners','Health','Dr. Ana Cruz','a.cruz@healthpartners.org','(512)555-1016','healthpartners.org','EVT-003','10 nurse stations, blood pressure/glucose',0,0,'','Signed','Paid in Full','Yes','contract_vdr016.pdf','','Volunteer nurses'],
  ['VDR-017','GreenLawn Supply','Decoration','Bob Ward','b.ward@greenlawn.com','(512)555-1017','greenlawn.com','EVT-003,EVT-011','Outdoor tables, tents, chairs',6000,2000,'','Signed','Deposit Paid','Yes','contract_vdr017.pdf','','Delivery 7 AM day-of'],
  ['VDR-018','Trophy House Austin','Printing','Cindy Ho','c.ho@trophyhouse.com','(512)555-1018','trophyhouse.com','EVT-006,EVT-022','Trophies, medals, plaques',2200,700,'','Signed','Deposit Paid','Yes','contract_vdr018.pdf','','Engraving due 05/28'],
  ['VDR-019','Zoom Corporate Events','AV/Tech','Nina Patel','n.patel@zoomevents.com','(512)555-1019','zoomevents.com','EVT-007','Zoom Events license + breakout rooms',1500,500,'','Signed','Paid in Full','Yes','contract_vdr019.pdf','','License valid 12 months'],
  ['VDR-020','Fresh Catering & Co','Catering','Andre Baptiste','a.baptiste@freshcatering.com','(512)555-1020','freshcatering.com','EVT-004,EVT-005','Box lunches & meeting refreshments',2400,0,'','Signed','Not Paid','Yes','contract_vdr020.pdf','','Invoice pending'],
  ['VDR-021','Heritage Event Design','Decoration','Rosa Alvarado','r.alvarado@heritageevent.com','(512)555-1021','heritageevent.com','EVT-019','Cultural booths, flags, staging',11000,3500,'','Sent','Not Paid','No','','','Awaiting insurance COI'],
  ['VDR-022','Austin Sounds DJ','Entertainment','DJ Marcus X','djmarcus@austinsounds.com','(512)555-1022','austinsounds.com','EVT-026','DJ set 8 PM–midnight',2500,800,'','Sent','Not Paid','No','','','Contract out for signature'],
  ['VDR-023','Lone Star Insurance','Insurance','George Pitts','g.pitts@lonestarins.com','(512)555-1023','lonestarins.com','EVT-001,EVT-002,EVT-013','Event liability insurance x3',4500,0,'','Signed','Paid in Full','Yes','contract_vdr023.pdf','5','Certificates on file'],
  ['VDR-024','Sunset Rentals','Other','Tammy Ford','t.ford@sunsetrentals.com','(512)555-1024','sunsetrentals.com','EVT-019,EVT-003','Staging, power, flooring',8500,2500,'','Signed','Deposit Paid','Yes','contract_vdr024.pdf','','Staging design approved'],
  ['VDR-025','Print & Pack Co','Printing','Susan Brown','s.brown@printpack.com','(512)555-1025','printpack.com','EVT-002,EVT-012','Programs, pamphlets, workbooks',1900,600,'','Signed','Deposit Paid','Yes','contract_vdr025.pdf','','Files submitted'],
];

(async () => {
  const values = [];
  const fmt = [];

  // ── Title ─────────────────────────────────────────────────────────
  values.push({ range:"'Vendors & Venues'!A1", values:[['VENDORS & VENUES']] });
  values.push({ range:"'Vendors & Venues'!A2", values:[['Manage all venues (top section) and vendors/suppliers (bottom section). Formulas auto-calculate balances.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,18), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,18), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,18), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,18), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:44}, fields:'pixelSize' }});

  // ── VENUES section header ─────────────────────────────────────────
  // Row 4 (0-indexed 3)
  values.push({ range:"'Vendors & Venues'!A4", values:[['VENUE DIRECTORY']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,3,4,0,18), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,3,4,0,18), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:11,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:3,endIndex:4},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  // Venue headers row 5 (0-indexed 4)
  values.push({ range:"'Vendors & Venues'!A5", values:[VENUE_HEADERS] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,17), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  // Venue data rows 6-17 (0-indexed 5-16)
  const venueData = VENUES.map((v,i)=>{
    const r = i+6;
    const row = [...v];
    // N=13 Total Cost = deposit + balance (user fills rate per day * # days separately; auto calc from rate)
    row[13] = `=IFERROR(L${r}+M${r},0)`;  // but actually Total = Rate * Days (user enters L=rate, M=deposit)
    // O=14 Balance Due = Total - Deposit
    // Wait, the columns are: A=VenueID,B=Name,C=Type,D=Address,E=City,F=State,G=Zip,H=Contact,I=Email,J=Phone,K=Capacity,L=RatePerDay,M=DepositPaid,N=TotalCost,O=BalanceDue,P=ContractStatus,Q=Notes
    // Let's keep TotalCost as user-entered (column N, index 13) and compute Balance Due
    row[13] = v[13]; // total cost (will be user-entered sample from data)
    // We need to recalculate from sample data: rate is index 11 (L), deposit is index 12 (M), total is in 13
    // Set total = rate (assuming 1 day, or just leave as formula)
    // Actually let's set total = user-entered from VENUES array
    row[14] = `=IFERROR(N${r}-M${r},0)`; // Balance = TotalCost - DepositPaid
    return row;
  });
  // Fill in total cost for venues
  VENUES.forEach((v,i)=>{
    // total cost = 2x rate for simplicity in sample data
    const totalMap = [30000,16000,24000,1600,500,2400,600,0,18000,12000,4000,15000];
    venueData[i][13] = totalMap[i] || (v[11]*2);
  });
  values.push({ range:"'Vendors & Venues'!A6", values: venueData });

  // Venue alt rows styling
  VENUES.forEach((_,i)=>{
    const ri = i+5;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,17), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#EEF4F8':'#F7F6F3'),
      textFormat:{foregroundColor:hex('#2F3336'),fontSize:9}, verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });
  // Formula col styling for venues: N=13, O=14
  VENUES.forEach((_,i)=>{
    const ri = i+5;
    [13,14].forEach(c=>{
      fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,c,c+1), cell:{userEnteredFormat:{
        backgroundColor:hex('#EEF4F8'), textFormat:{foregroundColor:hex('#3D6374'),fontSize:9},
        horizontalAlignment:'RIGHT',
      }}, fields:'userEnteredFormat' }});
    });
  });

  // Currency formats for venues
  [11,12,13,14].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,5,17,c,c+1), cell:{userEnteredFormat:{
      numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
    }}, fields:'userEnteredFormat.numberFormat' }});
  });

  // Venue contract status CF
  const contractCF = [
    { v:'Signed',      bg:'#CBDCC5', text:'#2A5C35' },
    { v:'Sent',        bg:'#E9DBA7', text:'#5C4A00' },
    { v:'Not Started', bg:'#F2F1EE', text:'#4A5056' },
    { v:'Cancelled',   bg:'#E8C0B8', text:'#5C1A1A' },
  ];
  contractCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,5,17,15,16)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // Balance Due > 0 → amber for venues
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,17,14,15)],
    booleanRule:{ condition:{ type:'NUMBER_GREATER', values:[{userEnteredValue:'0'}] },
      format:{ backgroundColor:hex('#F2D2B6'), textFormat:{bold:true} } }
  }, index:0 }});

  // ── Vendor section ────────────────────────────────────────────────
  // Gap rows 17-113 (users can add more venues)
  // Vendor header at row 113 (0-indexed 112)
  values.push({ range:"'Vendors & Venues'!A113", values:[['VENDOR DIRECTORY']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,112,113,0,18), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,112,113,0,18), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:11,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:112,endIndex:113},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  // Vendor column headers row 114 (0-indexed 113)
  values.push({ range:"'Vendors & Venues'!A114", values:[VENDOR_HEADERS] });
  fmt.push({ repeatCell:{ range: gridRange(SID,113,114,0,18), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:113,endIndex:114},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  // Vendor data rows 115-139 (0-indexed 114-138)
  const vendorData = VENDORS.map((v,i)=>{
    const r = i+115;
    const row = [...v];
    // L=11: Balance Due = Quote - Deposit
    // Col indices: A=0,B=1,C=2,D=3,E=4,F=5,G=6,H=7,I=8,J=9,K=10,L=11,M=12,N=13,O=14,P=15,Q=16,R=17
    row[11] = `=IFERROR(J${r}-K${r},0)`;
    return row;
  });
  values.push({ range:"'Vendors & Venues'!A115", values: vendorData });

  // Vendor alt rows
  VENDORS.forEach((_,i)=>{
    const ri = i+114;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,18), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#FDFCF9':'#F2F1EE'),
      textFormat:{foregroundColor:hex('#2F3336'),fontSize:9}, verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });
  // Balance Due formula col (L=11)
  VENDORS.forEach((_,i)=>{
    const ri = i+114;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,11,12), cell:{userEnteredFormat:{
      backgroundColor:hex('#EEF4F8'), textFormat:{foregroundColor:hex('#3D6374'),fontSize:9},
      horizontalAlignment:'RIGHT',
    }}, fields:'userEnteredFormat' }});
  });

  // Currency formats for vendors: J=9,K=10,L=11
  [9,10,11].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,114,140,c,c+1), cell:{userEnteredFormat:{
      numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
    }}, fields:'userEnteredFormat.numberFormat' }});
  });

  // Vendor contract status CF (M=12)
  contractCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,114,420,12,13)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // Vendor payment status CF (N=13)
  const paymentCF = [
    { v:'Paid in Full', bg:'#CBDCC5', text:'#2A5C35' },
    { v:'Deposit Paid', bg:'#D5E4F2', text:'#1E3A5F' },
    { v:'Partial',      bg:'#E9DBA7', text:'#5C4A00' },
    { v:'Not Paid',     bg:'#E8C0B8', text:'#5C1A1A' },
    { v:'Refunded',     bg:'#DDD5ED', text:'#3A1A5C' },
  ];
  paymentCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,114,420,13,14)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // Vendor balance due > 0 → amber (L=11)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,114,420,11,12)],
    booleanRule:{ condition:{ type:'NUMBER_GREATER', values:[{userEnteredValue:'0'}] },
      format:{ backgroundColor:hex('#F2D2B6'), textFormat:{bold:true} } }
  }, index:0 }});

  // Vendor insurance not on file → red (O=14)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,114,420,14,15)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'No'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{foregroundColor:hex('#5C1A1A'),bold:true} } }
  }, index:0 }});

  // Vendor rating (Q=16) CF: 5=green, <3=red
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,114,420,16,17)],
    booleanRule:{ condition:{ type:'NUMBER_EQ', values:[{userEnteredValue:'5'}] },
      format:{ textFormat:{foregroundColor:hex('#2A5C35'),bold:true} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,114,420,16,17)],
    booleanRule:{ condition:{ type:'NUMBER_LESS_THAN_EQ', values:[{userEnteredValue:'2'}] },
      format:{ textFormat:{foregroundColor:hex('#5C1A1A'),bold:true} } }
  }, index:0 }});

  // Vendor category dropdown (C=2)
  fmt.push({ setDataValidation:{ range: gridRange(SID,114,420,2,3), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Venue','Catering','AV/Tech','Photography','Videography',
      'Decoration','Entertainment','Security','Transportation','Marketing','Printing','Staffing',
      'Health','Insurance','Other'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:false,
  }}});
  // Contract status (M=12) and venue (P=15)
  [12].forEach(c=>{
    fmt.push({ setDataValidation:{ range: gridRange(SID,114,420,c,c+1), rule:{
      condition:{ type:'ONE_OF_LIST', values:['Not Started','Sent','Signed','Expired','Cancelled'].map(v=>({userEnteredValue:v})) },
      showCustomUi:true, strict:true,
    }}});
  });
  // Payment status (N=13)
  fmt.push({ setDataValidation:{ range: gridRange(SID,114,420,13,14), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Not Paid','Deposit Paid','Partial','Paid in Full','Refunded'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  // Insurance (O=14) yes/no
  fmt.push({ setDataValidation:{ range: gridRange(SID,114,420,14,15), rule:{
    condition:{ type:'ONE_OF_LIST', values:[{userEnteredValue:'Yes'},{userEnteredValue:'No'}] },
    showCustomUi:true, strict:true,
  }}});

  // Venue contract status validation
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,17,15,16), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Not Started','Sent','Signed','Expired','Cancelled'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  // Venue type validation
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,17,2,3), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Hotel','Conference Center','Community Hall','Outdoor','Virtual',
      'Restaurant','Office','School','Church','Other'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:false,
  }}});

  // Column widths
  const venueWidths = [70,200,130,200,100,60,70,130,200,110,80,90,90,90,90,110,180];
  venueWidths.forEach((w,i)=>{
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1},
      properties:{pixelSize:w}, fields:'pixelSize' }});
  });

  // Freeze row 5 (venue headers)
  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:5}},
    fields:'gridProperties.frozenRowCount' }});

  // Filters on venue and vendor sections
  fmt.push({ setBasicFilter:{ filter:{ range: gridRange(SID,4,17,0,17) }}});

  await valuesBatchUpdate(id, values, '07-vendors values');
  await batchUpdate(id, fmt, '07-vendors format');
  console.log('✅ Vendors & Venues done.');
})().catch(e => { console.error(e.errors||e.message||e); process.exit(1); });
