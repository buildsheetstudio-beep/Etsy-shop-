'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 4;

// 22 columns A:V
const HEADERS = [
  'Event ID','Full Name','Email','Phone','Organization','Ticket Type',
  'Registration Date','Attendance Status','Check-In Time','Meal Choice',
  'Dietary Restrictions','Accessibility Needs','Guest of','Table/Room Assignment',
  'Payment Status','Amount Paid','VIP?','Speaking?','Volunteer?',
  'Feedback Score','Notes','Follow-Up Needed',
];

const COL_WIDTHS = [80,180,200,120,160,110,100,110,90,110,140,140,140,120,110,90,60,70,70,90,180,110];

function r(ri) { return ri; } // just an alias for clarity

// 100 attendees across events
const ATTENDEES = [
  // EVT-001 (30)
  ['EVT-001','Sandra Okafor','sandra.okafor@nonprofitlead.org','(512)555-0101','Nonprofit Leaders USA','Speaker','03/01/2026','Confirmed','','Standard','None','None','','Table 1','Paid in Full',0,'No','Yes','No','','Keynote speaker','No'],
  ['EVT-001','Marcus Cole','marcus.cole@sunrisecf.org','(512)448-7701','Sunrise Community Foundation','Staff','03/01/2026','Confirmed','','Standard','None','None','','Staff Table','Paid in Full',0,'No','Yes','Yes','','Co-organizer','No'],
  ['EVT-001','Elena Park','elena.park@sunrisecf.org','(512)448-7702','Sunrise Community Foundation','Staff','03/01/2026','Confirmed','','Vegetarian','None','None','','Staff Table','Paid in Full',0,'No','No','Yes','','','No'],
  ['EVT-001','Priya Nair','priya.nair@austinunited.org','(512)555-0110','Austin United Way','VIP','03/05/2026','Confirmed','','Standard','None','Wheelchair','','Table 1','Paid in Full',150,'Yes','No','No','','Wheelchair access confirmed','No'],
  ['EVT-001','James Liu','james.liu@bizforward.com','(512)555-0120','BizForward Austin','General Admission','03/10/2026','Confirmed','','Standard','Gluten-Free','None','','Table 5','Paid in Full',150,'No','No','No','','','No'],
  ['EVT-001','Natalie Brooks','n.brooks@techbridge.io','(512)555-0130','TechBridge','VIP','03/10/2026','Confirmed','','Standard','None','None','','Table 2','Paid in Full',150,'Yes','No','No','','','No'],
  ['EVT-001','David Kim','d.kim@communityreach.org','(512)555-0140','Community Reach','General Admission','03/12/2026','Confirmed','','Standard','None','None','','Table 6','Paid in Full',150,'No','No','No','4','','No'],
  ['EVT-001','Aisha Mohamed','a.mohamed@cityofsf.org','(512)555-0150','City of Austin','VIP','03/12/2026','Confirmed','','Halal','None','None','','Table 2','Paid in Full',150,'Yes','No','No','5','','No'],
  ['EVT-001','Robert Torres','r.torres@healthyaustin.org','(512)555-0160','Healthy Austin','General Admission','03/15/2026','Confirmed','','Vegetarian','None','None','','Table 7','Paid in Full',150,'No','No','No','','','No'],
  ['EVT-001','Carmen Vega','c.vega@socialfoundation.net','(512)555-0170','Social Foundation','General Admission','03/15/2026','Confirmed','','Standard','None','None','','Table 7','Paid in Full',150,'No','No','No','3','','No'],
  ['EVT-001','Thomas Wright','t.wright@investaustin.org','(512)555-0180','Invest Austin','VIP','03/16/2026','Confirmed','','Standard','Nut Allergy','None','','Table 3','Paid in Full',150,'Yes','No','No','','Allergy noted','No'],
  ['EVT-001','Samantha Green','s.green@greenlead.org','(512)555-0190','Green Leadership','General Admission','03/17/2026','Confirmed','','Vegan','None','None','','Table 8','Paid in Full',150,'No','No','No','','','No'],
  ['EVT-001','Leon Baptiste','l.baptiste@faithcenter.org','(512)555-0191','Faith Center','Speaker','03/18/2026','Confirmed','','Standard','None','None','','Table 1','Paid in Full',0,'No','Yes','No','','Panel speaker','No'],
  ['EVT-001','Grace Huang','g.huang@stem4all.org','(512)555-0192','STEM for All','General Admission','03/19/2026','Confirmed','','Standard','None','None','','Table 9','Paid in Full',150,'No','No','No','','','No'],
  ['EVT-001','Marcus Diaz','m.diaz@mentornet.org','(512)555-0193','MentorNet','General Admission','03/20/2026','Confirmed','','Standard','None','None','','Table 9','Paid in Full',150,'No','No','No','','','No'],
  ['EVT-001','Yolanda Ferris','y.ferris@equityforward.org','(512)555-0194','Equity Forward','VIP','03/21/2026','Confirmed','','Standard','None','Hearing Loop','','Table 3','Paid in Full',150,'Yes','No','No','','Hearing loop needed','No'],
  ['EVT-001','Chris Reeves','c.reeves@cityparks.gov','(512)555-0195','City Parks Dept','General Admission','03/22/2026','Confirmed','','Standard','None','None','','Table 10','Paid in Full',150,'No','No','No','','','No'],
  ['EVT-001','Monica Alvarez','m.alvarez@workrights.org','(512)555-0196','Work Rights','General Admission','03/23/2026','Confirmed','','Gluten-Free','None','None','','Table 10','Paid in Full',150,'No','No','No','','','No'],
  ['EVT-001','Tony Chen','t.chen@digitalequity.org','(512)555-0197','Digital Equity','General Admission','03/24/2026','Confirmed','','Standard','None','None','','Table 11','Paid in Full',150,'No','No','No','','','No'],
  ['EVT-001','Rachel Osei','r.osei@youthempowers.org','(512)555-0198','Youth Empowers','General Admission','03/25/2026','Confirmed','','Vegetarian','None','None','','Table 11','Paid in Full',150,'No','No','No','','','No'],
  // EVT-002 (20)
  ['EVT-002','Diana Reyes','diana.reyes@sunrisecf.org','(512)448-7700','Sunrise Community Foundation','Staff','01/10/2026','Confirmed','','Standard','None','None','','Head Table','Paid in Full',0,'No','No','Yes','','Organizer','No'],
  ['EVT-002','Mayor Linda Park','mayor@cityofaustin.gov','(512)555-9001','City of Austin','VIP','01/15/2026','Confirmed','','Standard','None','None','','Head Table','Paid in Full',0,'Yes','No','No','5','Mayor attending','No'],
  ['EVT-002','Warren Hughes','w.hughes@foundationhub.org','(512)555-0201','Foundation Hub','VIP','01/16/2026','Confirmed','','Standard','None','None','','Table 2','Paid in Full',250,'Yes','No','No','','$10K donation pledge','No'],
  ['EVT-002','Cynthia Novak','c.novak@novakenterprises.com','(512)555-0202','Novak Enterprises','General Admission','01/17/2026','Confirmed','','Standard','None','None','','Table 3','Paid in Full',250,'No','No','No','','Auction bidder','No'],
  ['EVT-002','Peter Schultz','p.schultz@bankersfund.org','(512)555-0203','Bankers Fund','VIP','01/18/2026','Confirmed','','Standard','None','None','','Table 2','Paid in Full',250,'Yes','No','No','4','','No'],
  ['EVT-002','Angela Wright','a.wright@artsfund.org','(512)555-0204','Arts Fund','General Admission','01/19/2026','Confirmed','','Vegetarian','None','None','','Table 4','Paid in Full',250,'No','No','No','','','No'],
  ['EVT-002','Samuel Porter','s.porter@portergroup.com','(512)555-0205','Porter Group','General Admission','01/20/2026','Confirmed','','Standard','Kosher','None','','Table 5','Paid in Full',250,'No','No','No','','','No'],
  ['EVT-002','Grace Fontaine','g.fontaine@cityedge.org','(512)555-0206','City Edge','VIP','01/21/2026','Confirmed','','Standard','None','None','','Table 2','Paid in Full',250,'Yes','No','No','','','No'],
  ['EVT-002','Mike Chang','m.chang@changrealty.com','(512)555-0207','Chang Realty','General Admission','01/22/2026','Confirmed','','Standard','None','None','','Table 6','Paid in Full',250,'No','No','No','','Auction sponsor','No'],
  ['EVT-002','Lena Walsh','l.walsh@walshlegal.com','(512)555-0208','Walsh Legal','General Admission','01/23/2026','Confirmed','','Standard','None','None','','Table 6','Paid in Full',250,'No','No','No','','','No'],
  ['EVT-002','Oscar Fuentes','o.fuentes@fuendesign.com','(512)555-0209','FuenDesign','General Admission','01/24/2026','Confirmed','','Standard','None','None','','Table 7','Paid in Full',250,'No','No','No','','','No'],
  ['EVT-002','Ruth Kim','r.kim@globalgiving.org','(512)555-0210','Global Giving','VIP','01/25/2026','Confirmed','','Standard','None','None','','Table 3','Paid in Full',250,'Yes','No','No','5','Top donor relationship','No'],
  ['EVT-002','Patrick Okafor','p.okafor@lawclinic.org','(512)555-0211','Law Clinic','General Admission','01/26/2026','Confirmed','','Standard','None','Wheelchair','','Table 7','Paid in Full',250,'No','No','No','','Accessible seating needed','No'],
  ['EVT-002','Jennifer Adams','j.adams@adamsfoundation.org','(512)555-0212','Adams Foundation','General Admission','01/27/2026','Confirmed','','Vegan','None','None','','Table 8','Paid in Full',250,'No','No','No','','','No'],
  ['EVT-002','Brian Sato','b.sato@satoconsulting.com','(512)555-0213','Sato Consulting','General Admission','01/28/2026','Confirmed','','Standard','None','None','','Table 8','Paid in Full',250,'No','No','No','','','No'],
  ['EVT-002','Veronica Miles','v.miles@milescommunity.org','(512)555-0214','Miles Community','General Admission','02/01/2026','Confirmed','','Standard','None','None','','Table 9','Paid in Full',250,'No','No','No','','','No'],
  ['EVT-002','Henry Osei','h.osei@techgiveaustin.org','(512)555-0215','Tech Give Austin','General Admission','02/02/2026','Confirmed','','Standard','None','None','','Table 9','Paid in Full',250,'No','No','No','','','No'],
  ['EVT-002','Diane Flores','d.flores@greenfuture.org','(512)555-0216','Green Future','General Admission','02/03/2026','Confirmed','','Gluten-Free','None','None','','Table 10','Paid in Full',250,'No','No','No','','','No'],
  ['EVT-002','Nathan Burke','n.burke@burkephilanthropy.org','(512)555-0217','Burke Philanthropy','VIP','02/04/2026','Confirmed','','Standard','None','None','','Table 3','Paid in Full',250,'Yes','No','No','4','Major donor','No'],
  ['EVT-002','Sarah Ling','s.ling@lingfamily.org','(512)555-0218','Ling Family Fund','General Admission','02/05/2026','Confirmed','','Standard','None','None','','Table 10','Paid in Full',250,'No','No','No','','','No'],
  // EVT-003 (10)
  ['EVT-003','Sam Torres','s.torres@sunrisecf.org','(512)448-7703','Sunrise Community Foundation','Staff','03/01/2026','Confirmed','','Standard','None','None','','Staff','Paid in Full',0,'No','No','Yes','','Lead organizer','No'],
  ['EVT-003','Maria Lopez','m.lopez@communityhealth.org','(512)555-0301','Community Health','General Admission','04/01/2026','Confirmed','','Standard','None','None','','','Paid in Full',0,'No','No','No','','','No'],
  ['EVT-003','James Park','j.park@healthpromote.org','(512)555-0302','Health Promote','General Admission','04/02/2026','Registered','','Standard','None','None','','','Not Paid',0,'No','No','No','','','No'],
  ['EVT-003','Lucy Tran','l.tran@healthypflugerville.gov','(512)555-0303','City of Pflugerville','General Admission','04/05/2026','Registered','','Standard','None','Wheelchair','','','Not Paid',0,'No','No','No','','Accessible path required','No'],
  ['EVT-003','Derek Johnson','d.johnson@austinfitness.com','(512)555-0304','Austin Fitness','General Admission','04/06/2026','Registered','','Standard','None','None','','','Not Paid',0,'No','No','No','','','No'],
  ['EVT-003','Sharon Kwon','s.kwon@mentalhealth.org','(512)555-0305','Mental Health Alliance','Speaker','04/07/2026','Confirmed','','Standard','None','None','','','Paid in Full',0,'No','Yes','No','','Breakout speaker','No'],
  ['EVT-003','Frank Delgado','f.delgado@diabetesfund.org','(512)555-0306','Diabetes Fund','General Admission','04/08/2026','Registered','','Standard','None','None','','','Not Paid',0,'No','No','No','','','No'],
  ['EVT-003','Nina Obi','n.obi@ngo-nutrition.org','(512)555-0307','NGO Nutrition','General Admission','04/09/2026','Registered','','Vegetarian','None','None','','','Not Paid',0,'No','No','No','','','No'],
  ['EVT-003','Carl Washington','c.washington@seniorcentral.org','(512)555-0308','Senior Central','General Admission','04/10/2026','Registered','','Standard','None','Hearing Loop','','','Not Paid',0,'No','No','No','','','No'],
  ['EVT-003','Lisa Cheng','l.cheng@healthed.org','(512)555-0309','Health Ed','General Admission','04/11/2026','Registered','','Standard','None','None','','','Not Paid',0,'No','No','No','','','No'],
  // EVT-004 (10)
  ['EVT-004','Tara Singh','t.singh@mktgpros.com','(512)555-0401','MktgPros','General Admission','04/01/2026','Confirmed','','Standard','Vegan','None','','','Paid in Full',75,'No','No','No','','','No'],
  ['EVT-004','Justin Rowe','j.rowe@smallbizaustin.com','(512)555-0402','Small Biz Austin','General Admission','04/02/2026','Confirmed','','Standard','None','None','','','Paid in Full',75,'No','No','No','','','No'],
  ['EVT-004','Mei Lin','m.lin@linenterprises.biz','(512)555-0403','Lin Enterprises','General Admission','04/03/2026','Confirmed','','Standard','Gluten-Free','None','','','Paid in Full',75,'No','No','No','','','No'],
  ['EVT-004','Steve Owens','s.owens@owensco.com','(512)555-0404','OwensCo','General Admission','04/04/2026','Confirmed','','Standard','None','None','','','Paid in Full',75,'No','No','No','3','','No'],
  ['EVT-004','Bianca Cruz','b.cruz@cruzdigital.com','(512)555-0405','Cruz Digital','General Admission','04/05/2026','Confirmed','','Standard','None','None','','','Paid in Full',75,'No','No','No','','','No'],
  ['EVT-004','Noah Fischer','n.fischer@fischerdesign.com','(512)555-0406','Fischer Design','General Admission','04/06/2026','Confirmed','','Vegetarian','None','None','','','Paid in Full',75,'No','No','No','','','No'],
  ['EVT-004','Aliyah Brown','a.brown@brownpr.com','(512)555-0407','Brown PR','General Admission','04/07/2026','Confirmed','','Standard','None','None','','','Paid in Full',75,'No','No','No','','','No'],
  ['EVT-004','Kevin Ma','k.ma@maphotography.com','(512)555-0408','MA Photography','General Admission','04/08/2026','Registered','','Standard','None','None','','','Deposit Paid',37,'No','No','No','','','No'],
  ['EVT-004','Sandra Hill','s.hill@hillagency.com','(512)555-0409','Hill Agency','General Admission','04/09/2026','Registered','','Standard','None','None','','','Not Paid',0,'No','No','No','','','No'],
  ['EVT-004','Tom Brady','t.brady2@austinmedia.com','(512)555-0410','Austin Media','General Admission','04/10/2026','Waitlist','','Standard','None','None','','','Not Paid',0,'No','No','No','','Waitlisted','Yes'],
  // EVT-007 (10)
  ['EVT-007','Paula Stevens','p.stevens@remotework.io','(512)555-0701','Remote Work Pro','Online','04/15/2026','Confirmed','','No Meal','None','None','','','Paid in Full',0,'No','No','No','','Virtual attendee','No'],
  ['EVT-007','Ali Hassan','a.hassan@techcorp.com','(512)555-0702','TechCorp','Online','04/15/2026','Confirmed','','No Meal','None','None','','','Paid in Full',0,'No','No','No','','','No'],
  ['EVT-007','Megan Cross','m.cross@crosscareers.com','(512)555-0703','Cross Careers','Online','04/15/2026','Confirmed','','No Meal','None','None','','','Paid in Full',0,'No','No','No','','','No'],
  ['EVT-007','Doug Peterson','d.peterson@petersoncg.com','(512)555-0704','PetersonCG','Online','04/16/2026','Registered','','No Meal','None','None','','','Not Paid',0,'No','No','No','','','No'],
  ['EVT-007','Angie Watts','a.watts@wattsjobs.org','(512)555-0705','Watts Jobs','Online','04/16/2026','Confirmed','','No Meal','None','None','','','Paid in Full',0,'No','No','No','','','No'],
  ['EVT-007','Rafael Mendez','r.mendez@codeacademy.io','(512)555-0706','Code Academy','Online','04/17/2026','Confirmed','','No Meal','None','None','','','Paid in Full',0,'No','No','No','5','','No'],
  ['EVT-007','Toni Jackson','t.jackson@jacksonrecruit.com','(512)555-0707','Jackson Recruit','Online','04/18/2026','Confirmed','','No Meal','None','None','','','Paid in Full',0,'No','No','No','','','No'],
  ['EVT-007','Ella Zimmerman','e.zimm@zimmhire.com','(512)555-0708','ZimmHire','Online','04/18/2026','Registered','','No Meal','None','None','','','Not Paid',0,'No','No','No','','','No'],
  ['EVT-007','Carlos Rivera','c.rivera@riveraresume.com','(512)555-0709','Rivera Resume','Online','04/19/2026','Confirmed','','No Meal','None','None','','','Paid in Full',0,'No','No','No','','','No'],
  ['EVT-007','Ivy Patel','i.patel@patelconnects.com','(512)555-0710','Patel Connects','Online','04/19/2026','Registered','','No Meal','None','None','','','Not Paid',0,'No','No','No','','','No'],
];

(async () => {
  const values = [];
  const fmt = [];

  // Title
  values.push({ range:"'Guests & Attendees'!A1", values:[['GUESTS & ATTENDEES']] });
  values.push({ range:"'Guests & Attendees'!A2", values:[['Track all attendees across events. Link by Event ID. Auto-calculates capacity, dietary counts, and amounts paid.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,22), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,22), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,22), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,22), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:44}, fields:'pixelSize' }});

  // Info row 3
  values.push({ range:"'Guests & Attendees'!A3", values:[['ℹ VIP (Q), Speaking (R), Volunteer (S), and Follow-Up (V) columns use checkbox. Amount Paid (P) is a currency field.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,22), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,22), cell:{userEnteredFormat:{
    backgroundColor:hex('#EEF4F8'), textFormat:{italic:true,fontSize:9,foregroundColor:hex('#4A5056')},
    verticalAlignment:'MIDDLE', padding:{left:10},
  }}, fields:'userEnteredFormat' }});

  // Stats summary row 4 (merged)
  values.push({ range:"'Guests & Attendees'!A4", values:[
    [`=IFERROR("Total Attendees: "&COUNTA(B6:B500)&" | Confirmed: "&COUNTIF(H6:H500,"Confirmed")&" | VIP: "&COUNTIF(Q6:Q500,TRUE)&" | Total Paid: $"&TEXT(SUM(P6:P500),"#,##0"),"Loading...")`]
  ]});
  fmt.push({ mergeCells:{ range: gridRange(SID,3,4,0,22), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,3,4,0,22), cell:{userEnteredFormat:{
    backgroundColor:hex('#C9DCEF'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#1E3A5F')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:3,endIndex:4},
    properties:{pixelSize:28}, fields:'pixelSize' }});

  // Header row 5
  values.push({ range:"'Guests & Attendees'!A5", values:[HEADERS] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,22), cell:{userEnteredFormat:{
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

  // Data rows
  const dataRows = ATTENDEES.map((att,i) => {
    const row = [...att];
    // Q=16,R=17,S=18: TRUE/FALSE checkboxes
    row[16] = row[16] === 'Yes';
    row[17] = row[17] === 'Yes';
    row[18] = row[18] === 'Yes';
    // V=21: Follow-Up checkbox
    row[21] = row[21] === 'Yes';
    return row;
  });
  values.push({ range:"'Guests & Attendees'!A6", values: dataRows });

  // Checkboxes: Q=16, R=17, S=18, V=21
  [16,17,18,21].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,5,500,c,c+1), cell:{userEnteredFormat:{
      horizontalAlignment:'CENTER',
    }, dataValidation:{ condition:{ type:'BOOLEAN' }, showCustomUi:true }},
    fields:'userEnteredFormat.horizontalAlignment,dataValidation' }});
  });

  // Alt rows
  ATTENDEES.forEach((_,i)=>{
    const ri = i+5;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,22), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#FDFCF9':'#F2F1EE'),
      textFormat:{foregroundColor:hex('#2F3336'),fontSize:9},
      verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Validations
  // F=5: Ticket Type
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,5,6), rule:{
    condition:{ type:'ONE_OF_LIST', values:['General Admission','VIP','Staff','Speaker','Sponsor','Complimentary','Online'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  // H=7: Attendance Status
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,7,8), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Registered','Confirmed','Attended','No-Show','Cancelled','Waitlist'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});
  // J=9: Meal Choice
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,9,10), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Standard','Vegetarian','Vegan','Gluten-Free','Halal','Kosher','No Meal','Unknown'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:false,
  }}});
  // O=14: Payment Status
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,500,14,15), rule:{
    condition:{ type:'ONE_OF_LIST', values:['Not Paid','Deposit Paid','Partial','Paid in Full','Refunded'].map(v=>({userEnteredValue:v})) },
    showCustomUi:true, strict:true,
  }}});

  // Currency format: P=15 Amount Paid
  fmt.push({ repeatCell:{ range: gridRange(SID,5,500,15,16), cell:{userEnteredFormat:{
    numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0.00' }
  }}, fields:'userEnteredFormat.numberFormat' }});

  // CF: Attendance Status (H=col7)
  const attCF = [
    { v:'Confirmed', bg:'#CBDCC5', text:'#2A5C35' },
    { v:'Attended',  bg:'#92AE95', text:'#1A3A22' },
    { v:'Registered',bg:'#D5E4F2', text:'#1E3A5F' },
    { v:'Waitlist',  bg:'#E9DBA7', text:'#5C4A00' },
    { v:'No-Show',   bg:'#E8C0B8', text:'#5C1A1A' },
    { v:'Cancelled', bg:'#DDD5ED', text:'#3A1A5C' },
  ];
  attCF.forEach(s=>{
    fmt.push({ addConditionalFormatRule:{ rule:{
      ranges:[gridRange(SID,5,500,7,8)],
      booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:s.v}] },
        format:{ backgroundColor:hex(s.bg), textFormat:{foregroundColor:hex(s.text),bold:true} } }
    }, index:0 }});
  });

  // CF: Follow-Up Needed = TRUE
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,0,22)],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=AND($V6=TRUE,$A6<>"")`}] },
      format:{ backgroundColor:hex('#F2D2B6') } }
  }, index:0 }});

  // CF: Dietary Restrictions not "None" (K=col10)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,10,11)],
    booleanRule:{ condition:{ type:'TEXT_NOT_CONTAINS', values:[{userEnteredValue:'None'}] },
      format:{ backgroundColor:hex('#E9DBA7'), textFormat:{bold:true} } }
  }, index:0 }});

  // CF: VIP = TRUE
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,16,17)],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:'=$Q6=TRUE'}] },
      format:{ backgroundColor:hex('#DDD0DF'), textFormat:{bold:true} } }
  }, index:0 }});

  // CF: Payment Status
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,14,15)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Not Paid'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{foregroundColor:hex('#5C1A1A')} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,500,14,15)],
    booleanRule:{ condition:{ type:'TEXT_EQ', values:[{userEnteredValue:'Paid in Full'}] },
      format:{ backgroundColor:hex('#CBDCC5'), textFormat:{foregroundColor:hex('#2A5C35')} } }
  }, index:0 }});

  // Filter
  fmt.push({ setBasicFilter:{ filter:{ range: gridRange(SID,4,500,0,22) }}});

  await valuesBatchUpdate(id, values, '06-guests values');
  await batchUpdate(id, fmt, '06-guests format');
  console.log('✅ Guests & Attendees done.');
})().catch(e => { console.error(e.errors||e.message||e); process.exit(1); });
