'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GL = sheetMap['Guest List & RSVP'];
const S = "'Guest List & RSVP'";

// 27 sample guests
// [Household, Name, GuestType, ChildOrAdult, Phone, Email, InvStatus, RSVP, NumAttending, Dietary, AllergyNotes, GiftReceived, TYSent, Seating, ArrivalNotes, FollowUpDate, Notes]
const GUESTS = [
  ['Hendricks Family','Lily Hendricks','Child','Child','','','Invitation Sent','Confirmed',1,'None','','',false,false,'Table 1','','','Best friends since kindergarten'],
  ['Hendricks Family','Nora Hendricks','Parent / Guardian','Adult','(555)201-1234','nora.h@email.com','Invitation Sent','Confirmed',1,'None','','',false,false,'','','',''],
  ['Patel Family','Aarav Patel','Child','Child','','','Invitation Sent','Confirmed',1,'Nut-Free','No tree nuts','',true,true,'Table 1','','','Bring nut-free snack pack'],
  ['Patel Family','Priya Patel','Parent / Guardian','Adult','(555)304-5678','priya.p@email.com','Invitation Sent','Confirmed',1,'None','','',false,false,'','','',''],
  ['Robinson Family','Zoey Robinson','Classmate','Child','','','Invitation Sent','Confirmed',1,'Gluten-Free','Celiac disease','',true,false,'Table 2','','','Mum to confirm transport'],
  ['Robinson Family','Ben Robinson','Parent / Guardian','Adult','(555)445-9012','b.robinson@email.com','Invitation Sent','Confirmed',1,'None','','',false,false,'','','',''],
  ['Kim Family','Ethan Kim','Classmate','Child','','','Invitation Sent','Confirmed',1,'None','','',false,false,'Table 2','','',''],
  ['Alvarez Family','Sofia Alvarez','Friend','Child','','','Invitation Sent','Confirmed',1,'Dairy-Free','Lactose intolerant','',true,true,'Table 1','','','Loves craft activities'],
  ['Nguyen Family','Hana Nguyen','Classmate','Child','','','Invitation Sent','Maybe',1,'None','','',false,false,'','','Aug 5, 2026','Waiting on sports schedule'],
  ['Nguyen Family','Tran Nguyen','Parent / Guardian','Adult','(555)556-3456','','Invitation Sent','Maybe',1,'None','','',false,false,'','','Aug 5, 2026',''],
  ['Walker Family','Jack Walker','Neighbor','Child','','','Invitation Sent','Confirmed',1,'Vegan','Full vegan','',false,false,'Table 3','','',''],
  ['Walker Family','Dana Walker','Parent / Guardian','Adult','(555)667-7890','d.walker@email.com','Invitation Sent','Confirmed',1,'None','','',false,false,'','','',''],
  ['Campbell Family','Emma Campbell','Friend','Child','','','Invitation Sent','Declined',0,'None','','',false,false,'','','','Away on holiday'],
  ['Campbell Family','Ruth Campbell','Parent / Guardian','Adult','(555)778-2345','ruth.c@email.com','Invitation Sent','Declined',0,'None','','',false,false,'','','',''],
  ['Singh Family','Arjun Singh','Classmate','Child','','','Invitation Sent','Confirmed',1,'Halal','Halal meat only','',true,true,'Table 2','','',''],
  ['Singh Family','Meera Singh','Parent / Guardian','Adult','(555)889-6789','','Invitation Sent','Confirmed',1,'Halal','','',false,false,'','','',''],
  ['Thomas Family','Oliver Thomas','Classmate','Child','','','Invitation Sent','No Response',1,'None','','',false,false,'','','Jul 28, 2026','Send reminder'],
  ['Thomas Family','Karen Thomas','Parent / Guardian','Adult','(555)991-0123','k.thomas@email.com','Invitation Sent','No Response',1,'None','','',false,false,'','','Jul 28, 2026',''],
  ['Green Family','Isla Green','Friend','Child','','','Invitation Sent','Confirmed',1,'Egg-Free','Egg allergy confirmed','',true,false,'Table 3','','',''],
  ['Green Family','Mark Green','Parent / Guardian','Adult','(555)112-4567','','Invitation Sent','Confirmed',1,'None','','',false,false,'','','',''],
  ['Torres Family','Chloe Torres','Classmate','Child','','','Invitation Sent','Confirmed',1,'None','','',false,false,'Table 3','','',''],
  ['Torres Family','Angela Torres','Parent / Guardian','Adult','(555)223-8901','a.torres@email.com','Invitation Sent','Confirmed',1,'Vegetarian','','',false,false,'','','',''],
  ['Brooks Family','Leo Brooks','Neighbor','Child','','','Invitation Sent','No Response',1,'None','','',false,false,'','','Jul 30, 2026',''],
  ['Brooks Family','Phil Brooks','Parent / Guardian','Adult','(555)334-2345','p.brooks@email.com','Invitation Sent','No Response',1,'None','','',false,false,'','','Jul 30, 2026',''],
  ['Mia\'s Grandma','Ellen Hendricks','Family','Adult','(555)445-6789','','Invitation Sent','Confirmed',1,'None','','',true,false,'','Will bring cake stand','',''],
  ['Mia\'s Grandpa','Paul Hendricks','Family','Adult','(555)445-6789','','Invitation Sent','Confirmed',1,'None','','',false,false,'','','',''],
  ['Sharma Family','Anika Sharma','Classmate','Child','','','Invitation Ready','Invitation Ready',1,'Kosher','Strictly kosher','',false,false,'','','','Check kosher catering options'],
];

(async () => {
  const data = [];

  // Title & guidance
  data.push({ range:`${S}!A1`, values:[['GUEST LIST & RSVP TRACKER']] });
  data.push({ range:`${S}!A3`, values:[['Enter each invited household or guest. Invitation, food, dietary, and dashboard totals update automatically from this list.']] });

  // Headers row 6
  data.push({ range:`${S}!A6:R6`, values:[[
    'Guest ID','Household / Family','Guest Name','Guest Type','Child or Adult',
    'Phone','Email','Invitation Status','RSVP Status','Number Attending',
    'Dietary Restriction','Allergy Notes','Gift Received?','Thank-You Sent?',
    'Seating / Group','Arrival Notes','Follow-Up Date','Notes'
  ]] });

  // Guest ID formulas (col A)
  for (let r = 7; r <= 205; r++) {
    data.push({ range:`${S}!A${r}`, values:[[`=IF(C${r}="","","GST-"&TEXT(ROW()-6,"000"))`]] });
  }

  // Sample data rows 7-33
  GUESTS.forEach(([household,name,gtype,childAdult,phone,email,invSt,rsvp,numAtt,dietary,allergyNotes,_gift1,gift,ty,seating,arrival,followup,notes], i) => {
    const r = 7 + i;
    // B=Household, C=Name, D=GuestType, E=ChildOrAdult, F=Phone, G=Email,
    // H=InvStatus, I=RSVP, J=NumAtt, K=Dietary, L=Allergy, M=Gift(checkbox via validation),
    // N=TY(checkbox), O=Seating, P=Arrival, Q=FollowUp, R=Notes
    data.push({ range:`${S}!B${r}:L${r}`, values:[[household,name,gtype,childAdult,phone,email,invSt,rsvp,numAtt,dietary,allergyNotes]] });
    data.push({ range:`${S}!M${r}`, values:[[gift]] });
    data.push({ range:`${S}!N${r}`, values:[[ty]] });
    if (seating) data.push({ range:`${S}!O${r}`, values:[[seating]] });
    if (arrival) data.push({ range:`${S}!P${r}`, values:[[arrival]] });
    if (followup) data.push({ range:`${S}!Q${r}`, values:[[followup]] });
    if (notes) data.push({ range:`${S}!R${r}`, values:[[notes]] });
  });

  await valuesBatchUpdate(id, data, 'guestlist-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(GL,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };

  // Title rows 0-1
  reqs.push({ mergeCells:{ range:gridRange(GL,0,2,0,18), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,18,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Guidance row 2 (index 2)
  reqs.push({ mergeCells:{ range:gridRange(GL,2,4,0,18), mergeType:'MERGE_ALL' } });
  fmt(2,4,0,18,{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });

  // Blank row 5 (index 4) separator
  fmt(4,5,0,18,{ userEnteredFormat:{ backgroundColor:hex(C.bg) } });

  // Header row 6 (index 5)
  fmt(5,6,0,18,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });

  // Data rows indices 6-204
  for (let r = 6; r < 205; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,18,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }

  // Formula col A (0) — pale blue
  reqs.push({ repeatCell:{ range:gridRange(GL,6,205,0,1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });

  // Input cols B,C,D,E,F,G,H,I,J,K,L,O,P,R (1,2,3,4,5,6,7,8,9,10,11,14,15,17)
  for (const ci of [1,2,3,4,5,6,7,8,9,10,11,14,15,17]) {
    reqs.push({ repeatCell:{ range:gridRange(GL,6,205,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  }

  // Date col Q (16) — follow-up date
  reqs.push({ repeatCell:{ range:gridRange(GL,6,205,16,17), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER', backgroundColor:hex(C.input) } }, fields:'userEnteredFormat' } });

  // Checkboxes: M(12) Gift Received, N(13) Thank-You Sent
  for (let r = 6; r < 205; r++) {
    reqs.push({ setDataValidation:{ range:gridRange(GL,r,r+1,12,13), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
    reqs.push({ setDataValidation:{ range:gridRange(GL,r,r+1,13,14), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
  }

  // Dropdowns
  const listRule = (values) => ({ condition:{ type:'ONE_OF_LIST', values:values.map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true });
  // D=GuestType(3)
  reqs.push({ setDataValidation:{ range:gridRange(GL,6,205,3,4), rule:listRule(['Child','Parent / Guardian','Sibling','Family','Friend','Classmate','Neighbor','Other']) } });
  // E=ChildOrAdult(4)
  reqs.push({ setDataValidation:{ range:gridRange(GL,6,205,4,5), rule:listRule(['Child','Adult']) } });
  // H=InvStatus(7)
  reqs.push({ setDataValidation:{ range:gridRange(GL,6,205,7,8), rule:listRule(['Not Invited','Invitation Ready','Invitation Sent','Follow-Up Needed','Confirmed','Declined','Maybe','No Response']) } });
  // I=RSVP(8)
  reqs.push({ setDataValidation:{ range:gridRange(GL,6,205,8,9), rule:listRule(['Not Invited','Invitation Ready','Invitation Sent','Follow-Up Needed','Confirmed','Declined','Maybe','No Response']) } });
  // K=Dietary(10)
  reqs.push({ setDataValidation:{ range:gridRange(GL,6,205,10,11), rule:listRule(['None','Vegetarian','Vegan','Gluten-Free','Dairy-Free','Nut-Free','Egg-Free','Halal','Kosher','Other','Unknown']) } });

  // Conditional formatting — RSVP Status col I (index 8)
  const rsvpCF = [
    { val:'Confirmed',         bg:C.success,   fg:C.white },
    { val:'Declined',          bg:C.border,    fg:C.secText },
    { val:'Maybe',             bg:C.warning,   fg:C.mainText },
    { val:'No Response',       bg:C.lavender,  fg:C.mainText },
    { val:'Follow-Up Needed',  bg:C.attention, fg:C.white },
  ];
  rsvpCF.forEach(({ val, bg, fg }) => {
    reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(GL,6,205,8,9)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:val }] }, format:{ backgroundColor:hex(bg), textFormat:{ foregroundColor:hex(fg), bold:true } } } }, index:0 } });
  });

  // Thank-you tracking: gift received (M=true) but TY not sent (N=false) → amber row tint on N col
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(GL,6,205,13,14)], booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=AND($M7=TRUE,$N7=FALSE)` }] }, format:{ backgroundColor:hex(C.warning), textFormat:{ bold:true } } } }, index:0 } });
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(GL,6,205,13,14)], booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=AND($M7=TRUE,$N7=TRUE)` }] }, format:{ backgroundColor:hex(C.success), textFormat:{ foregroundColor:hex(C.white), bold:true } } } }, index:0 } });

  reqs.push({ updateBorders:{ range:gridRange(GL,5,205,0,18), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Column widths
  [[0,80],[1,130],[2,120],[3,110],[4,80],[5,110],[6,140],[7,110],[8,110],[9,80],[10,110],[11,130],[12,80],[13,80],[14,90],[15,110],[16,100],[17,140]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:GL, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:GL, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:GL, dimension:'ROWS', startIndex:2, endIndex:6 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:GL, dimension:'ROWS', startIndex:6, endIndex:205 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  // Freeze rows 1-5 (frozenRowCount:5) and columns A-C (frozenColumnCount:3)
  // Title: rows 0-1 (full merge), guidance rows 2-3, blank row 4, header row 5
  // Unmerge title then freeze then re-merge to handle column freeze + full-width merge
  reqs.push({ unmergeCells:{ range:gridRange(GL,0,2,0,18) } });
  reqs.push({ unmergeCells:{ range:gridRange(GL,2,4,0,18) } });
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:GL, gridProperties:{ frozenRowCount:5, frozenColumnCount:3 } }, fields:'gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });
  reqs.push({ mergeCells:{ range:gridRange(GL,0,2,3,18), mergeType:'MERGE_ALL' } });
  reqs.push({ mergeCells:{ range:gridRange(GL,2,4,3,18), mergeType:'MERGE_ALL' } });

  await batchUpdate(id, reqs, 'guestlist-format');
  console.log('Guest List & RSVP complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
