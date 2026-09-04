'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const IT = sheetMap['Invitation Tracker'];
const S = "'Invitation Tracker'";

// 22 sample invitation rows — Guest IDs GST-001 through GST-022 (first 22 guests from list)
// We use VLOOKUP from Guest List & RSVP to pull name, household, RSVP status, contact
// [GuestID, InvMethod, InvReady, SentDate, FollowUpDate, ReminderSent, Notes]
const INVITES = [
  ['GST-001','Printed Invitation',true,'Jul 5, 2026','','',false,'Hand delivered at school'],
  ['GST-002','Printed Invitation',true,'Jul 5, 2026','','',false,''],
  ['GST-003','Printed Invitation',true,'Jul 5, 2026','','',false,'Nut-free note included'],
  ['GST-004','Printed Invitation',true,'Jul 5, 2026','','',false,''],
  ['GST-005','Email',true,'Jul 6, 2026','','',true,'Sent PDF invitation'],
  ['GST-006','Email',true,'Jul 6, 2026','','',false,''],
  ['GST-007','Text Message',true,'Jul 6, 2026','','',false,''],
  ['GST-008','Printed Invitation',true,'Jul 5, 2026','','',false,'Dairy-free note included'],
  ['GST-009','Online Invitation',true,'Jul 7, 2026','Jul 28, 2026','Jul 28, 2026',false,'Follow up needed — maybe'],
  ['GST-010','Online Invitation',true,'Jul 7, 2026','Jul 28, 2026','Jul 28, 2026',false,''],
  ['GST-011','Text Message',true,'Jul 7, 2026','','',false,'Vegan note added'],
  ['GST-012','Text Message',true,'Jul 7, 2026','','',false,''],
  ['GST-013','Email',true,'Jul 6, 2026','','',false,'Declined — confirmed'],
  ['GST-014','Email',true,'Jul 6, 2026','','',false,''],
  ['GST-015','Printed Invitation',true,'Jul 5, 2026','','',false,'Halal meal required'],
  ['GST-016','Printed Invitation',true,'Jul 5, 2026','','',false,''],
  ['GST-017','Online Invitation',true,'Jul 8, 2026','Jul 28, 2026','Jul 28, 2026',false,'No response — send reminder'],
  ['GST-018','Online Invitation',true,'Jul 8, 2026','Jul 28, 2026','Jul 28, 2026',false,''],
  ['GST-019','Printed Invitation',true,'Jul 5, 2026','','',false,'Egg allergy noted'],
  ['GST-020','Printed Invitation',true,'Jul 5, 2026','','',false,''],
  ['GST-021','Text Message',true,'Jul 6, 2026','','',false,''],
  ['GST-022','Text Message',true,'Jul 6, 2026','','',false,'Vegetarian option needed'],
];

(async () => {
  const data = [];

  // Title & guidance
  data.push({ range:`${S}!A1`, values:[['INVITATION TRACKER']] });
  data.push({ range:`${S}!A3`, values:[['Track which guests have been invited, when invitations were sent, and who still needs a follow-up. Guest details pull from the Guest List & RSVP tab.']] });

  // Headers row 6
  data.push({ range:`${S}!A6:N6`, values:[[
    'Guest ID','Guest Name','Household / Family','Invitation Method','Invitation Ready?',
    'Sent Date','RSVP Deadline','Response Status','Follow-Up Needed?',
    'Follow-Up Date','Days Since Sent','Reminder Sent?','Contact Details','Notes'
  ]] });

  // Formulas for rows 7-206
  for (let r = 7; r <= 206; r++) {
    // B: Guest Name from Guest List
    data.push({ range:`${S}!B${r}`, values:[[`=IFERROR(IF(A${r}="","",VLOOKUP(A${r},'Guest List & RSVP'!$A$6:$R$205,3,FALSE)),"")`]] });
    // C: Household
    data.push({ range:`${S}!C${r}`, values:[[`=IFERROR(IF(A${r}="","",VLOOKUP(A${r},'Guest List & RSVP'!$A$6:$R$205,2,FALSE)),"")`]] });
    // G: RSVP Deadline from Party Setup
    data.push({ range:`${S}!G${r}`, values:[[`=IFERROR(IF(A${r}="","","August 1, 2026"),"")`]] });
    // H: Response Status from Guest List (RSVP col I = col 9 in the 18-col range)
    data.push({ range:`${S}!H${r}`, values:[[`=IFERROR(IF(A${r}="","",VLOOKUP(A${r},'Guest List & RSVP'!$A$6:$R$205,9,FALSE)),"")`]] });
    // I: Follow-Up Needed — Yes when sent, RSVP is No Response or Maybe, and not declined
    data.push({ range:`${S}!I${r}`, values:[[`=IFERROR(IF(A${r}="","",IF(AND(E${r}=TRUE,F${r}<>"",OR(H${r}="No Response",H${r}="Maybe"),H${r}<>"Declined"),"Yes","No")),"")`]] });
    // K: Days Since Sent
    data.push({ range:`${S}!K${r}`, values:[[`=IFERROR(IF(F${r}="","",TODAY()-F${r}),"")`]] });
    // M: Contact Details (Phone from col 6, Email col 7 = 6th and 7th data columns)
    data.push({ range:`${S}!M${r}`, values:[[`=IFERROR(IF(A${r}="","",VLOOKUP(A${r},'Guest List & RSVP'!$A$6:$R$205,6,FALSE)&IF(VLOOKUP(A${r},'Guest List & RSVP'!$A$6:$R$205,7,FALSE)<>"","  |  "&VLOOKUP(A${r},'Guest List & RSVP'!$A$6:$R$205,7,FALSE),"")),"")`]] });
  }

  // Sample data
  INVITES.forEach(([gid,method,ready,sent,followupDate,_,reminder,notes], i) => {
    const r = 7 + i;
    data.push({ range:`${S}!A${r}`, values:[[gid]] });
    data.push({ range:`${S}!D${r}`, values:[[method]] });
    data.push({ range:`${S}!E${r}`, values:[[ready]] });
    if (sent) data.push({ range:`${S}!F${r}`, values:[[sent]] });
    if (followupDate) data.push({ range:`${S}!J${r}`, values:[[followupDate]] });
    data.push({ range:`${S}!L${r}`, values:[[reminder]] });
    if (notes) data.push({ range:`${S}!N${r}`, values:[[notes]] });
  });

  await valuesBatchUpdate(id, data, 'invitation-values');

  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(IT,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };

  // Title rows 0-1
  reqs.push({ mergeCells:{ range:gridRange(IT,0,2,0,14), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:20, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Guidance rows 2-4
  reqs.push({ mergeCells:{ range:gridRange(IT,2,4,0,14), mergeType:'MERGE_ALL' } });
  fmt(2,4,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });

  // Blank row 5 (index 4)
  fmt(4,5,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.bg) } });

  // Header row 6 (index 5)
  fmt(5,6,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });

  // Data rows indices 6-205
  for (let r = 6; r < 206; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,14,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }

  // Formula cols: B(1),C(2),G(6),H(7),I(8),K(10),M(12)
  for (const ci of [1,2,6,7,8,10,12]) {
    reqs.push({ repeatCell:{ range:gridRange(IT,6,206,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  // Input cols: A(0),D(3),F(5),J(9),N(13)
  for (const ci of [0,3,5,9,13]) {
    reqs.push({ repeatCell:{ range:gridRange(IT,6,206,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' } });
  }
  // Date cols F(5), G(6), J(9)
  for (const ci of [5,6,9]) {
    reqs.push({ repeatCell:{ range:gridRange(IT,6,206,ci,ci+1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  }
  // Checkbox E(4) Invitation Ready, L(11) Reminder Sent
  for (let r = 6; r < 206; r++) {
    reqs.push({ setDataValidation:{ range:gridRange(IT,r,r+1,4,5), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
    reqs.push({ setDataValidation:{ range:gridRange(IT,r,r+1,11,12), rule:{ condition:{ type:'BOOLEAN' }, showCustomUi:true } } });
  }
  // Dropdown: Invitation Method D(3)
  reqs.push({ setDataValidation:{ range:gridRange(IT,6,206,3,4), rule:{ condition:{ type:'ONE_OF_LIST', values:['Printed Invitation','Email','Text Message','Online Invitation','Phone Call','Social Media','In Person'].map(v=>({ userEnteredValue:v })) }, showCustomUi:true, strict:true } } });

  // Conditional formatting: Follow-Up Needed col I (index 8)
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(IT,6,206,8,9)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'Yes' }] }, format:{ backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true } } } }, index:0 } });
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(IT,6,206,8,9)], booleanRule:{ condition:{ type:'TEXT_EQ', values:[{ userEnteredValue:'No' }] }, format:{ backgroundColor:hex(C.success), textFormat:{ foregroundColor:hex(C.white) } } } }, index:0 } });

  reqs.push({ updateBorders:{ range:gridRange(IT,5,206,0,14), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  [[0,80],[1,120],[2,120],[3,120],[4,90],[5,100],[6,100],[7,110],[8,110],[9,100],[10,90],[11,90],[12,200],[13,150]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:IT, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:IT, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:40 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:IT, dimension:'ROWS', startIndex:2, endIndex:6 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:IT, dimension:'ROWS', startIndex:6, endIndex:206 }, properties:{ pixelSize:22 }, fields:'pixelSize' } });

  reqs.push({ updateSheetProperties:{ properties:{ sheetId:IT, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' } });

  await batchUpdate(id, reqs, 'invitation-format');
  console.log('Invitation Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
