'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DB = sheetMap['Dashboard'];
const S = "'Dashboard'";

// ── Formula helpers ──
const GL = "'Guest List & RSVP'";
const BE = "'Budget & Expenses'";
const FD = "'Food, Drinks & Cake'";
const SL = "'Shopping List'";
const PC = "'Prep, Setup & Cleanup'";
const IT = "'Invitation Tracker'";

// Confirmed guests total (sum of Number Attending for Confirmed)
const confirmedGuests = `=IFERROR(SUMIF(${GL}!$I$6:$I$205,"Confirmed",${GL}!$J$6:$J$205),0)`;
const pendingRsvps    = `=IFERROR(COUNTIF(${GL}!$I$6:$I$205,"No Response")+COUNTIF(${GL}!$I$6:$I$205,"Maybe"),0)`;
const totalBudget     = `=IFERROR('Party Setup'!B16,0)`;
const amountSpent     = `=IFERROR(SUM(${BE}!$I$31:$I$229),0)`;
const remainingBudget = `=IFERROR(B10-D10,0)`;
const costPerGuest    = `=IFERROR(IF(B9=0,0,D10/B9),0)`;
const daysUntilParty  = `=IFERROR(MAX(0,DATEVALUE("August 15, 2026")-TODAY()),"—")`;
// Planning completion % (weighted):
// Invitations 20%: % of guests with Invitation Sent / total guests
// RSVPs 20%: % confirmed or declined / total invited
// Budget 15%: any budget set in category table
// Shopping 20%: % items purchased
// Prep Tasks 25%: % tasks complete
const planningPct = `=IFERROR(
  0.20*IFERROR(COUNTIF(${GL}!$H$6:$H$205,"Invitation Sent")/MAX(1,COUNTA(${GL}!$C$6:$C$205)),0)+
  0.20*IFERROR((COUNTIF(${GL}!$I$6:$I$205,"Confirmed")+COUNTIF(${GL}!$I$6:$I$205,"Declined"))/MAX(1,COUNTA(${GL}!$C$6:$C$205)),0)+
  0.15*IFERROR(IF(SUM(${BE}!$B$11:$B$25)>0,1,0),0)+
  0.20*IFERROR(COUNTIF(${SL}!$K$10:$K$209,TRUE)/MAX(1,COUNTA(${SL}!$C$10:$C$209)),0)+
  0.25*IFERROR(COUNTIF(${PC}!$J$10:$J$209,TRUE)/MAX(1,COUNTA(${PC}!$C$10:$C$209)),0)
,0)`.replace(/\n/g,'');

// Secondary cards
const guestCapRemaining = `=IFERROR('Party Setup'!B17-B9,0)`;
const invNotSent = `=IFERROR(COUNTIF(${GL}!$H$6:$H$205,"Invitation Ready")+COUNTIF(${GL}!$H$6:$H$205,"Not Invited"),0)`;
const followUpsNeeded = `=IFERROR(COUNTIF(${IT}!$I$7:$I$206,"Yes"),0)`;
const dietaryRestrictions = `=IFERROR(SUMPRODUCT((${GL}!$I$6:$I$205="Confirmed")*(${GL}!$K$6:$K$205<>"None")*(${GL}!$K$6:$K$205<>"")),0)`;
const shoppingRemaining = `=IFERROR(COUNTA(${SL}!$C$10:$C$209)-COUNTIF(${SL}!$K$10:$K$209,TRUE),0)`;
const prepTasksRemaining = `=IFERROR(COUNTA(${PC}!$C$10:$C$209)-COUNTIF(${PC}!$J$10:$J$209,TRUE),0)`;
const overdueTasks = `=IFERROR(COUNTIF(${PC}!$M$10:$M$209,"Yes"),0)`;
const cakeStatus = `=IFERROR(IF(${FD}!B136="Yes","Paid",IF(${FD}!B124<>"","Ordered","Not Ordered")),"—")`;

(async () => {
  const data = [];

  // Title & subtitle
  data.push({ range:`${S}!A1`, values:[['KIDS BIRTHDAY PARTY DASHBOARD']] });
  data.push({ range:`${S}!A2`, values:[['Guests  •  Budget  •  Invitations  •  Food  •  Shopping  •  Preparation']] });
  data.push({ range:`${S}!A3`, values:[['This dashboard summarises your party planning progress. All values update automatically from the other tabs.']] });

  // Primary KPI cards row 6 (labels) and row 7 (values)
  // 4 cards per row, 2 rows = 8 cards total across A:H
  const primaryKPIs = [
    ['Days Until Party', daysUntilParty],
    ['Confirmed Guests', confirmedGuests],
    ['Pending RSVPs', pendingRsvps],
    ['Total Budget', totalBudget],
    ['Amount Spent', amountSpent],
    ['Remaining Budget', remainingBudget],
    ['Cost per Guest', costPerGuest],
    ['Planning Progress', planningPct],
  ];

  // Row 6 labels (A6:H6), Row 7 values (A7:H7) - but we do 4+4 across two pair rows
  // Cards at A6-H6 and A7-H7 (4 cards each, 2 columns wide each = 8 cols)
  primaryKPIs.forEach(([label, formula], i) => {
    const col = String.fromCharCode(65 + i * 2); // A,C,E,G for row 6
    const col2 = String.fromCharCode(65 + i * 2 + 1);
    const row = i < 4 ? 6 : 10;
    const offset = i < 4 ? i : i - 4;
    const c = String.fromCharCode(65 + offset * 2);
    const c2 = String.fromCharCode(65 + offset * 2 + 1);
    data.push({ range:`${S}!${c}${row}`, values:[[label]] });
    data.push({ range:`${S}!${c2}${row}`, values:[[formula]] });
  });

  // Secondary insight cards row 13 (labels) and 14 (values)
  const secondaryKPIs = [
    ['Guest Capacity Left', guestCapRemaining],
    ['Invitations Not Sent', invNotSent],
    ['Follow-Ups Needed', followUpsNeeded],
    ['Dietary Restrictions', dietaryRestrictions],
    ['Shopping Remaining', shoppingRemaining],
    ['Prep Tasks Remaining', prepTasksRemaining],
    ['Overdue Tasks', overdueTasks],
    ['Cake Status', cakeStatus],
  ];
  secondaryKPIs.forEach(([label, formula], i) => {
    const offset = i < 4 ? i : i - 4;
    const row = i < 4 ? 13 : 16;
    const c = String.fromCharCode(65 + offset * 2);
    const c2 = String.fromCharCode(65 + offset * 2 + 1);
    data.push({ range:`${S}!${c}${row}`, values:[[label]] });
    data.push({ range:`${S}!${c2}${row}`, values:[[formula]] });
  });

  // Guest snapshot section (row 20)
  data.push({ range:`${S}!A20`, values:[['GUEST SNAPSHOT']] });
  data.push({ range:`${S}!A21:B21`, values:[['Metric','Count']] });
  const guestSnap = [
    ['Confirmed Children', `=IFERROR(SUMPRODUCT((${GL}!$I$6:$I$205="Confirmed")*(${GL}!$E$6:$E$205="Child")*(${GL}!$J$6:$J$205)),0)`],
    ['Confirmed Adults', `=IFERROR(SUMPRODUCT((${GL}!$I$6:$I$205="Confirmed")*(${GL}!$E$6:$E$205="Adult")*(${GL}!$J$6:$J$205)),0)`],
    ['Maybe', `=IFERROR(COUNTIF(${GL}!$I$6:$I$205,"Maybe"),0)`],
    ['Declined', `=IFERROR(COUNTIF(${GL}!$I$6:$I$205,"Declined"),0)`],
    ['No Response', `=IFERROR(COUNTIF(${GL}!$I$6:$I$205,"No Response"),0)`],
    ['Guest Capacity Remaining', guestCapRemaining],
  ];
  guestSnap.forEach(([label, formula], i) => {
    data.push({ range:`${S}!A${22+i}`, values:[[label]] });
    data.push({ range:`${S}!B${22+i}`, values:[[formula]] });
  });

  // Upcoming actions section (row 30)
  data.push({ range:`${S}!A30`, values:[['UPCOMING ACTIONS — NEXT 10 INCOMPLETE TASKS']] });
  data.push({ range:`${S}!A31:E31`, values:[['Due Date','Task','Assigned To','Priority','Status']] });
  for (let i = 1; i <= 10; i++) {
    const r = 31 + i;
    // Use SMALL to find the i-th soonest due date among incomplete tasks
    data.push({ range:`${S}!A${r}`, values:[[
      `=IFERROR(IF(ROW()-31>${i},"",INDEX(${PC}!$G$10:$G$209,MATCH(SMALL(IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),${i}),IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),0))),"")` ]] });
    data.push({ range:`${S}!B${r}`, values:[[
      `=IFERROR(IF(A${r}="","",INDEX(${PC}!$C$10:$C$209,MATCH(SMALL(IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),${i}),IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),0))),"")` ]] });
    data.push({ range:`${S}!C${r}`, values:[[
      `=IFERROR(IF(A${r}="","",INDEX(${PC}!$E$10:$E$209,MATCH(SMALL(IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),${i}),IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),0))),"")` ]] });
    data.push({ range:`${S}!D${r}`, values:[[
      `=IFERROR(IF(A${r}="","",INDEX(${PC}!$H$10:$H$209,MATCH(SMALL(IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),${i}),IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),0))),"")` ]] });
    data.push({ range:`${S}!E${r}`, values:[[
      `=IFERROR(IF(A${r}="","",INDEX(${PC}!$I$10:$I$209,MATCH(SMALL(IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),${i}),IF((${PC}!$J$10:$J$209<>TRUE)*(${PC}!$C$10:$C$209<>"")*(${PC}!$G$10:$G$209<>""),${PC}!$G$10:$G$209-DATE(2000,1,1)),0))),"")` ]] });
  }

  // Food & Shopping snapshot (row 45)
  data.push({ range:`${S}!A45`, values:[['FOOD & SHOPPING SNAPSHOT']] });
  data.push({ range:`${S}!A46:B46`, values:[['Item','Value']] });
  const foodSnap = [
    ['Planned Menu Items',      `=IFERROR(COUNTA(${FD}!$C$10:$C$109),0)`],
    ['Est. Food Cost',          `=IFERROR(SUM(${FD}!$J$10:$J$109),0)`],
    ['Dietary Coverage Issues', `=IFERROR(COUNTIF(${FD}!$D$113:$D$120,"Review Needed"),0)`],
    ['Shopping Items Remaining',shoppingRemaining],
    ['Cake Ordered?',           `=IFERROR(IF(${FD}!B128<>"","Yes","No"),"No")`],
    ['Cake Paid?',              `=IFERROR(${FD}!B136,"No")`],
  ];
  foodSnap.forEach(([label, formula], i) => {
    data.push({ range:`${S}!A${47+i}`, values:[[label]] });
    data.push({ range:`${S}!B${47+i}`, values:[[formula]] });
  });

  // Progress weight helper (row 55)
  data.push({ range:`${S}!A55`, values:[['PLANNING PROGRESS WEIGHTS']] });
  data.push({ range:`${S}!A56:C56`, values:[['Component','Weight','Note']] });
  const weights = [
    ['Invitations Sent','20%','% of guests with Invitation Sent status'],
    ['RSVPs Received','20%','% of guests confirmed or declined'],
    ['Budget Set','15%','At least one category budget entered'],
    ['Shopping Complete','20%','% of shopping items purchased'],
    ['Prep Tasks Done','25%','% of preparation tasks completed'],
  ];
  weights.forEach(([comp,w,note], i) => {
    data.push({ range:`${S}!A${57+i}`, values:[[comp]] });
    data.push({ range:`${S}!B${57+i}`, values:[[w]] });
    data.push({ range:`${S}!C${57+i}`, values:[[note]] });
  });
  data.push({ range:`${S}!A63`, values:[['Note: Planning Progress % is an estimate to help you stay on track. It does not guarantee party readiness.']] });

  // RSVP pivot for chart (col J-K hidden area)
  data.push({ range:`${S}!J6`, values:[['RSVP Status']] });
  data.push({ range:`${S}!K6`, values:[['Count']] });
  const rsvpStatuses = ['Confirmed','Declined','Maybe','No Response','Follow-Up Needed'];
  rsvpStatuses.forEach((s, i) => {
    data.push({ range:`${S}!J${7+i}`, values:[[s]] });
    data.push({ range:`${S}!K${7+i}`, values:[[`=IFERROR(COUNTIF(${GL}!$I$6:$I$205,"${s}"),0)`]] });
  });
  // Spending pivot for chart
  const cats = ['Venue','Invitations','Decorations','Food','Drinks','Cake & Desserts','Entertainment','Games & Activities','Party Favors','Supplies','Clothing','Photography','Transportation','Setup & Cleanup','Miscellaneous'];
  data.push({ range:`${S}!J13`, values:[['Category']] });
  data.push({ range:`${S}!K13`, values:[['Actual ($)']] });
  cats.forEach((cat, i) => {
    data.push({ range:`${S}!J${14+i}`, values:[[cat]] });
    data.push({ range:`${S}!K${14+i}`, values:[[`=IFERROR(SUMIF(${BE}!$C$31:$C$229,"${cat}",${BE}!$I$31:$I$229),0)`]] });
  });
  // Budget vs Actual for chart
  data.push({ range:`${S}!J30`, values:[['Category']] });
  data.push({ range:`${S}!K30`, values:[['Planned']] });
  data.push({ range:`${S}!L30`, values:[['Actual']] });
  cats.forEach((cat, i) => {
    data.push({ range:`${S}!J${31+i}`, values:[[cat]] });
    data.push({ range:`${S}!K${31+i}`, values:[[`=IFERROR(SUMIF(${BE}!$A$11:$A$25,"${cat}",${BE}!$B$11:$B$25),0)`]] });
    data.push({ range:`${S}!L${31+i}`, values:[[`=IFERROR(SUMIF(${BE}!$C$31:$C$229,"${cat}",${BE}!$I$31:$I$229),0)`]] });
  });
  // Prep task status pivot
  data.push({ range:`${S}!J47`, values:[['Status']] });
  data.push({ range:`${S}!K47`, values:[['Count']] });
  const taskStatuses = ['Not Started','In Progress','Waiting','Complete','Cancelled'];
  taskStatuses.forEach((s, i) => {
    data.push({ range:`${S}!J${48+i}`, values:[[s]] });
    data.push({ range:`${S}!K${48+i}`, values:[[`=IFERROR(COUNTIF(${PC}!$I$10:$I$209,"${s}"),0)`]] });
  });
  // Child vs Adult pivot
  data.push({ range:`${S}!J54`, values:[['Guest Type']] });
  data.push({ range:`${S}!K54`, values:[['Confirmed']] });
  data.push({ range:`${S}!J55`, values:[['Children']] });
  data.push({ range:`${S}!K55`, values:[[`=IFERROR(SUMPRODUCT((${GL}!$I$6:$I$205="Confirmed")*(${GL}!$E$6:$E$205="Child")*(${GL}!$J$6:$J$205)),0)`]] });
  data.push({ range:`${S}!J56`, values:[['Adults']] });
  data.push({ range:`${S}!K56`, values:[[`=IFERROR(SUMPRODUCT((${GL}!$I$6:$I$205="Confirmed")*(${GL}!$E$6:$E$205="Adult")*(${GL}!$J$6:$J$205)),0)`]] });
  // Dietary restriction pivot
  data.push({ range:`${S}!J58`, values:[['Restriction']] });
  data.push({ range:`${S}!K58`, values:[['Count']] });
  const dietRestrictions = ['Vegetarian','Vegan','Gluten-Free','Dairy-Free','Nut-Free','Egg-Free','Halal','Kosher'];
  dietRestrictions.forEach((d, i) => {
    data.push({ range:`${S}!J${59+i}`, values:[[d]] });
    data.push({ range:`${S}!K${59+i}`, values:[[`=IFERROR(SUMPRODUCT((${GL}!$I$6:$I$205="Confirmed")*(${GL}!$K$6:$K$205="${d}")),0)`]] });
  });

  await valuesBatchUpdate(id, data, 'dashboard-values');

  // ── Formatting ──
  const reqs = [];
  const fmt = (r1,r2,c1,c2,cell) => reqs.push({ repeatCell:{ range:gridRange(DB,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  const medium = { style:'SOLID_MEDIUM', color:hex(C.border) };
  const thin   = { style:'SOLID', color:hex(C.border) };
  const dateF  = { type:'DATE', pattern:'MMM D, YYYY' };
  const currF  = { type:'CURRENCY', pattern:'$#,##0.00' };
  const pctF   = { type:'PERCENT', pattern:'0%' };

  // Title rows 0-1
  reqs.push({ mergeCells:{ range:gridRange(DB,0,2,0,14), mergeType:'MERGE_ALL' } });
  fmt(0,2,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:24, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Subtitle row 2
  reqs.push({ mergeCells:{ range:gridRange(DB,2,3,0,14), mergeType:'MERGE_ALL' } });
  fmt(2,3,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  // Guidance row 3
  reqs.push({ mergeCells:{ range:gridRange(DB,3,5,0,14), mergeType:'MERGE_ALL' } });
  fmt(3,5,0,14,{ userEnteredFormat:{ backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText), fontSize:9, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // KPI section header row 5 (index 5)
  reqs.push({ mergeCells:{ range:gridRange(DB,5,6,0,8), mergeType:'MERGE_ALL' } });
  fmt(5,6,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });

  // Primary KPI cards rows 5-8 (2 rows of 4 cards)
  // Row 1 of cards: label at row 5 (index 5), value at row 6 (index 6) — wait, section header is at 5
  // Actually section header is at index 5, so cards start at index 6 (row 7)
  // Card row 1: indices 6-7 (label 6, value 7)
  // Card row 2: indices 9-10 (label 9, value 10) with separator at 8
  const cardLayout = [
    { labelRow:6, valRow:7 },  // 1st row of primary cards (indices)
    { labelRow:9, valRow:10 }, // 2nd row of primary cards
  ];
  cardLayout.forEach(({ labelRow, valRow }) => {
    for (let i = 0; i < 4; i++) {
      const c = i * 2;
      reqs.push({ mergeCells:{ range:gridRange(DB,labelRow,labelRow+1,c,c+2), mergeType:'MERGE_ALL' } });
      reqs.push({ mergeCells:{ range:gridRange(DB,valRow,valRow+1,c,c+2), mergeType:'MERGE_ALL' } });
      fmt(labelRow,labelRow+1,c,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.lavender), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
      fmt(valRow,valRow+1,c,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:18, fontFamily:'Arial', foregroundColor:hex(C.primary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    }
  });
  // Separator rows 8 (index 8)
  fmt(8,9,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.bg) } });
  // Currency formats for budget cards (row 6,7,9,10 for Total Budget/Spent/Remaining/Cost per Guest)
  // Primary row 1: D=Budget(col3), F=Spent(col5), H=Remaining(col7) → val row = index 7
  // Primary row 2: B=CostPerGuest(col1) → val row = index 10
  reqs.push({ repeatCell:{ range:gridRange(DB,7,8,3,4), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } }); // Total Budget
  reqs.push({ repeatCell:{ range:gridRange(DB,7,8,5,6), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } }); // Amount Spent
  reqs.push({ repeatCell:{ range:gridRange(DB,7,8,7,8), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } }); // Remaining
  reqs.push({ repeatCell:{ range:gridRange(DB,10,11,1,2), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } }); // Cost/guest
  reqs.push({ repeatCell:{ range:gridRange(DB,10,11,7,8), cell:{ userEnteredFormat:{ numberFormat:pctF } }, fields:'userEnteredFormat.numberFormat' } }); // Planning %
  reqs.push({ updateBorders:{ range:gridRange(DB,6,11,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Secondary insight section header row 11 (index 11)
  reqs.push({ mergeCells:{ range:gridRange(DB,11,12,0,8), mergeType:'MERGE_ALL' } });
  fmt(11,12,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.mutedBlue), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  // Secondary cards rows 12-13 and 15-16 (with separator 14)
  const secLayout = [
    { labelRow:12, valRow:13 },
    { labelRow:15, valRow:16 },
  ];
  secLayout.forEach(({ labelRow, valRow }) => {
    for (let i = 0; i < 4; i++) {
      const c = i * 2;
      reqs.push({ mergeCells:{ range:gridRange(DB,labelRow,labelRow+1,c,c+2), mergeType:'MERGE_ALL' } });
      reqs.push({ mergeCells:{ range:gridRange(DB,valRow,valRow+1,c,c+2), mergeType:'MERGE_ALL' } });
      fmt(labelRow,labelRow+1,c,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.mutedBlue), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.mainText) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
      fmt(valRow,valRow+1,c,c+2,{ userEnteredFormat:{ backgroundColor:hex(C.panel), textFormat:{ bold:true, fontSize:14, fontFamily:'Arial', foregroundColor:hex(C.secondary) }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
    }
  });
  fmt(14,15,0,8,{ userEnteredFormat:{ backgroundColor:hex(C.bg) } });
  reqs.push({ updateBorders:{ range:gridRange(DB,12,17,0,8), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // CF: Overdue tasks >0 → red
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(DB,16,17,7,8)], booleanRule:{ condition:{ type:'NUMBER_GREATER', values:[{ userEnteredValue:'0' }] }, format:{ textFormat:{ foregroundColor:hex(C.attention), bold:true } } } }, index:0 } });
  // Follow-ups >0 → amber
  reqs.push({ addConditionalFormatRule:{ rule:{ ranges:[gridRange(DB,13,14,5,6)], booleanRule:{ condition:{ type:'NUMBER_GREATER', values:[{ userEnteredValue:'0' }] }, format:{ textFormat:{ foregroundColor:hex(C.warning), bold:true } } } }, index:0 } });

  // Guest snapshot section row 19 (index 19)
  reqs.push({ mergeCells:{ range:gridRange(DB,19,20,0,2), mergeType:'MERGE_ALL' } });
  fmt(19,20,0,2,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(20,21,0,2,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 21; r < 27; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,2,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(DB,21,27,1,2), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(DB,20,27,0,2), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Upcoming actions section row 29 (index 29)
  reqs.push({ mergeCells:{ range:gridRange(DB,29,30,0,5), mergeType:'MERGE_ALL' } });
  fmt(29,30,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(30,31,0,5,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 31; r < 42; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,5,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  for (const ci of [0,1,2,3,4]) {
    reqs.push({ repeatCell:{ range:gridRange(DB,31,42,ci,ci+1), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 } } }, fields:'userEnteredFormat' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(DB,31,42,0,1), cell:{ userEnteredFormat:{ numberFormat:dateF, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(DB,30,42,0,5), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Food & Shopping snapshot row 44 (index 44)
  reqs.push({ mergeCells:{ range:gridRange(DB,44,45,0,2), mergeType:'MERGE_ALL' } });
  fmt(44,45,0,2,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(45,46,0,2,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 46; r < 53; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,2,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  reqs.push({ repeatCell:{ range:gridRange(DB,46,53,1,2), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula), textFormat:{ foregroundColor:hex(C.secText), fontSize:9 }, horizontalAlignment:'CENTER' } }, fields:'userEnteredFormat' } });
  // Currency for food cost (row 47 = index 47)
  reqs.push({ repeatCell:{ range:gridRange(DB,47,48,1,2), cell:{ userEnteredFormat:{ numberFormat:currF } }, fields:'userEnteredFormat.numberFormat' } });
  reqs.push({ updateBorders:{ range:gridRange(DB,45,53,0,2), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Progress weights table row 54 (index 54)
  reqs.push({ mergeCells:{ range:gridRange(DB,54,55,0,3), mergeType:'MERGE_ALL' } });
  fmt(54,55,0,3,{ userEnteredFormat:{ backgroundColor:hex(C.secondary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:11, fontFamily:'Arial' }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE' } });
  fmt(55,56,0,3,{ userEnteredFormat:{ backgroundColor:hex(C.primary), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9, fontFamily:'Arial' }, horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' } });
  for (let r = 56; r < 62; r++) {
    const bg = r%2===0 ? C.panel : C.bg;
    fmt(r,r+1,0,3,{ userEnteredFormat:{ backgroundColor:hex(bg), textFormat:{ fontSize:9, fontFamily:'Arial' }, verticalAlignment:'MIDDLE' } });
  }
  // Disclaimer row 62 (index 62)
  reqs.push({ mergeCells:{ range:gridRange(DB,62,63,0,9), mergeType:'MERGE_ALL' } });
  fmt(62,63,0,9,{ userEnteredFormat:{ backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true, fontSize:8, fontFamily:'Arial', italic:true }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP' } });
  reqs.push({ updateBorders:{ range:gridRange(DB,55,62,0,3), top:medium, bottom:medium, left:medium, right:medium, innerHorizontal:thin, innerVertical:thin } });

  // Column widths
  [[0,150],[1,130],[2,150],[3,130],[4,130],[5,130],[6,130],[7,130],[8,20],[9,120],[10,120],[11,20],[12,20],[13,20]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'COLUMNS', startIndex:ci, endIndex:ci+1 }, properties:{ pixelSize:w }, fields:'pixelSize' } });
  });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:0, endIndex:2 }, properties:{ pixelSize:48 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:5, endIndex:6 }, properties:{ pixelSize:28 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:6, endIndex:8 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:7, endIndex:8 }, properties:{ pixelSize:48 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:8, endIndex:9 }, properties:{ pixelSize:10 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:9, endIndex:11 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:10, endIndex:11 }, properties:{ pixelSize:48 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:11, endIndex:70 }, properties:{ pixelSize:24 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:13, endIndex:14 }, properties:{ pixelSize:48 }, fields:'pixelSize' } });
  reqs.push({ updateDimensionProperties:{ range:{ sheetId:DB, dimension:'ROWS', startIndex:16, endIndex:17 }, properties:{ pixelSize:48 }, fields:'pixelSize' } });

  // Freeze rows 1-5 (frozenRowCount:5)
  reqs.push({ updateSheetProperties:{ properties:{ sheetId:DB, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' } });

  // ── Charts ──
  // Chart 1: Donut — RSVP Status (pivot in J6:K11)
  reqs.push({ addChart:{ chart:{ spec:{
    title:'RSVP Status',
    pieChart:{
      legendPosition:'LABELED_LEGEND',
      domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:5, endRowIndex:12, startColumnIndex:9, endColumnIndex:10 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:5, endRowIndex:12, startColumnIndex:10, endColumnIndex:11 }] } },
      pieHole:0.45
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:19, columnIndex:9 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:260 } } } } });

  // Chart 2: Donut — Spending by Category
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Actual Spending by Category',
    pieChart:{
      legendPosition:'LABELED_LEGEND',
      domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:12, endRowIndex:28, startColumnIndex:9, endColumnIndex:10 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:12, endRowIndex:28, startColumnIndex:10, endColumnIndex:11 }] } },
      pieHole:0.45
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:19, columnIndex:11 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:260 } } } } });

  // Chart 3: Column — Planned vs Actual Budget
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Planned vs Actual Budget',
    basicChart:{
      chartType:'COLUMN',
      legendPosition:'BOTTOM_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Category' },
        { position:'LEFT_AXIS', title:'Amount ($)' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:29, endRowIndex:45, startColumnIndex:9, endColumnIndex:10 }] } } }],
      series:[
        { series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:29, endRowIndex:45, startColumnIndex:10, endColumnIndex:11 }] } }, targetAxis:'LEFT_AXIS' },
        { series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:29, endRowIndex:45, startColumnIndex:11, endColumnIndex:12 }] } }, targetAxis:'LEFT_AXIS' },
      ],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:29, columnIndex:9 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:280 } } } } });

  // Chart 4: Bar — Prep Tasks by Status
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Prep Tasks by Status',
    basicChart:{
      chartType:'BAR',
      legendPosition:'NO_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Count' },
        { position:'LEFT_AXIS', title:'Status' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:46, endRowIndex:53, startColumnIndex:9, endColumnIndex:10 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:46, endRowIndex:53, startColumnIndex:10, endColumnIndex:11 }] } }, targetAxis:'BOTTOM_AXIS' }],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:29, columnIndex:11 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:200 } } } } });

  // Chart 5: Donut — Guests by Child/Adult
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Confirmed Guests by Type',
    pieChart:{
      legendPosition:'LABELED_LEGEND',
      domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:53, endRowIndex:57, startColumnIndex:9, endColumnIndex:10 }] } },
      series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:53, endRowIndex:57, startColumnIndex:10, endColumnIndex:11 }] } },
      pieHole:0.45
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:38, columnIndex:9 }, offsetXPixels:0, offsetYPixels:0, widthPixels:260, heightPixels:200 } } } } });

  // Chart 6: Bar — Dietary Restrictions Count
  reqs.push({ addChart:{ chart:{ spec:{
    title:'Dietary Restrictions (Confirmed Guests)',
    basicChart:{
      chartType:'BAR',
      legendPosition:'NO_LEGEND',
      axis:[
        { position:'BOTTOM_AXIS', title:'Guests' },
        { position:'LEFT_AXIS', title:'Restriction' }
      ],
      domains:[{ domain:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:57, endRowIndex:66, startColumnIndex:9, endColumnIndex:10 }] } } }],
      series:[{ series:{ sourceRange:{ sources:[{ sheetId:DB, startRowIndex:57, endRowIndex:66, startColumnIndex:10, endColumnIndex:11 }] } }, targetAxis:'BOTTOM_AXIS' }],
      headerCount:1
    }
  }, position:{ overlayPosition:{ anchorCell:{ sheetId:DB, rowIndex:38, columnIndex:11 }, offsetXPixels:0, offsetYPixels:0, widthPixels:380, heightPixels:200 } } } } });

  await batchUpdate(id, reqs, 'dashboard-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
