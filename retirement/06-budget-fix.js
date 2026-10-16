'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const S = "'Budget & Expenses'";
const G = "'Guest List & RSVP'";
const R = "'Reference Data'";

const CATEGORIES = [
  ['Venue', 900],
  ['Invitations', 80],
  ['Decorations', 350],
  ['Food', 800],
  ['Drinks', 400],
  ['Cake & Desserts', 200],
  ['Entertainment', 300],
  ['Music', 150],
  ['Audio / Visual', 200],
  ['Gifts', 250],
  ['Awards', 100],
  ['Photography', 300],
  ['Supplies', 150],
  ['Transportation', 0],
  ['Setup & Cleanup', 100],
  ['Miscellaneous', 220],
];

const EXPENSES = [
  ['May 1, 2026','Venue','Lakeside Event Hall','Venue rental fee',1,900,900,'Deposit Paid','Jun 15, 2026','May 1, 2026',true,'50% deposit paid; balance due Jun 15'],
  ['May 3, 2026','Invitations','PrintSmart Co.','Premium card invitations - 120 pack',120,0.55,66,'Paid','May 3, 2026','May 3, 2026',true,'Includes envelopes'],
  ['May 3, 2026','Invitations','PrintSmart Co.','Postage stamps',50,0.68,34,'Paid','May 3, 2026','May 3, 2026',true,''],
  ['May 5, 2026','Decorations','Party Palace','Balloon arch kit',1,120,120,'Paid','May 5, 2026','May 5, 2026',true,'Deep indigo and champagne colors'],
  ['May 5, 2026','Decorations','Party Palace','Table centerpieces x12',12,18,216,'Paid','May 5, 2026','May 5, 2026',true,''],
  ['May 6, 2026','Decorations','Banner Express','Custom retirement banner - 6ft',1,75,75,'Paid','May 6, 2026','May 6, 2026',true,''],
  ['May 10, 2026','Food','Lakeside Catering','Plated dinner - 80 guests',80,18,1440,'Deposit Paid','Jun 15, 2026','May 10, 2026',true,'30% deposit; balance at delivery'],
  ['May 10, 2026','Food','Lakeside Catering','Vegetarian option uplift - 8 guests',8,5,40,'Planned','Jun 15, 2026','',false,''],
  ['May 12, 2026','Drinks','Wine & Spirits Co.','Champagne - 6 bottles',6,28,168,'Paid','May 12, 2026','May 12, 2026',true,'For toasts'],
  ['May 12, 2026','Drinks','Wine & Spirits Co.','Red & white wine - 12 bottles',12,15,180,'Paid','May 12, 2026','May 12, 2026',true,''],
  ['May 12, 2026','Drinks','Beverage Depot','Sparkling water and sodas - assorted',1,85,85,'Paid','May 12, 2026','May 12, 2026',true,''],
  ['May 14, 2026','Cake & Desserts','Sweet Endings Bakery','3-tier retirement cake',1,250,250,'Deposit Paid','Jun 20, 2026','May 14, 2026',true,'Balance due on delivery'],
  ['May 14, 2026','Cake & Desserts','Sweet Endings Bakery','Dessert table items',1,95,95,'Planned','Jun 20, 2026','',false,''],
  ['May 15, 2026','Entertainment','Riverside AV','Sound system rental',1,350,350,'Deposit Paid','Jun 25, 2026','May 15, 2026',true,''],
  ['May 15, 2026','Audio / Visual','Riverside AV','Projector and screen rental',1,200,200,'Deposit Paid','Jun 25, 2026','May 15, 2026',true,'Package deal with sound'],
  ['May 18, 2026','Music','DJ Smooth','DJ services - 3.5 hours',1,450,450,'Deposit Paid','Jun 25, 2026','May 18, 2026',true,'Includes playlist consultation'],
  ['May 20, 2026','Gifts','Trophy & Award Co.','Personalized crystal award',1,120,120,'Paid','May 20, 2026','May 20, 2026',true,'Engraved with name and years'],
  ['May 20, 2026','Gifts','Memory Lane Prints','Custom photo book',1,85,85,'Paid','May 20, 2026','May 20, 2026',true,'32 years photo collection'],
  ['May 20, 2026','Awards','Trophy & Award Co.','Retirement certificate - custom framed',1,75,75,'Paid','May 20, 2026','May 20, 2026',true,''],
  ['May 22, 2026','Photography','Riverside Photos','Event photography - 4 hours',1,400,400,'Deposit Paid','Jun 25, 2026','May 22, 2026',true,'Includes edited digital photos'],
  ['May 25, 2026','Supplies','Party Palace','Paper goods - plates, napkins, cups',1,95,95,'Paid','May 25, 2026','May 25, 2026',true,''],
  ['May 25, 2026','Supplies','Office Depot','Name tags, signage',1,35,35,'Paid','May 25, 2026','May 25, 2026',true,''],
  ['May 28, 2026','Decorations','Party Palace','Photo booth backdrop and props',1,150,150,'Paid','May 28, 2026','May 28, 2026',true,''],
  ['Jun 1, 2026','Setup & Cleanup','Event Services Inc.','Setup crew - 2 hours',2,45,90,'Planned','Jun 27, 2026','',false,'Day-of setup'],
  ['Jun 1, 2026','Setup & Cleanup','Event Services Inc.','Cleanup crew - 2 hours',2,35,70,'Planned','Jun 27, 2026','',false,'Post-party cleanup'],
  ['Jun 1, 2026','Miscellaneous','Michaels','Guest book and memory jar',1,45,45,'Paid','Jun 1, 2026','Jun 1, 2026',true,''],
  ['Jun 1, 2026','Miscellaneous','Amazon','Table numbers and place cards',1,28,28,'Paid','Jun 1, 2026','Jun 1, 2026',true,''],
  ['Jun 5, 2026','Food','Lakeside Catering','Appetizer add-on - 80 guests',80,4,320,'Planned','Jun 15, 2026','',false,'Added after headcount confirmed'],
  ['Jun 5, 2026','Entertainment','Event Supplies','Photo booth rental (4-hour)',1,250,250,'Deposit Paid','Jun 20, 2026','Jun 5, 2026',true,'Includes prints'],
  ['Jun 5, 2026','Miscellaneous','Canva (online)','Menu cards and program printing',1,32,32,'Paid','Jun 5, 2026','Jun 5, 2026',true,''],
  ['Jun 8, 2026','Gifts','Memory Lane Prints','Memory box for messages',1,55,55,'Paid','Jun 8, 2026','Jun 8, 2026',true,'Guests write retirement wishes'],
  ['Jun 8, 2026','Supplies','Party Palace','Table linens - 12 tables',12,12,144,'Planned','Jun 20, 2026','',false,''],
];

(async () => {
  const data = [];

  // ── Title + Subtitle ────────────────────────────────────────────────────────
  data.push({ range:`${S}!A1`, values:[['BUDGET & EXPENSES']] });
  data.push({ range:`${S}!A2`, values:[['Track planned and actual costs by category — budget status updates automatically']] });
  data.push({ range:`${S}!A3`, values:[['Actual Spent by category calculates automatically from the expense log below (row 31+). Estimated Cost = Qty × Unit Cost.']] });
  data.push({ range:`${S}!A4`, values:[['Log expenses in rows 31+. Enter Category, Item, Qty, Unit Cost, and Actual Cost. Payment Status dropdown controls budget tracking.']] });

  // ── Summary Cards ───────────────────────────────────────────────────────────
  const cardPositions = [
    [4, 0, 3, 3, 7],
    [4, 7, 10, 10, 14],
    [5, 0, 3, 3, 7],
    [5, 7, 10, 10, 14],
    [7, 0, 3, 3, 7],
    [7, 7, 10, 10, 14],
    [8, 0, 3, 3, 7],
    [8, 7, 10, 10, 14],
  ];
  const summaryCards = [
    { label:'Total Budget',            formula:`=IFERROR('Party Setup'!B18,0)` },
    { label:'Planned Spending',        formula:`=IFERROR(SUM(B12:B27),0)` },
    { label:'Actual Spending',         formula:`=IFERROR(SUM(I31:I329),0)` },
    { label:'Remaining Budget',        formula:`=IFERROR('Party Setup'!B18-SUM(I31:I329),0)` },
    { label:'Amount Over / Under',     formula:`=IFERROR(SUM(B12:B27)-SUM(I31:I329),0)` },
    { label:'Cost per Confirmed Guest',formula:`=IFERROR(IF(COUNTIF(${G}!$J$6:$J$5000,"Attending")=0,"—",SUM(I31:I329)/COUNTIF(${G}!$J$6:$J$5000,"Attending")),"—")` },
    { label:'Paid Expenses',           formula:`=IFERROR(SUMIF(J31:J329,"Paid",I31:I329),0)` },
    { label:'Unpaid / Planned',        formula:`=IFERROR(SUMIF(J31:J329,"Planned",I31:I329)+SUMIF(J31:J329,"Deposit Paid",I31:I329),0)` },
  ];
  summaryCards.forEach((card, i) => {
    const [ri, lc1,, vc1] = cardPositions[i];
    const sheetRow = ri + 1;
    const labelCell = String.fromCharCode(65 + lc1) + sheetRow;
    const valueCell = String.fromCharCode(65 + vc1) + sheetRow;
    data.push({ range:`${S}!${labelCell}`, values:[[card.label]] });
    data.push({ range:`${S}!${valueCell}`, values:[[card.formula]] });
  });

  // ── Category Budget Table ────────────────────────────────────────────────────
  data.push({ range:`${S}!A10`, values:[['CATEGORY BUDGET — Planned vs. Actual']] });
  data.push({ range:`${S}!A11`, values:[['Expense Category','Planned Budget','Actual Spent','Remaining','Variance %','Status']] });

  CATEGORIES.forEach(([cat, planned], i) => {
    const row = 12 + i;
    data.push({ range:`${S}!A${row}`, values:[[cat]] });
    data.push({ range:`${S}!B${row}`, values:[[planned]] });
    data.push({ range:`${S}!C${row}`, values:[[`=IFERROR(SUMIF($C$31:$C$329,"${cat}",$I$31:$I$329),0)`]] });
    data.push({ range:`${S}!D${row}`, values:[[`=IFERROR(B${row}-C${row},0)`]] });
    data.push({ range:`${S}!E${row}`, values:[[`=IFERROR(IF(B${row}=0,"—",(C${row}-B${row})/B${row}),"—")`]] });
    data.push({ range:`${S}!F${row}`, values:[[`=IFERROR(IF(B${row}=0,"No Budget Set",IF(C${row}>B${row},"Over Budget",IF(C${row}/B${row}>0.85,"Near Budget","Under Budget"))),"—")`]] });
  });

  // Budget totals row 28
  data.push({ range:`${S}!A28`, values:[['TOTALS']] });
  data.push({ range:`${S}!B28`, values:[[`=SUM(B12:B27)`]] });
  data.push({ range:`${S}!C28`, values:[[`=SUM(C12:C27)`]] });
  data.push({ range:`${S}!D28`, values:[[`=SUM(D12:D27)`]] });

  // ── Pivot data for charts ────────────────────────────────────────────────────
  data.push({ range:`${S}!H11`, values:[['Category']] });
  data.push({ range:`${S}!I11`, values:[['Actual Spent']] });
  CATEGORIES.forEach(([cat], i) => {
    const row = 12 + i;
    data.push({ range:`${S}!H${row}`, values:[[cat]] });
    data.push({ range:`${S}!I${row}`, values:[[`=IFERROR(SUMIF($C$31:$C$329,"${cat}",$I$31:$I$329),0)`]] });
  });

  data.push({ range:`${S}!K11`, values:[['Payment Status']] });
  data.push({ range:`${S}!L11`, values:[['Amount']] });
  [
    ['Paid',           `=IFERROR(SUMIF($J$31:$J$329,"Paid",$I$31:$I$329),0)`],
    ['Deposit Paid',   `=IFERROR(SUMIF($J$31:$J$329,"Deposit Paid",$I$31:$I$329),0)`],
    ['Planned',        `=IFERROR(SUMIF($J$31:$J$329,"Planned",$I$31:$I$329),0)`],
    ['Partially Paid', `=IFERROR(SUMIF($J$31:$J$329,"Partially Paid",$I$31:$I$329),0)`],
  ].forEach(([label, formula], i) => {
    data.push({ range:`${S}!K${12+i}`, values:[[label]] });
    data.push({ range:`${S}!L${12+i}`, values:[[formula]] });
  });

  // ── Expense Log section header ───────────────────────────────────────────────
  data.push({ range:`${S}!A29`, values:[['EXPENSE LOG']] });
  data.push({ range:`${S}!A30`, values:[['Expense ID','Date','Category','Vendor / Store','Item / Service','Quantity','Unit Cost','Estimated Cost','Actual Cost','Payment Status','Due Date','Paid Date','Receipt Saved?','Notes']] });

  // ── Expense data rows 31-62 ──────────────────────────────────────────────────
  EXPENSES.forEach(([date, cat, vendor, item, qty, unitCost, actualCost, payStatus, dueDate, paidDate, receipt, notes], i) => {
    const row = 31 + i;
    data.push({ range:`${S}!A${row}`, values:[[`=IF(E${row}="","","EXP-"&TEXT(ROW()-30,"000"))`]] });
    data.push({ range:`${S}!B${row}`, values:[[date]] });
    data.push({ range:`${S}!C${row}`, values:[[cat]] });
    data.push({ range:`${S}!D${row}`, values:[[vendor]] });
    data.push({ range:`${S}!E${row}`, values:[[item]] });
    data.push({ range:`${S}!F${row}`, values:[[qty]] });
    data.push({ range:`${S}!G${row}`, values:[[unitCost]] });
    data.push({ range:`${S}!H${row}`, values:[[`=IFERROR(IF(F${row}="","",F${row}*G${row}),"")`]] });
    data.push({ range:`${S}!I${row}`, values:[[actualCost]] });
    data.push({ range:`${S}!J${row}`, values:[[payStatus]] });
    if (dueDate)  data.push({ range:`${S}!K${row}`, values:[[dueDate]] });
    if (paidDate) data.push({ range:`${S}!L${row}`, values:[[paidDate]] });
    data.push({ range:`${S}!M${row}`, values:[[receipt]] });
    data.push({ range:`${S}!N${row}`, values:[[notes]] });
  });

  // ── Auto-formulas for blank rows 63-329 ─────────────────────────────────────
  for (let row = 63; row <= 329; row++) {
    data.push({ range:`${S}!A${row}`, values:[[`=IF(E${row}="","","EXP-"&TEXT(ROW()-30,"000"))`]] });
    data.push({ range:`${S}!H${row}`, values:[[`=IFERROR(IF(F${row}="","",F${row}*G${row}),"")`]] });
  }

  await valuesBatchUpdate(id, data);
  console.log('Budget & Expenses values written successfully');
})();
