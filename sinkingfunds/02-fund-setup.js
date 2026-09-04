'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Fund Setup & Goals'];
const S = "'Fund Setup & Goals'";
const CL = "'Contribution Log'";
const REF = "'Reference Data'";
const NC = 26; // A-Z

// Col mapping:
// A(0)=Fund ID  B(1)=Fund Name  C(2)=Fund Category  D(3)=Owner  E(4)=Priority
// F(5)=Funding Method  G(6)=Goal Amount  H(7)=Starting Balance  I(8)=Current Balance(fml)
// J(9)=Amount Remaining(fml)  K(10)=Progress %(fml)  L(11)=Target Date
// M(12)=Months Remaining(fml)  N(13)=Paychecks Remaining(fml)
// O(14)=Goal-Date Monthly Needed(fml)  P(15)=Goal-Date Paycheck Needed(fml)
// Q(16)=Fixed Monthly Amount  R(17)=Fixed Paycheck Amount
// S(18)=Selected Planned Monthly Funding(fml)  T(19)=Estimated Goal Date(fml)
// U(20)=Days Until Goal(fml)  V(21)=Status(fml)  W(22)=Funding Order(fml)
// X(23)=Active?  Y(24)=Goal Reached Date  Z(25)=Notes

// Fund helper: [name, cat, owner, priority, method, goal, targetDate, fixedMonthly, fixedPaycheck, active, notes]
const FUNDS = [
  ['Hawaii Vacation','Vacation','Joint / Household','High','Goal-Date Based',6000,'8/31/2026',null,null,true,'Maui trip for two — flights, hotel, and activities.'],
  ['Christmas Gifts 2026','Christmas','Joint / Household','High','Goal-Date Based',2000,'12/1/2026',null,null,true,'Reloading after 2025 holiday goal was met.'],
  ['Car Maintenance Reserve','Car Maintenance','Self','Medium','Fixed Monthly Amount',1200,'',100,null,true,'Oil changes, tires, and misc repairs throughout the year.'],
  ['Emergency Buffer Top-Up','Emergency Buffer','Joint / Household','Critical','Fixed Monthly Amount',5000,'',300,null,true,'Covers irregular surprises beyond main emergency fund.'],
  ['New Laptop — Marisol','Technology','Person 1','Low','Fixed Monthly Amount',1800,'12/31/2026',80,null,true,'MacBook Air replacement fund.'],
  ['Dental Work — Terrence','Dental','Person 2','High','Fixed Paycheck Amount',2500,'9/30/2026',null,150,true,'Crown and two fillings — estimated total from provider.'],
  ['Kitchen Renovation','Home Renovation','Joint / Household','Medium','Goal-Date Based',15000,'12/31/2027',null,null,true,'Full cabinet and countertop upgrade. Multi-year goal.'],
  ['Wedding Fund','Wedding','Joint / Household','Critical','Fixed Monthly Amount',25000,'10/15/2026',800,null,true,'Venue, catering, photography, florals, and rings.'],
  ['Annual Insurance Premiums','Insurance','Self','High','Fixed Monthly Amount',2400,'12/1/2026',200,null,true,'Home and auto annual policy renewals.'],
  ['Pet Medical Reserve','Pets','Self','Medium','Manual Contribution',1000,'',null,null,true,'Vet visits and unexpected pet care costs.'],
  ['Car Down Payment','New Vehicle','Joint / Household','High','Goal-Date Based',10000,'3/31/2027',null,null,true,'20% down payment on next vehicle purchase.'],
  ['Kids School Expenses','School Expenses','Joint / Household','Medium','Fixed Monthly Amount',3600,'9/1/2026',300,null,true,'Tuition, supplies, field trips, and after-school activities.'],
  ['Birthday Fund','Birthdays','Self','Low','Fixed Monthly Amount',600,'',50,null,true,'All family birthdays covered for the year. Goal reached!'],
  ['Annual Memberships','Memberships','Self','Low','Fixed Monthly Amount',480,'12/31/2026',40,null,true,'Gym, warehouse store, and streaming service bundles.'],
  ['Moving Costs — Marisol','Moving','Person 1','High','Goal-Date Based',4000,'8/31/2026',null,null,true,'Truck rental, security deposits, utility transfers.'],
  ['Self-Employment Tax Reserve','Taxes','Person 2','Critical','Fixed Paycheck Amount',6000,'4/15/2027',null,500,true,'Quarterly estimated taxes for Terrence freelance income.'],
  ['Home Maintenance Reserve','Home Maintenance','Joint / Household','Medium','Fixed Monthly Amount',3000,'',250,null,true,'1% rule savings for ongoing home upkeep.'],
  ['Living Room Furniture','Furniture','Joint / Household','Low','Manual Contribution',2800,'',null,null,false,'Sofa, accent chairs, and rug. Paused — prioritizing kitchen reno.'],
  ['Baby Essentials Fund','Baby','Joint / Household','High','Goal-Date Based',8000,'11/30/2026',null,null,true,'Nursery setup, stroller, car seat, and initial diaper supply.'],
  ['Anniversary Trip — Paris','Travel','Joint / Household','Medium','Fixed Monthly Amount',3500,'6/30/2026',300,null,true,'10th anniversary celebration trip abroad.'],
  ['New Smartphone — Marisol','Technology','Person 1','Low','Fixed Paycheck Amount',900,'',null,100,true,'iPhone upgrade fund. Goal reached!'],
  ['Business Software & Tools','Business','Person 2','Medium','Manual Contribution',2000,'',null,null,true,'Subscriptions, licenses, course access for Terrence freelance.'],
  ['Medical Deductible Fund','Medical','Joint / Household','High','Fixed Monthly Amount',3000,'',250,null,true,'Family annual deductible fully funded. Goal reached!'],
  ['Property Tax Reserve','Taxes','Joint / Household','Critical','Fixed Monthly Amount',8400,'12/31/2026',700,null,true,'Annual property tax bill — avoids lump-sum surprise.'],
  ['Annual Streaming Subscriptions','Annual Subscriptions','Self','Low','Fixed Monthly Amount',360,'',30,null,false,'Netflix, Spotify, and cloud storage bundle. Archived.'],
  ['Appliance Replacement Fund','Appliances','Joint / Household','Medium','Manual Contribution',5000,'',null,null,false,'Fridge and washer/dryer aging. Paused — revisit next year.'],
];

(async () => {
  const fmt = [];
  const vals = [];

  // ── Title ──────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['FUND SETUP & GOALS']] });

  // ── Subtitle ───────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.seafoam), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.darkText) }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Source of truth for every sinking fund. Fund IDs link all other tabs. Shaded blue cells are formula-calculated.']] });

  // ── Planning Controls (rows 3-6) ─────────────────────────────────────────
  const ctlBg = hex(C.dustyBlue);
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 4), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, 4), cell: { userEnteredFormat: { backgroundColor: ctlBg, textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.darkText) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  vals.push({ range: `${S}!A3`, values: [['PLANNING CONTROLS']] });

  const ctlLabels = [['Paycheck Frequency','Biweekly'],['Funding Order Weights','Priority 45%  |  Urgency 35%  |  Need 20%']];
  ctlLabels.forEach(([lbl, val], i) => {
    fmt.push({ repeatCell: { range: gridRange(SID, 3+i, 4+i, 0, 2), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3+i, 4+i, 2, 6), cell: { userEnteredFormat: { backgroundColor: hex(C.input), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
    vals.push({ range: `${S}!A${4+i}`, values: [[lbl]] });
    vals.push({ range: `${S}!C${4+i}`, values: [[val]] });
  });
  // Paycheck frequency dropdown
  fmt.push({ setDataValidation: { range: gridRange(SID, 3, 4, 2, 3), rule: { condition: { type: 'ONE_OF_LIST', values: ['Weekly','Biweekly','Semi-Monthly','Monthly'].map(v=>({userEnteredValue:v})) }, showCustomUi: true } } });

  // ── Summary Cards (rows 3-6, cols H-Z) ────────────────────────────────────
  const kpiDefs = [
    { label: 'Active Funds',        col: 8,  fml: `=COUNTIF($X$8:$X$507,TRUE)` },
    { label: 'Total Savings Goals', col: 11, fml: `=IFERROR(SUMIF($X$8:$X$507,TRUE,$G$8:$G$507),0)` },
    { label: 'Total Current Saved', col: 14, fml: `=IFERROR(SUMIF($X$8:$X$507,TRUE,$I$8:$I$507),0)` },
    { label: 'Total Remaining',     col: 17, fml: `=IFERROR(SUMIF($X$8:$X$507,TRUE,$J$8:$J$507),0)` },
    { label: 'Overall Progress %',  col: 20, fml: `=IFERROR(SUMIF($X$8:$X$507,TRUE,$I$8:$I$507)/SUMIF($X$8:$X$507,TRUE,$G$8:$G$507),0)` },
    { label: 'Monthly Funding Needed', col: 23, fml: `=IFERROR(SUMPRODUCT(($X$8:$X$507=TRUE)*($V$8:$V$507<>"Goal Reached")*($V$8:$V$507<>"Paused")*($O$8:$O$507)),0)` },
  ];
  kpiDefs.forEach(k => {
    const cardColor = C.primary;
    fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, k.col, k.col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 3, 5, k.col, k.col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, k.col, k.col+3), cell: { userEnteredFormat: { backgroundColor: hex(cardColor), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 3, 5, k.col, k.col+3), cell: { userEnteredFormat: { backgroundColor: hex(cardColor), textFormat: { bold: true, fontSize: 18, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${colL(k.col)}3`, values: [[k.label]] });
    vals.push({ range: `${S}!${colL(k.col)}4`, values: [[k.fml]] });
  });

  // ── Spacer row 7 ──────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 5, 6, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.bg) } }, fields: 'userEnteredFormat.backgroundColor' } });

  // ── Column headers row 7 (0-indexed row 6) ───────────────────────────────
  const HEADERS = [
    'Fund ID','Fund Name','Fund Category','Owner','Priority','Funding Method',
    'Goal Amount','Starting Balance','Current Balance','Amount Remaining','Progress %',
    'Target Date','Months Remaining','Paychecks Remaining','Goal-Date Monthly Needed',
    'Goal-Date Paycheck Needed','Fixed Monthly Amount','Fixed Paycheck Amount',
    'Selected Planned Monthly Funding','Estimated Goal Date','Days Until Goal',
    'Status','Funding Order','Active?','Goal Reached Date','Notes',
  ];
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy,verticalAlignment)' } });
  vals.push({ range: `${S}!A7`, values: [HEADERS] });

  // ── Data rows 8:507 (0-indexed 7:506) ─────────────────────────────────────
  const NDATA = 500;

  // Write sample fund rows first (data columns)
  FUNDS.forEach((fund, i) => {
    const r = 8 + i;
    const [name, cat, owner, priority, method, goal, targetDate, fixedMonthly, fixedPaycheck, active, notes] = fund;
    // Write input columns: B,C,D,E,F,G,H,L,Q,R,X,Y,Z
    vals.push({ range: `${S}!B${r}`, values: [[name]] });
    vals.push({ range: `${S}!C${r}`, values: [[cat]] });
    vals.push({ range: `${S}!D${r}`, values: [[owner]] });
    vals.push({ range: `${S}!E${r}`, values: [[priority]] });
    vals.push({ range: `${S}!F${r}`, values: [[method]] });
    vals.push({ range: `${S}!G${r}`, values: [[goal]] });
    vals.push({ range: `${S}!H${r}`, values: [[0]] }); // all starting balances = 0 (log covers full history)
    if (targetDate) vals.push({ range: `${S}!L${r}`, values: [[targetDate]] });
    if (fixedMonthly) vals.push({ range: `${S}!Q${r}`, values: [[fixedMonthly]] });
    if (fixedPaycheck) vals.push({ range: `${S}!R${r}`, values: [[fixedPaycheck]] });
    vals.push({ range: `${S}!X${r}`, values: [[active]] });
    vals.push({ range: `${S}!Z${r}`, values: [[notes]] });
  });

  // Formula columns: A, I, J, K, M, N, O, P, S, T, U, V, W
  // Fund ID
  const fundIdFmls = Array.from({ length: NDATA }, (_, i) => [`=IF(B${8+i}="","","FUND-"&TEXT(ROW()-7,"000"))`]);
  vals.push({ range: `${S}!A8:A507`, values: fundIdFmls });

  // Paycheck multiplier formula (named by row reference to planning control C4)
  const paycheckMult = `IF($C$4="Weekly",52/12,IF($C$4="Biweekly",26/12,IF($C$4="Semi-Monthly",24/12,1)))`;

  const iFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IFERROR(H${r}+SUMPRODUCT((${CL}!$E$8:$E$5007=A${r})*${CL}!$N$8:$N$5007),H${r})`]; });
  const jFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IFERROR(MAX(0,G${r}-I${r}),0)`]; });
  const kFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IFERROR(MIN(1,I${r}/G${r}),0)`]; });
  const mFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IF(L${r}="","",MAX(0,IFERROR(DATEDIF(TODAY(),L${r},"M"),0)))`]; });
  const nFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IF(M${r}="","",MAX(0,ROUND(M${r}*${paycheckMult},0)))`]; });
  const oFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IFERROR(IF(M${r}="",0,J${r}/MAX(1,M${r})),0)`]; });
  const pFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IFERROR(IF(N${r}="",0,J${r}/MAX(1,N${r})),0)`]; });
  const sFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IFERROR(IF(B${r}="","",IF(F${r}="Goal-Date Based",O${r},IF(F${r}="Fixed Monthly Amount",Q${r},IF(F${r}="Fixed Paycheck Amount",R${r}*${paycheckMult},"")))),"")` ]; });
  const tFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IFERROR(IF(B${r}="","",IF(F${r}="Goal-Date Based",IF(L${r}="","",L${r}),IF(F${r}="Manual Contribution","Manual",IF(S${r}="","",IFERROR(EDATE(TODAY(),ROUNDUP(J${r}/MAX(0.01,S${r}),0)),""))))),"")` ]; });
  const uFmls  = Array.from({ length: NDATA }, (_, i) => { const r=8+i; return [`=IF(OR(L${r}="",B${r}=""),"",L${r}-TODAY())`]; });

  // Status formula
  const vFmls = Array.from({ length: NDATA }, (_, i) => {
    const r = 8+i;
    return [`=IFERROR(IF(B${r}="","",IF(OR(X${r}=FALSE,X${r}=""),"Paused",IF(I${r}>=G${r},"Goal Reached",IF(I${r}=0,"Not Started",IF(OR(L${r}="",S${r}=""),"Active",IF(U${r}<0,"Goal Reached",IF(S${r}=0,"Active",IF(S${r}>O${r}*1.05,"Ahead of Plan",IF(S${r}>=O${r}*0.9,"On Track","Behind Plan"))))))))),"")`];
  });

  // Funding Order formula
  const wFmls = Array.from({ length: NDATA }, (_, i) => {
    const r = 8+i;
    const pri  = `IF(E${r}="Critical",4,IF(E${r}="High",3,IF(E${r}="Medium",2,IF(E${r}="Low",1,0))))`;
    const urg  = `IF(OR(L${r}="",U${r}=""),1,IF(U${r}<=30,4,IF(U${r}<=90,3,IF(U${r}<=180,2,1))))`;
    const need = `IF(K${r}>=0.75,1,IF(K${r}>=0.5,2,IF(K${r}>=0.25,3,4)))`;
    const score = `ROUND((${pri})*0.45+(${urg})*0.35+(${need})*0.20,2)`;
    return [`=IFERROR(IF(B${r}="","Insufficient Data",IF(OR(X${r}=FALSE,X${r}=""),"Paused",IF(I${r}>=G${r},"Goal Reached",IF(${score}>=3.5,"Fund First",IF(${score}>=2.5,"High Planning Priority",IF(${score}>=1.5,"Standard Planning Priority","Lower Planning Priority")))))),"Insufficient Data")`];
  });

  vals.push({ range: `${S}!A8:A507`, values: fundIdFmls });
  vals.push({ range: `${S}!I8:I507`, values: iFmls });
  vals.push({ range: `${S}!J8:J507`, values: jFmls });
  vals.push({ range: `${S}!K8:K507`, values: kFmls });
  vals.push({ range: `${S}!M8:M507`, values: mFmls });
  vals.push({ range: `${S}!N8:N507`, values: nFmls });
  vals.push({ range: `${S}!O8:O507`, values: oFmls });
  vals.push({ range: `${S}!P8:P507`, values: pFmls });
  vals.push({ range: `${S}!S8:S507`, values: sFmls });
  vals.push({ range: `${S}!T8:T507`, values: tFmls });
  vals.push({ range: `${S}!U8:U507`, values: uFmls });
  vals.push({ range: `${S}!V8:V507`, values: vFmls });
  vals.push({ range: `${S}!W8:W507`, values: wFmls });

  // ── Row styling ─────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  for (let r = 0; r < NDATA; r++) {
    if (r % 2 !== 0) fmt.push({ repeatCell: { range: gridRange(SID, 7+r, 8+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }

  // Formula column tint: A,I,J,K,M,N,O,P,S,T,U,V,W
  [0,8,9,10,12,13,14,15,18,19,20,21,22].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });
  // Input column tint
  [1,2,3,4,5,6,7,11,16,17,23,24,25].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Number formats
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 6, 10),   cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } }); // G-J
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 10, 11),  cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } }); // K
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 11, 12),  cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } }); // L
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 14, 18),  cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } }); // O-R
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 18, 19),  cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } }); // S
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 19, 20),  cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } }); // T
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 24, 25),  cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } }); // Y

  // Summary card number formats (row 3 values = row index 3)
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 5, 14, 17), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 5, 20, 23), cell: { userEnteredFormat: { numberFormat: { type: 'PERCENT', pattern: '0%' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Dropdowns
  const dv = (ci, items) => ({ setDataValidation: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), rule: { condition: { type: 'ONE_OF_LIST', values: items.map(v=>({userEnteredValue:v})) }, showCustomUi: true } } });
  fmt.push(dv(2, ['Vacation','Travel','Holidays','Christmas','Birthdays','Gifts','Car Maintenance','New Vehicle','Home Maintenance','Home Renovation','Furniture','Appliances','Medical','Dental','Insurance','Taxes','Education','School Expenses','Wedding','Baby','Pets','Technology','Annual Subscriptions','Memberships','Emergency Buffer','Moving','Business','Personal','Other']));
  fmt.push(dv(3, ['Self','Person 1','Person 2','Joint / Household']));
  fmt.push(dv(4, ['Low','Medium','High','Critical']));
  fmt.push(dv(5, ['Goal-Date Based','Fixed Monthly Amount','Fixed Paycheck Amount','Manual Contribution']));
  fmt.push({ setDataValidation: { range: gridRange(SID, 7, 7+NDATA, 23, 24), rule: { condition: { type: 'BOOLEAN' }, strict: true } } });

  // ── Freeze rows 1-7, cols A-D ─────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' } });

  // ── Row heights ─────────────────────────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 5 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 6 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 7+NDATA }, properties: { pixelSize: 20 }, fields: 'pixelSize' } });

  // ── Column widths ─────────────────────────────────────────────────────────
  const WIDTHS = [70,180,100,90,70,120,80,80,80,80,60,80,60,60,80,80,70,70,80,80,60,90,100,50,80,200];
  WIDTHS.forEach((w,i) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, fmt, 'fs-fmt');
  await valuesBatchUpdate(id, vals, 'fs-vals');
  console.log(`✓ Fund Setup & Goals complete — ${FUNDS.length} funds written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
