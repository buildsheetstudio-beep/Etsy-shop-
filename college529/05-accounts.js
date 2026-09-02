'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['529 Accounts'];
const S   = "'529 Accounts'";
const CL  = "'Contribution Log'";
const BS  = "'Beneficiary Setup'";
const NC  = 26;

const ACCOUNTS = [
  { id:'ACC-001', benId:'BEN-001', name:'Hartley — ScholarShare 529',        owner:'Parent / Guardian', inst:'ScholarShare 529 (CA)',    type:'529 Savings Plan', status:'Active', state:'CA', num:'SCS-0001', open:'1/15/2020', strat:'Aggressive Growth',  init:5000,  goal:3600 },
  { id:'ACC-002', benId:'BEN-002', name:'Hartley — Vanguard 529',            owner:'Parent / Guardian', inst:'Vanguard 529 Plan (NV)',   type:'529 Savings Plan', status:'Active', state:'NV', num:'VAN-0002', open:'3/1/2021',  strat:'Moderate Growth',   init:8000,  goal:4800 },
  { id:'ACC-003', benId:'BEN-003', name:'Delgado Education Fund',            owner:'Parent / Guardian', inst:'Fidelity UNIQUE 529 (NH)', type:'529 Savings Plan', status:'Active', state:'NH', num:'FID-0003', open:'9/1/2019',  strat:'Moderate',          init:12000, goal:4200 },
  { id:'ACC-004', benId:'BEN-004', name:'Washington — CollegeAdvantage',     owner:'Parent / Guardian', inst:'CollegeAdvantage (OH)',    type:'529 Savings Plan', status:'Active', state:'OH', num:'COH-0004', open:'6/1/2017',  strat:'Conservative',      init:45000, goal:4200 },
  { id:'ACC-005', benId:'BEN-005', name:'Beaumont — my529 Primary',          owner:'Grandparent',       inst:'my529 (UT)',               type:'529 Savings Plan', status:'Active', state:'UT', num:'MY5-0005', open:'11/1/2020', strat:'Moderate Growth',   init:15000, goal:4200 },
  { id:'ACC-006', benId:'BEN-001', name:'Chen Gift Account — ScholarShare',  owner:'Grandparent',       inst:'ScholarShare 529 (CA)',    type:'529 Savings Plan', status:'Active', state:'CA', num:'SCS-0006', open:'1/1/2022',  strat:'Balanced',          init:3000,  goal:2000 },
  { id:'ACC-007', benId:'BEN-002', name:'College Savings Iowa — Lucas',      owner:'Grandparent',       inst:'College Savings Iowa',    type:'529 Savings Plan', status:'Active', state:'IA', num:'CSI-0007', open:'7/1/2022',  strat:'Balanced',          init:4000,  goal:2000 },
  { id:'ACC-008', benId:'BEN-003', name:'Delgado — DreamAhead (Closed)',     owner:'Parent / Guardian', inst:'DreamAhead (WA)',          type:'529 Savings Plan', status:'Closed', state:'WA', num:'DAH-0008', open:'1/1/2023',  strat:'Moderate',          init:2000,  goal:0 },
  { id:'ACC-009', benId:'BEN-004', name:'Washington — MOST 529',             owner:'Parent / Guardian', inst:'MOST 529 (MO)',            type:'529 Savings Plan', status:'Active', state:'MO', num:'MOS-0009', open:'8/1/2016',  strat:'Conservative',      init:50000, goal:4200 },
  { id:'ACC-010', benId:'BEN-005', name:'Beaumont — NY 529 Secondary',       owner:'Parent / Guardian', inst:'NY 529 Direct Plan',      type:'529 Savings Plan', status:'Active', state:'NY', num:'NY5-0010', open:'6/1/2021',  strat:'Moderate Growth',   init:10000, goal:4200 },
];

const SCEN = `{"Conservative","Base","Higher Growth"}`;
const RATES = `0.04,0.06,0.08`;

(async () => {
  const vals = [];
  const fmt  = [];

  fmt.push({ repeatCell:{ range:gridRange(SID,0,1100,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.bg), textFormat:{ fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range:`${S}!A1`, values:[['529 Accounts']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,0,1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,0,1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.primary), textFormat:{ bold:true, fontSize:16, foregroundColor:hex(C.white) },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:42 }, fields:'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range:`${S}!A2`, values:[["Track all 529 and education savings accounts. Balances pull live from the Contribution Log. Yellow = editable. Blue-gray = formula."]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,1,2,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,1,2,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.aubergTint), textFormat:{ fontSize:9, foregroundColor:hex(C.text), italic:true },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // Rows 3-4: Summary Cards (4 cards, each 6-7 cols across 26)
  const CARDS = [
    { label:'Total Portfolio Value',   val:`=IFERROR(SUM(N8:N1007),"")`,                                                            cur:true,  pct:false },
    { label:'Active Accounts',         val:`=COUNTIF(G8:G1007,"Active")`,                                                            cur:false, pct:false },
    { label:'YTD Contributions',       val:`=IFERROR(SUM(R8:R1007),"")`,                                                            cur:true,  pct:false },
    { label:'% Accounts Fully Funded', val:`=IFERROR(COUNTIFS(G8:G1007,"Active",V8:V1007,">=1")/COUNTIF(G8:G1007,"Active"),"")`,   cur:false, pct:true  },
  ];
  // Approx col spans: 0-6, 6-13, 13-19, 19-26
  const cardSpans = [[0,6],[6,13],[13,19],[19,26]];
  CARDS.forEach(({ label, val, cur, pct }, i) => {
    const [cs,ce] = cardSpans[i];
    const cl = String.fromCharCode(65+cs);
    vals.push({ range:`${S}!${cl}3`, values:[[label]] });
    vals.push({ range:`${S}!${cl}4`, values:[[val]] });
    fmt.push({ mergeCells:{ range:gridRange(SID,2,3,cs,ce), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range:gridRange(SID,2,3,cs,ce), cell:{ userEnteredFormat:{
      backgroundColor:hex(C.eucalTint), textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.secText) },
      horizontalAlignment:'CENTER'
    }}, fields:'userEnteredFormat' }});
    fmt.push({ mergeCells:{ range:gridRange(SID,3,4,cs,ce), mergeType:'MERGE_ALL' }});
    const vf = { backgroundColor:hex(C.white), textFormat:{ bold:true, fontSize:20, foregroundColor:hex(C.primary) },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' };
    if (cur) vf.numberFormat = { type:'CURRENCY', pattern:'"$"#,##0' };
    if (pct) vf.numberFormat = { type:'PERCENT', pattern:'0.0%' };
    fmt.push({ repeatCell:{ range:gridRange(SID,3,4,cs,ce), cell:{ userEnteredFormat:vf }, fields:'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:36 }, fields:'pixelSize' }});

  // Row 6: section note
  vals.push({ range:`${S}!A6`, values:[["Account Register — 10 pre-loaded accounts. Add new rows as needed."]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,5,6,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,5,6,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrLight), textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.primary) },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // Row 7: Column headers
  vals.push({ range:`${S}!A7`, values:[[
    'Account ID','Beneficiary ID','Account Name','Account Owner',
    'Institution / Provider','Account Type','Account Status','State Plan',
    'Account Number','Open Date','Investment Strategy','Initial Deposit',
    'Annual Contribution Goal','Current Balance','Total Contributions',
    'Total Withdrawals','Total Earnings (Est.)','YTD Contributions',
    'YTD Withdrawals','Last Contribution Date','Projected Balance',
    'Goal % Funded','Tax Advantaged?','Beneficiary Name',
    'Years Until Need','Notes'
  ]] });
  fmt.push({ repeatCell:{ range:gridRange(SID,6,7,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:6, endIndex:7 }, properties:{ pixelSize:40 }, fields:'pixelSize' }});

  // Rows 8-1007: Data (10 seeded + 990 formula rows)
  const dataRows = [];
  for (let i = 0; i < 1000; i++) {
    const a  = ACCOUNTS[i] || null;
    const r  = i + 8;
    const ai = a ? a.id : '';
    const bi = a ? a.benId : '';

    const nF = `=IFERROR(SUMPRODUCT((${CL}!$D$6:$D$5005=$A${r})*(${CL}!$F$6:$F$5005<>"Withdrawal")*(${CL}!$H$6:$H$5005))-SUMPRODUCT((${CL}!$D$6:$D$5005=$A${r})*(${CL}!$F$6:$F$5005="Withdrawal")*(${CL}!$H$6:$H$5005)),0)`;
    const oF = `=IFERROR(SUMPRODUCT((${CL}!$D$6:$D$5005=$A${r})*(${CL}!$F$6:$F$5005="Contribution")*(${CL}!$H$6:$H$5005)),0)`;
    const pF = `=IFERROR(SUMPRODUCT((${CL}!$D$6:$D$5005=$A${r})*(${CL}!$F$6:$F$5005="Withdrawal")*(${CL}!$H$6:$H$5005)),0)`;
    const qF = `=IFERROR(IF(A${r}="","",N${r}-L${r}-O${r}+P${r}),0)`;
    const rF = `=IFERROR(SUMPRODUCT((YEAR(${CL}!$B$6:$B$5005)=YEAR(TODAY()))*(${CL}!$D$6:$D$5005=$A${r})*(${CL}!$F$6:$F$5005="Contribution")*(${CL}!$H$6:$H$5005)),0)`;
    const sF = `=IFERROR(SUMPRODUCT((YEAR(${CL}!$B$6:$B$5005)=YEAR(TODAY()))*(${CL}!$D$6:$D$5005=$A${r})*(${CL}!$F$6:$F$5005="Withdrawal")*(${CL}!$H$6:$H$5005)),0)`;
    const tF = `=IFERROR(TEXT(MAXIFS(${CL}!$B$6:$B$5005,${CL}!$D$6:$D$5005,$A${r},${CL}!$F$6:$F$5005,"Contribution"),"MM/DD/YYYY"),"")`;
    const uF = `=IFERROR(IF(A${r}="","",N${r}*(1+CHOOSE(MATCH(${BS}!$J$6,${SCEN},0),${RATES}))^Y${r}),"")`;
    const vF = `=IFERROR(IF(A${r}="","",N${r}/VLOOKUP($B${r},${BS}!$A$8:$L$507,12,FALSE)),"")`;
    const xF = `=IFERROR(VLOOKUP($B${r},${BS}!$A$8:$B$507,2,FALSE),"")`;
    const yF = `=IFERROR(VLOOKUP($B${r},${BS}!$A$8:$H$507,8,FALSE),"")`;

    if (a) {
      dataRows.push([
        a.id, a.benId, a.name, a.owner, a.inst, a.type, a.status, a.state,
        a.num, a.open, a.strat, a.init, a.goal,
        nF, oF, pF, qF, rF, sF, tF, uF, vF, true, xF, yF, ''
      ]);
    } else {
      dataRows.push(['','','','','','','','','','','','','',nF,oF,pF,qF,rF,sF,tF,uF,vF,'',xF,yF,'']);
    }
  }
  vals.push({ range:`${S}!A8`, values:dataRows });

  // Alternate rows
  for (let i = 0; i < 1000; i++) {
    fmt.push({ repeatCell:{ range:gridRange(SID,7+i,8+i,0,NC), cell:{ userEnteredFormat:{
      backgroundColor:hex(i%2===0 ? C.white : C.altRow)
    }}, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Input columns: A-M (0-12)
  fmt.push({ repeatCell:{ range:gridRange(SID,7,1007,0,13), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  // Formula columns: N-U (13-20), V (21)
  fmt.push({ repeatCell:{ range:gridRange(SID,7,1007,13,22), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.formula)
  }}, fields:'userEnteredFormat.backgroundColor' }});
  // Currency: N, O, P, Q, R, S, U (13-19, 20)
  const curCols = [13,14,15,16,17,18,20];
  curCols.forEach(ci => fmt.push({ repeatCell:{ range:gridRange(SID,7,1007,ci,ci+1), cell:{ userEnteredFormat:{
    numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
  }}, fields:'userEnteredFormat.numberFormat' }}));
  // Percent: V (21)
  fmt.push({ repeatCell:{ range:gridRange(SID,7,1007,21,22), cell:{ userEnteredFormat:{
    numberFormat:{ type:'PERCENT', pattern:'0.0%' }
  }}, fields:'userEnteredFormat.numberFormat' }});
  // Date: J (9)
  fmt.push({ repeatCell:{ range:gridRange(SID,7,1007,9,10), cell:{ userEnteredFormat:{
    numberFormat:{ type:'DATE', pattern:'MM/DD/YYYY' }
  }}, fields:'userEnteredFormat.numberFormat' }});

  // Freeze row 7
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:7 } }, fields:'gridProperties.frozenRowCount' }});

  // Column widths (A-Z)
  [80,90,220,120, 180,130,90,70, 95,90,140,90, 110, 110,110,110,110, 110,110,120,110, 100,100,140, 90,160]
    .forEach((w,i) => fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'COLUMNS', startIndex:i, endIndex:i+1 }, properties:{ pixelSize:w }, fields:'pixelSize' }}));

  await valuesBatchUpdate(id, vals, '05-accounts');
  await batchUpdate(id, fmt, '05-accounts');
  console.log('05-accounts done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
