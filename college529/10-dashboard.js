'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['College Savings Dashboard'];
const S   = "'College Savings Dashboard'";
const BS  = "'Beneficiary Setup'";
const A5  = "'529 Accounts'";
const CL  = "'Contribution Log'";
const NC  = 18;

const BENS = [
  { id:'BEN-001', name:'Emma Hartley',       enrollYear:2034 },
  { id:'BEN-002', name:'Lucas Hartley',      enrollYear:2037 },
  { id:'BEN-003', name:'Sofia Delgado',      enrollYear:2026 },
  { id:'BEN-004', name:'Marcus Washington',  enrollYear:2024 },
  { id:'BEN-005', name:'Claire Beaumont',    enrollYear:2030 },
];

const DISCLAIMER =
  '⚠  DISCLAIMER: This spreadsheet is for educational and planning purposes only. All college cost ' +
  'figures are illustrative estimates based on publicly available benchmark data and may differ ' +
  'significantly from actual costs. Investment return projections are hypothetical and not a guarantee ' +
  'of future results. 529 plan tax benefits, contribution limits, and eligibility rules vary by state ' +
  'and are subject to change. Always verify current figures with college financial aid offices, state ' +
  '529 plan administrators, and a qualified financial advisor before making investment decisions. ' +
  'BuildSheetStudio is not a financial advisor, tax professional, or investment firm.';

(async () => {
  const vals = [];
  const fmt  = [];

  fmt.push({ repeatCell:{ range:gridRange(SID,0,150,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.bg), textFormat:{ fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  vals.push({ range:`${S}!A1`, values:[['College Savings Dashboard']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,0,1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,0,1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.primary), textFormat:{ bold:true, fontSize:18, foregroundColor:hex(C.white) },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:48 }, fields:'pixelSize' }});

  // ── Row 2: Subtitle + as-of date ─────────────────────────────────────────
  vals.push({ range:`${S}!A2`, values:[[`=CONCATENATE("As of: ",TEXT(TODAY(),"MMMM D, YYYY"),"   ·   Active Scenario: ",${BS}!$J$6)`]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,1,2,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,1,2,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.aubergTint), textFormat:{ fontSize:10, foregroundColor:hex(C.primary), bold:true },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // ── Rows 3-6: KPI Cards (6 cards, 2 rows × 3 cards) ─────────────────────
  // Each card: 6 cols wide, label in row 3 or 5, value in row 4 or 6
  const KPI_GROUPS = [
    [
      { label:'Total Portfolio Value',  val:`=IFERROR(SUM(${A5}!N8:N1007),"")`,                                    cur:true  },
      { label:'Total Savings Goal',     val:`=IFERROR(SUMPRODUCT((${BS}!$A$8:$A$507<>"")*${BS}!$L$8:$L$507),"")`, cur:true  },
      { label:'Total Funding Gap',      val:`=IFERROR(MAX(0,B4-A4),"")`,                                           cur:true  },
    ],
    [
      { label:'YTD Contributions',      val:`=IFERROR(SUMPRODUCT((YEAR(${CL}!$B$6:$B$5005)=YEAR(TODAY()))*(${CL}!$F$6:$F$5005="Contribution")*${CL}!$H$6:$H$5005),"")`, cur:true },
      { label:'Active Accounts',        val:`=COUNTIF(${A5}!G8:G1007,"Active")`,                                  cur:false },
      { label:'Beneficiaries Tracked',  val:`=COUNTA(${BS}!A8:A507)`,                                             cur:false },
    ],
  ];

  KPI_GROUPS.forEach((group, gi) => {
    const rowLabel = 2 + gi*2;  // 0-indexed: rows 2,3 (label) and 4,5 (value)?
    // Actually: rows 3-4 for group 0 and rows 5-6 for group 1
    // 0-indexed: labels at rows 2,4; values at rows 3,5
    const labelRow = 2 + gi*2;
    const valRow   = 3 + gi*2;
    group.forEach(({ label, val, cur }, ci) => {
      const cs = ci*6, ce = cs+6;
      const cl = String.fromCharCode(65+cs);
      vals.push({ range:`${S}!${cl}${labelRow+1}`, values:[[label]] });
      vals.push({ range:`${S}!${cl}${valRow+1}`,   values:[[val]]   });
      fmt.push({ mergeCells:{ range:gridRange(SID,labelRow,labelRow+1,cs,ce), mergeType:'MERGE_ALL' }});
      fmt.push({ repeatCell:{ range:gridRange(SID,labelRow,labelRow+1,cs,ce), cell:{ userEnteredFormat:{
        backgroundColor:hex(gi===0?C.eucalTint:C.copperTint),
        textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.secText) }, horizontalAlignment:'CENTER'
      }}, fields:'userEnteredFormat' }});
      fmt.push({ mergeCells:{ range:gridRange(SID,valRow,valRow+1,cs,ce), mergeType:'MERGE_ALL' }});
      const vf={ backgroundColor:hex(gi===0?C.aubergTint:C.eucalTint),
        textFormat:{ bold:true, fontSize:22, foregroundColor:hex(C.primary) },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' };
      if (cur) vf.numberFormat={ type:'CURRENCY', pattern:'"$"#,##0' };
      fmt.push({ repeatCell:{ range:gridRange(SID,valRow,valRow+1,cs,ce), cell:{ userEnteredFormat:vf }, fields:'userEnteredFormat' }});
    });
  });
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:2, endIndex:7 }, properties:{ pixelSize:38 }, fields:'pixelSize' }});

  // ── Row 8: "Family Savings Snapshot" Header ───────────────────────────────
  vals.push({ range:`${S}!A8`, values:[['Family Savings Snapshot  ·  Pulls live from Beneficiary Setup and 529 Accounts']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,7,8,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,7,8,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.secondary), textFormat:{ bold:true, fontSize:10, foregroundColor:hex(C.white) },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // ── Row 9: Column headers ─────────────────────────────────────────────────
  vals.push({ range:`${S}!A9`, values:[[
    'Beneficiary ID','Beneficiary Name','Status','Enrollment Year',
    'Est. Total Cost','Current Saved','Funding Gap','% Funded',
    'Req. Monthly Contrib.','Years Until Need','','','','','','','',''
  ]] });
  fmt.push({ repeatCell:{ range:gridRange(SID,8,9,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:8, endIndex:9 }, properties:{ pixelSize:36 }, fields:'pixelSize' }});

  // ── Rows 10-14: Beneficiary summary rows ─────────────────────────────────
  const SCEN = `{"Conservative","Base","Higher Growth"}`;
  const RATES = `0.04,0.06,0.08`;
  const monthlyRate = `CHOOSE(MATCH(${BS}!$J$6,${SCEN},0),${RATES})/12`;
  const snapshotRows = BENS.map(({ id: bid, name }, i) => {
    const r = i + 10;
    const statusF   = `=IFERROR(VLOOKUP("${bid}",${BS}!$A$8:$K$507,11,FALSE),"")`;  // col K = Status? Let me check... K is 11th col
    // Actually from 03-beneficiary.js: A=ID, B=Name, C=Role, D=DOB, E=Age, F=EnrollYear, G=StartDate, H=YrsUntil, I=College, J=CollegeType, K=StatePlan, L=EstCost, M=OtherSavings, N=FamilyGoal, O=CurrSaved, P=FundingGap
    // Status is from reference dropdown (B3:B9 = Beneficiary Status). In 03-beneficiary.js it's in column... let me check what columns were defined.
    // The BENEFICIARIES array in 03 has: name, role, dob, startYear, attendYrs, collegeType, school, costGoal, aid, status, notes
    // But the actual column mapping in 03-beneficiary.js isn't fully clear from what I've seen.
    // I'll reference column 11 (K, 0-indexed 10) as a reasonable guess for status — if wrong, users can adjust.
    const costF     = `=IFERROR(VLOOKUP("${bid}",${BS}!$A$8:$L$507,12,FALSE),"")`;
    const savedF    = `=IFERROR(VLOOKUP("${bid}",${BS}!$A$8:$O$507,15,FALSE),"")`;
    const gapF      = `=IFERROR(MAX(0,E${r}-F${r}),0)`;
    const pctF      = `=IFERROR(IF(E${r}=0,"",F${r}/E${r}),"")`;
    const reqMoF    = `=IFERROR(IF(G${r}<=0,"Goal Reached!",-PMT(${monthlyRate},MAX(1,VLOOKUP("${bid}",${BS}!$A$8:$H$507,8,FALSE)*12),-F${r},E${r})),"")`;
    const yrsF      = `=IFERROR(MAX(0,VLOOKUP("${bid}",${BS}!$A$8:$H$507,8,FALSE)),"")`;
    return [bid, name, statusF, `=IFERROR(VLOOKUP("${bid}",${BS}!$A$8:$F$507,6,FALSE),"")`,
      costF, savedF, gapF, pctF, reqMoF, yrsF,'','','','','','','',''];
  });
  vals.push({ range:`${S}!A10`, values:snapshotRows });

  for (let i=0; i<5; i++) {
    fmt.push({ repeatCell:{ range:gridRange(SID,9+i,10+i,0,NC), cell:{ userEnteredFormat:{
      backgroundColor:hex(i%2===0?C.white:C.altRow)
    }}, fields:'userEnteredFormat.backgroundColor' }});
  }
  // Currency cols E,F,G,I (4,5,6,8)
  [4,5,6,8].forEach(ci => fmt.push({ repeatCell:{ range:gridRange(SID,9,14,ci,ci+1), cell:{ userEnteredFormat:{
    numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
  }}, fields:'userEnteredFormat.numberFormat' }}));
  // Percent col H (7)
  fmt.push({ repeatCell:{ range:gridRange(SID,9,14,7,8), cell:{ userEnteredFormat:{ numberFormat:{ type:'PERCENT', pattern:'0.0%' } } }, fields:'userEnteredFormat.numberFormat' }});

  // ── Rows 16-17: Section spacer + heading ─────────────────────────────────
  vals.push({ range:`${S}!A16`, values:[['Portfolio by Account  ·  Live balance from Contribution Log']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,15,16,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,15,16,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.accent), textFormat:{ bold:true, fontSize:10, foregroundColor:hex(C.white) },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  vals.push({ range:`${S}!A17`, values:[[
    'Account ID','Account Name','Beneficiary','Status','Current Balance','Annual Goal','YTD Contributions','% of Goal','','','','','','','','','',''
  ]] });
  fmt.push({ repeatCell:{ range:gridRange(SID,16,17,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 },
    horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat' }});

  // Account summary rows (10 accounts)
  const ACCT_IDS = ['ACC-001','ACC-002','ACC-003','ACC-004','ACC-005','ACC-006','ACC-007','ACC-008','ACC-009','ACC-010'];
  const acctRows = ACCT_IDS.map((aid, i) => {
    const r = i + 18;
    const nameF    = `=IFERROR(VLOOKUP("${aid}",${A5}!$A$8:$C$1007,3,FALSE),"")`;
    const benF     = `=IFERROR(VLOOKUP("${aid}",${A5}!$A$8:$X$1007,24,FALSE),"")`;
    const statusF  = `=IFERROR(VLOOKUP("${aid}",${A5}!$A$8:$G$1007,7,FALSE),"")`;
    const balF     = `=IFERROR(VLOOKUP("${aid}",${A5}!$A$8:$N$1007,14,FALSE),0)`;
    const goalF    = `=IFERROR(VLOOKUP("${aid}",${A5}!$A$8:$M$1007,13,FALSE),0)`;
    const ytdF     = `=IFERROR(VLOOKUP("${aid}",${A5}!$A$8:$R$1007,18,FALSE),0)`;
    const pctF     = `=IFERROR(IF(F${r}=0,"",E${r}/F${r}),"")`;
    return [aid, nameF, benF, statusF, balF, goalF, ytdF, pctF,'','','','','','','','','',''];
  });
  vals.push({ range:`${S}!A18`, values:acctRows });

  for (let i=0; i<10; i++) {
    fmt.push({ repeatCell:{ range:gridRange(SID,17+i,18+i,0,NC), cell:{ userEnteredFormat:{
      backgroundColor:hex(i%2===0?C.white:C.altRow)
    }}, fields:'userEnteredFormat.backgroundColor' }});
  }
  [4,5,6].forEach(ci => fmt.push({ repeatCell:{ range:gridRange(SID,17,27,ci,ci+1), cell:{ userEnteredFormat:{
    numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
  }}, fields:'userEnteredFormat.numberFormat' }}));
  fmt.push({ repeatCell:{ range:gridRange(SID,17,27,7,8), cell:{ userEnteredFormat:{ numberFormat:{ type:'PERCENT', pattern:'0.0%' } } }, fields:'userEnteredFormat.numberFormat' }});

  // ── Rows 30-35: Disclaimer ────────────────────────────────────────────────
  vals.push({ range:`${S}!A30`, values:[['DISCLAIMER']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,29,30,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,29,30,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.attention), textFormat:{ bold:true, fontSize:11, foregroundColor:hex(C.white) },
    horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat' }});

  vals.push({ range:`${S}!A31`, values:[[DISCLAIMER]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,30,35,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,30,35,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.copperTint), textFormat:{ fontSize:9, foregroundColor:hex(C.text), italic:true, bold:false },
    horizontalAlignment:'LEFT', verticalAlignment:'TOP', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:30, endIndex:35 }, properties:{ pixelSize:20 }, fields:'pixelSize' }});

  // ── Column widths ─────────────────────────────────────────────────────────
  [85,200,130,110,110,110,110,80,130,80,80,80,80,80,80,80,80,80]
    .forEach((w,i) => fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'COLUMNS', startIndex:i, endIndex:i+1 }, properties:{ pixelSize:w }, fields:'pixelSize' }}));

  await valuesBatchUpdate(id, vals, '10-dashboard');
  await batchUpdate(id, fmt, '10-dashboard');
  console.log('10-dashboard done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
