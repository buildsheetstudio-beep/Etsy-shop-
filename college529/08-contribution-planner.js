'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Contribution Planner'];
const S   = "'Contribution Planner'";
const BS  = "'Beneficiary Setup'";
const A5  = "'529 Accounts'";
const NC  = 16;

// 5 beneficiaries (matched to 03-beneficiary.js seed data)
const BENS = [
  { id:'BEN-001', name:'Emma Hartley',      enrollYear:2034 },
  { id:'BEN-002', name:'Lucas Hartley',     enrollYear:2037 },
  { id:'BEN-003', name:'Sofia Delgado',     enrollYear:2026 },
  { id:'BEN-004', name:'Marcus Washington', enrollYear:2024 },
  { id:'BEN-005', name:'Claire Beaumont',   enrollYear:2030 },
];

// Growth scenario helper (matches Reference Data dropdown + rates)
const SCEN = `{"Conservative","Base","Higher Growth"}`;
const RATES = `0.04,0.06,0.08`;
const monthlyRate = `CHOOSE(MATCH(${BS}!$J$6,${SCEN},0),${RATES})/12`;

(async () => {
  const vals = [];
  const fmt  = [];

  fmt.push({ repeatCell:{ range:gridRange(SID,0,200,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.bg), textFormat:{ fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  vals.push({ range:`${S}!A1`, values:[['Contribution Planner']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,0,1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,0,1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.primary), textFormat:{ bold:true, fontSize:16, foregroundColor:hex(C.white) },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:42 }, fields:'pixelSize' }});

  // ── Row 2: Subtitle ───────────────────────────────────────────────────────
  vals.push({ range:`${S}!A2`, values:[["Calculate how much to save monthly to close each beneficiary's funding gap. Uses the selected Growth Scenario from Beneficiary Setup."]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,1,2,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,1,2,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.aubergTint), textFormat:{ fontSize:9, foregroundColor:hex(C.text), italic:true },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // ── Rows 3-4: Summary Cards ───────────────────────────────────────────────
  const CARDS = [
    { label:'Active Scenario',    val:`=${BS}!$J$6`,                                         cur:false },
    { label:'Total Monthly Goal', val:`=IFERROR(SUM(H8:H12),"")`,                             cur:true  },
    { label:'Total Annual Goal',  val:`=IFERROR(SUM(I8:I12),"")`,                             cur:true  },
    { label:'Beneficiaries On Track', val:`=IFERROR(COUNTIF(J8:J12,"Ahead of Plan")+COUNTIF(J8:J12,"On Track"),"")`, cur:false },
  ];
  const cSpans=[[0,4],[4,8],[8,12],[12,16]];
  CARDS.forEach(({ label, val, cur }, i) => {
    const [cs,ce]=cSpans[i]; const cl=String.fromCharCode(65+cs);
    vals.push({ range:`${S}!${cl}3`, values:[[label]] });
    vals.push({ range:`${S}!${cl}4`, values:[[val]] });
    fmt.push({ mergeCells:{ range:gridRange(SID,2,3,cs,ce), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range:gridRange(SID,2,3,cs,ce), cell:{ userEnteredFormat:{
      backgroundColor:hex(C.eucalTint), textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.secText) }, horizontalAlignment:'CENTER'
    }}, fields:'userEnteredFormat' }});
    fmt.push({ mergeCells:{ range:gridRange(SID,3,4,cs,ce), mergeType:'MERGE_ALL' }});
    const vf={ backgroundColor:hex(C.white), textFormat:{ bold:true, fontSize:18, foregroundColor:hex(C.primary) },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' };
    if (cur) vf.numberFormat={ type:'CURRENCY', pattern:'"$"#,##0' };
    fmt.push({ repeatCell:{ range:gridRange(SID,3,4,cs,ce), cell:{ userEnteredFormat:vf }, fields:'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:36 }, fields:'pixelSize' }});

  // ── Row 6: Scenario note ──────────────────────────────────────────────────
  vals.push({ range:`${S}!A6`, values:[["Per-Beneficiary Savings Plan  ·  Required monthly contribution is computed using the PMT formula with the selected growth scenario and years remaining."]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,5,6,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,5,6,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrLight), textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.primary) }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // ── Row 7: Column headers ─────────────────────────────────────────────────
  vals.push({ range:`${S}!A7`, values:[[
    'Beneficiary ID','Beneficiary Name','Estimated Total Cost','Current Savings',
    'Funding Gap','Years Until Enrollment','Months Remaining',
    'Required Monthly Contribution','Required Annual Contribution',
    'Planner Status','Growth Scenario Used','Monthly Rate Used',
    'Notes','','',''
  ]] });
  fmt.push({ repeatCell:{ range:gridRange(SID,6,7,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:6, endIndex:7 }, properties:{ pixelSize:40 }, fields:'pixelSize' }});

  // ── Rows 8-12: One row per beneficiary ───────────────────────────────────
  const dataRows = BENS.map(({ id: bid, name }, i) => {
    const r = i + 8; // 1-indexed

    // Pull from Beneficiary Setup (lookup by BEN-ID in col A, data in cols A-P)
    const costF     = `=IFERROR(VLOOKUP("${bid}",${BS}!$A$8:$L$507,12,FALSE),"")`;   // col L = Estimated Total Cost
    const savedF    = `=IFERROR(VLOOKUP("${bid}",${BS}!$A$8:$O$507,15,FALSE),"")`;   // col O = Current Total Saved
    const gapF      = `=IFERROR(MAX(0,C${r}-D${r}),0)`;
    const yearsF    = `=IFERROR(MAX(0,VLOOKUP("${bid}",${BS}!$A$8:$H$507,8,FALSE)),0)`;  // col H = Years Until Enrollment
    const monthsF   = `=IFERROR(MAX(1,F${r}*12),1)`;
    // PMT(rate, nper, pv, fv): monthly payment to save from D (current) to C (target) in G months at scenario rate
    // PMT returns negative (outflow), so we negate it
    const pmtF      = `=IFERROR(IF(E${r}<=0,"Goal Reached!",-PMT(${monthlyRate},G${r},-D${r},C${r})),"")`;
    const annualF   = `=IFERROR(IF(E${r}<=0,"",H${r}*12),"")`;
    // Status: compare required monthly to actual current monthly contributions
    const actMonthly= `IFERROR(SUMPRODUCT((${A5}!$B$8:$B$1007="${bid}")*${A5}!$M$8:$M$1007)/12,0)`;
    const statusF   = `=IFERROR(IF(E${r}<=0,"Goal Reached!",IF(${actMonthly}>=H${r}*1.1,"Ahead of Plan",IF(${actMonthly}>=H${r}*0.9,"On Track",IF(${actMonthly}>=H${r}*0.5,"Behind Plan","Insufficient Data")))),"")`;
    const scenF     = `=${BS}!$J$6`;
    const rateF     = `=IFERROR(${monthlyRate},"")`;

    return [bid, name, costF, savedF, gapF, yearsF, monthsF, pmtF, annualF, statusF, scenF, rateF, '', '', '', ''];
  });
  vals.push({ range:`${S}!A8`, values:dataRows });

  // Formatting for data rows
  for (let i=0; i<5; i++) {
    const r = 7+i; // 0-indexed
    fmt.push({ repeatCell:{ range:gridRange(SID,r,r+1,0,NC), cell:{ userEnteredFormat:{
      backgroundColor:hex(i%2===0?C.white:C.altRow)
    }}, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Input cols A-B: input color
  fmt.push({ repeatCell:{ range:gridRange(SID,7,12,0,2), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  // Formula cols C-L: formula color
  fmt.push({ repeatCell:{ range:gridRange(SID,7,12,2,12), cell:{ userEnteredFormat:{ backgroundColor:hex(C.formula) } }, fields:'userEnteredFormat.backgroundColor' }});

  // Currency: C, D, E, H, I (2,3,4,7,8)
  [2,3,4,7,8].forEach(ci => fmt.push({ repeatCell:{ range:gridRange(SID,7,12,ci,ci+1), cell:{ userEnteredFormat:{
    numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
  }}, fields:'userEnteredFormat.numberFormat' }}));
  // Percent: L (11)
  fmt.push({ repeatCell:{ range:gridRange(SID,7,12,11,12), cell:{ userEnteredFormat:{ numberFormat:{ type:'PERCENT', pattern:'0.0%' } } }, fields:'userEnteredFormat.numberFormat' }});
  // Number: F, G (5, 6)
  [5,6].forEach(ci => fmt.push({ repeatCell:{ range:gridRange(SID,7,12,ci,ci+1), cell:{ userEnteredFormat:{
    numberFormat:{ type:'NUMBER', pattern:'0.0' }
  }}, fields:'userEnteredFormat.numberFormat' }}));

  // Status conditional colors (J col = 9): applied via 11-cf.js

  // ── Row 14+: Instructions ─────────────────────────────────────────────────
  const notes = [
    ['HOW TO USE THIS PLANNER'],
    ['1. Review the Required Monthly Contribution column (H) for each beneficiary.'],
    ['2. The calculation uses the Growth Scenario selected in Beneficiary Setup (cell J6).'],
    ['3. If "Goal Reached!" appears, the current savings already meet or exceed the estimated cost.'],
    ['4. "Ahead of Plan" = current monthly contributions exceed the required amount by 10%+.'],
    ['5. "On Track" = current contributions are within 10% of required. "Behind Plan" = below 50%.'],
    ['6. Adjust annual contribution goals in the 529 Accounts tab (column M) to update the status.'],
    ['7. Re-run this script after updating Beneficiary Setup estimates or account goals.'],
    [''],
    ['DISCLAIMER: PMT calculations assume a constant contribution rate and growth rate. Actual returns will vary.'],
  ];
  vals.push({ range:`${S}!A14`, values:notes });
  fmt.push({ repeatCell:{ range:gridRange(SID,13,14,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrLight), textFormat:{ bold:true, fontSize:10, foregroundColor:hex(C.primary) }
  }}, fields:'userEnteredFormat' }});
  fmt.push({ mergeCells:{ range:gridRange(SID,13,14,0,NC), mergeType:'MERGE_ALL' }});
  for (let i=1; i<notes.length; i++) {
    fmt.push({ repeatCell:{ range:gridRange(SID,13+i,14+i,0,NC), cell:{ userEnteredFormat:{
      backgroundColor:hex(i%2===0?C.altRow:C.white), textFormat:{ fontSize:9 }
    }}, fields:'userEnteredFormat' }});
    fmt.push({ mergeCells:{ range:gridRange(SID,13+i,14+i,0,NC), mergeType:'MERGE_ALL' }});
  }

  // Freeze row 7
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:7 } }, fields:'gridProperties.frozenRowCount' }});

  // Column widths
  [80,140,130,120,110,90,80, 150,140, 130,130,90, 120,80,80,80]
    .forEach((w,i) => fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'COLUMNS', startIndex:i, endIndex:i+1 }, properties:{ pixelSize:w }, fields:'pixelSize' }}));

  await valuesBatchUpdate(id, vals, '08-contribution-planner');
  await batchUpdate(id, fmt, '08-contribution-planner');
  console.log('08-contribution-planner done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
