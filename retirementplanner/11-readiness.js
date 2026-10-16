'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs'), path = require('path');
const SID   = JSON.parse(fs.readFileSync(path.join(__dirname,'spreadsheet.json'))).id;
const SHEET = 9;
const TAB   = 'Retirement Readiness';
const SETUP = "'Personal & Household Setup'";
const FCAST = "'Growth Forecast'";
const SCEN  = "'Scenario Planner'";
const INC   = "'Retirement Income'";

/* ── row index constants (0-indexed) ─────────────────────────────────────── */
const R_TITLE  = 0;
const R_SUB    = 1;
const R_SCRHD  = 2;   // "Overall Readiness Score" header
const R_SCRVL  = 3;   // Overall score values (big merged display)
const R_SP1    = 4;
const R_KPILBL = 5;
const R_KPIVAL = 6;
const R_SP2    = 7;
const R_SCHDR  = 8;   // Scorecard section header
const R_SCCOL  = 9;   // Scorecard column headers
const R_SC0    = 10;  // 8 scorecard rows (10-17) → 1-indexed 11-18
const R_SCTOT  = 18;  // Scorecard totals → 1-indexed 19
const R_SP3    = 19;
const R_SCNHDR = 20;  // Scenario Comparison section header
const R_SCNCOL = 21;
const R_SCN0   = 22;  // 6 scenario rows (22-27) → 1-indexed 23-28
const R_SP4    = 28;
const R_GAPHDR = 29;  // P1 vs P2 Gap Analysis
const R_GAPCOL = 30;
const R_GAP0   = 31;  // 7 gap rows (31-37) → 1-indexed 32-38
const R_SP5    = 38;
const R_ACTHDR = 39;  // Action Plan header
const R_ACTCOL = 40;
const R_ACT0   = 41;  // 11 action rows (41-51) → 1-indexed 42-52
const R_SP6    = 52;
const R_DISC   = 53;

/* ── column widths A-I (9 cols) ─────────────────────────────────────────── */
const COL_W = [55, 220, 165, 85, 75, 95, 200, 165, 175];

/* ── formula fragments (NO leading =) ─────────────────────────────────────── */
const portExpr    = `IFERROR(VLOOKUP(YEAR(${SETUP}!B8)+${SETUP}!B10,${FCAST}!$A$7:$K$47,11,0),0)`;
const incomeExpr  = `IFERROR(VLOOKUP(YEAR(${SETUP}!B8)+${SETUP}!B10,${FCAST}!$A$7:$H$47,8,0),0)`;
const annGoalExpr = `${SETUP}!B30`;
const safeRateExpr= `${SETUP}!B28`;
const nestEggExpr = `${SETUP}!B31`;

async function run() {
  const requests = [];
  const vals     = [];

  /* ── 1. Column widths & row heights ─────────────────────────────────────── */
  COL_W.forEach((w,i) => requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
    properties:{ pixelSize:w }, fields:'pixelSize' }
  }));
  [[R_TITLE,52],[R_SUB,22],[R_SCRHD,24],[R_SCRVL,50],[R_SP1,6],[R_SP2,6],[R_SP3,6],
   [R_SP4,6],[R_SP5,6],[R_SP6,6],[R_DISC,24]].forEach(([r,h]) => {
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r, endIndex:r+1 },
      properties:{ pixelSize:h }, fields:'pixelSize' }
    });
  });

  /* ── 2. Title & subtitle ─────────────────────────────────────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,R_TITLE,R_TITLE+1,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,R_TITLE,R_TITLE+1,0,9),
    cell:{ userEnteredValue:{ stringValue:`✅  ${TAB}` },
      userEnteredFormat:{ backgroundColor:hex(C.primary),
        textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:16, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });
  requests.push({ mergeCells:{ range:gridRange(SHEET,R_SUB,R_SUB+1,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,R_SUB,R_SUB+1,0,9),
    cell:{ userEnteredValue:{ stringValue:'Measure your readiness across 8 financial and life dimensions to reach a confident retirement' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrB),
        textFormat:{ foregroundColor:hex(C.primaryText), italic:true, fontSize:10, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── 3. Overall Readiness Score display (rows 2-3) ──────────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,R_SCRHD,R_SCRHD+1,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,R_SCRHD,R_SCRHD+1,0,9),
    cell:{ userEnteredValue:{ stringValue:'Overall Retirement Readiness Score' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrC),
        textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:10, fontFamily:'Montserrat' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:10 } }},
    fields:'userEnteredValue,userEnteredFormat' }
  });
  // 3 panels in the overall score row: Score, Rating, Portfolio vs Target
  const scorePanels = [
    { cols:[0,3], label:'Weighted Readiness Score (0–5.0)',  formula:`=IFERROR(F19,0)`, fmt:{ type:'NUMBER', pattern:'0.00" / 5.0"' }, color:C.primary, tc:C.primaryText },
    { cols:[3,6], label:'Readiness Rating',                  formula:`=IFERROR(IF(F19>=4,"Strong",IF(F19>=3,"On Track",IF(F19>=2,"Needs Attention","Critical"))),"")`, color:C.secondary, tc:C.text },
    { cols:[6,9], label:'Portfolio vs Target Nest Egg',      formula:`=IFERROR(TEXT(${portExpr}/${nestEggExpr},"0%")&" of target","")`, color:C.hdrB, tc:C.primaryText },
  ];
  scorePanels.forEach(({ cols:[c0,c1], label, formula, fmt, color, tc }) => {
    requests.push({ mergeCells:{ range:gridRange(SHEET,R_SCRVL,R_SCRVL+1,c0,c1), mergeType:'MERGE_ALL' }});
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_SCRVL,R_SCRVL+1,c0,c1),
      cell:{ userEnteredValue:{ formulaValue:formula },
        userEnteredFormat:{ backgroundColor:hex(color),
          textFormat:{ foregroundColor:hex(tc), bold:true, fontSize:20, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          ...(fmt ? { numberFormat:fmt } : {}) }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });

  /* ── 4. KPI strip (rows 5-6) ─────────────────────────────────────────────── */
  const kpis = [
    { label:'Target\nNest Egg',       formula:`=${nestEggExpr}`,                                         color:C.primary,  tc:C.primaryText, nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'Portfolio at\nRetirement',formula:`=${portExpr}`,                                           color:C.secondary,tc:C.text,        nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'Surplus /\nShortfall',    formula:`=IFERROR(${portExpr}-${nestEggExpr},0)`,                 color:C.hdrB,     tc:C.primaryText, nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'4% Annual\nWithdrawal',   formula:`=IFERROR(${portExpr}*${safeRateExpr},0)`,               color:C.hdrC,     tc:C.primaryText, nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'Income at\nRetirement',   formula:`=${incomeExpr}`,                                         color:C.info,     tc:C.text,        nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'Income\nCoverage %',      formula:`=IFERROR(${incomeExpr}/${annGoalExpr},0)`,              color:C.altRow,   tc:C.text,        nf:{ type:'PERCENT', pattern:'0%' } },
    { label:'P1 Yrs to\nRetirement',   formula:`=IFERROR(${SETUP}!B12,"")`,                              color:C.panel,    tc:C.text,        nf:{ type:'NUMBER', pattern:'0" yrs"' } },
    { label:'P2 Yrs to\nRetirement',   formula:`=IFERROR(${SETUP}!E12,"")`,                              color:C.panel,    tc:C.text,        nf:{ type:'NUMBER', pattern:'0" yrs"' } },
    { label:'Safe\nWithdrawal Rate',   formula:`=${safeRateExpr}`,                                       color:C.formula,  tc:C.text,        nf:{ type:'PERCENT', pattern:'0.0%' } },
  ];
  kpis.forEach((k,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_KPILBL,R_KPILBL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:k.label },
        userEnteredFormat:{ backgroundColor:hex(k.color),
          textFormat:{ foregroundColor:hex(k.tc), bold:true, fontSize:8, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
          borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_KPIVAL,R_KPIVAL+1,i,i+1),
      cell:{ userEnteredValue:{ formulaValue:k.formula },
        userEnteredFormat:{ backgroundColor:hex(k.color),
          textFormat:{ foregroundColor:hex(k.tc), bold:true, fontSize:13, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          numberFormat:k.nf }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });

  /* ── 5. Spacers ──────────────────────────────────────────────────────────── */
  [R_SP1,R_SP2,R_SP3,R_SP4,R_SP5,R_SP6].forEach(r => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,0,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(C.bg) }}, fields:'userEnteredFormat' }
    });
  });

  /* ── 6. Section headers ──────────────────────────────────────────────────── */
  function secHdr(r, text) {
    requests.push({ mergeCells:{ range:gridRange(SHEET,r,r+1,0,9), mergeType:'MERGE_ALL' }});
    requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,0,9),
      cell:{ userEnteredValue:{ stringValue:text },
        userEnteredFormat:{ backgroundColor:hex(C.hdrB),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:10, fontFamily:'Montserrat' },
          horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:10 } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r, endIndex:r+1 },
      properties:{ pixelSize:26 }, fields:'pixelSize' }
    });
  }
  secHdr(R_SCHDR,  'Readiness Scorecard — 8-Dimension Assessment');
  secHdr(R_SCNHDR, 'Scenario Comparison — Portfolio & Gap by Retirement Scenario');
  secHdr(R_GAPHDR, 'P1 vs P2 Retirement Profile Comparison');
  secHdr(R_ACTHDR, 'Priority Action Plan — Top 11 Items to Improve Readiness');

  /* ── 7. Scorecard ────────────────────────────────────────────────────────── */
  const SCHDR = ['#','Readiness Dimension','Assessment Basis','Score\n(1–5)','Weight','Weighted\nScore','Current Status','Gap / Improvement Needed','Owner'];
  SCHDR.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_SCCOL,R_SCCOL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:R_SCCOL, endIndex:R_SCCOL+1 },
    properties:{ pixelSize:32 }, fields:'pixelSize' }
  });

  // [dim, basis, score_formula, weight, status, gap, owner]
  const SCORECARD = [
    ['Portfolio Accumulation',   'Portfolio FV ÷ Target Nest Egg × 5',
     `=MIN(5,IFERROR(${portExpr}/${nestEggExpr}*5,0))`, 0.25,
     `=IF(D11>=4,"Strong","Needs Growth")`,
     `=IFERROR("Reach target nest egg of "&TEXT(${nestEggExpr},"$#,##0")&"; current projection "&TEXT(${portExpr}/${nestEggExpr},"0%")&" of goal","")`,
     'Both'],
    ['Income Replacement',       'External income at retire ÷ Annual goal × 5',
     `=MIN(5,IFERROR(${incomeExpr}/${annGoalExpr}*5,0))`, 0.20,
     `=IF(D12>=4,"Sufficient","Needs Sources")`,
     `=IFERROR("Income gap of "&TEXT(MAX(0,${annGoalExpr}-${incomeExpr}),"$#,##0")&" must be covered by withdrawals","")`,
     'Both'],
    ['Expense Coverage (4% Rule)', 'SWR withdrawal ÷ Annual goal × 5',
     `=MIN(5,IFERROR(${portExpr}*${safeRateExpr}/${annGoalExpr}*5,0))`, 0.15,
     `=IF(D13>=4,"Covered","Shortfall Risk")`,
     `=IFERROR("4% withdrawal "&TEXT(${portExpr}*${safeRateExpr},"$#,##0")&" vs goal "&TEXT(${annGoalExpr},"$#,##0"),"")`,
     'Both'],
    ['Healthcare Readiness',     'LTC policy, Medicare enrollment, health savings',
     3.0, 0.15,
     'In Progress — LTC shopping',
     'Purchase LTC insurance by age 57; enroll Medicare by 65',
     'Both'],
    ['Legal & Estate Prep',      'Will, trust, POA, beneficiary designations',
     4.0, 0.10,
     'Mostly Complete',
     'Update beneficiary designations; create letter of instruction',
     'Both'],
    ['Housing Stability',        'Mortgage payoff timeline vs retirement date',
     3.0, 0.05,
     'In Progress — payoff target age 62',
     'Accelerate mortgage payoff by $500/mo; confirm payoff at 62',
     'P1'],
    ['Longevity Protection',     'Inflation modeling, LTC, longevity to age 90',
     3.5, 0.05,
     'Partially Protected',
     'Increase inflation buffer; LTC insurance to protect assets',
     'Both'],
    ['Lifestyle Readiness',      'Social, travel, volunteer, hobby milestones',
     3.0, 0.05,
     'Planning Phase',
     'Define post-retirement routine; explore volunteer options',
     'Both'],
  ];

  SCORECARD.forEach(([dim, basis, score, weight, status, gap, owner], idx) => {
    const r0  = R_SC0 + idx;
    const r1  = r0 + 1; // 1-indexed
    const isAlt = idx % 2 === 1;
    const bg  = isAlt ? C.altRow : C.panel;

    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,0,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    // Score column D: bold, centered, number format
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,3,4),
      cell:{ userEnteredFormat:{ textFormat:{ bold:true, fontSize:12, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER',
        numberFormat:{ type:'NUMBER', pattern:'0.0' } }},
      fields:'userEnteredFormat' }
    });
    // Weight col E: percent
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,4,5),
      cell:{ userEnteredFormat:{ numberFormat:{ type:'PERCENT', pattern:'0%' },
        horizontalAlignment:'CENTER' }},
      fields:'userEnteredFormat' }
    });
    // Weighted score col F
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,5,6),
      cell:{ userEnteredFormat:{ numberFormat:{ type:'NUMBER', pattern:'0.00' },
        horizontalAlignment:'CENTER', textFormat:{ bold:true } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r0, endIndex:r0+1 },
      properties:{ pixelSize:40 }, fields:'pixelSize' }
    });

    const scoreVal  = typeof score === 'number' ? score : { formulaValue: score };
    const statusVal = typeof status === 'string' && !status.startsWith('=') ? status : { formulaValue: status };
    const gapVal    = typeof gap === 'string' && !gap.startsWith('=') ? gap : { formulaValue: gap };

    vals.push({ range:`${TAB}!A${r1}:I${r1}`, values:[[
      idx+1, dim, basis,
      typeof score === 'number' ? score : score,
      weight,
      `=D${r1}*E${r1}`,
      typeof status === 'string' && !status.startsWith('=') ? status : status,
      typeof gap === 'string' && !gap.startsWith('=') ? gap : gap,
      owner,
    ]] });
  });

  // Scorecard totals row
  const r_sctot = R_SCTOT;
  requests.push({ repeatCell:{ range:gridRange(SHEET,r_sctot,r_sctot+1,0,9),
    cell:{ userEnteredFormat:{ backgroundColor:hex(C.hdrC),
      textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontFamily:'Montserrat', fontSize:10 },
      verticalAlignment:'MIDDLE' }},
    fields:'userEnteredFormat' }
  });
  requests.push({ repeatCell:{ range:gridRange(SHEET,r_sctot,r_sctot+1,4,6),
    cell:{ userEnteredFormat:{ numberFormat:{ type:'NUMBER', pattern:'0.00' }, horizontalAlignment:'CENTER' }},
    fields:'userEnteredFormat' }
  });
  vals.push({ range:`${TAB}!A${r_sctot+1}:I${r_sctot+1}`, values:[[
    '','Overall Weighted Score','',
    '=SUMPRODUCT(D11:D18,E11:E18)','=SUM(E11:E18)','=SUM(F11:F18)',
    `=IF(F19>=4,"Strong — You are well-positioned for retirement",IF(F19>=3,"On Track — a few gaps to address",IF(F19>=2,"Needs Attention — significant gaps exist","Critical — urgent action required")))`,
    '','',
  ]] });
  requests.push({ repeatCell:{ range:gridRange(SHEET,r_sctot,r_sctot+1,3,4),
    cell:{ userEnteredFormat:{ numberFormat:{ type:'NUMBER', pattern:'0.0" / 5.0"' }, horizontalAlignment:'CENTER', textFormat:{ bold:true, fontSize:13 } }},
    fields:'userEnteredFormat' }
  });

  /* ── 8. Scenario Comparison ──────────────────────────────────────────────── */
  const SCNHDR = ['Scenario','Name','P1 Retire\nAge','P2 Retire\nAge','Return\nRate','Portfolio\nat Retirement','Annual\n4% SWR','Annual\nGoal','Gap / Surplus'];
  SCNHDR.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_SCNCOL,R_SCNCOL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:R_SCNCOL, endIndex:R_SCNCOL+1 },
    properties:{ pixelSize:30 }, fields:'pixelSize' }
  });

  const SCN_COLORS = [C.SCN01,C.SCN02,C.SCN03,C.SCN04,C.SCN05,C.SCN06];
  for (let i = 0; i < 6; i++) {
    const r0 = R_SCN0 + i;
    const r1 = r0 + 1;
    const sr  = 6 + i; // 1-indexed row in Scenario Planner (rows 6-11)
    const isAlt = i % 2 === 1;
    const bg = isAlt ? C.altRow : C.panel;
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,0,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateBorders:{ range:gridRange(SHEET,r0,r0+1,0,1),
      left:{ style:'SOLID_MEDIUM', color:hex(SCN_COLORS[i]) } }
    });
    [[4,'CURRENCY','$#,##0'],[5,'CURRENCY','$#,##0'],[6,'CURRENCY','$#,##0'],[7,'CURRENCY','$#,##0'],[8,'CURRENCY','+$#,##0;-$#,##0']].forEach(([col,type,pat]) => {
      requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,col,col+1),
        cell:{ userEnteredFormat:{ numberFormat:{ type, pattern:pat }, horizontalAlignment:'RIGHT' }},
        fields:'userEnteredFormat' }
      });
    });
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,2,4),
      cell:{ userEnteredFormat:{ numberFormat:{ type:'NUMBER', pattern:'0" yrs"' }, horizontalAlignment:'CENTER' }},
      fields:'userEnteredFormat' }
    });
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,3,4),
      cell:{ userEnteredFormat:{ numberFormat:{ type:'PERCENT', pattern:'0.0%' }, horizontalAlignment:'CENTER' }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r0, endIndex:r0+1 },
      properties:{ pixelSize:28 }, fields:'pixelSize' }
    });
    vals.push({ range:`${TAB}!A${r1}:I${r1}`, values:[[
      `=${SCEN}!A${sr}`, `=${SCEN}!B${sr}`, `=${SCEN}!D${sr}`, `=${SCEN}!E${sr}`,
      `=${SCEN}!G${sr}`, `=${SCEN}!K${sr}`,
      `=IFERROR(F${r1}*${safeRateExpr},0)`,
      `=${SCEN}!M${sr}`,
      `=${SCEN}!N${sr}`,
    ]] });
  }

  /* ── 9. P1 vs P2 Gap Analysis ────────────────────────────────────────────── */
  const GAPH = ['#','Metric','P1 — James','P2 — Patricia','Combined / Delta','Notes'];
  GAPH.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_GAPCOL,R_GAPCOL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment: i<=1 ? 'LEFT':'CENTER', verticalAlignment:'MIDDLE',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:R_GAPCOL, endIndex:R_GAPCOL+1 },
    properties:{ pixelSize:28 }, fields:'pixelSize' }
  });

  const gapData = [
    ['Date of Birth',         `=${SETUP}!B8`, `=${SETUP}!E8`, `=TEXT(${SETUP}!B8,"yyyy-mm-dd")&" / "&TEXT(${SETUP}!E8,"yyyy-mm-dd")`, 'P1 is 2.5 years older'],
    ['Current Age',           `=${SETUP}!B9`, `=${SETUP}!E9`, `=C${R_GAP0+1+1}+D${R_GAP0+1+1}`, ''],
    ['Target Retire Age',     `=${SETUP}!B10`, `=${SETUP}!E10`, `=MIN(C${R_GAP0+1+2},D${R_GAP0+1+2})`, 'P2 retires slightly earlier relative to age'],
    ['Projected Retire Year', `=${SETUP}!B11`, `=${SETUP}!E11`, `=${SETUP}!B11`, 'Both retire 2037'],
    ['Years to Retirement',   `=${SETUP}!B12`, `=${SETUP}!E12`, `=MAX(C${R_GAP0+1+4},D${R_GAP0+1+4})`, 'Household is "retired" when P1 retires'],
    ['Annual Gross Income',   `=${SETUP}!B14`, `=${SETUP}!E14`, `=${SETUP}!B32`, 'Combined household income'],
    ['Annual Savings',        `=${SETUP}!B16`, `=${SETUP}!E16`, `=${SETUP}!B33`, 'Total household savings per year'],
  ];
  gapData.forEach(([metric, p1, p2, combined, note], idx) => {
    const r0 = R_GAP0 + idx;
    const r1 = r0 + 1;
    const isAlt = idx % 2 === 1;
    const bg = isAlt ? C.altRow : C.panel;
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,0,6),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r0, endIndex:r0+1 },
      properties:{ pixelSize:26 }, fields:'pixelSize' }
    });
    vals.push({ range:`${TAB}!A${r1}:F${r1}`, values:[[idx+1, metric, p1, p2, combined, note]] });
  });

  /* ── 10. Action Plan ─────────────────────────────────────────────────────── */
  const ACTH = ['#','Action Item','Category','Owner','Priority','Target\nDate','Status','Impact','Notes'];
  ACTH.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_ACTCOL,R_ACTCOL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:R_ACTCOL, endIndex:R_ACTCOL+1 },
    properties:{ pixelSize:28 }, fields:'pixelSize' }
  });

  const PRIO_C = { High:C.attention, Medium:C.warning, Low:C.secText };
  // [action, category, owner, priority, target, status, impact, note]
  const ACTIONS = [
    ['Maximize 401(k) & HSA contributions annually',           'Financial','Both','High','2027','In Progress','High','Leverage tax-advantaged growth in accumulation phase'],
    ['Review & rebalance portfolio to target allocation',      'Financial','Both','High','Quarterly','In Progress','High','Ensure risk exposure matches 11-year horizon'],
    ['Execute Roth conversion ladder for P1 IRA',             'Financial','P1','High','2030','Not Started','High','Reduce future RMD burden; fill lower tax brackets'],
    ['Purchase long-term care insurance (both)',               'Health',   'Both','High','2027','In Progress','High','Premiums increase after age 57; buy now'],
    ['Accelerate mortgage payoff ($500/mo extra)',             'Housing',  'P1','High','2032','Not Started','Medium','Align payoff with age-62 milestone'],
    ['Update beneficiary designations on all accounts',       'Legal',    'Both','High','2026','In Progress','Medium','Post-trust creation, most accounts need updating'],
    ['Enroll P1 in Medicare Part A & B (3 mo before 65)',     'Health',   'P1','High','2035','Not Started','High','Penalty-free enrollment window closes quickly'],
    ['Establish monthly retirement income simulation',         'Financial','Both','Medium','2028','Not Started','Medium','Model the "paycheck" cash flow from multiple sources'],
    ['Open 529 accounts for grandchildren at retirement',     'Financial','Both','Medium','2037','Not Started','Low','Seed fund; annual gifting strategy'],
    ['Research and select retirement location / downsize',    'Housing',  'Both','Medium','2033','Not Started','Medium','55+ communities, coastal options, $650K budget'],
    ['Finalize volunteer / hobby / community plan',           'Lifestyle','Both','Low','2036','Not Started','Low','Structure post-retirement identity and routine'],
  ];
  ACTIONS.forEach(([action, cat, owner, prio, target, status, impact, note], idx) => {
    const r0 = R_ACT0 + idx;
    const r1 = r0 + 1;
    const isAlt = idx % 2 === 1;
    const bg = status === 'In Progress' ? '#E8EEF6' : (isAlt ? C.altRow : C.panel);
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,0,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    // Priority col E: colored text
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,4,5),
      cell:{ userEnteredFormat:{ textFormat:{ bold:true, foregroundColor:hex(PRIO_C[prio]||C.secText), fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER' }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r0, endIndex:r0+1 },
      properties:{ pixelSize:36 }, fields:'pixelSize' }
    });
    vals.push({ range:`${TAB}!A${r1}:I${r1}`, values:[[idx+1, action, cat, owner, prio, target, status, impact, note]] });
  });

  /* ── 11. Disclaimer ──────────────────────────────────────────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,R_DISC,R_DISC+1,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,R_DISC,R_DISC+1,0,9),
    cell:{ userEnteredValue:{ stringValue:'Readiness scores are estimates for planning purposes only. Consult a licensed financial planner before making retirement decisions.' },
      userEnteredFormat:{ backgroundColor:hex(C.bg),
        textFormat:{ foregroundColor:hex(C.secText), italic:true, fontSize:8, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── 12. Conditional formats ─────────────────────────────────────────────── */
  // Scorecard: score >= 4 → green, score < 3 → red
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET, R_SC0, R_SC0+8, 3, 4) ],
    booleanRule:{ condition:{ type:'NUMBER_GREATER_THAN_EQ', values:[{ userEnteredValue:'4' }] },
      format:{ textFormat:{ foregroundColor:hex(C.secondaryDk), bold:true }, backgroundColor:hex('#EAF5EE') } }
  }, index:0 }});
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET, R_SC0, R_SC0+8, 3, 4) ],
    booleanRule:{ condition:{ type:'NUMBER_LESS', values:[{ userEnteredValue:'3' }] },
      format:{ textFormat:{ foregroundColor:hex(C.attention), bold:true }, backgroundColor:hex('#FBF0EF') } }
  }, index:1 }});
  // Scenario gap: negative (shortfall) → red text
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET, R_SCN0, R_SCN0+6, 8, 9) ],
    booleanRule:{ condition:{ type:'NUMBER_LESS', values:[{ userEnteredValue:'0' }] },
      format:{ textFormat:{ foregroundColor:hex(C.attention), bold:true } } }
  }, index:2 }});

  /* ── 13. Charts ──────────────────────────────────────────────────────────── */
  // Chart 1: Readiness scorecard BAR (scores per dimension)
  const chart1 = {
    addChart:{ chart:{
      spec:{
        title:'Readiness Score by Dimension',
        basicChart:{
          chartType:'BAR',
          legendPosition:'NO_LEGEND',
          axis:[
            { position:'BOTTOM_AXIS', title:'Score (1–5)' },
            { position:'LEFT_AXIS',   title:'Dimension' },
          ],
          domains:[{ domain:{ sourceRange:{ sources:[{
            sheetId:SHEET, startRowIndex:R_SC0, endRowIndex:R_SC0+8,
            startColumnIndex:1, endColumnIndex:2,
          }]}}}],
          series:[{ series:{ sourceRange:{ sources:[{
            sheetId:SHEET, startRowIndex:R_SC0, endRowIndex:R_SC0+8,
            startColumnIndex:3, endColumnIndex:4,
          }]}}, targetAxis:'BOTTOM_AXIS', color:hex(C.secondary) }],
          headerCount:0,
        },
      },
      position:{ overlayPosition:{
        anchorCell:{ sheetId:SHEET, rowIndex:R_SCHDR, columnIndex:9 },
        offsetXPixels:10, offsetYPixels:0, widthPixels:400, heightPixels:280,
      }},
    }}
  };

  // Chart 2: Scenario portfolio comparison COLUMN
  const chart2 = {
    addChart:{ chart:{
      spec:{
        title:'Portfolio at Retirement — by Scenario',
        basicChart:{
          chartType:'COLUMN',
          legendPosition:'NO_LEGEND',
          axis:[
            { position:'BOTTOM_AXIS', title:'Scenario' },
            { position:'LEFT_AXIS',   title:'Portfolio Balance ($)' },
          ],
          domains:[{ domain:{ sourceRange:{ sources:[{
            sheetId:SHEET, startRowIndex:R_SCN0, endRowIndex:R_SCN0+6,
            startColumnIndex:1, endColumnIndex:2,
          }]}}}],
          series:[{ series:{ sourceRange:{ sources:[{
            sheetId:SHEET, startRowIndex:R_SCN0, endRowIndex:R_SCN0+6,
            startColumnIndex:5, endColumnIndex:6,
          }]}}, targetAxis:'LEFT_AXIS', color:hex(C.SCN01) }],
          headerCount:0,
        },
      },
      position:{ overlayPosition:{
        anchorCell:{ sheetId:SHEET, rowIndex:R_SCNHDR, columnIndex:9 },
        offsetXPixels:10, offsetYPixels:0, widthPixels:400, heightPixels:260,
      }},
    }}
  };

  /* ── 14. Freeze ──────────────────────────────────────────────────────────── */
  requests.push({ updateSheetProperties:{ properties:{ sheetId:SHEET,
    gridProperties:{ frozenRowCount:7 }}, fields:'gridProperties.frozenRowCount' }
  });

  /* ── Execute ─────────────────────────────────────────────────────────────── */
  await batchUpdate(SID, requests, '[11-readiness format]');
  await valuesBatchUpdate(SID, vals, '[11-readiness values]');
  await batchUpdate(SID, [chart1, chart2], '[11-readiness charts]');

  console.log(`✅ Retirement Readiness done — 8-dim scorecard, scenario comparison, action plan, 2 charts.`);
}
run().catch(e=>{ console.error(e); process.exit(1); });
