'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs'), path = require('path');
const SID   = JSON.parse(fs.readFileSync(path.join(__dirname,'spreadsheet.json'))).id;
const SHEET = 8;
const TAB   = 'Withdrawal Strategy';
const SETUP = "'Personal & Household Setup'";
const FCAST = "'Growth Forecast'";

/* ── row index constants (0-indexed) ─────────────────────────────────────── */
const R_TITLE  = 0;  // Title
const R_SUB    = 1;  // Subtitle
const R_KL     = 2;  // KPI labels
const R_KV     = 3;  // KPI values
const R_SP1    = 4;  // Spacer
const R_STRHD  = 5;  // Strategy section header
const R_STRCOL = 6;  // Strategy column headers
const R_STR0   = 7;  // Strategy rows 7-12
const R_SP2    = 13; // Spacer
const R_BKTHD  = 14; // Bucket section header
const R_BKTCOL = 15; // Bucket column headers
const R_BKT0   = 16; // Bucket rows 16-18
const R_BKTTOT = 19; // Bucket totals
const R_SP3    = 20; // Spacer
const R_SCHDHD = 21; // Schedule section header
const R_SCHCOL = 22; // Schedule column headers
const R_SCHD0  = 23; // Schedule data rows 23-51 (29 years 2037-2065)
const R_SP4    = 52; // Spacer after schedule
const R_RMDHD  = 53; // RMD section header
const R_RMDCOL = 54; // RMD column headers
const R_RMD0   = 55; // RMD data rows 55-75 (ages 70-90)
const R_SP5    = 76; // Spacer
const R_DISC   = 77; // Disclaimer

/* ── column widths (A-I, 9 cols) ─────────────────────────────────────────── */
const COL_W = [70, 185, 210, 115, 140, 130, 120, 140, 160];

/* ── formula fragments (NO leading =) ───────────────────────────────────── */
const portExpr    = `IFERROR(VLOOKUP(YEAR(${SETUP}!B8)+${SETUP}!B10,${FCAST}!$A$7:$K$47,11,0),0)`;
const annGoalExpr = `${SETUP}!B30`;
const safeRateExpr= `${SETUP}!B28`;

async function run() {
  const requests = [];
  const vals     = [];

  /* ── 1. Column widths & row heights ────────────────────────────────────── */
  COL_W.forEach((w,i) => requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
    properties:{ pixelSize:w }, fields:'pixelSize' }
  }));
  [[R_TITLE,52],[R_SUB,22],[R_SP1,6],[R_SP2,6],[R_SP3,6],[R_SP4,6],[R_SP5,6],
   [R_STRHD,26],[R_BKTHD,26],[R_SCHDHD,26],[R_RMDHD,26],[R_DISC,24]].forEach(([r,h]) => {
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r, endIndex:r+1 },
      properties:{ pixelSize:h }, fields:'pixelSize' }
    });
  });

  /* ── 2. Title & subtitle ────────────────────────────────────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,R_TITLE,R_TITLE+1,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,R_TITLE,R_TITLE+1,0,9),
    cell:{ userEnteredValue:{ stringValue:`💸  ${TAB}` },
      userEnteredFormat:{ backgroundColor:hex(C.primary),
        textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:16, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });
  requests.push({ mergeCells:{ range:gridRange(SHEET,R_SUB,R_SUB+1,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,R_SUB,R_SUB+1,0,9),
    cell:{ userEnteredValue:{ stringValue:'Compare withdrawal strategies and plan sustainable retirement income distributions' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrB),
        textFormat:{ foregroundColor:hex(C.primaryText), italic:true, fontSize:10, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── 3. KPI strip (9 cards, rows 2-3) ──────────────────────────────────── */
  const kpis = [
    { label:'Portfolio\nat Retirement',  formula:`=${portExpr}`,                                      color:C.primary,     tc:C.primaryText },
    { label:'4% Rule Annual\nWithdrawal',formula:`=IFERROR(${portExpr}*${safeRateExpr},0)`,           color:C.secondary,   tc:C.text },
    { label:'4% Rule\nMonthly',          formula:`=IFERROR(${portExpr}*${safeRateExpr}/12,0)`,        color:C.secondaryDk, tc:C.primaryText },
    { label:'Annual\nIncome Goal',       formula:`=${annGoalExpr}`,                                    color:C.hdrB,        tc:C.primaryText },
    { label:'Monthly\nIncome Goal',      formula:`=IFERROR(${annGoalExpr}/12,0)`,                     color:C.hdrC,        tc:C.primaryText },
    { label:'Surplus /\nDeficit (4%)',   formula:`=IFERROR(${portExpr}*${safeRateExpr}-${annGoalExpr},0)`, color:C.info,   tc:C.text },
    { label:'Est. Years\nFunded',        formula:`=IFERROR(${portExpr}/${annGoalExpr},0)`,            color:C.altRow,      tc:C.text },
    { label:'P1 RMD\nStart Year',        formula:`=IFERROR(YEAR(${SETUP}!B8)+73,"")`,                 color:C.panel,       tc:C.text },
    { label:'P2 RMD\nStart Year',        formula:`=IFERROR(YEAR(${SETUP}!E8)+73,"")`,                 color:C.panel,       tc:C.text },
  ];
  kpis.forEach((k,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_KL,R_KL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:k.label },
        userEnteredFormat:{ backgroundColor:hex(k.color),
          textFormat:{ foregroundColor:hex(k.tc), bold:true, fontSize:8, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          wrapStrategy:'WRAP',
          borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_KV,R_KV+1,i,i+1),
      cell:{ userEnteredValue:{ formulaValue:k.formula },
        userEnteredFormat:{ backgroundColor:hex(k.color),
          textFormat:{ foregroundColor:hex(k.tc), bold:true, fontSize:13, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          numberFormat: i<=5 ? { type:'CURRENCY', pattern:'"$"#,##0' } : (i===6 ? { type:'NUMBER', pattern:'0.0" yrs"' } : { type:'NUMBER', pattern:'0' }) }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });

  /* ── 4. Spacer rows (set bg) ───────────────────────────────────────────── */
  [R_SP1,R_SP2,R_SP3,R_SP4,R_SP5].forEach(r => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,0,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(C.bg) }}, fields:'userEnteredFormat' }
    });
  });

  /* ── 5. Section headers ─────────────────────────────────────────────────── */
  function sectionHdr(r, text) {
    requests.push({ mergeCells:{ range:gridRange(SHEET,r,r+1,0,9), mergeType:'MERGE_ALL' }});
    requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,0,9),
      cell:{ userEnteredValue:{ stringValue:text },
        userEnteredFormat:{ backgroundColor:hex(C.hdrB),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:10, fontFamily:'Montserrat' },
          horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:10 } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  }
  sectionHdr(R_STRHD,  'Withdrawal Strategy Comparison');
  sectionHdr(R_BKTHD,  'Bucket Strategy Allocation');
  sectionHdr(R_SCHDHD, 'Annual Distribution Schedule — Retirement Phase (2037–2065)');
  sectionHdr(R_RMDHD,  'IRS RMD Reference Table — Uniform Lifetime (SECURE 2.0)');

  /* ── 6. Strategy comparison headers ─────────────────────────────────────── */
  const STRHDR = ['#','Strategy','Method / Rule','Rate','Annual Withdrawal','Monthly','Est. Duration','Key Strength','Limitation'];
  STRHDR.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_STRCOL,R_STRCOL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment: i<=1||i>=6 ? 'LEFT':'CENTER', verticalAlignment:'MIDDLE',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:R_STRCOL, endIndex:R_STRCOL+1 },
    properties:{ pixelSize:30 }, fields:'pixelSize' }
  });

  /* ── 7. Strategy rows ────────────────────────────────────────────────────── */
  // [name, method, rateVal (formula or number), annualFormula, estYrs, strength, limitation]
  const strategies = [
    ['4% Rule (Bengen)',      'Withdraw 4% of retirement-date portfolio; inflate each year',
     `=${SETUP}!B28`,
     `=IFERROR(${portExpr}*${SETUP}!B28,0)`,
     '30+ yrs','Simple, research-backed, inflation-protected','Fixed draw ignores market performance'],
    ['Dynamic (Guyton–Klinger)','Adjust withdrawal by ±10% based on portfolio performance',
     `=0.045`,
     `=IFERROR(${portExpr}*0.045,0)`,
     '35+ yrs','Responsive to market; higher initial draw','Complex; requires annual recalculation'],
    ['Fixed Dollar Amount',    'Withdraw same nominal dollar amount; adjust at review years',
     `=IFERROR(${annGoalExpr}/${portExpr},0)`,
     `=${annGoalExpr}`,
     '25–30 yrs','Predictable income; easy to budget','Portfolio risk if returns disappoint'],
    ['Conservative 3.5% Rule', 'Withdraw 3.5% of current portfolio each year',
     `=0.035`,
     `=IFERROR(${portExpr}*0.035,0)`,
     '40+ yrs','Very high survival probability; preserves estate','Lower income; may leave excess wealth'],
    ['RMD-Based',              'Required Minimum Distributions per IRS table starting age 73',
     `=IFERROR(1/VLOOKUP(MAX(73,YEAR(TODAY())-YEAR(${SETUP}!B8)),$A$56:$B$76,2,FALSE),0)`,
     `=IFERROR(${portExpr}/VLOOKUP(73,$A$56:$B$76,2,FALSE),0)`,
     'Life','IRS-mandated; coordinates tax strategy','Fluctuates with portfolio; forced when large'],
    ['Bucket Strategy',        '3-bucket: cash (1–3 yrs), bonds (4–10 yrs), equities (10+ yrs)',
     `=IFERROR(${annGoalExpr}/${portExpr},0)`,
     `=${annGoalExpr}`,
     '30+ yrs','Psychological comfort; sequencing-risk control','Rebalancing complexity; drag in bull markets'],
  ];

  strategies.forEach(([name, method, rateF, annF, dur, strength, limit], idx) => {
    const r = R_STR0 + idx;
    const isAlt = idx % 2 === 1;
    const bg = isAlt ? C.altRow : C.panel;
    requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,0,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r, endIndex:r+1 },
      properties:{ pixelSize:40 }, fields:'pixelSize' }
    });
    // Number formatting for rate col (D=3) and annual/monthly (E=4, F=5)
    requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,3,4),
      cell:{ userEnteredFormat:{ numberFormat:{ type:'PERCENT', pattern:'0.00%' }, horizontalAlignment:'CENTER' }},
      fields:'userEnteredFormat' }
    });
    [[4,'CURRENCY','$#,##0'],[5,'CURRENCY','$#,##0']].forEach(([col,type,pat]) => {
      requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,col,col+1),
        cell:{ userEnteredFormat:{ numberFormat:{ type, pattern:pat }, horizontalAlignment:'CENTER' }},
        fields:'userEnteredFormat' }
      });
    });
    vals.push({ range:`${TAB}!A${r+1}:I${r+1}`, values:[[idx+1, name, method, rateF, annF, `=E${r+1}/12`, dur, strength, limit]] });
  });

  /* ── 8. Bucket strategy table ───────────────────────────────────────────── */
  const BKTHDR = ['Bucket','Label','Time Horizon','Allocation','Amount','Asset Class','Purpose','Return Target','Rebalance Trigger'];
  BKTHDR.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_BKTCOL,R_BKTCOL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:R_BKTCOL, endIndex:R_BKTCOL+1 },
    properties:{ pixelSize:30 }, fields:'pixelSize' }
  });

  const BUCKET_COLORS = [C.attention, C.secondary, C.primary];
  const buckets = [
    ['1','Cash & Short-Term','Years 1–3','0.10',`=IFERROR(${portExpr}*0.10,0)`,
     'Cash, CDs, T-Bills','Immediate liquidity; no sequence risk','1–2%','Bucket emptied by year 3'],
    ['2','Conservative / Income','Years 4–10','0.25',`=IFERROR(${portExpr}*0.25,0)`,
     'Bonds, Div. Stocks','Income generation; refill Bucket 1','3–5%','Bucket 1 down to 6 months'],
    ['3','Growth / Equities','Years 10+','0.65',`=IFERROR(${portExpr}*0.65,0)`,
     'Stocks, REITs, Alternatives','Long-term growth; replenish Bucket 2','7–9%','Annual or after 20%+ gain'],
  ];
  buckets.forEach(([num,label,horizon,allocPct,allocAmt,assetClass,purpose,returnTgt,trigger], idx) => {
    const r = R_BKT0 + idx;
    const bg = [C.taxable, C.formula, C.preTax][idx];
    requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,0,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateBorders:{ range:gridRange(SHEET,r,r+1,0,1),
      left:{ style:'SOLID_MEDIUM', color:hex(BUCKET_COLORS[idx]) } }
    });
    requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,3,4),
      cell:{ userEnteredFormat:{ numberFormat:{ type:'PERCENT', pattern:'0%' }, horizontalAlignment:'CENTER' }},
      fields:'userEnteredFormat' }
    });
    requests.push({ repeatCell:{ range:gridRange(SHEET,r,r+1,4,5),
      cell:{ userEnteredFormat:{ numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }, horizontalAlignment:'CENTER' }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r, endIndex:r+1 },
      properties:{ pixelSize:38 }, fields:'pixelSize' }
    });
    vals.push({ range:`${TAB}!A${r+1}:I${r+1}`, values:[[num,label,horizon,allocPct,allocAmt,assetClass,purpose,returnTgt,trigger]] });
  });

  // Bucket totals row
  const btr = R_BKTTOT;
  requests.push({ repeatCell:{ range:gridRange(SHEET,btr,btr+1,0,9),
    cell:{ userEnteredFormat:{ backgroundColor:hex(C.hdrC),
      textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontFamily:'Montserrat', fontSize:9 },
      verticalAlignment:'MIDDLE' }},
    fields:'userEnteredFormat' }
  });
  vals.push({ range:`${TAB}!A${btr+1}:I${btr+1}`, values:[['','Total','','=SUM(D17:D19)',`=IFERROR(${portExpr},0)`,'','','','']] });
  requests.push({ repeatCell:{ range:gridRange(SHEET,btr,btr+1,3,4),
    cell:{ userEnteredFormat:{ numberFormat:{ type:'PERCENT', pattern:'0%' }}}, fields:'userEnteredFormat' }
  });
  requests.push({ repeatCell:{ range:gridRange(SHEET,btr,btr+1,4,5),
    cell:{ userEnteredFormat:{ numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }}}, fields:'userEnteredFormat' }
  });

  /* ── 9. Annual Distribution Schedule ───────────────────────────────────── */
  const SCHHDR = ['Year','P1 Age','P2 Age','Portfolio Balance','Annual Expenses','External Income','4% Withdrawal','Net Income Need','RMD Estimate'];
  SCHHDR.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_SCHCOL,R_SCHCOL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:R_SCHCOL, endIndex:R_SCHCOL+1 },
    properties:{ pixelSize:30 }, fields:'pixelSize' }
  });

  for (let i = 0; i < 29; i++) {
    const yr  = 2037 + i;
    const r0  = R_SCHD0 + i; // 0-indexed
    const r1  = r0 + 1;       // 1-indexed
    const isAlt = i % 2 === 1;
    const bg  = isAlt ? C.altRow : C.panel;

    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,0,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r0, endIndex:r0+1 },
      properties:{ pixelSize:24 }, fields:'pixelSize' }
    });

    // Number formats for currency cols D-I (indices 3-8)
    [[3,'$#,##0'],[4,'$#,##0'],[5,'$#,##0'],[6,'$#,##0'],[7,'$#,##0'],[8,'$#,##0']].forEach(([col,pat]) => {
      requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,col,col+1),
        cell:{ userEnteredFormat:{ numberFormat:{ type:'CURRENCY', pattern:pat }, horizontalAlignment:'RIGHT' }},
        fields:'userEnteredFormat' }
      });
    });

    const portBal = `IFERROR(VLOOKUP(A${r1},${FCAST}!$A$7:$K$47,11,0),0)`;
    const expenses = `IFERROR(VLOOKUP(A${r1},${FCAST}!$A$7:$I$47,9,0),0)`;
    const income   = `IFERROR(VLOOKUP(A${r1},${FCAST}!$A$7:$H$47,8,0),0)`;
    const rmdFactor= `IFERROR(VLOOKUP(B${r1},$A$56:$B$76,2,FALSE),27.4)`;

    vals.push({ range:`${TAB}!A${r1}:I${r1}`, values:[[
      yr,
      `=IFERROR(A${r1}-YEAR(${SETUP}!B8),"")`,
      `=IFERROR(A${r1}-YEAR(${SETUP}!E8),"")`,
      `=${portBal}`,
      `=${expenses}`,
      `=${income}`,
      `=IFERROR(${portBal}*${SETUP}!B28,0)`,
      `=IFERROR(MAX(0,E${r1}-F${r1}),0)`,
      `=IFERROR(IF(B${r1}>=73,${portBal}/${rmdFactor},0),0)`,
    ]]});
  }

  // Totals row for schedule
  const schtot = R_SP4 - 1; // row just before spacer
  // no totals needed — leave as data only

  /* ── 10. RMD reference table ────────────────────────────────────────────── */
  const RMDHDR = ['Age','Distribution Factor','Age','Factor','Age','Factor','Age','Factor','Note'];
  RMDHDR.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,R_RMDCOL,R_RMDCOL+1,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:R_RMDCOL, endIndex:R_RMDCOL+1 },
    properties:{ pixelSize:26 }, fields:'pixelSize' }
  });

  // IRS Uniform Lifetime Table (ages 70–90)
  const RMD_TABLE = [
    [70,27.4],[71,26.5],[72,25.5],[73,24.6],[74,23.7],
    [75,22.9],[76,22.0],[77,21.1],[78,20.2],[79,19.4],
    [80,18.5],[81,17.7],[82,16.8],[83,16.0],[84,15.2],
    [85,14.4],[86,13.7],[87,12.9],[88,12.2],[89,11.4],
    [90,10.8],
  ];
  RMD_TABLE.forEach(([age, factor], idx) => {
    const r0 = R_RMD0 + idx;
    const r1 = r0 + 1;
    const isAlt = idx % 2 === 1;
    const bg = isAlt ? C.altRow : C.panel;
    const note = age === 73 ? 'RMD starts (SECURE 2.0)' : age === 75 ? 'Higher factor applies' : '';
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,0,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r0, endIndex:r0+1 },
      properties:{ pixelSize:22 }, fields:'pixelSize' }
    });
    vals.push({ range:`${TAB}!A${r1}:I${r1}`, values:[[age, factor, '', '', '', '', '', '', note]] });
  });

  /* ── 11. Disclaimer ─────────────────────────────────────────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,R_DISC,R_DISC+1,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,R_DISC,R_DISC+1,0,9),
    cell:{ userEnteredValue:{ stringValue:'Withdrawal estimates are for planning purposes only. Tax treatment, RMD rules, and portfolio returns will vary. Consult a licensed financial advisor.' },
      userEnteredFormat:{ backgroundColor:hex(C.bg),
        textFormat:{ foregroundColor:hex(C.secText), italic:true, fontSize:8, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── 12. Conditional formats ────────────────────────────────────────────── */
  // Schedule: high RMD (col I > 4% withdrawal col G) → amber row
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET, R_SCHD0, R_SCHD0+29, 0, 9) ],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=AND(I24>0,I24>G24)` }] },
      format:{ backgroundColor:hex('#FBF4E4') } }
  }, index:0 }});
  // Schedule: net income need > 4% withdrawal (possible shortfall) → attention red text
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET, R_SCHD0, R_SCHD0+29, 7, 8) ],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=H24>G24` }] },
      format:{ textFormat:{ foregroundColor:hex(C.attention), bold:true } } }
  }, index:1 }});

  /* ── 13. Freeze ─────────────────────────────────────────────────────────── */
  requests.push({ updateSheetProperties:{ properties:{ sheetId:SHEET,
    gridProperties:{ frozenRowCount:4 }}, fields:'gridProperties.frozenRowCount' }
  });

  /* ── 14. Charts ─────────────────────────────────────────────────────────── */
  // Chart 1: Strategy annual withdrawal comparison (bar)
  // Source: strategy table E8:E13 (annual withdrawal) vs B8:B13 (name)
  // Use the strategy rows, anchor to the right of strategy table
  const chart1 = {
    addChart:{ chart:{
      spec:{
        title:'Annual Withdrawal by Strategy',
        basicChart:{
          chartType:'BAR',
          legendPosition:'BOTTOM_LEGEND',
          axis:[
            { position:'BOTTOM_AXIS', title:'Annual Withdrawal ($)' },
            { position:'LEFT_AXIS',   title:'Strategy' },
          ],
          domains:[{ domain:{ sourceRange:{ sources:[{
            sheetId:SHEET, startRowIndex:R_STR0, endRowIndex:R_STR0+6,
            startColumnIndex:1, endColumnIndex:2,
          }]}}}],
          series:[{ series:{ sourceRange:{ sources:[{
            sheetId:SHEET, startRowIndex:R_STR0, endRowIndex:R_STR0+6,
            startColumnIndex:4, endColumnIndex:5,
          }]}}, targetAxis:'BOTTOM_AXIS', color:hex(C.secondary) }],
          headerCount:0,
        },
      },
      position:{ overlayPosition:{
        anchorCell:{ sheetId:SHEET, rowIndex:R_STRHD, columnIndex:9 },
        offsetXPixels:10, offsetYPixels:0, widthPixels:420, heightPixels:260,
      }},
    }}
  };

  // Chart 2: Annual schedule — 4% withdrawal vs net income need over time
  // Source: schedule col A (year) + cols G and H
  // Build a helper data strip to the right (col J-L = 10-12) for clean chart source
  const HELP_COL = 10; // col K (0-indexed)
  const schedYrCol  = 10; // col K: Year
  const sched4PctCol= 11; // col L: 4% withdrawal
  const schedNeedCol= 12; // col M: Net Income Need
  for (let i = 0; i < 29; i++) {
    const r0 = R_SCHD0 + i;
    const r1 = r0 + 1;
    vals.push({ range:`${TAB}!K${r1}:M${r1}`, values:[[`=A${r1}`, `=G${r1}`, `=H${r1}`]] });
  }

  const chart2 = {
    addChart:{ chart:{
      spec:{
        title:'Withdrawal vs. Income Need Over Retirement (2037–2065)',
        basicChart:{
          chartType:'COLUMN',
          legendPosition:'BOTTOM_LEGEND',
          axis:[
            { position:'BOTTOM_AXIS', title:'Year' },
            { position:'LEFT_AXIS',   title:'Amount ($)' },
          ],

          domains:[{ domain:{ sourceRange:{ sources:[{
            sheetId:SHEET, startRowIndex:R_SCHD0, endRowIndex:R_SCHD0+29,
            startColumnIndex:schedYrCol, endColumnIndex:schedYrCol+1,
          }]}}}],
          series:[
            { series:{ sourceRange:{ sources:[{
              sheetId:SHEET, startRowIndex:R_SCHD0, endRowIndex:R_SCHD0+29,
              startColumnIndex:sched4PctCol, endColumnIndex:sched4PctCol+1,
            }]}}, targetAxis:'LEFT_AXIS', color:hex(C.secondary) },
            { series:{ sourceRange:{ sources:[{
              sheetId:SHEET, startRowIndex:R_SCHD0, endRowIndex:R_SCHD0+29,
              startColumnIndex:schedNeedCol, endColumnIndex:schedNeedCol+1,
            }]}}, targetAxis:'LEFT_AXIS', color:hex(C.attention) },
          ],
          headerCount:0,
        },
      },
      position:{ overlayPosition:{
        anchorCell:{ sheetId:SHEET, rowIndex:R_BKTHD, columnIndex:9 },
        offsetXPixels:10, offsetYPixels:0, widthPixels:420, heightPixels:280,
      }},
    }}
  };

  /* ── Execute ─────────────────────────────────────────────────────────── */
  await batchUpdate(SID, requests, '[10-withdrawal format]');
  await valuesBatchUpdate(SID, vals, '[10-withdrawal values]');
  await batchUpdate(SID, [chart1, chart2], '[10-withdrawal charts]');

  console.log(`✅ Withdrawal Strategy done — 6 strategies, bucket model, 29-year schedule, 2 charts.`);
}
run().catch(e=>{ console.error(e); process.exit(1); });
