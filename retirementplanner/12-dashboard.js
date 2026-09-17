'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs'), path = require('path');
const SID   = JSON.parse(fs.readFileSync(path.join(__dirname,'spreadsheet.json'))).id;
const SHEET = 10;
const TAB   = 'Retirement Dashboard';
const SETUP = "'Personal & Household Setup'";
const ACCTS = "'Retirement Accounts'";
const INC   = "'Retirement Income'";
const MST   = "'Milestone Tracker'";
const SCEN  = "'Scenario Planner'";
const FCAST = "'Growth Forecast'";
const READY = "'Retirement Readiness'";

// Other sheet IDs for cross-sheet chart references
const SID_FCAST = 6;
const SID_SCEN  = 5;
const SID_READY = 9;

/* ── column widths A-I ───────────────────────────────────────────────────── */
const COL_W = [85, 165, 120, 125, 110, 115, 110, 120, 140];

/* ── formula fragments (no leading =) ──────────────────────────────────────*/
const portNow    = `IFERROR(SUMIF(${ACCTS}!L7:L100,"Active",${ACCTS}!F7:F100),0)`;
const portFV     = `IFERROR(VLOOKUP(YEAR(${SETUP}!B8)+${SETUP}!B10,${FCAST}!$A$7:$K$47,11,0),0)`;
const annGoal    = `${SETUP}!B30`;
const nestEgg    = `${SETUP}!B31`;
const safeRate   = `${SETUP}!B28`;
const readyScore = `${READY}!D19`;

async function run() {
  const requests = [];
  const vals     = [];

  /* ── 1. Column widths & title row height ────────────────────────────────── */
  COL_W.forEach((w,i) => requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
    properties:{ pixelSize:w }, fields:'pixelSize' }
  }));
  // Hide helper cols K-Z from view (width 1)
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'COLUMNS', startIndex:9, endIndex:30 },
    properties:{ pixelSize:1 }, fields:'pixelSize' }
  });
  [[0,58],[1,22],[2,48],[3,48]].forEach(([r,h]) => {
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r, endIndex:r+1 },
      properties:{ pixelSize:h }, fields:'pixelSize' }
    });
  });

  /* ── 2. Title & subtitle ─────────────────────────────────────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,0,1,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,0,1,0,9),
    cell:{ userEnteredValue:{ stringValue:`📊  ${TAB}` },
      userEnteredFormat:{ backgroundColor:hex(C.primary),
        textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:18, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });
  requests.push({ mergeCells:{ range:gridRange(SHEET,1,2,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,1,2,0,9),
    cell:{ userEnteredValue:{ stringValue:'James & Patricia Whitmore — Retirement Command Center' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrB),
        textFormat:{ foregroundColor:hex(C.primaryText), italic:true, fontSize:11, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── 3. Big KPI panels (rows 2-3) ───────────────────────────────────────── */
  const bigKpis = [
    { label:'Current Portfolio Value',     formula:`=${portNow}`,      color:C.primary,     tc:C.primaryText, nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'Portfolio at Retirement 2037',formula:`=${portFV}`,       color:C.secondaryDk, tc:C.primaryText, nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'Readiness Score (0–5.0)',     formula:`=${readyScore}`,   color:C.hdrA,        tc:C.primaryText, nf:{ type:'NUMBER',   pattern:'0.0" / 5.0"' } },
  ];
  bigKpis.forEach(({ label, formula, color, tc, nf }, i) => {
    const c0 = i*3, c1 = c0+3;
    requests.push({ mergeCells:{ range:gridRange(SHEET,2,3,c0,c1), mergeType:'MERGE_ALL' }});
    requests.push({ repeatCell:{ range:gridRange(SHEET,2,3,c0,c1),
      cell:{ userEnteredValue:{ stringValue:label },
        userEnteredFormat:{ backgroundColor:hex(color),
          textFormat:{ foregroundColor:hex(tc), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          borders:{ bottom:{ style:'SOLID', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
    requests.push({ mergeCells:{ range:gridRange(SHEET,3,4,c0,c1), mergeType:'MERGE_ALL' }});
    requests.push({ repeatCell:{ range:gridRange(SHEET,3,4,c0,c1),
      cell:{ userEnteredValue:{ formulaValue:formula },
        userEnteredFormat:{ backgroundColor:hex(color),
          textFormat:{ foregroundColor:hex(tc), bold:true, fontSize:22, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          numberFormat:nf }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });

  /* ── 4. Spacer between big KPIs and medium KPIs ─────────────────────────── */
  requests.push({ repeatCell:{ range:gridRange(SHEET,4,5,0,9),
    cell:{ userEnteredFormat:{ backgroundColor:hex(C.bg) }}, fields:'userEnteredFormat' }
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:4, endIndex:5 },
    properties:{ pixelSize:5 }, fields:'pixelSize' }
  });

  /* ── 5. Medium KPI strip (rows 5-6) ─────────────────────────────────────── */
  const medKpis = [
    { label:'Target\nNest Egg',      formula:`=${nestEgg}`,                             color:C.hdrC,     tc:C.primaryText, nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'Annual\nIncome Goal',   formula:`=${annGoal}`,                             color:C.hdrB,     tc:C.primaryText, nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'Monthly\nIncome Goal',  formula:`=IFERROR(${annGoal}/12,0)`,              color:C.hdrC,     tc:C.primaryText, nf:{ type:'CURRENCY', pattern:'"$"#,##0' } },
    { label:'Surplus /\nShortfall',  formula:`=IFERROR(${portFV}-${nestEgg},0)`,       color:C.info,     tc:C.text,        nf:{ type:'CURRENCY', pattern:'"$"#,##0;[Red]-"$"#,##0' } },
    { label:'P1 Current\nAge',       formula:`=${SETUP}!B9`,                            color:C.altRow,   tc:C.text,        nf:{ type:'NUMBER', pattern:'0" yrs"' } },
    { label:'P2 Current\nAge',       formula:`=${SETUP}!E9`,                            color:C.altRow,   tc:C.text,        nf:{ type:'NUMBER', pattern:'0" yrs"' } },
    { label:'P1 Retire\nYear',       formula:`=${SETUP}!B11`,                           color:C.panel,    tc:C.text,        nf:{ type:'NUMBER', pattern:'0' } },
    { label:'P2 Retire\nYear',       formula:`=${SETUP}!E11`,                           color:C.panel,    tc:C.text,        nf:{ type:'NUMBER', pattern:'0' } },
    { label:'Safe\nWithdrawal',      formula:`=${safeRate}`,                            color:C.formula,  tc:C.text,        nf:{ type:'PERCENT', pattern:'0.0%' } },
  ];
  medKpis.forEach((k,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,5,6,i,i+1),
      cell:{ userEnteredValue:{ stringValue:k.label },
        userEnteredFormat:{ backgroundColor:hex(k.color),
          textFormat:{ foregroundColor:hex(k.tc), bold:true, fontSize:8, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
          borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
    requests.push({ repeatCell:{ range:gridRange(SHEET,6,7,i,i+1),
      cell:{ userEnteredValue:{ formulaValue:k.formula },
        userEnteredFormat:{ backgroundColor:hex(k.color),
          textFormat:{ foregroundColor:hex(k.tc), bold:true, fontSize:13, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', numberFormat:k.nf }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:5, endIndex:6 },
    properties:{ pixelSize:28 }, fields:'pixelSize' }
  });

  /* ── 6. Thin spacer ──────────────────────────────────────────────────────── */
  requests.push({ repeatCell:{ range:gridRange(SHEET,7,8,0,9),
    cell:{ userEnteredFormat:{ backgroundColor:hex(C.bg) }}, fields:'userEnteredFormat' }
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:7, endIndex:8 },
    properties:{ pixelSize:5 }, fields:'pixelSize' }
  });

  /* ── 7. Left section: Portfolio by Account Type (rows 8-14, cols A-D) ────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,8,9,0,4), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,8,9,0,4),
    cell:{ userEnteredValue:{ stringValue:'Portfolio by Account Type' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrB),
        textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:10, fontFamily:'Montserrat' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:8 } }},
    fields:'userEnteredValue,userEnteredFormat' }
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:8, endIndex:9 },
    properties:{ pixelSize:25 }, fields:'pixelSize' }
  });

  /* ── Right section: Income Sources (rows 8-14, cols E-I) ─────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,8,9,4,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,8,9,4,9),
    cell:{ userEnteredValue:{ stringValue:'Retirement Income Sources' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrC),
        textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:10, fontFamily:'Montserrat' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:8 } }},
    fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── Column headers row 9 ────────────────────────────────────────────────── */
  const acctHdrs  = ['Type','Current Balance','FV at Retire','% of Total'];
  const incHdrs   = ['Source','Owner','Annual Amount','COLA Adj.','Status'];
  [...acctHdrs, ...incHdrs].forEach((h,i) => {
    const col = i < 4 ? i : i;
    requests.push({ repeatCell:{ range:gridRange(SHEET,9,10,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(i<4?C.hdrA:C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:9, endIndex:10 },
    properties:{ pixelSize:28 }, fields:'pixelSize' }
  });

  /* ── Account type data rows 10-14 ────────────────────────────────────────── */
  const actTypes = [
    { label:'Pre-Tax (401k / IRA)',     bg:C.preTax,
      balF:`=SUMPRODUCT((ISNUMBER(SEARCH("Pre-Tax",${ACCTS}!D7:D100)))*(${ACCTS}!L7:L100="Active")*(${ACCTS}!F7:F100))`,
      fvF: `=SUMPRODUCT((ISNUMBER(SEARCH("Pre-Tax",${ACCTS}!D7:D100)))*(${ACCTS}!L7:L100="Active")*(${ACCTS}!K7:K100))` },
    { label:'Roth (Roth IRA / 401k)',   bg:C.roth,
      balF:`=SUMPRODUCT((ISNUMBER(SEARCH("Roth",${ACCTS}!D7:D100)))*(${ACCTS}!L7:L100="Active")*(${ACCTS}!F7:F100))`,
      fvF: `=SUMPRODUCT((ISNUMBER(SEARCH("Roth",${ACCTS}!D7:D100)))*(${ACCTS}!L7:L100="Active")*(${ACCTS}!K7:K100))` },
    { label:'Taxable Brokerage',        bg:C.taxable,
      balF:`=SUMPRODUCT((${ACCTS}!D7:D100="Taxable Brokerage")*(${ACCTS}!L7:L100="Active")*(${ACCTS}!F7:F100))`,
      fvF: `=SUMPRODUCT((${ACCTS}!D7:D100="Taxable Brokerage")*(${ACCTS}!L7:L100="Active")*(${ACCTS}!K7:K100))` },
    { label:'Tax-Deferred Annuity',     bg:C.taxDeferred,
      balF:`=SUMPRODUCT((${ACCTS}!D7:D100="Tax-Deferred Annuity")*(${ACCTS}!F7:F100))`,
      fvF: `=SUMPRODUCT((${ACCTS}!D7:D100="Tax-Deferred Annuity")*(${ACCTS}!K7:K100))` },
    { label:'Total (Active Accounts)',  bg:C.hdrC, tc:C.primaryText, bold:true,
      balF:`=SUMIF(${ACCTS}!L7:L100,"Active",${ACCTS}!F7:F100)`,
      fvF: `=SUMIF(${ACCTS}!L7:L100,"Active",${ACCTS}!K7:K100)` },
  ];
  const totalBalRow = 14; // 1-indexed row 15 = row for totals
  actTypes.forEach(({ label, bg, tc=C.text, bold=false, balF, fvF }, idx) => {
    const r0 = 10 + idx;
    const r1 = r0 + 1;
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,0,4),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(tc), bold },
        verticalAlignment:'MIDDLE',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    [[1,'CURRENCY','"$"#,##0'],[2,'CURRENCY','"$"#,##0'],[3,'PERCENT','0%']].forEach(([ci,type,pat]) => {
      requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,ci,ci+1),
        cell:{ userEnteredFormat:{ numberFormat:{ type, pattern:pat }, horizontalAlignment:'RIGHT' }},
        fields:'userEnteredFormat' }
      });
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r0, endIndex:r0+1 },
      properties:{ pixelSize:26 }, fields:'pixelSize' }
    });
    vals.push({ range:`${TAB}!A${r1}:D${r1}`, values:[[
      label, balF, fvF,
      `=IFERROR(B${r1}/B15,0)`,
    ]] });
  });

  /* ── Income source data rows 10-15, cols E-I ─────────────────────────────── */
  for (let i = 0; i < 6; i++) {
    const r0 = 10 + i;
    const r1 = r0 + 1;
    const srcR = 7 + i; // Retirement Income row (1-indexed)
    const isAlt = i % 2 === 1;
    const bg = isAlt ? C.altRow : C.panel;
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,4,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(bg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,6,7),
      cell:{ userEnteredFormat:{ numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }, horizontalAlignment:'RIGHT' }},
      fields:'userEnteredFormat' }
    });
    vals.push({ range:`${TAB}!E${r1}:I${r1}`, values:[[
      `=${INC}!B${srcR}`,
      `=${INC}!C${srcR}`,
      `=${INC}!E${srcR}`,
      `=IFERROR(IF(${INC}!J${srcR}=TRUE,${INC}!K${srcR},"None"),"")`,
      `=${INC}!L${srcR}`,
    ]] });
  }

  /* ── 8. Spacer row 16 ────────────────────────────────────────────────────── */
  requests.push({ repeatCell:{ range:gridRange(SHEET,16,17,0,9),
    cell:{ userEnteredFormat:{ backgroundColor:hex(C.bg) }}, fields:'userEnteredFormat' }
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:16, endIndex:17 },
    properties:{ pixelSize:5 }, fields:'pixelSize' }
  });

  /* ── 9. Left: Key Milestones (rows 17-24, cols A-D) ─────────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,17,18,0,4), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,17,18,0,4),
    cell:{ userEnteredValue:{ stringValue:'Key Milestones' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrB),
        textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:10, fontFamily:'Montserrat' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:8 } }},
    fields:'userEnteredValue,userEnteredFormat' }
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:17, endIndex:18 },
    properties:{ pixelSize:25 }, fields:'pixelSize' }
  });

  /* ── Right: Scenario Summary (rows 17-24, cols E-I) ─────────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,17,18,4,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,17,18,4,9),
    cell:{ userEnteredValue:{ stringValue:'Scenario Summary' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrC),
        textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:10, fontFamily:'Montserrat' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:8 } }},
    fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── Column headers row 18 ───────────────────────────────────────────────── */
  const mstHdrs = ['Milestone','Category','Target Yr','Status'];
  const scnHdrs = ['Scenario','Portfolio FV','4% Annual','Gap / Surplus','Funded Yrs'];
  [...mstHdrs, ...scnHdrs].forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,18,19,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA),
          textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
          borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:18, endIndex:19 },
    properties:{ pixelSize:28 }, fields:'pixelSize' }
  });

  /* ── Milestone data rows 19-24 (first 6 milestones) ─────────────────────── */
  for (let i = 0; i < 6; i++) {
    const r0 = 19 + i;
    const r1 = r0 + 1;
    const mstR = 7 + i;
    const isAlt = i % 2 === 1;
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,0,4),
      cell:{ userEnteredFormat:{ backgroundColor:hex(isAlt?C.altRow:C.panel),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r0, endIndex:r0+1 },
      properties:{ pixelSize:28 }, fields:'pixelSize' }
    });
    vals.push({ range:`${TAB}!A${r1}:D${r1}`, values:[[
      `=${MST}!B${mstR}`,
      `=${MST}!C${mstR}`,
      `=${MST}!F${mstR}`,
      `=${MST}!H${mstR}`,
    ]] });
  }

  /* ── Scenario data rows 19-24 (6 scenarios), cols E-I ───────────────────── */
  const SCN_COLORS = [C.SCN01,C.SCN02,C.SCN03,C.SCN04,C.SCN05,C.SCN06];
  for (let i = 0; i < 6; i++) {
    const r0  = 19 + i;
    const r1  = r0 + 1;
    const scnR = 6 + i; // Scenario Planner row (1-indexed)
    requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,4,9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(C.panel),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateBorders:{ range:gridRange(SHEET,r0,r0+1,4,5),
      left:{ style:'SOLID_MEDIUM', color:hex(SCN_COLORS[i]) } }
    });
    [[5,'CURRENCY','"$"#,##0'],[6,'CURRENCY','"$"#,##0'],[7,'CURRENCY','+$#,##0;-$#,##0'],[8,'NUMBER','0.0" yrs"']].forEach(([ci,type,pat]) => {
      requests.push({ repeatCell:{ range:gridRange(SHEET,r0,r0+1,ci,ci+1),
        cell:{ userEnteredFormat:{ numberFormat:{ type, pattern:pat }, horizontalAlignment:'RIGHT' }},
        fields:'userEnteredFormat' }
      });
    });
    vals.push({ range:`${TAB}!E${r1}:I${r1}`, values:[[
      `=${SCEN}!B${scnR}`,
      `=${SCEN}!K${scnR}`,
      `=IFERROR(F${r1}*${safeRate},0)`,
      `=${SCEN}!N${scnR}`,
      `=IFERROR(F${r1}/${annGoal},0)`,
    ]] });
  }

  /* ── 10. Spacer + Disclaimer ──────────────────────────────────────────────── */
  requests.push({ repeatCell:{ range:gridRange(SHEET,25,26,0,9),
    cell:{ userEnteredFormat:{ backgroundColor:hex(C.bg) }}, fields:'userEnteredFormat' }
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:25, endIndex:26 },
    properties:{ pixelSize:5 }, fields:'pixelSize' }
  });
  requests.push({ mergeCells:{ range:gridRange(SHEET,26,27,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,26,27,0,9),
    cell:{ userEnteredValue:{ stringValue:'This dashboard is for planning purposes only. All projections are estimates. Consult a licensed financial advisor before making retirement decisions.' },
      userEnteredFormat:{ backgroundColor:hex(C.bg),
        textFormat:{ foregroundColor:hex(C.secText), italic:true, fontSize:8, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:26, endIndex:27 },
    properties:{ pixelSize:22 }, fields:'pixelSize' }
  });

  /* ── 11. Helper data for Account Type chart (off-screen cols K-L) ──────── */
  // K-L data at rows 30-35 (0-indexed), used as chart source on this tab
  // Headers at row 30, data rows 31-34
  vals.push({ range:`${TAB}!K31:L35`, values:[
    ['Account Type','Balance'],
    ['Pre-Tax',     `=SUMPRODUCT((ISNUMBER(SEARCH("Pre-Tax",${ACCTS}!D7:D100)))*(${ACCTS}!L7:L100="Active")*(${ACCTS}!F7:F100))`],
    ['Roth',        `=SUMPRODUCT((ISNUMBER(SEARCH("Roth",${ACCTS}!D7:D100)))*(${ACCTS}!L7:L100="Active")*(${ACCTS}!F7:F100))`],
    ['Taxable',     `=SUMIF(${ACCTS}!D7:D100,"Taxable Brokerage",${ACCTS}!F7:F100)`],
    ['Tax-Deferred',`=SUMIF(${ACCTS}!D7:D100,"Tax-Deferred Annuity",${ACCTS}!F7:F100)`],
  ]});

  /* ── 12. Conditional formats ──────────────────────────────────────────────── */
  // Scenario gap: negative → red, positive → green
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET, 19, 25, 7, 8) ],
    booleanRule:{ condition:{ type:'NUMBER_LESS', values:[{ userEnteredValue:'0' }] },
      format:{ textFormat:{ foregroundColor:hex(C.attention), bold:true } } }
  }, index:0 }});
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET, 19, 25, 7, 8) ],
    booleanRule:{ condition:{ type:'NUMBER_GREATER_THAN_EQ', values:[{ userEnteredValue:'0' }] },
      format:{ textFormat:{ foregroundColor:hex(C.secondaryDk), bold:true } } }
  }, index:1 }});
  // Milestone status: Completed → green, In Progress → blue
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET, 19, 25, 3, 4) ],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:'=D20="Completed"' }] },
      format:{ textFormat:{ foregroundColor:hex(C.secondaryDk), bold:true } } }
  }, index:2 }});

  /* ── 13. Freeze (title + subtitle + big KPIs) ─────────────────────────────── */
  requests.push({ updateSheetProperties:{ properties:{ sheetId:SHEET,
    gridProperties:{ frozenRowCount:7 }}, fields:'gridProperties.frozenRowCount' }
  });

  /* ── 14. Charts ───────────────────────────────────────────────────────────── */
  // Chart 1: Portfolio Growth Over Time (cross-sheet from Growth Forecast, sheetId 6)
  const chart1 = {
    addChart:{ chart:{
      spec:{
        title:'Portfolio Growth 2026–2065',
        basicChart:{
          chartType:'COLUMN',
          legendPosition:'NO_LEGEND',
          axis:[
            { position:'BOTTOM_AXIS', title:'Year' },
            { position:'LEFT_AXIS',   title:'Portfolio Value ($)' },
          ],
          domains:[{ domain:{ sourceRange:{ sources:[{
            sheetId:SID_FCAST, startRowIndex:6, endRowIndex:46,
            startColumnIndex:0, endColumnIndex:1,
          }]}}}],
          series:[{ series:{ sourceRange:{ sources:[{
            sheetId:SID_FCAST, startRowIndex:6, endRowIndex:46,
            startColumnIndex:10, endColumnIndex:11,
          }]}}, targetAxis:'LEFT_AXIS', color:hex(C.secondary) }],
          headerCount:0,
        },
      },
      position:{ overlayPosition:{
        anchorCell:{ sheetId:SHEET, rowIndex:8, columnIndex:9 },
        offsetXPixels:10, offsetYPixels:0, widthPixels:400, heightPixels:270,
      }},
    }}
  };

  // Chart 2: Account Type Distribution (helper data on this tab, rows 30-34, 0-indexed)
  const chart2 = {
    addChart:{ chart:{
      spec:{
        title:'Current Portfolio by Account Type',
        basicChart:{
          chartType:'BAR',
          legendPosition:'NO_LEGEND',
          axis:[
            { position:'BOTTOM_AXIS', title:'Balance ($)' },
            { position:'LEFT_AXIS',   title:'Account Type' },
          ],
          domains:[{ domain:{ sourceRange:{ sources:[{
            sheetId:SHEET, startRowIndex:30, endRowIndex:35,
            startColumnIndex:10, endColumnIndex:11,
          }]}}}],
          series:[{ series:{ sourceRange:{ sources:[{
            sheetId:SHEET, startRowIndex:30, endRowIndex:35,
            startColumnIndex:11, endColumnIndex:12,
          }]}}, targetAxis:'BOTTOM_AXIS', color:hex(C.preTax) }],
          headerCount:1,
        },
      },
      position:{ overlayPosition:{
        anchorCell:{ sheetId:SHEET, rowIndex:8, columnIndex:14 },
        offsetXPixels:10, offsetYPixels:0, widthPixels:380, heightPixels:270,
      }},
    }}
  };

  // Chart 3: Scenario Portfolio Comparison (cross-sheet from Scenario Planner, sheetId 5)
  const chart3 = {
    addChart:{ chart:{
      spec:{
        title:'Portfolio at Retirement — by Scenario',
        basicChart:{
          chartType:'COLUMN',
          legendPosition:'NO_LEGEND',
          axis:[
            { position:'BOTTOM_AXIS', title:'Scenario' },
            { position:'LEFT_AXIS',   title:'Portfolio ($)' },
          ],
          domains:[{ domain:{ sourceRange:{ sources:[{
            sheetId:SID_SCEN, startRowIndex:5, endRowIndex:11,
            startColumnIndex:1, endColumnIndex:2,
          }]}}}],
          series:[{ series:{ sourceRange:{ sources:[{
            sheetId:SID_SCEN, startRowIndex:5, endRowIndex:11,
            startColumnIndex:10, endColumnIndex:11,
          }]}}, targetAxis:'LEFT_AXIS', color:hex(C.SCN01) }],
          headerCount:0,
        },
      },
      position:{ overlayPosition:{
        anchorCell:{ sheetId:SHEET, rowIndex:17, columnIndex:9 },
        offsetXPixels:10, offsetYPixels:0, widthPixels:400, heightPixels:270,
      }},
    }}
  };

  // Chart 4: Readiness Dimension Scores (cross-sheet from Readiness, sheetId 9)
  const chart4 = {
    addChart:{ chart:{
      spec:{
        title:'Retirement Readiness by Dimension',
        basicChart:{
          chartType:'BAR',
          legendPosition:'NO_LEGEND',
          axis:[
            { position:'BOTTOM_AXIS', title:'Score (1–5)' },
            { position:'LEFT_AXIS',   title:'Dimension' },
          ],
          domains:[{ domain:{ sourceRange:{ sources:[{
            sheetId:SID_READY, startRowIndex:10, endRowIndex:18,
            startColumnIndex:1, endColumnIndex:2,
          }]}}}],
          series:[{ series:{ sourceRange:{ sources:[{
            sheetId:SID_READY, startRowIndex:10, endRowIndex:18,
            startColumnIndex:3, endColumnIndex:4,
          }]}}, targetAxis:'BOTTOM_AXIS', color:hex(C.secondary) }],
          headerCount:0,
        },
      },
      position:{ overlayPosition:{
        anchorCell:{ sheetId:SHEET, rowIndex:17, columnIndex:14 },
        offsetXPixels:10, offsetYPixels:0, widthPixels:380, heightPixels:270,
      }},
    }}
  };

  /* ── Execute ──────────────────────────────────────────────────────────────── */
  await batchUpdate(SID, requests, '[12-dashboard format]');
  await valuesBatchUpdate(SID, vals, '[12-dashboard values]');
  await batchUpdate(SID, [chart1, chart2, chart3, chart4], '[12-dashboard charts]');

  console.log(`✅ Retirement Dashboard done — 3 big KPIs, 9 medium KPIs, account/income/milestone/scenario snapshots, 4 charts.`);
}
run().catch(e=>{ console.error(e); process.exit(1); });
