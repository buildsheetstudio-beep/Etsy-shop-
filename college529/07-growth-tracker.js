'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Growth Tracker'];
const S   = "'Growth Tracker'";
const CL  = "'Contribution Log'";
const A5  = "'529 Accounts'";
const NC  = 14;

// Monthly dates: Jan 2024 → Sep 2026 (33 months)
const HIST_MONTHS = [];
for (let y=2024; y<=2026; y++) {
  const maxM = (y===2026) ? 9 : 12;
  for (let m=1; m<=maxM; m++) HIST_MONTHS.push({ y, m, label:`${String(m).padStart(2,'0')}/01/${y}` });
}
const N_HIST = HIST_MONTHS.length; // 33

// Projection: Oct 2026 → Sep 2031 (60 months)
const PROJ_MONTHS = [];
let py=2026, pm=10;
for (let k=0; k<60; k++) {
  PROJ_MONTHS.push({ y:py, m:pm, label:`${String(pm).padStart(2,'0')}/01/${py}`, seq:k+1 });
  if (++pm>12) { pm=1; py++; }
}

// Scenario rates
const SCENARIOS = [
  { name:'Conservative', rate:0.04, col:'D' },
  { name:'Base',         rate:0.06, col:'E' },
  { name:'Higher Growth',rate:0.08, col:'F' },
];

(async () => {
  const vals = [];
  const fmt  = [];

  fmt.push({ repeatCell:{ range:gridRange(SID,0,2500,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.bg), textFormat:{ fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  vals.push({ range:`${S}!A1`, values:[['Growth Tracker']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,0,1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,0,1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.primary), textFormat:{ bold:true, fontSize:16, foregroundColor:hex(C.white) },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:42 }, fields:'pixelSize' }});

  // ── Row 2: Subtitle ───────────────────────────────────────────────────────
  vals.push({ range:`${S}!A2`, values:[["Monitor historical contribution totals and project future portfolio growth across three scenarios. All figures are illustrative."]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,1,2,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,1,2,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.aubergTint), textFormat:{ fontSize:9, foregroundColor:hex(C.text), italic:true },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // ── Rows 3-4: Summary Cards ───────────────────────────────────────────────
  const CARDS = [
    { label:'Current Portfolio Value', val:`=IFERROR(SUM(${A5}!N8:N1007),"")`, cur:true, pct:false },
    { label:'YTD Net Contributions',   val:`=IFERROR(SUMPRODUCT((YEAR(${CL}!$B$6:$B$5005)=YEAR(TODAY()))*(${CL}!$F$6:$F$5005<>"Withdrawal")*${CL}!$H$6:$H$5005)-SUMPRODUCT((YEAR(${CL}!$B$6:$B$5005)=YEAR(TODAY()))*(${CL}!$F$6:$F$5005="Withdrawal")*${CL}!$H$6:$H$5005),"")`, cur:true, pct:false },
    { label:'5-Yr Base Projection',    val:`=IFERROR(E${N_HIST+55},"")`,  cur:true,  pct:false },
    { label:'5-Yr Higher Growth Proj.',val:`=IFERROR(F${N_HIST+55},"")`,  cur:true,  pct:false },
  ];
  const cSpans=[[0,3],[3,7],[7,10],[10,14]];
  CARDS.forEach(({ label, val, cur }, i) => {
    const [cs,ce] = cSpans[i]; const cl=String.fromCharCode(65+cs);
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

  // ── Rows 7-14: Scenario Assumptions Table (index 6-13) ────────────────────
  const SA_ROW = 6; // 0-indexed
  vals.push({ range:`${S}!A7`, values:[['SCENARIO ASSUMPTIONS']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,SA_ROW,SA_ROW+1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,SA_ROW,SA_ROW+1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrLight), textFormat:{ bold:true, fontSize:10, foregroundColor:hex(C.primary) }, horizontalAlignment:'LEFT'
  }}, fields:'userEnteredFormat' }});

  vals.push({ range:`${S}!A8`, values:[['Scenario','Annual Return','Monthly Return','Description']] });
  fmt.push({ repeatCell:{ range:gridRange(SID,SA_ROW+1,SA_ROW+2,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 }, horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat' }});

  const scenRows = [
    ['Conservative', 0.04, '=B9/12',  'Lower-risk approach; prioritizes capital preservation'],
    ['Base',         0.06, '=B10/12', 'Balanced growth; diversified index fund strategy'],
    ['Higher Growth',0.08, '=B11/12', 'Growth-oriented; higher equity allocation; greater volatility'],
  ];
  vals.push({ range:`${S}!A9`, values:scenRows });
  fmt.push({ repeatCell:{ range:gridRange(SID,SA_ROW+2,SA_ROW+5,0,4), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.white), textFormat:{ fontSize:9 }
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,SA_ROW+2,SA_ROW+5,1,3), cell:{ userEnteredFormat:{
    numberFormat:{ type:'PERCENT', pattern:'0.0%' }
  }}, fields:'userEnteredFormat.numberFormat' }});

  // ── Row 15-16: Historical Section Header ──────────────────────────────────
  const HIST_HEADER = 14; // 0-indexed row 15
  vals.push({ range:`${S}!A15`, values:[['HISTORICAL MONTHLY CONTRIBUTION SUMMARY (Jan 2024 – Sep 2026)']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,HIST_HEADER,HIST_HEADER+1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,HIST_HEADER,HIST_HEADER+1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.secondary), textFormat:{ bold:true, fontSize:10, foregroundColor:hex(C.white) }, horizontalAlignment:'LEFT'
  }}, fields:'userEnteredFormat' }});

  vals.push({ range:`${S}!A16`, values:[['Month','Contributions','Withdrawals','Net Contributions','Cumulative Balance (No Growth)','','','','','','','','','']] });
  fmt.push({ repeatCell:{ range:gridRange(SID,HIST_HEADER+1,HIST_HEADER+2,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 }, horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat' }});

  // Historical data rows (rows 17-49, 0-indexed 16-48)
  const HIST_DATA_START = 16; // 0-indexed
  const histRows = HIST_MONTHS.map(({ label }, i) => {
    const r = i + 17; // 1-indexed row
    const dateExpr = `DATE(YEAR("${label}"),MONTH("${label}"),1)`;
    const eomExpr  = `EOMONTH("${label}",0)`;
    const contribF = `=IFERROR(SUMPRODUCT((YEAR(${CL}!$B$6:$B$5005)=YEAR("${label}"))*(MONTH(${CL}!$B$6:$B$5005)=MONTH("${label}"))*(${CL}!$F$6:$F$5005="Contribution")*${CL}!$H$6:$H$5005),0)`;
    const withdrawF= `=IFERROR(SUMPRODUCT((YEAR(${CL}!$B$6:$B$5005)=YEAR("${label}"))*(MONTH(${CL}!$B$6:$B$5005)=MONTH("${label}"))*(${CL}!$F$6:$F$5005="Withdrawal")*${CL}!$H$6:$H$5005),0)`;
    const netF     = `=IFERROR(B${r}-C${r},0)`;
    const cumF     = i===0
      ? `=IFERROR(D${r},0)`
      : `=IFERROR(E${r-1}+D${r},0)`;
    return [label, contribF, withdrawF, netF, cumF,'','','','','','','','',''];
  });
  vals.push({ range:`${S}!A17`, values:histRows });

  // Format historical section
  const HIST_END = HIST_DATA_START + N_HIST; // 49
  fmt.push({ repeatCell:{ range:gridRange(SID,HIST_DATA_START,HIST_END,1,5), cell:{ userEnteredFormat:{ numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' } } }, fields:'userEnteredFormat.numberFormat' }});
  for (let i=0; i<N_HIST; i++) {
    fmt.push({ repeatCell:{ range:gridRange(SID,HIST_DATA_START+i,HIST_DATA_START+i+1,0,5), cell:{ userEnteredFormat:{ backgroundColor:hex(i%2===0?C.white:C.altRow) } }, fields:'userEnteredFormat.backgroundColor' }});
  }

  // ── 5-Year Projection Section ─────────────────────────────────────────────
  const PROJ_HEADER = HIST_END + 1; // row after historical
  vals.push({ range:`${S}!A${PROJ_HEADER+1}`, values:[['5-YEAR MONTHLY PORTFOLIO PROJECTION (Oct 2026 – Sep 2031)']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,PROJ_HEADER,PROJ_HEADER+1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,PROJ_HEADER,PROJ_HEADER+1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.accent), textFormat:{ bold:true, fontSize:10, foregroundColor:hex(C.white) }, horizontalAlignment:'LEFT'
  }}, fields:'userEnteredFormat' }});

  const PROJ_HDR_ROW = PROJ_HEADER + 1;
  vals.push({ range:`${S}!A${PROJ_HDR_ROW+1}`, values:[['Month #','Date','Monthly Contribution','Conservative (4%)','Base (6%)','Higher Growth (8%)','','','','','','','','']] });
  fmt.push({ repeatCell:{ range:gridRange(SID,PROJ_HDR_ROW,PROJ_HDR_ROW+1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 }, horizontalAlignment:'CENTER'
  }}, fields:'userEnteredFormat' }});

  // Starting balance and monthly contribution (pulled from 529 Accounts)
  const startBal = `IFERROR(SUMPRODUCT((${A5}!$G$8:$G$1007="Active")*${A5}!$N$8:$N$1007),0)`;
  const monthlyContrib = `IFERROR(SUMPRODUCT((${A5}!$G$8:$G$1007="Active")*${A5}!$M$8:$M$1007)/12,0)`;

  const PROJ_DATA_START = PROJ_HDR_ROW + 1; // 0-indexed
  const projRows = PROJ_MONTHS.map(({ label, seq }, k) => {
    const r = PROJ_DATA_START + k + 1; // 1-indexed row
    const mc = `=${monthlyContrib}`;
    let dF, cF, bF, hF;
    if (k === 0) {
      dF = `=${startBal}*(1+0.04/12)+C${r}`;
      bF = `=${startBal}*(1+0.06/12)+C${r}`;
      hF = `=${startBal}*(1+0.08/12)+C${r}`;
    } else {
      const pr = r - 1;
      dF = `=IFERROR(D${pr}*(1+0.04/12)+C${r},"")`;
      bF = `=IFERROR(E${pr}*(1+0.06/12)+C${r},"")`;
      hF = `=IFERROR(F${pr}*(1+0.08/12)+C${r},"")`;
    }
    return [seq, label, mc, dF, bF, hF,'','','','','','','',''];
  });
  vals.push({ range:`${S}!A${PROJ_DATA_START+1}`, values:projRows });

  // Format projection section
  const PROJ_END = PROJ_DATA_START + 60;
  fmt.push({ repeatCell:{ range:gridRange(SID,PROJ_DATA_START,PROJ_END,2,6), cell:{ userEnteredFormat:{ numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' } } }, fields:'userEnteredFormat.numberFormat' }});
  for (let i=0; i<60; i++) {
    fmt.push({ repeatCell:{ range:gridRange(SID,PROJ_DATA_START+i,PROJ_DATA_START+i+1,0,6), cell:{ userEnteredFormat:{ backgroundColor:hex(i%2===0?C.white:C.altRow) } }, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Freeze row 5
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:5 } }, fields:'gridProperties.frozenRowCount' }});

  // Column widths
  [70,120,130,120,120,130,80,80,80,80,80,80,80,80]
    .forEach((w,i) => fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'COLUMNS', startIndex:i, endIndex:i+1 }, properties:{ pixelSize:w }, fields:'pixelSize' }}));

  await valuesBatchUpdate(id, vals, '07-growth-tracker');
  await batchUpdate(id, fmt, '07-growth-tracker');
  console.log('07-growth-tracker done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
