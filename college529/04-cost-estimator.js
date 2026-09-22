'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['College Cost Estimator'];
const S   = "'College Cost Estimator'";
const NC  = 24;
const BENCH = "'Reference Data'!$A$24:$J$30";

// Expression (no leading =): use override if non-empty, else VLOOKUP benchmark
const expr = (col, r, idx) =>
  `IF(${col}${r}<>"",${col}${r},IFERROR(VLOOKUP(B${r},${BENCH},${idx},FALSE),0))`;
const tInfl = (r) =>
  `IF(L${r}<>"",L${r},IFERROR(VLOOKUP(B${r},${BENCH},9,FALSE),0.04))`;
const oInfl = (r) =>
  `IF(M${r}<>"",M${r},IFERROR(VLOOKUP(B${r},${BENCH},10,FALSE),0.03))`;
// Geometric series: FirstYear * ((1+r)^n - 1) / r; IFERROR catches r=0 → FirstYear*n
const geo = (yrCol, inf, nCol, r) =>
  `IFERROR(${yrCol}${r}*((1+(${inf}))^${nCol}${r}-1)/(${inf}),${yrCol}${r}*${nCol}${r})`;

// 9 colleges — types must exactly match Reference Data dropdown
const COLLEGES = [
  { name:'Maple Ridge State University',  type:'In-State Public',          state:'CA', pp:'Public',  n:8,  o:4, ben:'BEN-001' },
  { name:'Pacific Northwest University',  type:'Out-of-State Public',      state:'WA', pp:'Public',  n:11, o:4, ben:'BEN-002' },
  { name:'Fairview University',           type:'Out-of-State Public',      state:'TX', pp:'Public',  n:0,  o:4, ben:'BEN-003' },
  { name:'Crestwood University',          type:'Private Nonprofit',        state:'MA', pp:'Private', n:11, o:4, ben:'BEN-002' },
  { name:'Heritage University',           type:'Private Nonprofit',        state:'VT', pp:'Private', n:4,  o:4, ben:'BEN-005' },
  { name:'Westbrook Community College',   type:'Community College',        state:'OR', pp:'Public',  n:8,  o:2, ben:'BEN-001' },
  { name:'Northern Technical Institute',  type:'Trade / Technical School', state:'MN', pp:'Public',  n:0,  o:2, ben:'BEN-004' },
  { name:'Lakewood Career College',       type:'Private For-Profit',       state:'FL', pp:'Private', n:0,  o:4, ben:'BEN-003' },
  { name:'Summit Learning Academy',       type:'Custom School',            state:'CO', pp:'Private', n:0,  o:4, ben:'BEN-004' },
];

(async () => {
  const vals = [];
  const fmt  = [];

  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,NC), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range:`${S}!A1`, values:[['College Cost Estimator']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,0,1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,0,1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.primary), textFormat:{ bold:true, fontSize:16, foregroundColor:hex(C.white) },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:42 }, fields:'pixelSize' }});

  // Row 2: Subtitle
  vals.push({ range:`${S}!A2`, values:[["Compare schools, project inflation-adjusted costs, and track savings gaps. Leave override fields blank to use benchmark data from Reference Data."]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,1,2,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,1,2,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.aubergTint), textFormat:{ fontSize:9, foregroundColor:hex(C.text), italic:true },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // Rows 3-4: 4 Summary Cards (each spans 6 cols)
  const CARDS = [
    { label:'Schools Analyzed',   val:`=COUNTA(A10:A109)`,                                               cur:false },
    { label:'Avg Year 1 Total',   val:`=IFERROR(AVERAGEIF(R10:R109,">"&0),"")`,                          cur:true  },
    { label:'Lowest 4-Yr Total',  val:`=IFERROR(MINIFS(S10:S109,A10:A109,"<>"&""),"")`,                  cur:true  },
    { label:'Highest 4-Yr Total', val:`=IFERROR(MAXIFS(S10:S109,A10:A109,"<>"&""),"")`,                  cur:true  },
  ];
  CARDS.forEach(({ label, val, cur }, i) => {
    const cs = i*6, ce = cs+6;
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
    fmt.push({ repeatCell:{ range:gridRange(SID,3,4,cs,ce), cell:{ userEnteredFormat:vf }, fields:'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:36 }, fields:'pixelSize' }});

  // Row 6: Section Header
  vals.push({ range:`${S}!A6`, values:[["School Comparison Table  ·  Leave override fields blank to use benchmark data  ·  Yellow = editable  ·  Blue-gray = formula"]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,5,6,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,5,6,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrLight), textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.primary) },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // Row 9: Column headers
  vals.push({ range:`${S}!A9`, values:[[
    'School Name','College Type','State','Public / Private',
    'Tuition (Override)','Fees (Override)','Housing (Override)','Food (Override)',
    'Books (Override)','Transport (Override)','Personal (Override)',
    'Tuition Infl % (Override)','Other Infl % (Override)',
    'Yrs Until Enrollment','Yrs of Study',
    'Year 1 Tuition (Proj.)','Year 1 Other Costs (Proj.)','Total Year 1',
    '4-Yr Projected Total','Est. Total incl. 2 Grad Yrs',
    'Assigned Beneficiary','Current Savings','Funding Gap','Notes'
  ]] });
  fmt.push({ repeatCell:{ range:gridRange(SID,8,9,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:8, endIndex:9 }, properties:{ pixelSize:40 }, fields:'pixelSize' }});

  // Rows 10-109: Data
  const dataRows = [];
  for (let i = 0; i < 100; i++) {
    const c = COLLEGES[i] || null;
    const r = i + 10;
    const ti = tInfl(r), oi = oInfl(r);
    const pF = `=IFERROR(IF(A${r}="","",${expr('E',r,2)}*(1+(${ti}))^N${r}),"")`;
    const qF = `=IFERROR(IF(A${r}="","",( ${expr('F',r,3)}+${expr('G',r,4)}+${expr('H',r,5)}+${expr('I',r,6)}+${expr('J',r,7)}+${expr('K',r,8)} )*(1+(${oi}))^N${r}),"")`;
    const rF = `=IFERROR(IF(A${r}="","",P${r}+Q${r}),"")`;
    const sF = `=IFERROR(IF(A${r}="","",${geo('P',ti,'O',r)}+${geo('Q',oi,'O',r)}),"")`;
    const tF = `=IFERROR(IF(A${r}="","",S${r}+P${r}*(1+(${ti}))^O${r}+P${r}*(1+(${ti}))^(O${r}+1)+Q${r}*(1+(${oi}))^O${r}+Q${r}*(1+(${oi}))^(O${r}+1)),"")`;
    const vF = `=IFERROR(VLOOKUP(U${r},'Beneficiary Setup'!$A$8:$O$507,15,FALSE),"")`;
    const wF = `=IFERROR(IF(A${r}="","",MAX(0,S${r}-V${r})),"")`;
    if (c) {
      dataRows.push([c.name,c.type,c.state,c.pp,'','','','','','','','','',c.n,c.o,pF,qF,rF,sF,tF,c.ben,vF,wF,'']);
    } else {
      dataRows.push(['','','','','','','','','','','','','','','',pF,qF,rF,sF,tF,'',vF,wF,'']);
    }
  }
  vals.push({ range:`${S}!A10`, values:dataRows });

  // Row colors (alternate), then overlay override/formula colors
  for (let i = 0; i < 100; i++) {
    fmt.push({ repeatCell:{ range:gridRange(SID,9+i,10+i,0,NC), cell:{ userEnteredFormat:{
      backgroundColor:hex(i%2===0 ? C.white : C.altRow)
    }}, fields:'userEnteredFormat.backgroundColor' }});
  }
  // Override input columns E-M (4-12) and N-O (13-14): input yellow
  fmt.push({ repeatCell:{ range:gridRange(SID,9,109,4,15), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  // U (20): assigned beneficiary — input
  fmt.push({ repeatCell:{ range:gridRange(SID,9,109,20,21), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  // P-T (15-19), V-W (21-22): formula color + currency
  fmt.push({ repeatCell:{ range:gridRange(SID,9,109,15,20), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.formula), numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
  }}, fields:'userEnteredFormat(backgroundColor,numberFormat)' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,9,109,21,23), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.formula), numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' }
  }}, fields:'userEnteredFormat(backgroundColor,numberFormat)' }});

  // Freeze row 9
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:9 } }, fields:'gridProperties.frozenRowCount' }});

  // Column widths
  [220,130,55,90, 90,75,85,70,75,80,80, 100,100, 80,70, 115,120,105,120,130, 120,110,105,150]
    .forEach((w,i) => fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'COLUMNS', startIndex:i, endIndex:i+1 }, properties:{ pixelSize:w }, fields:'pixelSize' }}));

  await valuesBatchUpdate(id, vals, '04-cost-estimator');
  await batchUpdate(id, fmt, '04-cost-estimator');
  console.log('04-cost-estimator done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
