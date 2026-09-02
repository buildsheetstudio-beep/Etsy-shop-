'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Goals & Milestones'];
const S   = "'Goals & Milestones'";
const NC  = 14;

// 35+ fictional goals and milestones across all 5 beneficiaries
const GOALS = [
  // BEN-001 Emma Hartley (enrollment 2034)
  ['BEN-001','Emma Hartley',  'Savings Milestone', '10% Funded',        10000,  '6/30/2025',  'Achieved',      'Emma fund hit 10% of goal'],
  ['BEN-001','Emma Hartley',  'Savings Milestone', '25% Funded',        25000,  '6/30/2027',  'Active',        ''],
  ['BEN-001','Emma Hartley',  'Savings Milestone', '50% Funded',        50000,  '6/30/2030',  'Not Started',   ''],
  ['BEN-001','Emma Hartley',  'Savings Milestone', '75% Funded',        75000,  '6/30/2032',  'Not Started',   ''],
  ['BEN-001','Emma Hartley',  'Savings Milestone', '100% Funded',       95000,  '9/1/2034',   'Not Started',   'Full goal by enrollment'],
  ['BEN-001','Emma Hartley',  'Custom Amount',     'Annual $3,600 Goal', 3600,  '12/31/2026', 'On Track',      'Full-year contribution goal'],
  ['BEN-001','Emma Hartley',  'Enrollment Date',   'Enrollment Fall 2034', 0,   '9/1/2034',   'Not Started',   'First day of classes'],
  // BEN-002 Lucas Hartley (enrollment 2037)
  ['BEN-002','Lucas Hartley', 'Savings Milestone', '10% Funded',        22000,  '12/31/2026', 'Active',        ''],
  ['BEN-002','Lucas Hartley', 'Savings Milestone', '25% Funded',        55000,  '12/31/2029', 'Not Started',   ''],
  ['BEN-002','Lucas Hartley', 'Savings Milestone', '50% Funded',       110000,  '12/31/2032', 'Not Started',   ''],
  ['BEN-002','Lucas Hartley', 'Savings Milestone', '75% Funded',       165000,  '12/31/2035', 'Not Started',   ''],
  ['BEN-002','Lucas Hartley', 'Savings Milestone', '100% Funded',      220000,  '9/1/2037',   'Not Started',   ''],
  ['BEN-002','Lucas Hartley', 'Custom Amount',     'Annual $4,800 Goal', 4800,  '12/31/2026', 'On Track',      ''],
  ['BEN-002','Lucas Hartley', 'Custom Amount',     'Open Second Account',    0, '6/30/2025',  'Achieved',      'ACC-007 opened Jun 2024'],
  // BEN-003 Sofia Delgado (enrollment 2026 — near/enrolled)
  ['BEN-003','Sofia Delgado', 'Savings Milestone', '10% Funded',        14000,  '6/30/2024',  'Achieved',      ''],
  ['BEN-003','Sofia Delgado', 'Savings Milestone', '25% Funded',        35000,  '9/1/2026',   'Not Started',   'Pre-enrollment target'],
  ['BEN-003','Sofia Delgado', 'Enrollment Date',   'Enrollment Fall 2026', 0,   '9/1/2026',   'Active',        'Fairview University — Fall 2026'],
  ['BEN-003','Sofia Delgado', 'Custom Amount',     'Annual $4,200 Goal', 4200,  '12/31/2026', 'On Track',      ''],
  ['BEN-003','Sofia Delgado', 'Custom Amount',     'Apply for FAFSA',       0,  '10/1/2025',  'Achieved',      'FAFSA submitted Oct 2025'],
  ['BEN-003','Sofia Delgado', 'Custom Amount',     'Financial Aid Award Review', 0, '4/1/2026','Achieved',    'Aid package reviewed; $20k awarded'],
  // BEN-004 Marcus Washington (enrolled 2024)
  ['BEN-004','Marcus Washington','Savings Milestone','Enrollment Balance $60k', 60000,'9/1/2024','Achieved',  'Starting balance at enrollment'],
  ['BEN-004','Marcus Washington','Custom Amount',  'Year 1 Withdrawal',    10000,'12/31/2024', 'Achieved',    'Q1 tuition + room/board paid'],
  ['BEN-004','Marcus Washington','Custom Amount',  'Year 2 Withdrawal',    20000,'12/31/2025', 'Achieved',    'Full year 2 tuition/expenses paid'],
  ['BEN-004','Marcus Washington','Custom Amount',  'Year 3 Withdrawal',    20000,'12/31/2026', 'Active',      'On track'],
  ['BEN-004','Marcus Washington','Custom Amount',  'Year 4 Withdrawal',    20000,'12/31/2027', 'Not Started', ''],
  ['BEN-004','Marcus Washington','Enrollment Date','Graduation Target',       0,'5/15/2028',   'Not Started', 'Expected graduation'],
  // BEN-005 Claire Beaumont (enrollment 2030)
  ['BEN-005','Claire Beaumont', 'Savings Milestone','10% Funded',        18000,  '6/30/2024',  'Achieved',    ''],
  ['BEN-005','Claire Beaumont', 'Savings Milestone','25% Funded',        45000,  '12/31/2025', 'Achieved',    ''],
  ['BEN-005','Claire Beaumont', 'Savings Milestone','50% Funded',        90000,  '12/31/2027', 'Not Started', ''],
  ['BEN-005','Claire Beaumont', 'Savings Milestone','75% Funded',       135000,  '12/31/2029', 'Not Started', ''],
  ['BEN-005','Claire Beaumont', 'Savings Milestone','100% Funded',      180000,  '9/1/2030',   'Not Started', ''],
  ['BEN-005','Claire Beaumont', 'Custom Amount',   'Annual $4,200 Goal',  4200,  '12/31/2026', 'On Track',    ''],
  ['BEN-005','Claire Beaumont', 'Custom Amount',   'Annual Holiday Gift',  5000, '12/25/2026', 'Not Started', 'Rose Beaumont annual gift'],
  ['BEN-005','Claire Beaumont', 'Custom Amount',   'Open Third Account',      0, '6/30/2027',  'Not Started', 'Consider UTMA or additional 529'],
  ['BEN-005','Claire Beaumont', 'Enrollment Date', 'Enrollment Fall 2030',    0, '9/1/2030',   'Not Started', 'Heritage University — target date'],
];

(async () => {
  const vals = [];
  const fmt  = [];

  fmt.push({ repeatCell:{ range:gridRange(SID,0,1600,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.bg), textFormat:{ fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) }
  }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  vals.push({ range:`${S}!A1`, values:[['Goals & Milestones']] });
  fmt.push({ mergeCells:{ range:gridRange(SID,0,1,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,0,1,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.primary), textFormat:{ bold:true, fontSize:16, foregroundColor:hex(C.white) },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:42 }, fields:'pixelSize' }});

  // ── Row 2: Subtitle ───────────────────────────────────────────────────────
  vals.push({ range:`${S}!A2`, values:[["Track savings milestones, enrollment dates, and financial goals for each beneficiary. Update status as you progress."]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,1,2,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,1,2,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.aubergTint), textFormat:{ fontSize:9, foregroundColor:hex(C.text), italic:true },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // ── Rows 3-4: Summary Cards ───────────────────────────────────────────────
  const CARDS = [
    { label:'Total Goals',    val:`=COUNTA(A8:A1600)`,                       cur:false },
    { label:'Achieved',       val:`=COUNTIF(G8:G1600,"Achieved")`,           cur:false },
    { label:'Active / On Track', val:`=COUNTIF(G8:G1600,"Active")+COUNTIF(G8:G1600,"On Track")`, cur:false },
    { label:'Behind / Not Started', val:`=COUNTIF(G8:G1600,"Behind Plan")+COUNTIF(G8:G1600,"Not Started")`, cur:false },
  ];
  const cSpans=[[0,3],[3,7],[7,10],[10,14]];
  CARDS.forEach(({ label, val, cur }, i) => {
    const [cs,ce]=cSpans[i]; const cl=String.fromCharCode(65+cs);
    vals.push({ range:`${S}!${cl}3`, values:[[label]] });
    vals.push({ range:`${S}!${cl}4`, values:[[val]] });
    fmt.push({ mergeCells:{ range:gridRange(SID,2,3,cs,ce), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range:gridRange(SID,2,3,cs,ce), cell:{ userEnteredFormat:{
      backgroundColor:hex(C.eucalTint), textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.secText) }, horizontalAlignment:'CENTER'
    }}, fields:'userEnteredFormat' }});
    fmt.push({ mergeCells:{ range:gridRange(SID,3,4,cs,ce), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range:gridRange(SID,3,4,cs,ce), cell:{ userEnteredFormat:{
      backgroundColor:hex(C.white), textFormat:{ bold:true, fontSize:20, foregroundColor:hex(C.primary) },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }}, fields:'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:2, endIndex:5 }, properties:{ pixelSize:36 }, fields:'pixelSize' }});

  // ── Rows 5-6: Spacer + section label ─────────────────────────────────────
  vals.push({ range:`${S}!A6`, values:[["Goal & Milestone Register   ·   Yellow = editable. Use Milestone Types and Goal Statuses from Reference Data dropdowns."]] });
  fmt.push({ mergeCells:{ range:gridRange(SID,5,6,0,NC), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range:gridRange(SID,5,6,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrLight), textFormat:{ bold:true, fontSize:9, foregroundColor:hex(C.primary) }, horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
  }}, fields:'userEnteredFormat' }});

  // ── Row 7: Column headers (1-indexed row 7 = 0-indexed 6) ────────────────
  vals.push({ range:`${S}!A7`, values:[[
    'Beneficiary ID','Beneficiary Name','Milestone Type','Goal / Milestone Title',
    'Target Amount ($)','Target Date','Status','Notes',
    '','','','','',''
  ]] });
  fmt.push({ repeatCell:{ range:gridRange(SID,6,7,0,NC), cell:{ userEnteredFormat:{
    backgroundColor:hex(C.hdrDark), textFormat:{ foregroundColor:hex(C.white), bold:true, fontSize:9 },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'ROWS', startIndex:6, endIndex:7 }, properties:{ pixelSize:36 }, fields:'pixelSize' }});

  // ── Rows 8+: Goal data ────────────────────────────────────────────────────
  // Add Goal ID column (auto-generated)
  const goalData = GOALS.map((row, i) => {
    const goalId = `GOL-${String(i+1).padStart(3,'0')}`;
    return [row[0], row[1], row[2], row[3], row[4]||'', row[5], row[6], row[7], '','','','','',''];
  });

  // Prepend Goal ID to headers
  vals.push({ range:`${S}!A7`, values:[[
    'Beneficiary ID','Beneficiary Name','Milestone Type','Goal / Milestone Title',
    'Target Amount ($)','Target Date','Status','Notes',
    '','','','','',''
  ]] });

  vals.push({ range:`${S}!A8`, values:goalData });

  // Alternate row colors
  for (let i=0; i<GOALS.length; i++) {
    fmt.push({ repeatCell:{ range:gridRange(SID,7+i,8+i,0,NC), cell:{ userEnteredFormat:{
      backgroundColor:hex(i%2===0?C.white:C.altRow)
    }}, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Input cols A-H (0-7)
  fmt.push({ repeatCell:{ range:gridRange(SID,7,7+GOALS.length,0,8), cell:{ userEnteredFormat:{ backgroundColor:hex(C.input) } }, fields:'userEnteredFormat.backgroundColor' }});
  // Target Amount: currency
  fmt.push({ repeatCell:{ range:gridRange(SID,7,7+GOALS.length,4,5), cell:{ userEnteredFormat:{ numberFormat:{ type:'CURRENCY', pattern:'"$"#,##0' } } }, fields:'userEnteredFormat.numberFormat' }});
  // Target Date: date format
  fmt.push({ repeatCell:{ range:gridRange(SID,7,7+GOALS.length,5,6), cell:{ userEnteredFormat:{ numberFormat:{ type:'DATE', pattern:'MM/DD/YYYY' } } }, fields:'userEnteredFormat.numberFormat' }});

  // Freeze row 7
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:7 } }, fields:'gridProperties.frozenRowCount' }});

  // Column widths
  [80,140,150,220, 100,90,120,200, 80,80,80,80,80,80]
    .forEach((w,i) => fmt.push({ updateDimensionProperties:{ range:{ sheetId:SID, dimension:'COLUMNS', startIndex:i, endIndex:i+1 }, properties:{ pixelSize:w }, fields:'pixelSize' }}));

  await valuesBatchUpdate(id, vals, '09-goals');
  await batchUpdate(id, fmt, '09-goals');
  console.log('09-goals done ✓');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
