'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs'), path = require('path');
const SID = JSON.parse(fs.readFileSync(path.join(__dirname,'spreadsheet.json'))).id;
const SHEET = 7;
const TAB   = 'Milestone Tracker';
const SETUP = "'Personal & Household Setup'";
const REF   = "'Reference Data'";

/* ── helpers ─────────────────────────────────────────────────────────────── */
const colWidths = [90, 255, 145, 85, 100, 130, 110, 130, 295]; // A-I

function mstId(r)  { return `=IF(B${r}="","","MST-"&TEXT(ROW()-6,"000"))`; }
function tgtYear(r){ return `=IFERROR(YEAR(IF(D${r}="P2",${SETUP}!E8,${SETUP}!B8))+E${r},"")`; }

/* ── Milestone data ──────────────────────────────────────────────────────── */
// [Name, Category, Owner, TargetAge, Priority, Status, Notes]
const MILESTONES = [
  // Financial
  ['Maximize 401(k) contributions for both',   'Financial', 'Both', 56, 'High',   'In Progress', 'Increase deductions; aim for catch-up limits after age 50'],
  ['Eliminate all auto loan debt',              'Financial', 'Both', 54, 'High',   'Completed',   'Final payment made; redirect $1,200/mo to brokerage'],
  ['Fund HSA to annual maximum',                'Financial', 'Both', 55, 'Medium', 'In Progress', 'Family limit $8,300 (2024); invest excess in index funds'],
  ['Reach $1M total retirement portfolio',      'Financial', 'Both', 58, 'High',   'In Progress', 'Current trajectory on track per Growth Forecast'],
  ['Pay off primary mortgage',                  'Financial', 'P1',   62, 'High',   'Not Started', 'Accelerate payments at $500/mo extra starting 2028'],
  ['Reach $2M total retirement portfolio',      'Financial', 'Both', 63, 'High',   'Not Started', 'Milestone year per Growth Forecast'],
  ['Reach $3M total retirement portfolio',      'Financial', 'Both', 67, 'Medium', 'Not Started', 'Target nest egg per safe withdrawal strategy'],
  ['Convert traditional IRA to Roth (partial)', 'Financial', 'P1',  60, 'Medium', 'Not Started', 'Execute Roth ladder conversion before RMD age'],
  // Health
  ['Complete long-term care insurance review',  'Health',    'Both', 57, 'High',   'In Progress', 'Compare policies; purchase before premiums spike'],
  ['P1 Medicare enrollment (Part A & B)',       'Health',    'P1',   65, 'High',   'Not Started', 'Enroll 3 months before 65th birthday to avoid penalties'],
  ['P2 Medicare enrollment (Part A & B)',       'Health',    'P2',   65, 'High',   'Not Started', 'P2 turns 65 in 2037; coordinate with COBRA if retiring early'],
  ['Establish wellness baseline & health plan', 'Health',    'Both', 55, 'Medium', 'Completed',   'Annual physicals, colonoscopy, cardiac screening scheduled'],
  // Housing
  ['Evaluate downsizing options',               'Housing',   'Both', 64, 'Medium', 'Not Started', 'Research 55+ communities and coastal markets'],
  ['Downsize to retirement home',               'Housing',   'Both', 68, 'Medium', 'Not Started', 'Target within 2 years of retirement; budget $650,000'],
  ['Set up smart-home / aging-in-place upgrades','Housing',  'Both', 70, 'Low',    'Not Started', 'Grab bars, wider doorways, home lift — $25,000 estimate'],
  // Social
  ['Join community volunteer organization',     'Social',    'Both', 67, 'Low',    'Not Started', 'Habitat for Humanity chapter or Meals on Wheels interest'],
  ['Establish regular social activities plan',  'Social',    'P2',   66, 'Low',    'Not Started', 'Book club, garden club, or fitness class schedule'],
  // Travel
  ['European cultural tour (2-week)',           'Travel',    'Both', 60, 'Medium', 'In Progress', 'Italy, France, Spain — saving $1,500/mo in sinking fund'],
  ['U.S. National Parks road trip',             'Travel',    'Both', 63, 'Low',    'Not Started', 'Yellowstone, Glacier, Grand Canyon; target spring'],
  ['International bucket-list adventure',       'Travel',    'Both', 68, 'Low',    'Not Started', 'New Zealand / Southeast Asia; budget $18,000'],
  ['Annual cruise tradition',                   'Travel',    'Both', 70, 'Low',    'On Hold',     'On hold until mortgage paid; resume post-retirement'],
  // Legal
  ['Update will & revocable living trust',      'Legal',     'Both', 55, 'High',   'Completed',   'Signed with estate attorney; review every 5 years'],
  ['Designate healthcare & financial POA',      'Legal',     'Both', 55, 'High',   'Completed',   'Documents filed; copies to both attorneys'],
  ['Review all beneficiary designations',       'Legal',     'Both', 56, 'High',   'In Progress', 'Update 401k, IRA, life insurance after trust creation'],
  ['Create letter of instruction',              'Legal',     'Both', 58, 'Medium', 'Not Started', 'Account locations, passwords, funeral wishes'],
  // Personal
  ['Fund 529 for grandchildren',                'Personal',  'Both', 65, 'Medium', 'Not Started', 'Open accounts at retirement; seed $25,000 each'],
  ['Establish hobby / passion project',         'Personal',  'P2',   65, 'Low',    'Not Started', 'Photography or pottery — research local studios'],
  ['Family reunion planning (every 3 yrs)',     'Personal',  'Both', 68, 'Low',    'Skipped',     'Skipped 2026 event due to renovation; reschedule 2029'],
];

async function run() {
  const requests = [];
  const vals     = [];

  /* ── 1. Column widths ──────────────────────────────────────────────────── */
  colWidths.forEach((w, i) => {
    requests.push({ updateDimensionProperties: {
      range: { sheetId: SHEET, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });
  requests.push({ updateDimensionProperties: {
    range: { sheetId: SHEET, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});

  /* ── 2. Title row ─────────────────────────────────────────────────────── */
  requests.push({
    mergeCells: { range: gridRange(SHEET,0,1,0,9), mergeType:'MERGE_ALL' }
  });
  requests.push({ repeatCell: { range: gridRange(SHEET,0,1,0,9),
    cell: { userEnteredValue: { stringValue: `🏁  ${TAB}` },
      userEnteredFormat: {
        backgroundColor: hex(C.primary), textFormat: { foregroundColor: hex(C.primaryText), bold:true, fontSize:16, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      }}, fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── 3. Subtitle ──────────────────────────────────────────────────────── */
  requests.push({ mergeCells:{ range:gridRange(SHEET,1,2,0,9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET,1,2,0,9),
    cell:{ userEnteredValue:{ stringValue:'Track key life & financial milestones on your path to retirement' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrB), textFormat:{ foregroundColor:hex(C.primaryText), italic:true, fontSize:10, fontFamily:'Montserrat' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── 4. KPI summary strip (rows 3-4) ────────────────────────────────── */
  const kpis = [
    { label:'Total Milestones', formula:`=COUNTA(B7:B200)`, color:C.primary, tcolor:C.primaryText },
    { label:'Completed',        formula:`=COUNTIF(H7:H200,"Completed")`, color:C.success, tcolor:C.text },
    { label:'In Progress',      formula:`=COUNTIF(H7:H200,"In Progress")`, color:C.info, tcolor:C.text },
    { label:'Not Started',      formula:`=COUNTIF(H7:H200,"Not Started")`, color:C.altRow, tcolor:C.text },
    { label:'On Hold',          formula:`=COUNTIF(H7:H200,"On Hold")`, color:C.warning, tcolor:C.text },
    { label:'High Priority',    formula:`=COUNTIF(G7:G200,"High")`, color:C.attention, tcolor:C.primaryText },
    { label:'Skipped',          formula:`=COUNTIF(H7:H200,"Skipped")`, color:C.border, tcolor:C.secText },
    { label:'% Complete',       formula:`=IFERROR(TEXT(COUNTIF(H7:H200,"Completed")/COUNTA(B7:B200),"0%"),"0%")`, color:C.secondaryDk, tcolor:C.primaryText },
    { label:'Completion Rate',  formula:`=IFERROR(COUNTIF(H7:H200,"Completed")/COUNTA(B7:B200),0)`, color:C.hdrB, tcolor:C.primaryText },
  ];
  // Lay out 4 KPI cards across cols A-H (2 rows = label + value) using cols A-H in pairs
  // Actually, let's do 9 KPIs across 9 single columns (one KPI per column)
  kpis.slice(0,9).forEach((k, i) => {
    requests.push({ mergeCells:{ range:gridRange(SHEET,2,3,i,i+1), mergeType:'MERGE_ALL' }});
    requests.push({ repeatCell:{ range:gridRange(SHEET,2,3,i,i+1),
      cell:{ userEnteredValue:{ stringValue: k.label },
        userEnteredFormat:{ backgroundColor:hex(k.color), textFormat:{ foregroundColor:hex(k.tcolor), bold:true, fontSize:8, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', borders:{ bottom:{ style:'SOLID', color:hex(C.border) }} }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
    requests.push({ repeatCell:{ range:gridRange(SHEET,3,4,i,i+1),
      cell:{ userEnteredValue:{ formulaValue: k.formula },
        userEnteredFormat:{ backgroundColor:hex(k.color), textFormat:{ foregroundColor:hex(k.tcolor), bold:true, fontSize:14, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });

  /* ── 5. Spacer row 5 ────────────────────────────────────────────────── */
  requests.push({ repeatCell:{ range:gridRange(SHEET,4,5,0,9),
    cell:{ userEnteredFormat:{ backgroundColor:hex(C.bg) }}, fields:'userEnteredFormat' }
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:4, endIndex:5 },
    properties:{ pixelSize:6 }, fields:'pixelSize' }
  });

  /* ── 6. Column headers (row 6, 0-indexed row 5) ─────────────────────── */
  const HEADERS = ['ID','Milestone Name','Category','Owner','Target Age','Target Year','Priority','Status','Notes / Action Steps'];
  HEADERS.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET,5,6,i,i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrA), textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:9, fontFamily:'Montserrat' },
        horizontalAlignment: i===1||i===8 ? 'LEFT' : 'CENTER', verticalAlignment:'MIDDLE',
        borders:{ bottom:{ style:'SOLID_MEDIUM', color:hex(C.secondaryDk) } } }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:5, endIndex:6 },
    properties:{ pixelSize:32 }, fields:'pixelSize' }
  });

  /* ── 7. Data rows ────────────────────────────────────────────────────── */
  // Status → row bg map
  const STATUS_BG = {
    'Completed':   '#EAF5EE',
    'In Progress': '#E8EEF6',
    'Not Started': C.panel,
    'On Hold':     '#FBF4E4',
    'Skipped':     C.altRow,
  };
  // Category → left border color
  const CAT_COLOR = {
    'Financial': C.primary,
    'Health':    C.attention,
    'Housing':   C.warning,
    'Social':    C.secondary,
    'Travel':    C.info,
    'Legal':     C.SCN03,
    'Personal':  C.SCN06,
  };
  // Priority → text color for priority cell
  const PRIO_COLOR = { 'High': C.attention, 'Medium': C.warning, 'Low': C.secText, 'None': C.border };

  const dataRows = [];
  MILESTONES.forEach(([name, cat, owner, tgtAge, priority, status, notes], idx) => {
    const r = idx + 7; // 1-indexed row
    const rowBg = STATUS_BG[status] || C.panel;
    const isAlt  = idx % 2 === 1;
    const bg     = isAlt ? rowBg.replace('FF','') : rowBg; // same bg, alt rows handled by status
    // Use slightly lighter alt shade for alt rows within a status group
    const effectiveBg = isAlt && status==='Not Started' ? C.altRow : rowBg;

    // Full row background
    requests.push({ repeatCell:{ range:gridRange(SHEET, r-1, r, 0, 9),
      cell:{ userEnteredFormat:{ backgroundColor:hex(effectiveBg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9, foregroundColor:hex(C.text) },
        verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
        borders:{ bottom:{ style:'SOLID', color:hex(C.border) } } }},
      fields:'userEnteredFormat' }
    });

    // Left accent border on col A = category color
    requests.push({ updateBorders:{ range:gridRange(SHEET, r-1, r, 0, 1),
      left:{ style:'SOLID_MEDIUM', color:hex(CAT_COLOR[cat]||C.border) } }
    });

    // Col A: ID
    requests.push({ repeatCell:{ range:gridRange(SHEET, r-1, r, 0, 1),
      cell:{ userEnteredValue:{ formulaValue: mstId(r) },
        userEnteredFormat:{ backgroundColor:hex(effectiveBg),
          textFormat:{ foregroundColor:hex(C.secText), fontFamily:'Montserrat', fontSize:8 },
          horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
      fields:'userEnteredValue,userEnteredFormat' }
    });

    // Priority cell highlight
    requests.push({ repeatCell:{ range:gridRange(SHEET, r-1, r, 6, 7),
      cell:{ userEnteredFormat:{ textFormat:{ bold: priority==='High', foregroundColor:hex(PRIO_COLOR[priority]||C.text), fontFamily:'Montserrat', fontSize:9 },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
      fields:'userEnteredFormat' }
    });

    // Status cell — bold
    requests.push({ repeatCell:{ range:gridRange(SHEET, r-1, r, 7, 8),
      cell:{ userEnteredFormat:{ textFormat:{ bold:true, fontFamily:'Montserrat', fontSize:9 },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
      fields:'userEnteredFormat' }
    });

    // Row height
    requests.push({ updateDimensionProperties:{
      range:{ sheetId:SHEET, dimension:'ROWS', startIndex:r-1, endIndex:r },
      properties:{ pixelSize:36 }, fields:'pixelSize' }
    });

    dataRows.push({
      row: r,
      values: [ mstId(r), name, cat, owner, tgtAge, tgtYear(r), priority, status, notes ]
    });
  });

  /* ── 8. Dropdowns ────────────────────────────────────────────────────── */
  const lastDataRow = 6 + MILESTONES.length;
  function dropdown(col, src) {
    requests.push({ setDataValidation:{ range: gridRange(SHEET,6,lastDataRow,col,col+1),
      rule:{ condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:src }] }, showCustomUi:true } }
    });
  }
  dropdown(2, `=${REF}!$A$55:$A$61`); // Category
  // Owner as ONE_OF_LIST
  requests.push({ setDataValidation:{ range: gridRange(SHEET,6,lastDataRow,3,4),
    rule:{ condition:{ type:'ONE_OF_LIST', values:[
      {userEnteredValue:'P1'},{userEnteredValue:'P2'},{userEnteredValue:'Both'}
    ]}, showCustomUi:true } }
  });
  dropdown(6, `=${REF}!$A$101:$A$104`); // Priority
  dropdown(7, `=${REF}!$A$64:$A$68`);   // Status

  /* ── 9. Freeze header ────────────────────────────────────────────────── */
  requests.push({ updateSheetProperties:{ properties:{ sheetId:SHEET, gridProperties:{ frozenRowCount:6 }}, fields:'gridProperties.frozenRowCount' }});

  /* ── 10. Conditional formats ─────────────────────────────────────────── */
  // Completed → green text in milestone name col
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET,6,lastDataRow,1,2) ],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=H7="Completed"` }] },
      format:{ textFormat:{ foregroundColor:hex(C.secondaryDk), bold:true } } }
  }, index:0 }});
  // Skipped → strikethrough
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET,6,lastDataRow,1,2) ],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=H7="Skipped"` }] },
      format:{ textFormat:{ foregroundColor:hex(C.secText), strikethrough:true } } }
  }, index:1 }});
  // High priority row — amber left tint already handled per-row; add slight bold on name
  requests.push({ addConditionalFormatRule:{ rule:{
    ranges:[ gridRange(SHEET,6,lastDataRow,1,2) ],
    booleanRule:{ condition:{ type:'CUSTOM_FORMULA', values:[{ userEnteredValue:`=G7="High"` }] },
      format:{ textFormat:{ bold:true } } }
  }, index:2 }});

  /* ── 11. Summary / progress section (below data) ────────────────────── */
  const summaryStartRow = lastDataRow + 1; // 1-indexed
  const si = summaryStartRow - 1;          // 0-indexed
  requests.push({ mergeCells:{ range:gridRange(SHEET, si, si+1, 0, 9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET, si, si+1, 0, 9),
    cell:{ userEnteredValue:{ stringValue:'Category Summary' },
      userEnteredFormat:{ backgroundColor:hex(C.hdrB), textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:10, fontFamily:'Montserrat' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:8 } }},
    fields:'userEnteredValue,userEnteredFormat' }
  });
  requests.push({ updateDimensionProperties:{
    range:{ sheetId:SHEET, dimension:'ROWS', startIndex:si, endIndex:si+1 },
    properties:{ pixelSize:28 }, fields:'pixelSize' }
  });

  // Summary row headers
  const sumHdrs = ['Category','Total','Completed','In Progress','Not Started','On Hold','Skipped'];
  sumHdrs.forEach((h,i) => {
    requests.push({ repeatCell:{ range:gridRange(SHEET, si+1, si+2, i, i+1),
      cell:{ userEnteredValue:{ stringValue:h },
        userEnteredFormat:{ backgroundColor:hex(C.hdrC), textFormat:{ foregroundColor:hex(C.primaryText), bold:true, fontSize:8, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
      fields:'userEnteredValue,userEnteredFormat' }
    });
  });

  const CATS = ['Financial','Health','Housing','Social','Travel','Legal','Personal'];
  CATS.forEach((cat, ci) => {
    const ri = si + 2 + ci;
    const isAltCat = ci % 2 === 1;
    const catBg = isAltCat ? C.altRow : C.panel;
    requests.push({ repeatCell:{ range:gridRange(SHEET, ri, ri+1, 0, 7),
      cell:{ userEnteredFormat:{ backgroundColor:hex(catBg),
        textFormat:{ fontFamily:'Montserrat', fontSize:9 },
        verticalAlignment:'MIDDLE' }},
      fields:'userEnteredFormat' }
    });
    requests.push({ updateBorders:{ range:gridRange(SHEET, ri, ri+1, 0, 1),
      left:{ style:'SOLID_MEDIUM', color:hex(CAT_COLOR[cat]||C.border) } }
    });
    // Values via batchUpdate
    const catStatuses = ['Completed','In Progress','Not Started','On Hold','Skipped'];
    vals.push({ range:`${TAB}!A${ri+1}:G${ri+1}`, values:[[cat,
      `=COUNTIF(C7:C${lastDataRow},"${cat}")`,
      ...catStatuses.map(s=>`=SUMPRODUCT((C7:C${lastDataRow}="${cat}")*(H7:H${lastDataRow}="${s}"))`)
    ]]});
  });

  /* ── 12. Disclaimer ──────────────────────────────────────────────────── */
  const discRow = si + 2 + CATS.length + 1;
  requests.push({ mergeCells:{ range:gridRange(SHEET, discRow, discRow+1, 0, 9), mergeType:'MERGE_ALL' }});
  requests.push({ repeatCell:{ range:gridRange(SHEET, discRow, discRow+1, 0, 9),
    cell:{ userEnteredValue:{ stringValue:'This worksheet is for planning purposes only. Consult a licensed financial advisor before making major retirement decisions.' },
      userEnteredFormat:{ backgroundColor:hex(C.bg),
        textFormat:{ foregroundColor:hex(C.secText), italic:true, fontSize:8, fontFamily:'Montserrat' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE' }},
    fields:'userEnteredValue,userEnteredFormat' }
  });

  /* ── Execute ─────────────────────────────────────────────────────────── */
  await batchUpdate(SID, requests, '[09-milestones format]');

  // Values
  const valueData = [];
  dataRows.forEach(({ row, values }) => {
    valueData.push({ range:`${TAB}!A${row}:I${row}`, values:[values] });
  });
  await valuesBatchUpdate(SID, valueData, '[09-milestones values]');
  await valuesBatchUpdate(SID, vals, '[09-milestones summary]');

  console.log(`✅ Milestone Tracker done — ${MILESTONES.length} milestones across 7 categories.`);
}
run().catch(e=>{ console.error(e); process.exit(1); });
