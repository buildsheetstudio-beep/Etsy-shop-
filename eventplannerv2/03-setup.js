'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 1;

const CATEGORIES = ['Business','Conference','Workshop','Networking','Community','Nonprofit',
  'School','Fundraiser','Social','Birthday','Wedding','Family','Holiday','Sports','Virtual','Other'];
const CAT_COLORS = {
  Business:'#C9DCEF', Conference:'#D8CDE8', Workshop:'#F2D2B6', Networking:'#C5E0DD',
  Community:'#CBDCC5', Nonprofit:'#E9DBA7', School:'#D5E4F2',  Fundraiser:'#E8C0B8',
  Social:'#E9CCD8',   Birthday:'#F3D0BB',   Wedding:'#DDD0DF',  Family:'#EEE2C8',
  Holiday:'#D0E5D8',  Sports:'#C6D9E8',     Virtual:'#DDD5ED',  Other:'#E3E1DE',
};

(async () => {
  const values = [];
  const fmt = [];

  // ── Title ──────────────────────────────────────────────────────────
  values.push({ range: "'Event Setup'!A1", values: [['ULTIMATE EVENT PLANNER']] });
  values.push({ range: "'Event Setup'!A2", values: [['Organization & Planning Command Center — Configure your event details here']] });

  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,12), mergeType:'MERGE_ALL' } });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,12), mergeType:'MERGE_ALL' } });

  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex('#2F3336'),
    textFormat: { bold:true, fontSize:18, foregroundColor: hex('#FFFFFF') },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex('#4A5056'),
    textFormat: { fontSize:10, foregroundColor: hex('#D5D8DB'), italic:true },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  fmt.push({ updateDimensionProperties: { range:{ sheetId:SID, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties:{ pixelSize:50 }, fields:'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range:{ sheetId:SID, dimension:'ROWS', startIndex:1, endIndex:2 },
    properties:{ pixelSize:28 }, fields:'pixelSize' }});

  // ── Section A: Organization Info (rows 3-14, 0-indexed 3-14) ───────
  values.push({ range: "'Event Setup'!A4", values: [['ORGANIZATION INFORMATION']] });
  fmt.push({ mergeCells: { range: gridRange(SID,3,4,0,12), mergeType:'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex('#3D6374'),
    textFormat: { bold:true, fontSize:11, foregroundColor: hex('#FFFFFF') },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE',
    padding:{ left:12 },
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range:{ sheetId:SID, dimension:'ROWS', startIndex:3, endIndex:4 },
    properties:{ pixelSize:30 }, fields:'pixelSize' }});

  const orgFields = [
    ['Organization / Host Name',   'Sunrise Community Foundation',    'B5'],
    ['Contact Person',              'Diana Reyes',                     'B6'],
    ['Email',                       'events@sunrisecf.org',            'B6'],
    ['Phone',                       '(512) 448-7700',                  'B7'],
    ['Website',                     'www.sunrisecf.org',               'B8'],
    ['Address',                     '4200 Lamar Blvd, Austin, TX 78756','B9'],
    ['City / State / Zip',          'Austin, TX 78756',                'B10'],
    ['Default Event Category',      'Nonprofit',                       'B11'],
    ['Fiscal Year Start Month',     '1',                               'B12'],
    ['Currency Symbol',             '$',                               'B13'],
  ];

  const orgData = orgFields.map((f, i) => {
    values.push({ range: `'Event Setup'!A${5+i}`, values: [[f[0]]] });
    values.push({ range: `'Event Setup'!B${5+i}`, values: [[f[1]]] });
    return i;
  });

  // Label and input styling for org rows 4-13 (0-indexed)
  fmt.push({ repeatCell: { range: gridRange(SID,4,14,0,1), cell: { userEnteredFormat: {
    textFormat: { bold:true, foregroundColor: hex('#2F3336') },
    backgroundColor: hex('#F2F1EE'),
    verticalAlignment:'MIDDLE',
    padding:{ left:10 },
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,4,14,1,5), cell: { userEnteredFormat: {
    backgroundColor: hex('#FDFCF9'),
    textFormat: { foregroundColor: hex('#2F3336') },
    verticalAlignment:'MIDDLE',
    padding:{ left:6 },
    borders:{ bottom:{ style:'SOLID', color: hex('#C9C8C4'), width:1 } },
  }}, fields:'userEnteredFormat' }});

  // Validation: Default Event Category
  fmt.push({ setDataValidation: { range: gridRange(SID,11,12,1,2), rule: {
    condition: { type:'ONE_OF_LIST', values: CATEGORIES.map(v=>({ userEnteredValue:v })) },
    showCustomUi:true, strict:true,
  }}});

  // ── Section B: Planning Controls (rows 15-25) ────────────────────────
  values.push({ range: "'Event Setup'!A16", values: [['PLANNING CONTROLS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,15,16,0,12), mergeType:'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID,15,16,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex('#3D6374'),
    textFormat: { bold:true, fontSize:11, foregroundColor: hex('#FFFFFF') },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE',
    padding:{ left:12 },
  }}, fields:'userEnteredFormat' }});

  const ctrlFields = [
    ['Planning Start Date', '01/15/2026'],
    ['Default Event Duration (hours)', '4'],
    ['Max Attendees Per Event (default)', '200'],
    ['Task Lead Time (days)', '7'],
    ['Budget Overrun Alert (%)', '10'],
    ['Reminder Threshold (days before)', '14'],
    ['Equipment Buffer (%)', '15'],
    ['Show Cancelled Events?', 'Yes'],
  ];
  ctrlFields.forEach((f, i) => {
    values.push({ range: `'Event Setup'!A${17+i}`, values: [[f[0]]] });
    values.push({ range: `'Event Setup'!B${17+i}`, values: [[f[1]]] });
  });

  fmt.push({ repeatCell: { range: gridRange(SID,16,24,0,1), cell: { userEnteredFormat: {
    textFormat: { bold:true, foregroundColor: hex('#2F3336') },
    backgroundColor: hex('#F2F1EE'), verticalAlignment:'MIDDLE', padding:{ left:10 },
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,16,24,1,5), cell: { userEnteredFormat: {
    backgroundColor: hex('#FDFCF9'), verticalAlignment:'MIDDLE', padding:{ left:6 },
    borders:{ bottom:{ style:'SOLID', color: hex('#C9C8C4'), width:1 } },
  }}, fields:'userEnteredFormat' }});

  // Yes/No validation for Show Cancelled
  fmt.push({ setDataValidation: { range: gridRange(SID,23,24,1,2), rule: {
    condition: { type:'ONE_OF_LIST', values:[{userEnteredValue:'Yes'},{userEnteredValue:'No'}] },
    showCustomUi:true, strict:true,
  }}});

  // ── Category Color Legend (rows 26-45) ────────────────────────────
  values.push({ range: "'Event Setup'!A27", values: [['EVENT CATEGORY COLOR LEGEND']] });
  fmt.push({ mergeCells: { range: gridRange(SID,26,27,0,12), mergeType:'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID,26,27,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex('#3D6374'),
    textFormat: { bold:true, fontSize:11, foregroundColor: hex('#FFFFFF') },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:12 },
  }}, fields:'userEnteredFormat' }});

  // 2 columns of 8
  const half = CATEGORIES.length / 2;
  CATEGORIES.forEach((cat, i) => {
    const row = 27 + (i % half);
    const colA = i < half ? 0 : 6;
    const colB = i < half ? 1 : 7;
    values.push({ range: `'Event Setup'!${colToLetter(colA)}${row+1}`, values: [[cat]] });
    fmt.push({ repeatCell: { range: gridRange(SID, row, row+1, colA, colB+1), cell: {
      userEnteredFormat: { backgroundColor: hex(CAT_COLORS[cat]),
        textFormat: { bold: true, foregroundColor: hex('#2F3336') },
        verticalAlignment:'MIDDLE', padding:{ left:8 },
      }}, fields:'userEnteredFormat' }});
  });

  // ── 8 KPI Summary Cards (rows 47-60) ─────────────────────────────
  values.push({ range: "'Event Setup'!A47", values: [['PLANNER SUMMARY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,46,47,0,12), mergeType:'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID,46,47,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex('#3D6374'),
    textFormat: { bold:true, fontSize:11, foregroundColor: hex('#FFFFFF') },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:12 },
  }}, fields:'userEnteredFormat' }});

  const kpis = [
    { label:'Total Events',      formula:`=IFERROR(COUNTA('Master Event Log'!B6:B500),"—")`,   col:0 },
    { label:'Active Events',     formula:`=IFERROR(COUNTIF('Master Event Log'!F6:F500,"Confirmed"),"—")`, col:3 },
    { label:'Total Budget',      formula:`=IFERROR(SUM('Budget & Expenses'!C6:C25),"—")`,      col:6 },
    { label:'Total Expenses',    formula:`=IFERROR(SUM('Budget & Expenses'!G41:G1040),"—")`,   col:9 },
    { label:'Total Attendees',   formula:`=IFERROR(COUNTA('Guests & Attendees'!B6:B500),"—")`, col:0 },
    { label:'Tasks Complete',    formula:`=IFERROR(COUNTIF('Event Tasks'!F6:F500,"Completed"),"—")`, col:3 },
    { label:'Vendors Booked',    formula:`=IFERROR(COUNTIF('Vendors & Venues'!K121:K420,"Signed"),"—")`, col:6 },
    { label:'Next Event',        formula:`=IFERROR(INDEX('Master Event Log'!B6:B500,MATCH(MIN(IF('Master Event Log'!D6:D500>=TODAY(),'Master Event Log'!D6:D500)),'Master Event Log'!D6:D500,0)),"—")`, col:9 },
  ];

  kpis.forEach((kpi, i) => {
    const baseRow = i < 4 ? 47 : 52;
    const r = baseRow;
    const c = kpi.col;
    fmt.push({ mergeCells: { range: gridRange(SID, r, r+1, c, c+3), mergeType:'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID, r+1, r+3, c, c+3), mergeType:'MERGE_ALL' }});

    values.push({ range: `'Event Setup'!${colToLetter(c)}${r+1}`, values: [[kpi.label]] });
    values.push({ range: `'Event Setup'!${colToLetter(c)}${r+2}`, values: [[kpi.formula]] });

    const cardColor = i % 4 === 0 ? '#C9DCEF' : i % 4 === 1 ? '#CBDCC5' : i % 4 === 2 ? '#F2D2B6' : '#DDD5ED';
    fmt.push({ repeatCell: { range: gridRange(SID, r, r+1, c, c+3), cell: { userEnteredFormat: {
      backgroundColor: hex('#4A5056'),
      textFormat: { bold:true, foregroundColor: hex('#FFFFFF'), fontSize:9 },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID, r+1, r+3, c, c+3), cell: { userEnteredFormat: {
      backgroundColor: hex(cardColor),
      textFormat: { bold:true, fontSize:16, foregroundColor: hex('#2F3336') },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Card row heights
  for (let r = 47; r <= 57; r++) {
    fmt.push({ updateDimensionProperties: { range:{ sheetId:SID, dimension:'ROWS', startIndex:r, endIndex:r+1 },
      properties:{ pixelSize: r===47||r===52 ? 24 : 50 }, fields:'pixelSize' }});
  }

  // ── Instructions (rows 60-74) ──────────────────────────────────────
  values.push({ range: "'Event Setup'!A61", values: [['SETUP INSTRUCTIONS']] });
  fmt.push({ mergeCells: { range: gridRange(SID,60,61,0,12), mergeType:'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID,60,61,0,12), cell: { userEnteredFormat: {
    backgroundColor: hex('#3D6374'),
    textFormat: { bold:true, fontSize:11, foregroundColor: hex('#FFFFFF') },
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{ left:12 },
  }}, fields:'userEnteredFormat' }});

  const instructions = [
    'STEP 1: Fill in your Organization Information above. This data populates the Event Dashboard header.',
    'STEP 2: Review and adjust Planning Controls to match your workflow defaults.',
    'STEP 3: Open Master Event Log to add all your upcoming events (single-day, multi-day, virtual, hybrid, or recurring).',
    'STEP 4: Use Event Details to add agenda items and sub-sessions for each event.',
    'STEP 5: Add attendees in Guests & Attendees — track registration, dietary needs, and accessibility requirements.',
    'STEP 6: Record all venues and vendors in Vendors & Venues with contract and payment status.',
    'STEP 7: Build your task list in Event Tasks. Tasks link to events and phases.',
    'STEP 8: Track supplies and equipment in Equipment & Supplies with quantity and readiness flags.',
    'STEP 9: Enter your category budgets and expenses in Budget & Expenses. Revenue tracking is included.',
    'STEP 10: Use Weekly Schedule to view and manage conflicts for any 7-day window.',
    'STEP 11: Monthly Calendar displays all events with multi-day spans and pastel category colors.',
    'STEP 12: Event Dashboard auto-calculates all KPIs, readiness scores, and charts — no editing required.',
    '⚠ DISCLAIMER: This planner uses Google Sheets formulas only. No macros, scripts, or internet access required.',
  ];
  instructions.forEach((text, i) => {
    values.push({ range: `'Event Setup'!A${62+i}`, values: [[text]] });
    fmt.push({ mergeCells: { range: gridRange(SID, 61+i, 62+i, 0, 12), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID, 61+i, 62+i, 0, 12), cell: { userEnteredFormat: {
      backgroundColor: hex(i % 2 === 0 ? '#F7F6F3' : '#FFFFFF'),
      textFormat: { foregroundColor: hex(i === instructions.length-1 ? '#BC7770' : '#2F3336'),
        italic: i === instructions.length-1, fontSize:10 },
      verticalAlignment:'MIDDLE', padding:{ left:12 },
      wrapStrategy:'WRAP',
    }}, fields:'userEnteredFormat' }});
  });

  // ── Global column widths ──────────────────────────────────────────
  fmt.push({ updateDimensionProperties: { range:{ sheetId:SID, dimension:'COLUMNS', startIndex:0, endIndex:1 },
    properties:{ pixelSize:240 }, fields:'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range:{ sheetId:SID, dimension:'COLUMNS', startIndex:1, endIndex:6 },
    properties:{ pixelSize:160 }, fields:'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range:{ sheetId:SID, dimension:'COLUMNS', startIndex:6, endIndex:12 },
    properties:{ pixelSize:140 }, fields:'pixelSize' }});

  // Freeze row 1
  fmt.push({ updateSheetProperties: { properties:{ sheetId:SID, gridProperties:{ frozenRowCount:1 } },
    fields:'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, values, '03-setup values');
  await batchUpdate(id, fmt, '03-setup format');
  console.log('✅ Event Setup done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });

function colToLetter(c) { return String.fromCharCode(65+c); }
