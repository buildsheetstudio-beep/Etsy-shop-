'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 11;
function cl(i) { if (i<26) return String.fromCharCode(65+i); return 'A'+String.fromCharCode(65+i-26); }

(async () => {
  const values = [];
  const fmt = [];

  // ── Title ────────────────────────────────────────────────────────
  values.push({ range:"'Event Dashboard'!A1", values:[["=IFERROR('Event Setup'!B5&\" — Event Dashboard\",\"ULTIMATE EVENT PLANNER — DASHBOARD\")"]] });
  values.push({ range:"'Event Dashboard'!A2", values:[["=IFERROR(\"Powered by: \"&'Event Setup'!B6&\" | Last Updated: \"&TEXT(TODAY(),\"MMMM D, YYYY\"),\"Organization Event Planning Dashboard\")"]] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:18,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:10,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:52}, fields:'pixelSize' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:2},
    properties:{pixelSize:28}, fields:'pixelSize' }});

  // ── 8 Primary KPI Cards (rows 3-8, 0-indexed 2-7) ────────────────
  // 4 per row, each 4 cols wide
  const kpis = [
    { label:'TOTAL EVENTS',    formula:`=IFERROR(COUNTA('Master Event Log'!B6:B500),0)`,         bg:'#C9DCEF', row:0 },
    { label:'CONFIRMED EVENTS',formula:`=IFERROR(COUNTIF('Master Event Log'!F6:F500,"Confirmed"),0)`, bg:'#CBDCC5', row:0 },
    { label:'TOTAL ATTENDEES', formula:`=IFERROR(COUNTA('Guests & Attendees'!B6:B500),0)`,       bg:'#F2D2B6', row:0 },
    { label:'TOTAL REVENUE',   formula:`=IFERROR(SUM('Budget & Expenses'!H1047:H1346),0)`,       bg:'#DDD5ED', row:0 },
    { label:'TOTAL BUDGET',    formula:`=IFERROR(SUM('Budget & Expenses'!C7:C23),0)`,            bg:'#DDD0DF', row:1 },
    { label:'TOTAL EXPENSES',  formula:`=IFERROR(SUM('Budget & Expenses'!H37:H1040),0)`,         bg:'#E9DBA7', row:1 },
    { label:'TASKS COMPLETE',  formula:`=IFERROR(COUNTIF('Event Tasks'!F6:F500,"Completed"),0)`, bg:'#C5E0DD', row:1 },
    { label:'VENDORS BOOKED',  formula:`=IFERROR(COUNTIF('Vendors & Venues'!M115:M420,"Signed"),0)`,bg:'#E8C0B8', row:1 },
  ];

  const kpiCurrencyIdx = [3,4,5]; // indices that should show as $
  kpis.forEach((kpi,i)=>{
    const baseRowIdx = kpi.row === 0 ? 2 : 6;
    const col = (i%4)*4;
    // Label row
    fmt.push({ mergeCells:{ range: gridRange(SID,baseRowIdx,baseRowIdx+1,col,col+4), mergeType:'MERGE_ALL' }});
    // Value rows
    fmt.push({ mergeCells:{ range: gridRange(SID,baseRowIdx+1,baseRowIdx+4,col,col+4), mergeType:'MERGE_ALL' }});

    values.push({ range:`'Event Dashboard'!${cl(col)}${baseRowIdx+1}`, values:[[kpi.label]] });
    values.push({ range:`'Event Dashboard'!${cl(col)}${baseRowIdx+2}`, values:[[kpi.formula]] });

    fmt.push({ repeatCell:{ range: gridRange(SID,baseRowIdx,baseRowIdx+1,col,col+4), cell:{userEnteredFormat:{
      backgroundColor:hex('#3D6374'),
      textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    fmt.push({ repeatCell:{ range: gridRange(SID,baseRowIdx+1,baseRowIdx+4,col,col+4), cell:{userEnteredFormat:{
      backgroundColor:hex(kpi.bg),
      textFormat:{bold:true,fontSize:22,foregroundColor:hex('#2F3336')},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      numberFormat: kpiCurrencyIdx.includes(i) ? { type:'CURRENCY', pattern:'"$"#,##0' } : { type:'NUMBER', pattern:'#,##0' },
    }}, fields:'userEnteredFormat' }});

    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:baseRowIdx,endIndex:baseRowIdx+1},
      properties:{pixelSize:24}, fields:'pixelSize' }});
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:baseRowIdx+1,endIndex:baseRowIdx+4},
      properties:{pixelSize:20}, fields:'pixelSize' }});
  });

  // ── 8 Secondary KPI flat pairs (rows 10-12, 0-indexed 9-11) ──────
  const secKpis = [
    { label:'Completed Events',  formula:`=IFERROR(COUNTIF('Master Event Log'!F6:F500,"Completed"),0)` },
    { label:'Events In Progress',formula:`=IFERROR(COUNTIF('Master Event Log'!F6:F500,"In Progress"),0)` },
    { label:'Overdue Tasks',     formula:`=IFERROR(COUNTIF('Event Tasks'!R6:R500,"YES"),0)` },
    { label:'Tasks In Progress', formula:`=IFERROR(COUNTIF('Event Tasks'!F6:F500,"In Progress"),0)` },
    { label:'Revenue Received',  formula:`=IFERROR(SUM('Budget & Expenses'!I1047:I1346),0)` },
    { label:'Expenses Paid',     formula:`=IFERROR(SUM('Budget & Expenses'!I37:I1040),0)` },
    { label:'VIP Attendees',     formula:`=IFERROR(COUNTIF('Guests & Attendees'!Q6:Q500,TRUE),0)` },
    { label:'Equipment Shortages',formula:`=IFERROR(COUNTIF('Equipment & Supplies'!T6:T500,"SHORTAGE"),0)` },
  ];

  // Row 11 (0-indexed 10): section header
  values.push({ range:"'Event Dashboard'!A11", values:[['AT-A-GLANCE METRICS']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,10,11,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,10,11,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:10,endIndex:11},
    properties:{pixelSize:26}, fields:'pixelSize' }});

  secKpis.forEach((kpi,i)=>{
    const col = (i%4)*4;
    const baseRow = 11; // row 12 (0-indexed 11)
    const r = baseRow + Math.floor(i/4)*2;
    fmt.push({ mergeCells:{ range: gridRange(SID,r,r+1,col,col+2), mergeType:'MERGE_ALL' }});
    fmt.push({ mergeCells:{ range: gridRange(SID,r,r+1,col+2,col+4), mergeType:'MERGE_ALL' }});
    values.push({ range:`'Event Dashboard'!${cl(col)}${r+1}`, values:[[kpi.label]] });
    values.push({ range:`'Event Dashboard'!${cl(col+2)}${r+1}`, values:[[kpi.formula]] });
    fmt.push({ repeatCell:{ range: gridRange(SID,r,r+1,col,col+2), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#F2F1EE':'#FDFCF9'),
      textFormat:{bold:true,fontSize:9,foregroundColor:hex('#4A5056')},
      verticalAlignment:'MIDDLE', padding:{left:8},
    }}, fields:'userEnteredFormat' }});
    fmt.push({ repeatCell:{ range: gridRange(SID,r,r+1,col+2,col+4), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#EEF4F8':'#FDFCF9'),
      textFormat:{bold:true,fontSize:12,foregroundColor:hex('#2F3336')},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      numberFormat: [4,5].includes(i) ? { type:'CURRENCY', pattern:'"$"#,##0' } : { type:'NUMBER', pattern:'#,##0' },
    }}, fields:'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:11,endIndex:15},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  // ── Event Readiness Scorecard (rows 15-24, 0-indexed 14-23) ───────
  values.push({ range:"'Event Dashboard'!A15", values:[['EVENT READINESS SCORECARD']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,14,15,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,14,15,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:11,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:14,endIndex:15},
    properties:{pixelSize:28}, fields:'pixelSize' }});

  // Readiness components (weight × score)
  const readinessHdrs = ['Readiness Area','Weight','Formula','Score','Weighted Score','Status'];
  values.push({ range:"'Event Dashboard'!A16", values:[readinessHdrs] });
  fmt.push({ repeatCell:{ range: gridRange(SID,15,16,0,6), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat' }});

  const readinessItems = [
    ['Event Setup & Planning', '10%',
      `=IFERROR(IF('Event Setup'!B5<>"",1,0),0)`,
      '', '', ''],
    ['Venue & Vendors Confirmed', '15%',
      `=IFERROR(COUNTIF('Vendors & Venues'!M115:M420,"Signed")/MAX(COUNTA('Vendors & Venues'!A115:A420)-1,1),0)`,
      '', '', ''],
    ['Task Completion Rate', '20%',
      `=IFERROR(COUNTIF('Event Tasks'!F6:F500,"Completed")/MAX(COUNTA('Event Tasks'!C6:C500),1),0)`,
      '', '', ''],
    ['Budget Allocated & Tracked', '25%',
      `=IFERROR(IF(SUM('Budget & Expenses'!C7:C23)>0,MIN(SUM('Budget & Expenses'!H37:H1040)/SUM('Budget & Expenses'!C7:C23),1),0),0)`,
      '', '', ''],
    ['Guest Registration Progress', '15%',
      `=IFERROR(COUNTIF('Guests & Attendees'!H6:H500,"Confirmed")/MAX(COUNTA('Guests & Attendees'!B6:B500),1),0)`,
      '', '', ''],
    ['Equipment Ready', '15%',
      `=IFERROR(COUNTIF('Equipment & Supplies'!I6:I500,"Available")/MAX(COUNTA('Equipment & Supplies'!C6:C500),1),0)`,
      '', '', ''],
  ];
  const weights = [0.10,0.15,0.20,0.25,0.15,0.15];

  readinessItems.forEach((item,i)=>{
    const r = 17+i;
    const weightFmt = `${(weights[i]*100).toFixed(0)}%`;
    // Score = formula in col C (index 2), Weighted = weight * score
    values.push({ range:`'Event Dashboard'!A${r}`, values:[[item[0]]] });
    values.push({ range:`'Event Dashboard'!B${r}`, values:[[weightFmt]] });
    values.push({ range:`'Event Dashboard'!C${r}`, values:[[item[2]]] }); // Formula → Score col (D)
    // Actually: A=label, B=weight, C=formula/score, D=weighted, E=status
    values.push({ range:`'Event Dashboard'!D${r}`, values:[[`=IFERROR(ROUND(C${r}*${weights[i]},3),0)`]] });
    values.push({ range:`'Event Dashboard'!E${r}`, values:[[`=IFERROR(IF(C${r}>=0.8,"✓ Ready",IF(C${r}>=0.5,"In Progress","Needs Attention")),"—")`]] });

    fmt.push({ repeatCell:{ range: gridRange(SID,r-1,r,0,6), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#F7F6F3':'#FDFCF9'),
      textFormat:{fontSize:9,foregroundColor:hex('#2F3336')},
      verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    // C and D: percent format
    fmt.push({ repeatCell:{ range: gridRange(SID,r-1,r,2,5), cell:{userEnteredFormat:{
      numberFormat:{ type:'PERCENT', pattern:'0%' },
      backgroundColor:hex('#EEF4F8'), textFormat:{foregroundColor:hex('#3D6374')},
      horizontalAlignment:'CENTER',
    }}, fields:'userEnteredFormat' }});
  });

  // Total readiness score row
  const readTotalRow = 17 + readinessItems.length; // 0-indexed 22
  values.push({ range:`'Event Dashboard'!A${readTotalRow}`, values:[['OVERALL READINESS SCORE']] });
  values.push({ range:`'Event Dashboard'!D${readTotalRow}`, values:[[`=IFERROR(SUM(D17:D${readTotalRow-1}),0)`]] });
  values.push({ range:`'Event Dashboard'!E${readTotalRow}`, values:[[`=IFERROR(IF(D${readTotalRow}>=0.8,"✓ READY TO EXECUTE",IF(D${readTotalRow}>=0.5,"⚡ IN PROGRESS","⚠ NEEDS ATTENTION")),"—")`]] });
  fmt.push({ repeatCell:{ range: gridRange(SID,readTotalRow-1,readTotalRow,0,6), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF')},
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,readTotalRow-1,readTotalRow,3,5), cell:{userEnteredFormat:{
    numberFormat:{ type:'PERCENT', pattern:'0%' },
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:14,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat' }});

  // CF: Readiness status (E=col4)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,16,readTotalRow,4,5)],
    booleanRule:{ condition:{ type:'TEXT_CONTAINS', values:[{userEnteredValue:'Ready'}] },
      format:{ backgroundColor:hex('#CBDCC5'), textFormat:{foregroundColor:hex('#2A5C35'),bold:true} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,16,readTotalRow,4,5)],
    booleanRule:{ condition:{ type:'TEXT_CONTAINS', values:[{userEnteredValue:'Progress'}] },
      format:{ backgroundColor:hex('#E9DBA7'), textFormat:{foregroundColor:hex('#5C4A00')} } }
  }, index:0 }});
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,16,readTotalRow,4,5)],
    booleanRule:{ condition:{ type:'TEXT_CONTAINS', values:[{userEnteredValue:'Attention'}] },
      format:{ backgroundColor:hex('#E8C0B8'), textFormat:{foregroundColor:hex('#5C1A1A'),bold:true} } }
  }, index:0 }});

  // ── Upcoming Events (rows 26-36, 0-indexed 25-35) ─────────────────
  const upcomingStart = readTotalRow + 2; // 0-indexed
  values.push({ range:`'Event Dashboard'!A${upcomingStart+1}`, values:[['UPCOMING EVENTS (NEXT 12)']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,upcomingStart,upcomingStart+1,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,upcomingStart,upcomingStart+1,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:11,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:upcomingStart,endIndex:upcomingStart+1},
    properties:{pixelSize:28}, fields:'pixelSize' }});

  const upcomingHdrs = ['Event ID','Event Name','Category','Start Date','End Date','Status','Venue/Location','Attendees','Priority'];
  values.push({ range:`'Event Dashboard'!A${upcomingStart+2}`, values:[upcomingHdrs] });
  fmt.push({ repeatCell:{ range: gridRange(SID,upcomingStart+1,upcomingStart+2,0,9), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat' }});

  // Formula-based upcoming events (sorted by date, future only)
  // Using IFERROR INDEX/MATCH approach for first 12 upcoming
  const upcomingData = [];
  for (let i=0; i<12; i++) {
    upcomingData.push([
      `=IFERROR(INDEX('Master Event Log'!A$6:A$500,MATCH(SMALL(IF('Master Event Log'!D$6:D$500>=TODAY(),'Master Event Log'!D$6:D$500),${i+1}),'Master Event Log'!D$6:D$500,0)),"")`,
      `=IFERROR(INDEX('Master Event Log'!B$6:B$500,MATCH(SMALL(IF('Master Event Log'!D$6:D$500>=TODAY(),'Master Event Log'!D$6:D$500),${i+1}),'Master Event Log'!D$6:D$500,0)),"")`,
      `=IFERROR(INDEX('Master Event Log'!G$6:G$500,MATCH(SMALL(IF('Master Event Log'!D$6:D$500>=TODAY(),'Master Event Log'!D$6:D$500),${i+1}),'Master Event Log'!D$6:D$500,0)),"")`,
      `=IFERROR(SMALL(IF('Master Event Log'!D$6:D$500>=TODAY(),'Master Event Log'!D$6:D$500),${i+1}),"")`,
      `=IFERROR(INDEX('Master Event Log'!E$6:E$500,MATCH(SMALL(IF('Master Event Log'!D$6:D$500>=TODAY(),'Master Event Log'!D$6:D$500),${i+1}),'Master Event Log'!D$6:D$500,0)),"")`,
      `=IFERROR(INDEX('Master Event Log'!F$6:F$500,MATCH(SMALL(IF('Master Event Log'!D$6:D$500>=TODAY(),'Master Event Log'!D$6:D$500),${i+1}),'Master Event Log'!D$6:D$500,0)),"")`,
      `=IFERROR(INDEX('Master Event Log'!H$6:H$500,MATCH(SMALL(IF('Master Event Log'!D$6:D$500>=TODAY(),'Master Event Log'!D$6:D$500),${i+1}),'Master Event Log'!D$6:D$500,0)),"")`,
      `=IFERROR(INDEX('Master Event Log'!O$6:O$500,MATCH(SMALL(IF('Master Event Log'!D$6:D$500>=TODAY(),'Master Event Log'!D$6:D$500),${i+1}),'Master Event Log'!D$6:D$500,0)),"")`,
      `=IFERROR(INDEX('Master Event Log'!AB$6:AB$500,MATCH(SMALL(IF('Master Event Log'!D$6:D$500>=TODAY(),'Master Event Log'!D$6:D$500),${i+1}),'Master Event Log'!D$6:D$500,0)),"")`,
    ]);
  }
  values.push({ range:`'Event Dashboard'!A${upcomingStart+3}`, values: upcomingData });

  // Date format for columns D, E (cols 3, 4 = 0-indexed)
  [3,4].forEach(c=>{
    fmt.push({ repeatCell:{ range: gridRange(SID,upcomingStart+2,upcomingStart+15,c,c+1), cell:{userEnteredFormat:{
      numberFormat:{ type:'DATE', pattern:'mmm d, yyyy' },
    }}, fields:'userEnteredFormat.numberFormat' }});
  });

  // Upcoming row styling
  for (let i=0; i<12; i++) {
    const ri = upcomingStart+2+i;
    fmt.push({ repeatCell:{ range: gridRange(SID,ri,ri+1,0,9), cell:{userEnteredFormat:{
      backgroundColor:hex(i%2===0?'#FDFCF9':'#F2F1EE'),
      textFormat:{fontSize:9,foregroundColor:hex('#2F3336')},
      verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  }

  // ── Chart Data Area (rows upcomingStart+17 through +30) ──────────
  const chartDataStart = upcomingStart + 17; // 0-indexed

  // Chart data header
  values.push({ range:`'Event Dashboard'!A${chartDataStart+1}`, values:[['CHART DATA (AUTO-CALCULATED — DO NOT EDIT)']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,chartDataStart,chartDataStart+1,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,chartDataStart,chartDataStart+1,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#EEF4F8'), textFormat:{italic:true,fontSize:9,foregroundColor:hex('#4A5056')},
    horizontalAlignment:'LEFT', padding:{left:12},
  }}, fields:'userEnteredFormat' }});

  // Chart data: col A-B = Status counts, col D-E = Top categories, col G-H = Budget bars
  const chartRow = chartDataStart + 1;
  const statusLabels = ['Planning','Confirmed','In Progress','Completed','Cancelled','Postponed'];
  values.push({ range:`'Event Dashboard'!A${chartRow+1}`, values:[['Status','Count']] });
  statusLabels.forEach((s,i)=>{
    values.push({ range:`'Event Dashboard'!A${chartRow+2+i}`, values:[[s,
      `=IFERROR(COUNTIF('Master Event Log'!F$6:F$500,"${s}"),0)`
    ]]});
  });

  // Category counts
  const catLabels = ['Conference','Community','Fundraiser','Business','Workshop','Social','Sports','Other'];
  values.push({ range:`'Event Dashboard'!D${chartRow+1}`, values:[['Category','Event Count']] });
  catLabels.forEach((cat,i)=>{
    values.push({ range:`'Event Dashboard'!D${chartRow+2+i}`, values:[[cat,
      `=IFERROR(COUNTIF('Master Event Log'!G$6:G$500,"${cat}"),0)`
    ]]});
  });

  // Budget vs actual
  values.push({ range:`'Event Dashboard'!G${chartRow+1}`, values:[['Category','Budget','Spent']] });
  const budgetCats = ['Venue & Facilities','Catering & Food','AV & Technology','Marketing & Advertising','Staffing'];
  budgetCats.forEach((cat,i)=>{
    values.push({ range:`'Event Dashboard'!G${chartRow+2+i}`, values:[[cat,
      `=IFERROR(SUMPRODUCT(('Budget & Expenses'!B$7:B$23=G${chartRow+1+i+1})*('Budget & Expenses'!C$7:C$23)),0)`,
      `=IFERROR(SUMPRODUCT(('Budget & Expenses'!C$41:C$1040=G${chartRow+1+i+1})*('Budget & Expenses'!H$41:H$1040)),0)`,
    ]]});
  });

  // Task completion
  values.push({ range:`'Event Dashboard'!J${chartRow+1}`, values:[['Phase','% Complete']] });
  const phases = ['Pre-Planning','Venue & Vendors','Marketing & Comms','Logistics & Ops','Day-Of Setup','Post-Event'];
  phases.forEach((ph,i)=>{
    values.push({ range:`'Event Dashboard'!J${chartRow+2+i}`, values:[[ph,
      `=IFERROR(SUMPRODUCT(('Event Tasks'!B$6:B$500="${ph}")*('Event Tasks'!F$6:F$500="Completed"))/MAX(SUMPRODUCT(('Event Tasks'!B$6:B$500="${ph}")*('Event Tasks'!C$6:C$500<>"")),1),0)`
    ]]});
  });

  // Column widths
  const colWidths = [80,160,120,90,110,100,120,120,90,120,180,120,100,100,100,100];
  colWidths.forEach((w,i)=>{
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1},
      properties:{pixelSize:w}, fields:'pixelSize' }});
  });

  // Freeze title rows
  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:2}},
    fields:'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, values, '13-dashboard values');
  await batchUpdate(id, fmt, '13-dashboard format');

  // ── 4 Charts ──────────────────────────────────────────────────────
  const chartRequests = [];

  // Get actual 1-indexed rows for chart data
  const cd = chartRow + 1; // 1-indexed header row of chart data

  // Chart 1: Event Status Distribution (Pie)
  chartRequests.push({ addChart: { chart: {
    spec: {
      title: 'Event Status Distribution',
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        series: { sourceRange: { sources: [{ sheetId: SID,
          startRowIndex: cd, endRowIndex: cd+statusLabels.length,
          startColumnIndex: 1, endColumnIndex: 2 }] }},
        domain: { sourceRange: { sources: [{ sheetId: SID,
          startRowIndex: cd, endRowIndex: cd+statusLabels.length,
          startColumnIndex: 0, endColumnIndex: 1 }] }},
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: 2, columnIndex: 0 },
      offsetXPixels: 0, offsetYPixels: 0,
      widthPixels: 400, heightPixels: 280,
    }},
  }}});

  // Chart 2: Events by Category (Bar)
  chartRequests.push({ addChart: { chart: {
    spec: {
      title: 'Events by Category',
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Count' },
          { position: 'LEFT_AXIS', title: 'Category' },
        ],
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: SID,
          startRowIndex: cd, endRowIndex: cd+catLabels.length,
          startColumnIndex: 3, endColumnIndex: 4 }] }}}],
        series: [{ series: { sourceRange: { sources: [{ sheetId: SID,
          startRowIndex: cd, endRowIndex: cd+catLabels.length,
          startColumnIndex: 4, endColumnIndex: 5 }] }},
          targetAxis: 'BOTTOM_AXIS', type: 'COLUMN' }],
        headerCount: 0,
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: 2, columnIndex: 4 },
      offsetXPixels: 0, offsetYPixels: 0,
      widthPixels: 400, heightPixels: 280,
    }},
  }}});

  // Chart 3: Budget vs Actual by Category (Bar)
  chartRequests.push({ addChart: { chart: {
    spec: {
      title: 'Budget vs Actual Spend',
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Amount ($)' },
          { position: 'LEFT_AXIS', title: 'Category' },
        ],
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: SID,
          startRowIndex: cd, endRowIndex: cd+budgetCats.length,
          startColumnIndex: 6, endColumnIndex: 7 }] }}}],
        series: [
          { series: { sourceRange: { sources: [{ sheetId: SID,
            startRowIndex: cd, endRowIndex: cd+budgetCats.length,
            startColumnIndex: 7, endColumnIndex: 8 }] }},
            targetAxis: 'BOTTOM_AXIS', type: 'COLUMN' },
          { series: { sourceRange: { sources: [{ sheetId: SID,
            startRowIndex: cd, endRowIndex: cd+budgetCats.length,
            startColumnIndex: 8, endColumnIndex: 9 }] }},
            targetAxis: 'BOTTOM_AXIS', type: 'COLUMN' },
        ],
        headerCount: 0,
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: 2, columnIndex: 8 },
      offsetXPixels: 0, offsetYPixels: 0,
      widthPixels: 400, heightPixels: 280,
    }},
  }}});

  // Chart 4: Task Completion by Phase (Column)
  chartRequests.push({ addChart: { chart: {
    spec: {
      title: 'Task Completion by Phase',
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Phase' },
          { position: 'LEFT_AXIS', title: '% Complete' },
        ],
        domains: [{ domain: { sourceRange: { sources: [{ sheetId: SID,
          startRowIndex: cd, endRowIndex: cd+phases.length,
          startColumnIndex: 9, endColumnIndex: 10 }] }}}],
        series: [{ series: { sourceRange: { sources: [{ sheetId: SID,
          startRowIndex: cd, endRowIndex: cd+phases.length,
          startColumnIndex: 10, endColumnIndex: 11 }] }},
          targetAxis: 'LEFT_AXIS', type: 'COLUMN' }],
        headerCount: 0,
      },
    },
    position: { overlayPosition: {
      anchorCell: { sheetId: SID, rowIndex: 2, columnIndex: 12 },
      offsetXPixels: 0, offsetYPixels: 0,
      widthPixels: 400, heightPixels: 280,
    }},
  }}});

  await batchUpdate(id, chartRequests, '13-dashboard charts');

  // ── Disclaimer at bottom ──────────────────────────────────────────
  const disclaimerRow = chartDataStart + 20;
  values.push({ range:`'Event Dashboard'!A${disclaimerRow+1}`, values:[[
    '⚠ DISCLAIMER: This planner uses Google Sheets formulas only. No macros, external connections, or internet access required. All data is stored locally in this spreadsheet. Formulas recalculate automatically when data is updated.'
  ]]});
  fmt.push({ mergeCells:{ range: gridRange(SID,disclaimerRow,disclaimerRow+1,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,disclaimerRow,disclaimerRow+1,0,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#F7F6F3'),
    textFormat:{italic:true,fontSize:9,foregroundColor:hex('#BC7770')},
    verticalAlignment:'MIDDLE', padding:{left:12}, wrapStrategy:'WRAP',
  }}, fields:'userEnteredFormat' }});

  // Need a second values batch for disclaimer
  await valuesBatchUpdate(id, [
    { range:`'Event Dashboard'!A${disclaimerRow+1}`, values:[[
      '⚠ DISCLAIMER: This planner uses Google Sheets formulas only. No macros, external connections, or internet access required. All data is stored locally in this spreadsheet. Formulas recalculate automatically when data is updated.'
    ]]}
  ], '13-dashboard disclaimer');
  await batchUpdate(id, fmt.splice(0), '13-dashboard format-2');

  console.log('✅ Event Dashboard done.');
})().catch(e => { console.error(e.errors||e.message||e); process.exit(1); });
