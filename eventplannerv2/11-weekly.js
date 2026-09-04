'use strict';
const { sheets, C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const path = require('path');
const { id } = JSON.parse(fs.readFileSync(path.join(__dirname, 'spreadsheet.json')));

const SID = 9;

const CAT_COLORS = {
  Business:'#C9DCEF', Conference:'#D8CDE8', Workshop:'#F2D2B6', Networking:'#C5E0DD',
  Community:'#CBDCC5', Nonprofit:'#E9DBA7', School:'#D5E4F2',  Fundraiser:'#E8C0B8',
  Social:'#E9CCD8',   Birthday:'#F3D0BB',   Wedding:'#DDD0DF',  Family:'#EEE2C8',
  Holiday:'#D0E5D8',  Sports:'#C6D9E8',     Virtual:'#DDD5ED',  Other:'#E3E1DE',
};

function cl(i) { if (i<26) return String.fromCharCode(65+i); return 'A'+String.fromCharCode(65+i-26); }

(async () => {
  const values = [];
  const fmt = [];

  // Title
  values.push({ range:"'Weekly Schedule'!A1", values:[['WEEKLY SCHEDULE']] });
  values.push({ range:"'Weekly Schedule'!A2", values:[['View and manage events for any 7-day window. Set your week start date in B5 and the schedule auto-populates.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,14), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,14), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,14), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,14), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1},
    properties:{pixelSize:44}, fields:'pixelSize' }});

  // Week start date control row 4
  values.push({ range:"'Weekly Schedule'!A4", values:[['Week Starting:']] });
  values.push({ range:"'Weekly Schedule'!B4", values:[['04/13/2026']] }); // Monday of summit week
  values.push({ range:"'Weekly Schedule'!D4", values:[['Week Ending:']] });
  values.push({ range:"'Weekly Schedule'!E4", values:[['=IFERROR(B4+6,"")']]} );
  values.push({ range:"'Weekly Schedule'!G4", values:[['Events This Week:']] });
  values.push({ range:"'Weekly Schedule'!H4", values:[[
    `=IFERROR(SUMPRODUCT((WEEKNUM('Master Event Log'!D6:D500)=WEEKNUM(B4))*(YEAR('Master Event Log'!D6:D500)=YEAR(B4))*('Master Event Log'!F6:F500<>"Cancelled")),0)`
  ]]});

  fmt.push({ repeatCell:{ range: gridRange(SID,3,4,0,14), cell:{userEnteredFormat:{
    backgroundColor:hex('#C9DCEF'), textFormat:{bold:true,foregroundColor:hex('#1E3A5F')},
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,3,4,1,2), cell:{userEnteredFormat:{
    backgroundColor:hex('#FDFCF9'), textFormat:{bold:true,foregroundColor:hex('#2F3336')},
    numberFormat:{ type:'DATE', pattern:'ddd, mmm d yyyy' },
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,3,4,4,5), cell:{userEnteredFormat:{
    backgroundColor:hex('#FDFCF9'), textFormat:{foregroundColor:hex('#2F3336')},
    numberFormat:{ type:'DATE', pattern:'ddd, mmm d yyyy' },
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:3,endIndex:4},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  // Day column headers row 6 (0-indexed 5)
  // Columns: A=Time, B=Mon, C=Tue, D=Wed, E=Thu, F=Fri, G=Sat, H=Sun
  // Wide grid: 2 cols per day (main + details)
  // Layout: A=Time, B-C=Mon, D-E=Tue, F-G=Wed, H-I=Thu, J-K=Fri, L-M=Sat, N-O=Sun
  const dayLabels = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
  values.push({ range:"'Weekly Schedule'!A6", values:[['Time']] });
  dayLabels.forEach((day,di)=>{
    const col = di*2+1; // B=1, D=3, F=5, H=7, J=9, L=11, N=13
    // Day name formula
    values.push({ range:`'Weekly Schedule'!${cl(col)}6`, values:[[
      `=IFERROR(TEXT($B$4+${di},"ddd mmm d"),"${day}")`
    ]]});
    // Merge 2 cols for day
    fmt.push({ mergeCells:{ range: gridRange(SID,5,6,col,col+2), mergeType:'MERGE_ALL' }});
    fmt.push({ repeatCell:{ range: gridRange(SID,5,6,col,col+2), cell:{userEnteredFormat:{
      backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF')},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });
  fmt.push({ repeatCell:{ range: gridRange(SID,5,6,0,1), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3336'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:5,endIndex:6},
    properties:{pixelSize:32}, fields:'pixelSize' }});

  // Time slots from 6:00 AM to 10:00 PM in 30-min increments
  const timeSlots = [];
  for (let h = 6; h <= 22; h++) {
    for (let m = 0; m < 60; m += 30) {
      const hStr = h <= 12 ? h : h-12;
      const mStr = m === 0 ? '00' : '30';
      const ampm = h < 12 ? 'AM' : 'PM';
      const label = h === 12 ? `12:${mStr} PM` : `${hStr}:${mStr} ${ampm}`;
      timeSlots.push(label);
    }
  }

  // Write time slots in col A (rows 7-40, 0-indexed 6-39)
  const timeValues = timeSlots.map(t=>[t]);
  values.push({ range:"'Weekly Schedule'!A7", values: timeValues });

  fmt.push({ repeatCell:{ range: gridRange(SID,6,6+timeSlots.length,0,1), cell:{userEnteredFormat:{
    backgroundColor:hex('#F2F1EE'), textFormat:{bold:true,fontSize:8,foregroundColor:hex('#4A5056')},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  // Time slot row heights
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:6,endIndex:6+timeSlots.length},
    properties:{pixelSize:24}, fields:'pixelSize' }});

  // Column widths: A=80, day cols=110 each
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:0,endIndex:1},
    properties:{pixelSize:80}, fields:'pixelSize' }});
  for (let i=1; i<=14; i++) {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1},
      properties:{pixelSize:i%2===1?140:80}, fields:'pixelSize' }});
  }

  // Populate event blocks: for each day (Mon–Sun of the sample week)
  // Sample week: Apr 13-19, 2026 (nothing; so let's use the summit week Apr 13-19)
  // EVT-001 Apr 15-17, EVT-005 May 27 (different week), use summit week data

  // Pre-populate known events for the default week (Apr 13-19, 2026)
  // Day 0=Mon Apr 13, Day 2=Wed Apr 15 (Summit start)
  const weekEvents = [
    // [dayIndex(0=Mon), startSlot(0-based index), endSlot, eventName, eventID, category, venue]
    [2, 4, 6, 'EVT-001 Summit — Registration & Welcome', 'EVT-001', 'Conference', 'ACC Atrium'],   // Apr15 08:00-09:30
    [2, 7, 10,'EVT-001 Opening Keynote: Leading Through Change', 'EVT-001','Conference','Main Ballroom'],
    [2,11,13,'EVT-001 Breakout A: Strategy & Vision','EVT-001','Conference','Room 101'],
    [2,13,14,'EVT-001 Networking Lunch','EVT-001','Conference','Dining Hall'],
    [2,15,18,'EVT-001 Panel: Nonprofit Governance','EVT-001','Conference','Main Ballroom'],
    [2,19,21,'EVT-001 Day 1 Wrap-Up & Networking','EVT-001','Conference','Patio'],
    [3, 1, 3,'EVT-001 Morning Yoga (Optional)','EVT-001','Conference','Pool Deck'],  // Apr16 06:30-07:30
    [3, 6,10,'EVT-001 Keynote Day 2','EVT-001','Conference','Main Ballroom'],
    [3,11,13,'EVT-001 Workshop: Grant Writing','EVT-001','Conference','Room 201'],
    [3,12,16,'EVT-001 Awards Luncheon','EVT-001','Conference','Ballroom C'],
    [3,17,20,'EVT-001 Afternoon Keynote','EVT-001','Conference','Main Ballroom'],
    [3,22,26,'EVT-001 Closing Reception Day 2','EVT-001','Conference','Rooftop'],
    [4, 6,10,'EVT-001 Closing Keynote','EVT-001','Conference','Main Ballroom'],   // Apr17
    [4,11,13,'EVT-001 Workshop: Action Plan','EVT-001','Conference','Breakout Rooms'],
    [4,12,15,'EVT-001 Farewell Lunch','EVT-001','Conference','Dining Hall'],
  ];

  weekEvents.forEach((ev,i)=>{
    const [dayIdx, startSlot, endSlot, label, evId, cat, venue] = ev;
    const col = dayIdx*2+1; // col index for main (B=1 for Mon, etc)
    const rowStart = 6+startSlot;
    const cat_color = CAT_COLORS[cat] || '#E3E1DE';

    // Write event name and venue into start row only (no merging — avoids conflicts with overlapping events)
    values.push({ range:`'Weekly Schedule'!${cl(col)}${rowStart+1}`, values:[[`▶ ${evId} ${label}`]] });
    values.push({ range:`'Weekly Schedule'!${cl(col+1)}${rowStart+1}`, values:[[venue]] });

    // Color only the start row
    fmt.push({ repeatCell:{ range: gridRange(SID,rowStart,rowStart+1,col,col+2), cell:{userEnteredFormat:{
      backgroundColor:hex(cat_color),
      textFormat:{bold:true,fontSize:8,foregroundColor:hex('#2F3336')},
      verticalAlignment:'MIDDLE', wrapStrategy:'CLIP',
      borders:{
        top:{ style:'SOLID', color:hex('#FFFFFF'), width:2 },
        left:{ style:'SOLID', color:hex('#FFFFFF'), width:2 },
      },
    }}, fields:'userEnteredFormat' }});
  });

  // Conflict detection section below schedule (row ~45, 0-indexed 44)
  const conflictRow = 6 + timeSlots.length + 2; // after time slots + gap
  values.push({ range:`'Weekly Schedule'!A${conflictRow+1}`, values:[['CONFLICT ALERTS']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,conflictRow,conflictRow+1,0,14), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,conflictRow,conflictRow+1,0,14), cell:{userEnteredFormat:{
    backgroundColor:hex('#BC7770'), textFormat:{bold:true,fontSize:11,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:conflictRow,endIndex:conflictRow+1},
    properties:{pixelSize:30}, fields:'pixelSize' }});

  const conflictHdrs = ['Event ID','Event Name','Conflict Type','Date','Time','Notes'];
  values.push({ range:`'Weekly Schedule'!A${conflictRow+2}`, values:[conflictHdrs] });
  fmt.push({ repeatCell:{ range: gridRange(SID,conflictRow+1,conflictRow+2,0,6), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat' }});

  // Conflict formula rows: check if multiple events fall in the same week
  const conflictFormulas = [
    [`=IFERROR(IF(SUMPRODUCT((IFERROR(WEEKNUM('Master Event Log'!D6:D500),0)=WEEKNUM($B$4))*(IFERROR(YEAR('Master Event Log'!D6:D500),0)=YEAR($B$4))*('Master Event Log'!F6:F500<>"Cancelled"))>1,"Multiple events this week - review for date conflicts","No conflicts detected this week"),"")`,],
  ];
  values.push({ range:`'Weekly Schedule'!A${conflictRow+3}`, values: conflictFormulas });
  fmt.push({ mergeCells:{ range: gridRange(SID,conflictRow+2,conflictRow+3,0,14), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,conflictRow+2,conflictRow+3,0,14), cell:{userEnteredFormat:{
    backgroundColor:hex('#F7F6F3'), textFormat:{fontSize:10,foregroundColor:hex('#2F3336'),italic:true},
    verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});

  // Category legend below conflicts
  const legendRow = conflictRow + 5;
  values.push({ range:`'Weekly Schedule'!A${legendRow+1}`, values:[['CATEGORY COLOR LEGEND']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,legendRow,legendRow+1,0,14), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,legendRow,legendRow+1,0,14), cell:{userEnteredFormat:{
    backgroundColor:hex('#3D6374'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF')},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:12},
  }}, fields:'userEnteredFormat' }});

  const cats = Object.keys(CAT_COLORS);
  cats.forEach((cat,i)=>{
    const col = (i%8)*1;  // 8 per row, 1 col each
    const row = legendRow + 1 + Math.floor(i/8);
    values.push({ range:`'Weekly Schedule'!${cl(col)}${row+1}`, values:[[cat]] });
    fmt.push({ repeatCell:{ range: gridRange(SID,row,row+1,col,col+1), cell:{userEnteredFormat:{
      backgroundColor:hex(CAT_COLORS[cat]),
      textFormat:{bold:true,fontSize:8,foregroundColor:hex('#2F3336')},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
  });

  // Freeze row 6 (day headers)
  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:6}},
    fields:'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, values, '11-weekly values');
  await batchUpdate(id, fmt, '11-weekly format');
  console.log('✅ Weekly Schedule done.');
})().catch(e => { console.error(e.errors||e.message||e); process.exit(1); });
