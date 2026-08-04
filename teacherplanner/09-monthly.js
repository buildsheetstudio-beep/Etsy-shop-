'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Monthly Planner'];
const S = "'Monthly Planner'";

// Layout: 3 months (Sept, Oct, Nov 2025), each on its own section
// Each month = header row + day-of-week row + 6 calendar rows × 7 cols (Mon-Sun) + 2-row notes + spacer
// Cols A-G = Mon through Sun
// Cells contain: day number + event/note text

// School events from 03-setup.js (subset for these months)
const EVENTS = {
  '2025-09-02': 'First Day of School',
  '2025-09-03': 'Curriculum Night Info sent home',
  '2025-09-12': 'School Picture Day',
  '2025-09-19': 'Professional Development (no students)',
  '2025-09-25': 'Back-to-School Night 6:30 PM',
  '2025-10-13': 'Fall Break',
  '2025-10-14': 'Fall Break',
  '2025-10-17': 'Fall Break ends',
  '2025-10-31': 'Halloween Parade 2:00 PM',
  '2025-11-05': 'Election Day — no school',
  '2025-11-11': 'Veterans Day — no school',
  '2025-11-20': 'Parent-Teacher Conferences (half day)',
  '2025-11-21': 'Parent-Teacher Conferences (half day)',
  '2025-11-26': 'Thanksgiving Break',
  '2025-11-27': 'Thanksgiving — no school',
  '2025-11-28': 'Thanksgiving Break',
};

// Teacher personal notes per date
const TEACHER_NOTES = {
  '2025-09-04': 'Collect supply lists',
  '2025-09-09': 'Running records: Group A',
  '2025-09-10': 'Running records: Group B',
  '2025-09-15': 'Grade reading logs',
  '2025-09-22': 'Reports due to office',
  '2025-09-26': 'Order science supplies',
  '2025-10-06': 'Submit grades Q1',
  '2025-10-08': 'Call Julia Brown parent re: missing work',
  '2025-10-15': 'Copy research project rubric',
  '2025-10-22': 'Prepare parent conference notes',
  '2025-10-28': 'Final project grades due',
  '2025-11-03': 'Benchmark assessments begin',
  '2025-11-07': 'Benchmark results to admin',
  '2025-11-12': 'Grade science notebooks',
  '2025-11-14': 'Order holiday party supplies',
  '2025-11-17': 'Progress reports',
  '2025-11-19': 'Set up conference schedule',
};

const MONTHS = [
  { name:'September 2025', year:2025, month:9 },
  { name:'October 2025',   year:2025, month:10 },
  { name:'November 2025',  year:2025, month:11 },
];

function getFirstDayOffset(year, month) {
  const d = new Date(year, month - 1, 1);
  let dow = d.getDay(); // 0=Sun, 1=Mon ... 6=Sat
  return dow === 0 ? 6 : dow - 1; // convert to Mon=0 ... Sun=6
}

function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function padDate(n) { return n < 10 ? '0' + n : '' + n; }

function buildCalendar(year, month) {
  const offset = getFirstDayOffset(year, month);
  const days = getDaysInMonth(year, month);
  const grid = []; // 6 rows × 7 cols
  for (let r = 0; r < 6; r++) grid.push(Array(7).fill(''));
  let dayNum = 1 - offset;
  for (let r = 0; r < 6; r++) {
    for (let c = 0; c < 7; c++) {
      if (dayNum >= 1 && dayNum <= days) {
        const key = `${year}-${padDate(month)}-${padDate(dayNum)}`;
        const event = EVENTS[key] || '';
        const note  = TEACHER_NOTES[key] || '';
        const combined = [event, note].filter(Boolean).join('\n');
        grid[r][c] = `${dayNum}${combined ? '\n' + combined : ''}`;
      }
      dayNum++;
    }
  }
  return grid;
}

(async () => {
  const vals = [];
  const fmt  = [];

  // BG
  fmt.push({ repeatCell:{ range: gridRange(SID,0,200,0,16), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});

  // Title row 1
  vals.push({ range:`${S}!A1`, values:[['MONTHLY PLANNER']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1}, properties:{pixelSize:44}, fields:'pixelSize' }});

  // Subtitle row 2
  vals.push({ range:`${S}!A2`, values:[['Month-at-a-glance view. School events from calendar auto-populated. Add your own notes directly in any date cell.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:2}, properties:{pixelSize:28}, fields:'pixelSize' }});

  // Column widths (Mon-Sun, 7 cols + 1 extra)
  const colWidths = [100,120,120,120,120,120,120,80];
  colWidths.forEach((px, c) => {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:c,endIndex:c+1}, properties:{pixelSize:px}, fields:'pixelSize' }});
  });

  await batchUpdate(id, fmt, '09-monthly format');

  const monthFmt = [];
  const monthVals = [];
  const DOW_COLORS = [C.ELA,C.ELA,C.Math,C.Science,C.SocialStudies,'#E6E6E6','#D5D5D5'];
  const MONTH_ACCENT_COLORS = [C.ELA, C.SocialStudies, C.Science];

  const startSheetRow = 3; // 0-indexed
  const BLOCK_ROWS = 11; // 1 month header + 1 dow header + 6 cal rows + 1 notes label + 1 notes content + 1 spacer

  MONTHS.forEach((mon, mi) => {
    const block = startSheetRow + mi * BLOCK_ROWS;

    // Month header
    monthVals.push({ range:`${S}!A${block+1}`, values:[[mon.name,'','','','','','','']]});
    monthFmt.push({ mergeCells:{ range: gridRange(SID,block,block+1,0,7), mergeType:'MERGE_ALL' }});
    monthFmt.push({ repeatCell:{ range: gridRange(SID,block,block+1,0,7), cell:{userEnteredFormat:{
      backgroundColor:hex(MONTH_ACCENT_COLORS[mi]),
      textFormat:{bold:true,fontSize:13,foregroundColor:hex(C.text),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      borders:{ bottom:{style:'SOLID',color:hex(C.border)} },
    }}, fields:'userEnteredFormat' }});
    monthFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:block,endIndex:block+1}, properties:{pixelSize:32}, fields:'pixelSize' }});

    // Day of week header
    const dowRow = block + 1;
    monthVals.push({ range:`${S}!A${dowRow+1}`, values:[['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']]});
    for (let c = 0; c < 7; c++) {
      monthFmt.push({ repeatCell:{ range: gridRange(SID,dowRow,dowRow+1,c,c+1), cell:{userEnteredFormat:{
        backgroundColor:hex(DOW_COLORS[c]),
        textFormat:{bold:true,fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
        borders:{ bottom:{style:'SOLID',color:hex(C.border)} },
      }}, fields:'userEnteredFormat' }});
    }
    monthFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:dowRow,endIndex:dowRow+1}, properties:{pixelSize:22}, fields:'pixelSize' }});

    // Calendar rows (6 rows)
    const calGrid = buildCalendar(mon.year, mon.month);
    calGrid.forEach((row, ri) => {
      const calRow = block + 2 + ri;
      monthVals.push({ range:`${S}!A${calRow+1}`, values:[row]});
      // Style each cell
      for (let c = 0; c < 7; c++) {
        const isEmpty = row[c] === '';
        monthFmt.push({ repeatCell:{ range: gridRange(SID,calRow,calRow+1,c,c+1), cell:{userEnteredFormat:{
          backgroundColor: isEmpty ? hex(C.altRow) : hex(C.panel),
          textFormat:{fontSize:8,foregroundColor:hex(isEmpty ? C.border : C.text),fontFamily:'Arial'},
          verticalAlignment:'TOP',
          wrapStrategy:'WRAP',
          borders:{
            top:{style:'SOLID',color:hex(C.border)},
            right:{style:'SOLID',color:hex(C.border)},
          },
        }}, fields:'userEnteredFormat' }});
      }
      monthFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:calRow,endIndex:calRow+1}, properties:{pixelSize:65}, fields:'pixelSize' }});
    });

    // Notes row (month notes)
    const notesLabelRow = block + 8;
    const notesContentRow = block + 9;
    monthVals.push({ range:`${S}!A${notesLabelRow+1}`, values:[['Monthly Notes','','','','','','','']]});
    monthFmt.push({ mergeCells:{ range: gridRange(SID,notesLabelRow,notesLabelRow+1,0,7), mergeType:'MERGE_ALL' }});
    monthFmt.push({ repeatCell:{ range: gridRange(SID,notesLabelRow,notesLabelRow+1,0,7), cell:{userEnteredFormat:{
      backgroundColor:hex(C.warning), textFormat:{bold:true,fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:8},
    }}, fields:'userEnteredFormat' }});
    monthFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:notesLabelRow,endIndex:notesLabelRow+1}, properties:{pixelSize:22}, fields:'pixelSize' }});

    const monthNotes = [
      'Sept goals: Establish routines, complete running records, begin guided reading groups.',
      'Oct goals: Unit 1 assessments, parent conferences, Q1 grades. Watch weather for outdoor activities.',
      'Nov goals: Benchmark assessments, progress reports, parent conferences 11/20-21. Begin holiday projects.',
    ];
    monthVals.push({ range:`${S}!A${notesContentRow+1}`, values:[[monthNotes[mi],'','','','','','','']]});
    monthFmt.push({ mergeCells:{ range: gridRange(SID,notesContentRow,notesContentRow+1,0,7), mergeType:'MERGE_ALL' }});
    monthFmt.push({ repeatCell:{ range: gridRange(SID,notesContentRow,notesContentRow+1,0,7), cell:{userEnteredFormat:{
      backgroundColor:hex('#FFFBF0'), textFormat:{fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
      verticalAlignment:'MIDDLE', wrapStrategy:'WRAP',
    }}, fields:'userEnteredFormat' }});
    monthFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:notesContentRow,endIndex:notesContentRow+1}, properties:{pixelSize:30}, fields:'pixelSize' }});

    // Spacer
    const spacerRow = block + 10;
    monthFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:spacerRow,endIndex:spacerRow+1}, properties:{pixelSize:12}, fields:'pixelSize' }});
    monthFmt.push({ repeatCell:{ range: gridRange(SID,spacerRow,spacerRow+1,0,7), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});
  });

  await valuesBatchUpdate(id, monthVals, '09-monthly values');
  await batchUpdate(id, monthFmt, '09-monthly calendar fmt');

  // Freeze top 2 rows
  await batchUpdate(id, [{ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:2 }}, fields:'gridProperties.frozenRowCount' }}], '09-monthly freeze');

  console.log('✅ Monthly Planner done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
