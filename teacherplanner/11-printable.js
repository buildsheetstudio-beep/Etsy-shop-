'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Printable Weekly Overview'];
const S = "'Printable Weekly Overview'";

// A clean one-page printable view of Week 1 lesson plans
// Layout optimized for printing on one sheet (portrait or landscape)
// Rows:
// 1-2: Title + teacher info
// 3: Column headers (Class | Mon | Tue | Wed | Thu | Fri)
// 4-8: 5 class rows
// 9: blank spacer
// 10-12: Reminders / Notes section
// 13: Standards quick reference

// Uses static values from Week 1 of the Weekly Planner
const WEEK_LABEL = 'Week 1 — September 1–5, 2025';
const TEACHER_NAME = 'Sarah Mitchell';
const SCHOOL = 'Maplewood Elementary | Grade 4';

const CLASSES = [
  { name:'ELA (CLS-001)', color: C.ELA,
    plans:['Intro Reading Workshop; establish routines','Phonics: short vowels review','Read-Aloud: Charlotte\'s Web Ch 1-2','Character traits discussion','Vocab journals: 10 words'],
    std:'CCSS.ELA-RF.4.4; RL.4.3', hw:'Reading log 20 min/night' },
  { name:'Reading (CLS-002)', color: C.ELA,
    plans:['Establish reading groups; running records','Fluency partner reading','Guided reading: Group A','Visualizing strategy','Response journals'],
    std:'CCSS.ELA-RF.4.4; RL.4.1', hw:'Read 20 min + log' },
  { name:'Mathematics (CLS-003)', color: C.Math,
    plans:['Place value intro + anchor chart','Expanded form practice','Compare + order numbers','Round to nearest 10, 100','Place value quiz prep'],
    std:'CCSS.MATH.4.NBT.A.2', hw:'Place value worksheet' },
  { name:'Science (CLS-004)', color: C.Science,
    plans:['Safety rules + science notebooks','Forms of energy intro','Energy transformations activity','Lab: rolling objects + energy','Vocab cards'],
    std:'NGSS.4-PS3', hw:'Science notebook reflection' },
  { name:'Social Studies (CLS-005)', color: C.SocialStudies,
    plans:['Map skills: compass rose + legend','Political vs physical maps','Latitude + longitude intro','Practice: find places on map','Map quiz prep'],
    std:'C3.D2.Geo.1', hw:'Map worksheet' },
];

const REMINDERS = [
  'Back to School Night: Thursday Sept 4, 6:30 PM — send home reminder letter',
  'Collect student supply lists by Friday',
  'Run records: begin Group A (reading) Monday — complete all groups by Sept 19',
  'Submit class roster to office by September 5',
  'Fire drill practice: check schedule for this week',
];

(async () => {
  const vals = [];
  const fmt  = [];

  // Set page break / print area hint by sizing columns for portrait print
  // 10 cols: A=Class(90), B-F=Days(110 each), G=Standards(120), H=HW(110), I-J=blank
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:0,endIndex:1}, properties:{pixelSize:105}, fields:'pixelSize' }});
  [1,2,3,4,5].forEach(c => {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:c,endIndex:c+1}, properties:{pixelSize:118}, fields:'pixelSize' }});
  });
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:6,endIndex:7}, properties:{pixelSize:140}, fields:'pixelSize' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:7,endIndex:8}, properties:{pixelSize:120}, fields:'pixelSize' }});

  // BG
  fmt.push({ repeatCell:{ range: gridRange(SID,0,100,0,10), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});

  // Title block rows 1-2
  vals.push({ range:`${S}!A1`, values:[['WEEKLY LESSON PLAN OVERVIEW']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:15,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1}, properties:{pixelSize:40}, fields:'pixelSize' }});

  vals.push({ range:`${S}!A2`, values:[[`${TEACHER_NAME} | ${SCHOOL} | ${WEEK_LABEL}`]] });
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:false,fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:2}, properties:{pixelSize:24}, fields:'pixelSize' }});

  // Column headers row 3
  vals.push({ range:`${S}!A3`, values:[['Class','Monday\nSep 1','Tuesday\nSep 2','Wednesday\nSep 3','Thursday\nSep 4','Friday\nSep 5','Standards','Homework Assigned']] });
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex(C.altRow), textFormat:{bold:true,fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    borders:{ bottom:{style:'SOLID',color:hex(C.border)}, top:{style:'SOLID',color:hex(C.border)} },
    wrapStrategy:'WRAP',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:2,endIndex:3}, properties:{pixelSize:30}, fields:'pixelSize' }});

  // 5 class rows (rows 4-8)
  CLASSES.forEach((cls, ci) => {
    const r = 3 + ci;
    vals.push({ range:`${S}!A${r+1}`, values:[[cls.name, ...cls.plans, cls.std, cls.hw]] });

    // Class name cell
    fmt.push({ repeatCell:{ range: gridRange(SID,r,r+1,0,1), cell:{userEnteredFormat:{
      backgroundColor:hex(cls.color), textFormat:{bold:true,fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      borders:{ bottom:{style:'SOLID',color:hex(C.border)}, right:{style:'SOLID',color:hex(C.border)} },
    }}, fields:'userEnteredFormat' }});

    // Plan cells (cols 1-5)
    fmt.push({ repeatCell:{ range: gridRange(SID,r,r+1,1,6), cell:{userEnteredFormat:{
      backgroundColor:hex(C.panel), textFormat:{fontSize:8,foregroundColor:hex(C.text),fontFamily:'Arial'},
      verticalAlignment:'TOP', wrapStrategy:'WRAP',
      borders:{ bottom:{style:'SOLID',color:hex(C.border)}, right:{style:'SOLID',color:hex(C.border)} },
    }}, fields:'userEnteredFormat' }});

    // Standards + HW cols
    fmt.push({ repeatCell:{ range: gridRange(SID,r,r+1,6,8), cell:{userEnteredFormat:{
      backgroundColor:hex(C.altRow), textFormat:{fontSize:8,foregroundColor:hex(C.secText),fontFamily:'Arial'},
      verticalAlignment:'TOP', wrapStrategy:'WRAP',
      borders:{ bottom:{style:'SOLID',color:hex(C.border)}, right:{style:'SOLID',color:hex(C.border)} },
    }}, fields:'userEnteredFormat' }});

    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:r,endIndex:r+1}, properties:{pixelSize:60}, fields:'pixelSize' }});
  });

  // Spacer row 9
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:8,endIndex:9}, properties:{pixelSize:10}, fields:'pixelSize' }});

  // Reminders section rows 10-12
  vals.push({ range:`${S}!A10`, values:[['THIS WEEK\'S REMINDERS & NOTES']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,9,10,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,9,10,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex(C.warning), textFormat:{bold:true,fontSize:10,foregroundColor:hex(C.text),fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:8},
    borders:{ bottom:{style:'SOLID',color:hex(C.border)} },
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:9,endIndex:10}, properties:{pixelSize:24}, fields:'pixelSize' }});

  // Reminder bullet lines
  const reminderText = REMINDERS.map(r => `• ${r}`).join('\n');
  vals.push({ range:`${S}!A11`, values:[[reminderText]] });
  fmt.push({ mergeCells:{ range: gridRange(SID,10,11,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,10,11,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex('#FFFBF0'), textFormat:{fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
    verticalAlignment:'TOP', wrapStrategy:'WRAP',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:10,endIndex:11}, properties:{pixelSize:70}, fields:'pixelSize' }});

  // Notes line row 12
  vals.push({ range:`${S}!A12`, values:[['Additional Notes:']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,11,12,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,11,12,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex(C.input), textFormat:{fontSize:9,foregroundColor:hex(C.secText),fontFamily:'Arial'},
    verticalAlignment:'MIDDLE',
    borders:{ bottom:{style:'SOLID',color:hex(C.border)} },
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:11,endIndex:12}, properties:{pixelSize:30}, fields:'pixelSize' }});

  // Print disclaimer row 13
  vals.push({ range:`${S}!A13`, values:[['Printed from Ultimate Teacher Planner | Maplewood Elementary | 2025–26 School Year | For classroom use only']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,12,13,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,12,13,0,8), cell:{userEnteredFormat:{
    backgroundColor:hex(C.bg), textFormat:{fontSize:7,foregroundColor:hex(C.border),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:12,endIndex:13}, properties:{pixelSize:18}, fields:'pixelSize' }});

  await valuesBatchUpdate(id, vals, '11-printable values');
  await batchUpdate(id, fmt, '11-printable format');

  console.log('✅ Printable Weekly Overview done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
