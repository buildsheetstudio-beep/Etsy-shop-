'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['School Year Setup'];
const S = "'School Year Setup'";
const REF = "'Reference Data'";

(async () => {
  const vals = [];
  const fmt  = [];

  // ── Title banner ────────────────────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['ULTIMATE TEACHER PLANNER — SCHOOL YEAR SETUP']] });
  vals.push({ range: `${S}!A2`, values: [['Configure your teacher details, school year dates, classes, grading weights, and school calendar.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,20), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,20), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex('#2F3437'), textFormat: { bold:true, fontSize:16, foregroundColor: hex('#FFFFFF'), fontFamily:'Arial' },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex('#4A5056'), textFormat: { fontSize:9, foregroundColor: hex('#D5D8DB'), italic:true, fontFamily:'Arial' },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1}, properties:{pixelSize:48}, fields:'pixelSize' }});
  fmt.push({ updateDimensionProperties: { range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:2}, properties:{pixelSize:24}, fields:'pixelSize' }});

  // ── Section A: Teacher & School Info (rows 3-18) ─────────────────
  vals.push({ range: `${S}!A3`, values: [['TEACHER & SCHOOL INFORMATION']] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.ELA), textFormat:{ bold:true, fontSize:11, foregroundColor:hex(C.text), fontFamily:'Arial' },
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range:{sheetId:SID,dimension:'ROWS',startIndex:2,endIndex:3}, properties:{pixelSize:30}, fields:'pixelSize' }});

  const teacherFields = [
    ['Teacher Name','Sarah Mitchell'],
    ['School Name','Maplewood Elementary School'],
    ['School Level','Elementary'],
    ['Grade / Grade Band','Grade 4'],
    ['School Year Label','2025–2026'],
    ['School Year Start Date','08/25/2025'],
    ['School Year End Date','06/05/2026'],
    ['Term Type','Full Year'],
    ['Number of Terms','1'],
    ['Current Term','1'],
    ['Default Week Start','Monday'],
    ['Primary Contact Email','smitchell@maplewoodelementary.edu'],
    ['Notes','Fourth grade homeroom teacher — 3 core subjects plus specials rotation'],
  ];

  teacherFields.forEach(([label, val], i) => {
    const row = i + 4;
    vals.push({ range: `${S}!A${row}`, values: [[label]] });
    vals.push({ range: `${S}!B${row}`, values: [[val]] });
  });

  // Label col A styling
  fmt.push({ repeatCell: { range: gridRange(SID,3,3+teacherFields.length,0,1), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) },
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  // Input col B styling
  fmt.push({ repeatCell: { range: gridRange(SID,3,3+teacherFields.length,1,4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input), textFormat:{ fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) },
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  // Dropdown for School Level
  fmt.push({ setDataValidation: { range: gridRange(SID,5,6,1,2), rule: {
    condition: { type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!A2:A6`}] },
    showCustomUi: true, strict: false,
  }}});
  // Dropdown for Term Type
  fmt.push({ setDataValidation: { range: gridRange(SID,10,11,1,2), rule: {
    condition: { type:'ONE_OF_LIST', values:[
      {userEnteredValue:'Full Year'},{userEnteredValue:'Semester'},
      {userEnteredValue:'Trimester'},{userEnteredValue:'Quarter'},{userEnteredValue:'Custom'}
    ]},
    showCustomUi:true, strict:false,
  }}});

  // Date format for Start/End Date
  [8,9].forEach(r => {
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,1,2), cell: { userEnteredFormat: {
      numberFormat: { type:'DATE', pattern:'mmm d, yyyy' },
    }}, fields:'userEnteredFormat.numberFormat' }});
  });

  // ── Section B: Terms (rows 19-26) ────────────────────────────────
  const termRow = 3 + teacherFields.length + 2; // ~19
  vals.push({ range: `${S}!A${termRow}`, values: [['TERM SETUP']] });
  fmt.push({ mergeCells: { range: gridRange(SID,termRow-1,termRow,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,termRow-1,termRow,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.Math), textFormat:{ bold:true, fontSize:11, foregroundColor:hex(C.text), fontFamily:'Arial' },
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range:{sheetId:SID,dimension:'ROWS',startIndex:termRow-1,endIndex:termRow}, properties:{pixelSize:30}, fields:'pixelSize' }});

  const termHdrs = ['Term ID','Term Name','Start Date','End Date','Active?','Notes'];
  vals.push({ range: `${S}!A${termRow+1}`, values: [termHdrs] });
  fmt.push({ repeatCell: { range: gridRange(SID,termRow,termRow+1,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat' }});

  const termData = [
    ['TRM-001','Full Year 2025–26','08/25/2025','06/05/2026',true,'Primary school year'],
  ];
  termData.forEach((row, i) => {
    vals.push({ range: `${S}!A${termRow+2+i}`, values: [row] });
  });
  // Add 5 empty term rows with checkbox
  for (let i = 1; i < 6; i++) {
    fmt.push({ setDataValidation: { range: gridRange(SID,termRow+1+i,termRow+2+i,4,5), rule: {
      condition: { type:'BOOLEAN' }, showCustomUi:true,
    }}});
  }
  // Date format for term start/end
  for (let i = 0; i < 6; i++) {
    [2,3].forEach(c => {
      fmt.push({ repeatCell: { range: gridRange(SID,termRow+1+i,termRow+2+i,c,c+1), cell: { userEnteredFormat: {
        numberFormat:{type:'DATE',pattern:'mmm d, yyyy'},
      }}, fields:'userEnteredFormat.numberFormat' }});
    });
  }

  // ── Section C: Class Setup (rows ~27 – ~48) ───────────────────────
  const clsRow = termRow + 9; // ~28
  vals.push({ range: `${S}!A${clsRow}`, values: [['CLASS / SUBJECT SETUP']] });
  fmt.push({ mergeCells: { range: gridRange(SID,clsRow-1,clsRow,0,16), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,clsRow-1,clsRow,0,16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.Science), textFormat:{ bold:true, fontSize:11, foregroundColor:hex(C.text), fontFamily:'Arial' },
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range:{sheetId:SID,dimension:'ROWS',startIndex:clsRow-1,endIndex:clsRow}, properties:{pixelSize:30}, fields:'pixelSize' }});

  const clsHdrs = ['Class ID','Class Name','Class Type','Subject','Period/Block','Grade Level',
                   'Room','Days Met','Start Time','End Time','Teacher / Co-Teacher',
                   'Capacity','Grading Scale','Active?','Notes'];
  vals.push({ range: `${S}!A${clsRow+1}`, values: [clsHdrs] });
  fmt.push({ repeatCell: { range: gridRange(SID,clsRow,clsRow+1,0,15), cell: { userEnteredFormat: {
    backgroundColor: hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat' }});

  // Class ID formula in col A
  for (let i = 0; i < 20; i++) {
    const r = clsRow + 1 + i;
    vals.push({ range: `${S}!A${r+1}`, values: [[`=IF(B${r+1}="","","CLS-"&TEXT(ROW()-${clsRow+1},"000"))`]] });
  }

  // 10 sample classes
  const classData = [
    ['','ELA – Grade 4',      'Homeroom','English Language Arts','Block 1','Grade 4','204','Mon–Fri','8:15 AM','9:00 AM','Sarah Mitchell','24','Standard','',''    ],
    ['','Reading – Grade 4',  'Subject', 'Reading',              'Block 2','Grade 4','204','Mon–Fri','9:05 AM','9:50 AM','Sarah Mitchell','24','Standard',true,''  ],
    ['','Math – Grade 4',     'Subject', 'Mathematics',          'Block 3','Grade 4','204','Mon–Fri','10:00 AM','10:50 AM','Sarah Mitchell','24','Standard',true,''],
    ['','Science – Grade 4',  'Subject', 'Science',              'Block 4','Grade 4','204','Mon/Wed/Fri','11:00 AM','11:45 AM','Sarah Mitchell','24','Standard',true,''],
    ['','Social Studies G4',  'Subject', 'Social Studies',       'Block 5','Grade 4','204','Tue/Thu','11:00 AM','11:45 AM','Sarah Mitchell','24','Standard',true,'' ],
    ['','Writing – Grade 4',  'Subject', 'Writing',              'Block 6','Grade 4','204','Mon–Fri','1:30 PM','2:15 PM','Sarah Mitchell','24','Standard',true,''   ],
    ['','Art – Grade 4',      'Subject', 'Art',                  'Specials','Grade 4','Art Room','Mon','2:20 PM','3:05 PM','Mrs. Rivera','24','Standard',true,''    ],
    ['','Music – Grade 4',    'Subject', 'Music',                'Specials','Grade 4','Music Room','Wed','2:20 PM','3:05 PM','Mr. Torres','24','Standard',true,''   ],
    ['','PE – Grade 4',       'Subject', 'Physical Education',   'Specials','Grade 4','Gym','Fri','2:20 PM','3:05 PM','Coach Davis','24','Standard',true,''         ],
    ['','Advisory – Grade 4', 'Homeroom','Advisory',             'Morning','Grade 4','204','Mon–Fri','8:00 AM','8:15 AM','Sarah Mitchell','24','Standard',true,''   ],
  ];

  classData.forEach((row, i) => {
    const r = clsRow + 2 + i;
    vals.push({ range: `${S}!B${r}:O${r}`, values: [row.slice(1)] });
  });

  // Active checkbox validation
  for (let i = 0; i < 20; i++) {
    fmt.push({ setDataValidation: { range: gridRange(SID,clsRow+1+i,clsRow+2+i,13,14), rule: {
      condition:{ type:'BOOLEAN' }, showCustomUi:true,
    }}});
  }
  // Class Type dropdown
  for (let i = 0; i < 20; i++) {
    fmt.push({ setDataValidation: { range: gridRange(SID,clsRow+1+i,clsRow+2+i,2,3), rule: {
      condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!C2:C8`}] }, showCustomUi:true, strict:false,
    }}});
  }
  // Subject dropdown
  for (let i = 0; i < 20; i++) {
    fmt.push({ setDataValidation: { range: gridRange(SID,clsRow+1+i,clsRow+2+i,3,4), rule: {
      condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!D2:D18`}] }, showCustomUi:true, strict:false,
    }}});
  }

  // Alternating row fill for class table
  for (let i = 0; i < 20; i++) {
    fmt.push({ repeatCell: { range: gridRange(SID,clsRow+1+i,clsRow+2+i,0,15), cell: { userEnteredFormat: {
      backgroundColor: hex(i%2===0 ? C.panel : C.altRow),
      textFormat:{ fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) },
    }}, fields:'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // ── Section D: Grade Category Weights (below class table) ─────────
  const wRow = clsRow + 22; // below class table
  vals.push({ range: `${S}!A${wRow}`, values: [['GRADING CATEGORY WEIGHTS (per class)']] });
  fmt.push({ mergeCells: { range: gridRange(SID,wRow-1,wRow,0,6), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,wRow-1,wRow,0,6), cell: { userEnteredFormat: {
    backgroundColor: hex(C.Writing), textFormat:{ bold:true, fontSize:11, foregroundColor:hex(C.text), fontFamily:'Arial' },
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  const wHdrs = ['Class ID','Grade Category','Weight %','Included?','Notes'];
  vals.push({ range: `${S}!A${wRow+1}`, values: [wHdrs] });
  fmt.push({ repeatCell: { range: gridRange(SID,wRow,wRow+1,0,5), cell: { userEnteredFormat: {
    backgroundColor: hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat' }});

  const weightData = [
    ['CLS-001','Homework',10,true,''],
    ['CLS-001','Classwork',20,true,''],
    ['CLS-001','Quizzes',20,true,''],
    ['CLS-001','Tests',30,true,''],
    ['CLS-001','Projects',20,true,''],
    ['CLS-002','Homework',15,true,''],
    ['CLS-002','Classwork',25,true,''],
    ['CLS-002','Quizzes',25,true,''],
    ['CLS-002','Tests',35,true,'Weights intentionally do not total 100% — review needed'],  // intentional gap
    ['CLS-003','Homework',10,true,''],
    ['CLS-003','Classwork',20,true,''],
    ['CLS-003','Quizzes',20,true,''],
    ['CLS-003','Tests',30,true,''],
    ['CLS-003','Projects',20,true,''],
    ['CLS-004','Homework',10,true,''],
    ['CLS-004','Classwork',20,true,''],
    ['CLS-004','Labs',30,true,''],
    ['CLS-004','Tests',25,true,''],
    ['CLS-004','Projects',15,true,''],
  ];

  weightData.forEach((row, i) => {
    vals.push({ range: `${S}!A${wRow+2+i}`, values: [row] });
  });
  for (let i = 0; i < 30; i++) {
    fmt.push({ setDataValidation: { range: gridRange(SID,wRow+1+i,wRow+2+i,3,4), rule: {
      condition:{ type:'BOOLEAN' }, showCustomUi:true,
    }}});
    fmt.push({ repeatCell: { range: gridRange(SID,wRow+1+i,wRow+2+i,2,3), cell: { userEnteredFormat: {
      numberFormat:{ type:'PERCENT', pattern:'0%' },
    }}, fields:'userEnteredFormat.numberFormat' }});
  }

  // ── Section E: School Calendar Dates ─────────────────────────────
  const calRow = wRow + 22;
  vals.push({ range: `${S}!A${calRow}`, values: [['SCHOOL CALENDAR DATES']] });
  fmt.push({ mergeCells: { range: gridRange(SID,calRow-1,calRow,0,8), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,calRow-1,calRow,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.SocialStudies), textFormat:{ bold:true, fontSize:11, foregroundColor:hex(C.text), fontFamily:'Arial' },
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  const calHdrs = ['Date ID','Date','End Date','Category','Title','School Closed?','Notes'];
  vals.push({ range: `${S}!A${calRow+1}`, values: [calHdrs] });
  fmt.push({ repeatCell: { range: gridRange(SID,calRow,calRow+1,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex('#4A5056'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER',
  }}, fields:'userEnteredFormat' }});

  const calDates = [
    [`=IF(B${calRow+3}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '09/01/2025','','Holiday','Labor Day',true,'No school'],
    [`=IF(B${calRow+4}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '10/13/2025','','Holiday','Columbus Day',true,'No school'],
    [`=IF(B${calRow+5}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '11/11/2025','','Holiday','Veterans Day',true,'No school'],
    [`=IF(B${calRow+6}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '11/26/2025','11/28/2025','Holiday','Thanksgiving Break',true,''],
    [`=IF(B${calRow+7}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '12/22/2025','01/02/2026','Holiday','Winter Break',true,''],
    [`=IF(B${calRow+8}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '01/19/2026','','Holiday','MLK Day',true,''],
    [`=IF(B${calRow+9}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '02/16/2026','','Holiday','Presidents Day',true,''],
    [`=IF(B${calRow+10}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '03/20/2026','','Teacher Workday','Spring PD Day',false,'Prof. development'],
    [`=IF(B${calRow+11}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '04/03/2026','04/10/2026','Holiday','Spring Break',true,''],
    [`=IF(B${calRow+12}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '05/25/2026','','Holiday','Memorial Day',true,''],
    [`=IF(B${calRow+13}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '11/20/2025','11/21/2025','Conferences','Parent–Teacher Conferences',false,''],
    [`=IF(B${calRow+14}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '01/15/2026','','Report Cards','Quarter 2 Report Cards Due',false,''],
    [`=IF(B${calRow+15}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '03/02/2026','03/06/2026','Testing','State Assessment Week',false,'Reading & Math tests'],
    [`=IF(B${calRow+16}="","","CAL-"&TEXT(ROW()-${calRow+1},"000"))`, '05/15/2026','','School Event','Spring Concert',false,'Evening event'],
  ];

  calDates.forEach((row, i) => {
    vals.push({ range: `${S}!A${calRow+2+i}`, values: [row] });
  });
  // Date format and closed checkbox
  for (let i = 0; i < 20; i++) {
    [1,2].forEach(c => {
      fmt.push({ repeatCell: { range: gridRange(SID,calRow+1+i,calRow+2+i,c,c+1), cell: { userEnteredFormat: {
        numberFormat:{type:'DATE',pattern:'mmm d, yyyy'},
      }}, fields:'userEnteredFormat.numberFormat' }});
    });
    fmt.push({ setDataValidation: { range: gridRange(SID,calRow+1+i,calRow+2+i,5,6), rule: {
      condition:{ type:'BOOLEAN' }, showCustomUi:true,
    }}});
  }

  // ── Section F: Summary Cards ──────────────────────────────────────
  // Place summary cards at columns J–P of teacher info section (rows 4-9)
  const sumCol = 9; // col J (0-indexed)
  const sumLabels = [
    ['Active Classes',       `=IFERROR(COUNTIF('School Year Setup'!N${clsRow+2}:N${clsRow+21},TRUE),0)`],
    ['Active Students',      `=IFERROR(COUNTIF('Student Roster'!Q6:Q1005,TRUE),0)`],
    ['Current Term',         `=IFERROR(B12,"")`],
    ['Days Remaining',       `=MAX(0,B9-TODAY())`],
    ['Assignments Planned',  `=IFERROR(COUNTIF('Assignments & Assessments'!O6:O1505,"Planned"),0)`],
    ['Tasks Remaining',      `=IFERROR(COUNTIFS('Teacher To-Do'!J6:J1505,"<>Complete",'Teacher To-Do'!J6:J1505,"<>Cancelled"),0)`],
  ];

  vals.push({ range: `${S}!J3`, values: [['SUMMARY']] });
  fmt.push({ mergeCells: { range: gridRange(SID,2,3,sumCol,sumCol+4), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,sumCol,sumCol+4), cell: { userEnteredFormat: {
    backgroundColor: hex('#2F3437'), textFormat:{ bold:true, fontSize:11, foregroundColor:hex('#FFFFFF'), fontFamily:'Arial' },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  sumLabels.forEach(([label, formula], i) => {
    vals.push({ range: `${S}!J${4+i}`, values: [[label]] });
    vals.push({ range: `${S}!L${4+i}`, values: [[formula]] });
  });
  fmt.push({ repeatCell: { range: gridRange(SID,3,9,sumCol,sumCol+2), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow), textFormat:{ bold:true, fontSize:9, fontFamily:'Arial', foregroundColor:hex(C.text) },
    verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell: { range: gridRange(SID,3,9,sumCol+2,sumCol+4), cell: { userEnteredFormat: {
    backgroundColor: hex(C.formula), textFormat:{ bold:true, fontSize:12, fontFamily:'Arial', foregroundColor:hex(C.text) },
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  // ── Column widths ─────────────────────────────────────────────────
  const colWidths = [80,180,120,120,90,90,80,100,90,90,140,80,110,70,200];
  colWidths.forEach((w,i) => {
    fmt.push({ updateDimensionProperties: { range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1}, properties:{pixelSize:w}, fields:'pixelSize' }});
  });

  // Freeze rows 1:3
  fmt.push({ updateSheetProperties: { properties:{ sheetId:SID, gridProperties:{ frozenRowCount:3 } }, fields:'gridProperties.frozenRowCount' }});

  // BG fill
  fmt.push({ repeatCell: { range: gridRange(SID,0,200,0,20), cell:{ userEnteredFormat:{ backgroundColor:hex(C.bg) } }, fields:'userEnteredFormat.backgroundColor' }});

  await valuesBatchUpdate(id, vals, '03-setup values');
  await batchUpdate(id, fmt, '03-setup format');
  console.log('✅ School Year Setup done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
