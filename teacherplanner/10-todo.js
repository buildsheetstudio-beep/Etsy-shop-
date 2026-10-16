'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Teacher To-Do'];
const S = "'Teacher To-Do'";
const REF = "'Reference Data'";

// Columns (header row 5, data from row 6):
// A: Task ID   B: Task   C: Category   D: Priority   E: Status
// F: Due Date  G: Related Class  H: Estimated Time  I: Actual Time
// J: Completed?  K: Completion Date  L: Notes  M: blank...

// 85 tasks across all categories and statuses
const TASKS = [
  // [task, category, priority, status, due, class, estTime, notes]
  // Lesson Planning
  ['Plan Week 1 ELA unit — Reading Workshop launch','Lesson Planning','High','Complete','08/28/2025','CLS-001','2h',''],
  ['Plan Week 1 Math unit — Place Value','Lesson Planning','High','Complete','08/28/2025','CLS-003','1.5h',''],
  ['Create Week 2 ELA lesson plans','Lesson Planning','High','Complete','09/05/2025','CLS-001','2h',''],
  ['Plan Reading guided reading groups rotations','Lesson Planning','High','Complete','09/05/2025','CLS-002','1h',''],
  ['Develop Week 3 Science energy lesson sequence','Lesson Planning','High','Complete','09/12/2025','CLS-004','1.5h',''],
  ['Create Week 4 Social Studies map skills activities','Lesson Planning','Medium','Complete','09/19/2025','CLS-005','1h',''],
  ['Plan mid-unit ELA comprehension check','Lesson Planning','Medium','Complete','09/22/2025','CLS-001','1h',''],
  ['Create differentiated math practice sheets','Lesson Planning','Medium','Complete','09/26/2025','CLS-003','2h','Created for 3 levels'],
  ['Plan science lab: Energy forms','Lesson Planning','High','Complete','09/12/2025','CLS-004','1h',''],
  ['Write rubric for personal narrative essay','Lesson Planning','High','Complete','09/18/2025','CLS-001','1.5h',''],
  ['Create poetry unit anchor charts','Lesson Planning','Medium','Complete','01/30/2026','CLS-001','1h',''],
  ['Plan research project scaffold','Lesson Planning','High','In Progress','02/06/2026','CLS-001','2h',''],
  ['Develop geometry unit pacing','Lesson Planning','Medium','In Progress','02/06/2026','CLS-003','1h',''],
  ['Plan spring benchmark prep activities','Lesson Planning','High','Not Started','03/01/2026','CLS-003','2h',''],
  ['Create author study project guidelines','Lesson Planning','Medium','Not Started','04/01/2026','CLS-001','1h',''],
  // Grading
  ['Grade Week 1 reading logs — CLS-001','Grading','Medium','Complete','09/08/2025','CLS-001','1h',''],
  ['Grade vocabulary quiz Unit 1','Grading','High','Complete','09/17/2025','CLS-001','1h',''],
  ['Enter grades for ELA comprehension test 1','Grading','High','Complete','10/08/2025','CLS-001','1.5h',''],
  ['Grade personal narrative drafts (24 essays)','Grading','High','Complete','10/20/2025','CLS-001','4h','Long process'],
  ['Grade math test — Operations unit','Grading','High','Complete','10/10/2025','CLS-003','1.5h',''],
  ['Enter science lab rubric scores','Grading','Medium','Complete','09/22/2025','CLS-004','1h',''],
  ['Grade mid-year math assessment','Grading','High','Complete','01/24/2026','CLS-003','2h',''],
  ['Grade research project — animal habitats','Grading','High','In Progress','03/13/2026','CLS-001','4h','Grading in progress'],
  ['Enter attendance for last 2 weeks','Grading','Low','In Progress','02/05/2026','','30m',''],
  ['Grade geometry quiz','Grading','Medium','Not Started','02/25/2026','CLS-003','45m',''],
  ['Finalize Q2 grades in gradebook','Grading','High','Not Started','02/01/2026','','2h',''],
  // Parent Communication
  ['Send welcome letter to CLS-001 families','Parent Communication','High','Complete','09/01/2025','CLS-001','1h',''],
  ['Call parent — Gabriel Santos re: IEP check-in','Parent Communication','High','Complete','09/15/2025','CLS-001','30m','Positive check-in'],
  ['Email parent — Oliver Park behavior update','Parent Communication','High','Complete','10/01/2025','CLS-001','20m',''],
  ['Send newsletter: October update','Parent Communication','Medium','Complete','10/06/2025','','1h',''],
  ['Call Julia Brown parent — missing assignments','Parent Communication','High','Complete','10/09/2025','CLS-001','30m',''],
  ['Prepare conference notes for 24 students','Parent Communication','High','Complete','11/18/2025','CLS-001','3h',''],
  ['Send progress report home — Q1','Parent Communication','High','Complete','11/17/2025','','1h',''],
  ['Email parent Sofia Rodriguez — ELL resources','Parent Communication','Medium','In Progress','02/10/2026','CLS-001','20m',''],
  ['Send February newsletter','Parent Communication','Medium','Not Started','02/02/2026','','1h',''],
  ['Schedule spring parent check-ins','Parent Communication','Medium','Not Started','03/01/2026','','30m',''],
  // Meetings
  ['Grade-level team meeting — Week 1 planning','Meetings','High','Complete','09/03/2025','','1h',''],
  ['IEP meeting — Gabriel Santos','Meetings','High','Complete','09/16/2025','','1h','Team meeting'],
  ['504 meeting — Diana Patel','Meetings','High','Complete','09/23/2025','','30m',''],
  ['Professional Learning Community — Q1','Meetings','Medium','Complete','10/07/2025','','2h',''],
  ['Principal observation debrief','Meetings','High','Complete','10/15/2025','','30m','Positive feedback'],
  ['Grade-level curriculum alignment meeting','Meetings','Medium','Complete','11/04/2025','','1h',''],
  ['Staff meeting — Feb PD focus','Meetings','Medium','Not Started','02/09/2026','','1.5h',''],
  ['IEP annual review — Gabriel Santos','Meetings','High','Not Started','02/25/2026','','1h',''],
  // Classroom Management
  ['Set up classroom library bins by genre','Classroom Management','High','Complete','08/22/2025','','2h',''],
  ['Create classroom rules anchor chart with students','Classroom Management','High','Complete','09/02/2025','CLS-001','30m',''],
  ['Assign classroom jobs — rotate weekly','Classroom Management','Medium','Complete','09/03/2025','CLS-001','20m',''],
  ['Revise behavior support plan — Oliver Park','Classroom Management','High','Complete','10/02/2025','CLS-001','1h',''],
  ['Rearrange seating for Q2','Classroom Management','Medium','Complete','11/03/2025','','30m',''],
  ['Update class community agreements poster','Classroom Management','Low','In Progress','02/12/2026','CLS-001','30m',''],
  // Materials
  ['Order copy paper + pencils for Q2','Materials','Medium','Complete','10/01/2025','','20m',''],
  ['Request science lab materials from supply room','Materials','High','Complete','09/10/2025','CLS-004','30m',''],
  ['Print and laminate anchor charts for ELA','Materials','Medium','Complete','09/04/2025','CLS-001','1h',''],
  ['Order fraction tiles for math unit','Materials','Medium','Complete','10/15/2025','CLS-003','20m',''],
  ['Prepare research project folders for students','Materials','Medium','In Progress','02/09/2026','CLS-001','1h',''],
  ['Order spring benchmark test copies','Materials','High','Not Started','02/15/2026','','30m',''],
  ['Set up science fair display boards','Materials','Medium','Not Started','03/15/2026','CLS-004','1h',''],
  // Professional Development
  ['Complete reading workshop training (online)','Professional Development','High','Complete','09/05/2025','','3h','Completed modules 1-4'],
  ['Read: The Writing Strategies Book by Jennifer Serravallo','Professional Development','Medium','Complete','10/01/2025','','2h','Great resource'],
  ['Attend district ELL strategies workshop','Professional Development','High','Complete','10/10/2025','','6h','Full day PD'],
  ['Complete IEP compliance training','Professional Development','High','Complete','11/01/2025','','1.5h',''],
  ['Observe colleague\'s math class (inquiry-based)','Professional Development','Medium','In Progress','02/13/2026','','1h','Scheduled with Ms. Reed'],
  ['Register for spring ASCD conference','Professional Development','Medium','Not Started','02/28/2026','','30m',''],
  // Administrative
  ['Submit lesson plan binder for October','Administrative','Medium','Complete','10/06/2025','','30m',''],
  ['File student accommodation forms','Administrative','High','Complete','09/04/2025','','1h',''],
  ['Complete Q1 grade reports','Administrative','High','Complete','11/14/2025','','2h',''],
  ['Update emergency contact binder','Administrative','High','Complete','09/08/2025','','1h',''],
  ['Submit field trip permission slips to office','Administrative','Medium','In Progress','02/20/2026','','30m','Science center field trip'],
  ['Complete student information update forms','Administrative','Medium','Not Started','02/20/2026','','1h',''],
  ['Prepare substitute lesson plans folder','Administrative','Medium','Not Started','03/01/2026','','1h',''],
  // Student Support
  ['Check in with Oliver Park daily — behavior plan','Student Support','High','Complete','09/08/2025','CLS-001','15m/day','Going well'],
  ['Arrange ELL reading materials for Sofia Rodriguez','Student Support','High','Complete','09/10/2025','CLS-001','30m',''],
  ['Connect Gabriel Santos with reading intervention','Student Support','High','Complete','09/17/2025','CLS-001','30m','Referral submitted'],
  ['Monitor Diana Patel\'s extended time accommodations','Student Support','High','Complete','09/22/2025','CLS-001','15m','504 in place'],
  ['Check-in meeting with Wendy Chang — GT extensions','Student Support','Medium','Complete','10/14/2025','CLS-001','30m','Enrichment project assigned'],
  ['Review Ethan Kim\'s Q1 progress — below grade level','Student Support','High','Complete','11/07/2025','CLS-003','30m','Intervention plan started'],
  ['Connect Victor Lopez with tutoring resources','Student Support','Medium','In Progress','02/12/2026','CLS-004','30m',''],
  ['Schedule reading intervention assessment','Student Support','High','Not Started','02/20/2026','CLS-002','45m',''],
  ['Update IEP goal progress notes — Gabriel Santos','Student Support','High','Not Started','02/27/2026','CLS-001','1h','Due before March 1'],
  // Other
  ['Decorate bulletin board — October theme','Other','Low','Complete','10/01/2025','','1h',''],
  ['Order holiday gifts for classroom helpers','Other','Low','Complete','12/05/2025','','30m',''],
  ['Photograph student work for portfolio','Other','Medium','In Progress','02/14/2026','CLS-001','1h',''],
  ['Write end-of-year letter to students\' families','Other','Medium','Not Started','05/01/2026','','1h',''],
  ['Reserve gym for spring project presentations','Other','Medium','Not Started','03/01/2026','','20m',''],
];

(async () => {
  const vals = [];
  const fmt  = [];

  // BG
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1510,0,18), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});

  // Title row 1
  vals.push({ range:`${S}!A1`, values:[['TEACHER TO-DO']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,13), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1}, properties:{pixelSize:44}, fields:'pixelSize' }});

  // Subtitle row 2
  vals.push({ range:`${S}!A2`, values:[['Track all tasks, deadlines, and follow-ups. Task ID auto-fills. Sort by Priority or Due Date. Filter by Category or Status.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,13), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  // Priority legend row 3
  vals.push({ range:`${S}!A3`, values:[['Priority: Urgent = red ● High = amber ● Medium = green ● Low = blue  |  Completed rows are dimmed']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,13), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex(C.Writing), textFormat:{fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:8},
  }}, fields:'userEnteredFormat' }});

  vals.push({ range:`${S}!A4`, values:[['']] });
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:3}, properties:{pixelSize:28}, fields:'pixelSize' }});

  // Headers row 5
  const hdrs = ['Task ID','Task / Description','Category','Priority','Status',
                'Due Date','Related Class','Est. Time','Actual Time',
                'Done?','Completion Date','Notes','',''];
  vals.push({ range:`${S}!A5`, values:[hdrs] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,13), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5}, properties:{pixelSize:30}, fields:'pixelSize' }});

  // Freeze rows 1-5
  fmt.push({ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:5 }}, fields:'gridProperties.frozenRowCount' }});

  // Column widths
  const colWidths = [80,260,130,80,100,90,100,80,80,55,100,180,40,40];
  colWidths.forEach((px, c) => {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:c,endIndex:c+1}, properties:{pixelSize:px}, fields:'pixelSize' }});
  });

  await batchUpdate(id, fmt, '10-todo format');

  // Formula column A (Task ID) for 1500 rows
  const aForms = [];
  for (let i = 0; i < 1500; i++) {
    const r = i + 6;
    aForms.push([`=IF(B${r}="","","TSK-"&TEXT(ROW()-5,"0000"))`]);
  }
  await valuesBatchUpdate(id, [{ range:`${S}!A6`, values: aForms }], '10-todo id formulas');

  // Sample task data
  const taskVals = [];
  TASKS.forEach((task, i) => {
    const r = i + 6;
    const [desc, cat, pri, status, due, cls, estTime, notes] = task;
    const done = status === 'Complete';
    const completedDate = done ? due : '';
    taskVals.push({ range:`${S}!B${r}:L${r}`, values:[[desc, cat, pri, status, due, cls, estTime, '', done, completedDate, notes]] });
  });
  await valuesBatchUpdate(id, taskVals, '10-todo task data');

  // Data formatting
  const fmtData = [];

  // Data row height
  fmtData.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:5,endIndex:1505}, properties:{pixelSize:21}, fields:'pixelSize' }});

  // Alternate row tint
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,1505,0,13)],
    booleanRule:{ condition:{type:'CUSTOM_FORMULA',values:[{userEnteredValue:'=AND(MOD(ROW(),2)=0,B6<>"")'}]},
      format:{backgroundColor:hex(C.altRow)} },
  }, index:0}});

  // Complete rows = dimmed (light gray text)
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,1505,0,13)],
    booleanRule:{ condition:{type:'CUSTOM_FORMULA',values:[{userEnteredValue:'=$E6="Complete"'}]},
      format:{textFormat:{foregroundColor:hex(C.secText)},backgroundColor:hex('#F0F0EE')} },
  }, index:1}});

  // Urgent priority = red tint
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,1505,3,4)], // Priority col D
    booleanRule:{ condition:{type:'TEXT_EQ',values:[{userEnteredValue:'Urgent'}]},
      format:{backgroundColor:hex('#F8DCDC'),textFormat:{bold:true}} },
  }, index:2}});

  // High priority = amber tint
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,1505,3,4)],
    booleanRule:{ condition:{type:'TEXT_EQ',values:[{userEnteredValue:'High'}]},
      format:{backgroundColor:hex(C.warning)} },
  }, index:3}});

  // Medium priority = green tint
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,1505,3,4)],
    booleanRule:{ condition:{type:'TEXT_EQ',values:[{userEnteredValue:'Medium'}]},
      format:{backgroundColor:hex(C.Math)} },
  }, index:4}});

  // Low priority = blue tint
  fmtData.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,1505,3,4)],
    booleanRule:{ condition:{type:'TEXT_EQ',values:[{userEnteredValue:'Low'}]},
      format:{backgroundColor:hex(C.ELA)} },
  }, index:5}});

  // Checkbox for Done column (J = index 9)
  fmtData.push({ setDataValidation:{ range: gridRange(SID,5,1505,9,10), rule:{
    condition:{ type:'BOOLEAN' }, showCustomUi:true,
  }}});

  // Dropdowns
  fmtData.push({ setDataValidation:{ range: gridRange(SID,5,1505,2,3), rule:{
    condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!G2:G11`}] }, showCustomUi:true, strict:false,
  }}});
  fmtData.push({ setDataValidation:{ range: gridRange(SID,5,1505,3,4), rule:{
    condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!L2:L5`}] }, showCustomUi:true, strict:false,
  }}});
  fmtData.push({ setDataValidation:{ range: gridRange(SID,5,1505,4,5), rule:{
    condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!H2:H6`}] }, showCustomUi:true, strict:false,
  }}});

  // Center align ID, priority, status, done columns
  [0,3,4,9].forEach(c => {
    fmtData.push({ repeatCell:{ range: gridRange(SID,5,1505,c,c+1), cell:{userEnteredFormat:{
      horizontalAlignment:'CENTER',
    }}, fields:'userEnteredFormat.horizontalAlignment' }});
  });

  await batchUpdate(id, fmtData, '10-todo data format');
  console.log(`✅ Teacher To-Do done. ${TASKS.length} tasks.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
