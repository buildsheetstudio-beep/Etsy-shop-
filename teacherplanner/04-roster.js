'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Student Roster'];
const S = "'Student Roster'";
const REF = "'Reference Data'";

// 65 fictional students across 10 classes
const STUDENTS = [
  // [StudentName, PreferredName, GradeLevel, ClassID, StudentNumber, Email, Parent, ParentEmail, Phone, EmergencyContact, Support?, SupportSummary, Medical, Transport, Active, StartDate, Notes]
  ['Aiden Torres','Aiden','Grade 4','CLS-001','40001','atorres@maplewoodelem.edu','Maria Torres','mtorres@email.com','555-0101','Carlos Torres','','','','Bus Route 14',true,'08/25/2025',''],
  ['Bella Nguyen','Bella','Grade 4','CLS-001','40002','bnguyen@maplewoodelem.edu','Linh Nguyen','lnguyen@email.com','555-0102','David Nguyen','','',' Carries EpiPen — bee allergy','Car pickup',true,'08/25/2025','EpiPen in nurse office'],
  ['Carlos Rivera','Carlos','Grade 4','CLS-001','40003','crivera@maplewoodelem.edu','Ana Rivera','arivera@email.com','555-0103','Jorge Rivera','','','','Bus Route 7',true,'08/25/2025',''],
  ['Diana Patel','Diana','Grade 4','CLS-001','40004','dpatel@maplewoodelem.edu','Priya Patel','ppatel@email.com','555-0104','Raj Patel',true,'Extended time on tests; preferential seating','','Car pickup',true,'08/25/2025','504 Plan'],
  ['Ethan Kim','Ethan','Grade 4','CLS-001','40005','ekim@maplewoodelem.edu','Soo Kim','skim@email.com','555-0105','Jin Kim','','','','Bus Route 14',true,'08/25/2025',''],
  ['Fiona Chen','Fiona','Grade 4','CLS-001','40006','fchen@maplewoodelem.edu','Wei Chen','wchen@email.com','555-0106','Mei Chen','','','','Car pickup',true,'08/25/2025',''],
  ['Gabriel Santos','Gabe','Grade 4','CLS-001','40007','gsantos@maplewoodelem.edu','Rosa Santos','rsantos@email.com','555-0107','Luis Santos',true,'Small group testing; check-in support','','Bus Route 7',true,'08/25/2025','IEP — reading support'],
  ['Hannah Lee','Hannah','Grade 4','CLS-001','40008','hlee@maplewoodelem.edu','Jin Lee','jlee@email.com','555-0108','Sam Lee','','','','Walks home',true,'08/25/2025',''],
  ['Isaac Martinez','Isaac','Grade 4','CLS-001','40009','imartinez@maplewoodelem.edu','Gloria Martinez','gmartinez@email.com','555-0109','Hector Martinez','','','Asthma — inhaler in bag','Bus Route 3',true,'08/25/2025',''],
  ['Julia Brown','Julia','Grade 4','CLS-001','40010','jbrown@maplewoodelem.edu','Lisa Brown','lbrown@email.com','555-0110','Mike Brown','','','','Car pickup',true,'08/25/2025',''],
  ['Kevin Wilson','Kevin','Grade 4','CLS-001','40011','kwilson@maplewoodelem.edu','Sarah Wilson','swilson@email.com','555-0111','Tom Wilson',true,'Preferential seating; extended time','','Bus Route 14',true,'08/25/2025','504 Plan'],
  ['Lily Davis','Lily','Grade 4','CLS-001','40012','ldavis@maplewoodelem.edu','Karen Davis','kdavis@email.com','555-0112','Bob Davis','','','','Car pickup',true,'08/25/2025',''],
  ['Marco Hernandez','Marco','Grade 4','CLS-001','40013','mhernandez@maplewoodelem.edu','Elena Hernandez','ehernandez@email.com','555-0113','Pedro Hernandez','','','','Bus Route 7',true,'08/25/2025',''],
  ['Nadia Johnson','Nadia','Grade 4','CLS-001','40014','njohnson@maplewoodelem.edu','Tracy Johnson','tjohnson@email.com','555-0114','Alan Johnson','','','','Walks home',true,'08/25/2025',''],
  ['Oliver Park','Ollie','Grade 4','CLS-001','40015','opark@maplewoodelem.edu','Mina Park','mpark@email.com','555-0115','Steve Park',true,'Check-in / check-out support','','Car pickup',true,'08/25/2025','Behavior support plan'],
  ['Penelope Adams','Penny','Grade 4','CLS-001','40016','padams@maplewoodelem.edu','Jane Adams','jadams@email.com','555-0116','Bill Adams','','','','Bus Route 3',true,'08/25/2025',''],
  ['Quinn Thomas','Quinn','Grade 4','CLS-001','40017','qthomas@maplewoodelem.edu','Dale Thomas','dthomas@email.com','555-0117','Gail Thomas','','','','Car pickup',true,'08/25/2025',''],
  ['Ryan Garcia','Ryan','Grade 4','CLS-001','40018','rgarcia@maplewoodelem.edu','Carmen Garcia','cgarcia@email.com','555-0118','Jose Garcia','','','','Bus Route 14',true,'08/25/2025',''],
  ['Sofia Rodriguez','Sofia','Grade 4','CLS-001','40019','srodriguez@maplewoodelem.edu','Lucia Rodriguez','lrodriguez@email.com','555-0119','Marco Rodriguez',true,'ELL support; bilingual resources','','Car pickup',true,'08/25/2025','ELL — Spanish home language'],
  ['Tyler White','Tyler','Grade 4','CLS-001','40020','twhite@maplewoodelem.edu','Anne White','awhite@email.com','555-0120','Frank White','','','','Walks home',true,'08/25/2025',''],
  ['Uma Patel','Uma','Grade 4','CLS-001','40021','upatel@maplewoodelem.edu','Deepa Patel','deepap@email.com','555-0121','Sanjay Patel','','','','Car pickup',true,'08/25/2025',''],
  ['Victor Lopez','Victor','Grade 4','CLS-001','40022','vlopez@maplewoodelem.edu','Isabel Lopez','ilopez@email.com','555-0122','Eduardo Lopez','','','','Bus Route 7',true,'08/25/2025',''],
  ['Wendy Chang','Wendy','Grade 4','CLS-001','40023','wchang@maplewoodelem.edu','Amy Chang','achang@email.com','555-0123','David Chang',true,'Gifted program participation','','Car pickup',true,'08/25/2025','Attends GT pull-out Thursdays'],
  ['Xavier Moore','Xavi','Grade 4','CLS-001','40024','xmoore@maplewoodelem.edu','Dana Moore','dmoore@email.com','555-0124','Rich Moore','','','','Bus Route 3',true,'08/25/2025',''],
  // Reading class (CLS-002) — same students, different class
  ['Aiden Torres','Aiden','Grade 4','CLS-002','40001','atorres@maplewoodelem.edu','Maria Torres','mtorres@email.com','555-0101','Carlos Torres','','','','Bus Route 14',true,'08/25/2025',''],
  ['Bella Nguyen','Bella','Grade 4','CLS-002','40002','bnguyen@maplewoodelem.edu','Linh Nguyen','lnguyen@email.com','555-0102','David Nguyen','','','Carries EpiPen — bee allergy','Car pickup',true,'08/25/2025',''],
  ['Carlos Rivera','Carlos','Grade 4','CLS-002','40003','crivera@maplewoodelem.edu','Ana Rivera','arivera@email.com','555-0103','Jorge Rivera','','','','Bus Route 7',true,'08/25/2025',''],
  ['Diana Patel','Diana','Grade 4','CLS-002','40004','dpatel@maplewoodelem.edu','Priya Patel','ppatel@email.com','555-0104','Raj Patel',true,'Extended time on tests','','Car pickup',true,'08/25/2025','504 Plan'],
  ['Ethan Kim','Ethan','Grade 4','CLS-002','40005','ekim@maplewoodelem.edu','Soo Kim','skim@email.com','555-0105','Jin Kim','','','','Bus Route 14',true,'08/25/2025',''],
  ['Fiona Chen','Fiona','Grade 4','CLS-002','40006','fchen@maplewoodelem.edu','Wei Chen','wchen@email.com','555-0106','Mei Chen','','','','Car pickup',true,'08/25/2025',''],
  ['Gabriel Santos','Gabe','Grade 4','CLS-002','40007','gsantos@maplewoodelem.edu','Rosa Santos','rsantos@email.com','555-0107','Luis Santos',true,'Small group testing; check-in support','','Bus Route 7',true,'08/25/2025','IEP'],
  ['Hannah Lee','Hannah','Grade 4','CLS-002','40008','hlee@maplewoodelem.edu','Jin Lee','jlee@email.com','555-0108','Sam Lee','','','','Walks home',true,'08/25/2025',''],
  ['Isaac Martinez','Isaac','Grade 4','CLS-002','40009','imartinez@maplewoodelem.edu','Gloria Martinez','gmartinez@email.com','555-0109','Hector Martinez','','','Asthma — inhaler in bag','Bus Route 3',true,'08/25/2025',''],
  ['Julia Brown','Julia','Grade 4','CLS-002','40010','jbrown@maplewoodelem.edu','Lisa Brown','lbrown@email.com','555-0110','Mike Brown','','','','Car pickup',true,'08/25/2025',''],
  // Math class (CLS-003)
  ['Kevin Wilson','Kevin','Grade 4','CLS-003','40011','kwilson@maplewoodelem.edu','Sarah Wilson','swilson@email.com','555-0111','Tom Wilson',true,'Extended time; preferential seating','','Bus Route 14',true,'08/25/2025','504 Plan'],
  ['Lily Davis','Lily','Grade 4','CLS-003','40012','ldavis@maplewoodelem.edu','Karen Davis','kdavis@email.com','555-0112','Bob Davis','','','','Car pickup',true,'08/25/2025',''],
  ['Marco Hernandez','Marco','Grade 4','CLS-003','40013','mhernandez@maplewoodelem.edu','Elena Hernandez','ehernandez@email.com','555-0113','Pedro Hernandez','','','','Bus Route 7',true,'08/25/2025',''],
  ['Nadia Johnson','Nadia','Grade 4','CLS-003','40014','njohnson@maplewoodelem.edu','Tracy Johnson','tjohnson@email.com','555-0114','Alan Johnson','','','','Walks home',true,'08/25/2025',''],
  ['Oliver Park','Ollie','Grade 4','CLS-003','40015','opark@maplewoodelem.edu','Mina Park','mpark@email.com','555-0115','Steve Park',true,'Check-in / check-out','','Car pickup',true,'08/25/2025',''],
  ['Penelope Adams','Penny','Grade 4','CLS-003','40016','padams@maplewoodelem.edu','Jane Adams','jadams@email.com','555-0116','Bill Adams','','','','Bus Route 3',true,'08/25/2025',''],
  ['Quinn Thomas','Quinn','Grade 4','CLS-003','40017','qthomas@maplewoodelem.edu','Dale Thomas','dthomas@email.com','555-0117','Gail Thomas','','','','Car pickup',true,'08/25/2025',''],
  ['Ryan Garcia','Ryan','Grade 4','CLS-003','40018','rgarcia@maplewoodelem.edu','Carmen Garcia','cgarcia@email.com','555-0118','Jose Garcia','','','','Bus Route 14',true,'08/25/2025',''],
  ['Sofia Rodriguez','Sofia','Grade 4','CLS-003','40019','srodriguez@maplewoodelem.edu','Lucia Rodriguez','lrodriguez@email.com','555-0119','Marco Rodriguez',true,'ELL support','','Car pickup',true,'08/25/2025','ELL'],
  ['Tyler White','Tyler','Grade 4','CLS-003','40020','twhite@maplewoodelem.edu','Anne White','awhite@email.com','555-0120','Frank White','','','','Walks home',true,'08/25/2025',''],
  // Science class (CLS-004)
  ['Uma Patel','Uma','Grade 4','CLS-004','40021','upatel@maplewoodelem.edu','Deepa Patel','deepap@email.com','555-0121','Sanjay Patel','','','','Car pickup',true,'08/25/2025',''],
  ['Victor Lopez','Victor','Grade 4','CLS-004','40022','vlopez@maplewoodelem.edu','Isabel Lopez','ilopez@email.com','555-0122','Eduardo Lopez','','','','Bus Route 7',true,'08/25/2025',''],
  ['Wendy Chang','Wendy','Grade 4','CLS-004','40023','wchang@maplewoodelem.edu','Amy Chang','achang@email.com','555-0123','David Chang',true,'GT program','','Car pickup',true,'08/25/2025',''],
  ['Xavier Moore','Xavi','Grade 4','CLS-004','40024','xmoore@maplewoodelem.edu','Dana Moore','dmoore@email.com','555-0124','Rich Moore','','','','Bus Route 3',true,'08/25/2025',''],
  ['Aiden Torres','Aiden','Grade 4','CLS-004','40001','atorres@maplewoodelem.edu','Maria Torres','mtorres@email.com','555-0101','Carlos Torres','','','','Bus Route 14',true,'08/25/2025',''],
  ['Bella Nguyen','Bella','Grade 4','CLS-004','40002','bnguyen@maplewoodelem.edu','Linh Nguyen','lnguyen@email.com','555-0102','David Nguyen','','','Carries EpiPen','Car pickup',true,'08/25/2025',''],
  ['Carlos Rivera','Carlos','Grade 4','CLS-004','40003','crivera@maplewoodelem.edu','Ana Rivera','arivera@email.com','555-0103','Jorge Rivera','','','','Bus Route 7',true,'08/25/2025',''],
  ['Diana Patel','Diana','Grade 4','CLS-004','40004','dpatel@maplewoodelem.edu','Priya Patel','ppatel@email.com','555-0104','Raj Patel',true,'Extended time','','Car pickup',true,'08/25/2025','504'],
  // Social Studies (CLS-005)
  ['Ethan Kim','Ethan','Grade 4','CLS-005','40005','ekim@maplewoodelem.edu','Soo Kim','skim@email.com','555-0105','Jin Kim','','','','Bus Route 14',true,'08/25/2025',''],
  ['Fiona Chen','Fiona','Grade 4','CLS-005','40006','fchen@maplewoodelem.edu','Wei Chen','wchen@email.com','555-0106','Mei Chen','','','','Car pickup',true,'08/25/2025',''],
  ['Gabriel Santos','Gabe','Grade 4','CLS-005','40007','gsantos@maplewoodelem.edu','Rosa Santos','rsantos@email.com','555-0107','Luis Santos',true,'Small group testing','','Bus Route 7',true,'08/25/2025','IEP'],
  ['Hannah Lee','Hannah','Grade 4','CLS-005','40008','hlee@maplewoodelem.edu','Jin Lee','jlee@email.com','555-0108','Sam Lee','','','','Walks home',true,'08/25/2025',''],
  ['Isaac Martinez','Isaac','Grade 4','CLS-005','40009','imartinez@maplewoodelem.edu','Gloria Martinez','gmartinez@email.com','555-0109','Hector Martinez','','','Asthma','Bus Route 3',true,'08/25/2025',''],
  ['Julia Brown','Julia','Grade 4','CLS-005','40010','jbrown@maplewoodelem.edu','Lisa Brown','lbrown@email.com','555-0110','Mike Brown','','','','Car pickup',true,'08/25/2025',''],
  ['Kevin Wilson','Kevin','Grade 4','CLS-005','40011','kwilson@maplewoodelem.edu','Sarah Wilson','swilson@email.com','555-0111','Tom Wilson',true,'Extended time','','Bus Route 14',true,'08/25/2025',''],
  ['Lily Davis','Lily','Grade 4','CLS-005','40012','ldavis@maplewoodelem.edu','Karen Davis','kdavis@email.com','555-0112','Bob Davis','','','','Car pickup',true,'08/25/2025',''],
  ['Marco Hernandez','Marco','Grade 4','CLS-005','40013','mhernandez@maplewoodelem.edu','Elena Hernandez','ehernandez@email.com','555-0113','Pedro Hernandez','','','','Bus Route 7',true,'08/25/2025',''],
  ['Nadia Johnson','Nadia','Grade 4','CLS-005','40014','njohnson@maplewoodelem.edu','Tracy Johnson','tjohnson@email.com','555-0114','Alan Johnson','','','','Walks home',true,'08/25/2025',''],
  ['Oliver Park','Ollie','Grade 4','CLS-005','40015','opark@maplewoodelem.edu','Mina Park','mpark@email.com','555-0115','Steve Park',true,'Behavior plan','','Car pickup',true,'08/25/2025',''],
];

(async () => {
  const vals = [];
  const fmt  = [];

  // BG
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1010,0,22), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});

  // Title
  vals.push({ range:`${S}!A1`, values:[['STUDENT ROSTER']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,20), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,20), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1}, properties:{pixelSize:44}, fields:'pixelSize' }});

  // Subtitle
  vals.push({ range:`${S}!A2`, values:[['Track your students, contacts, accommodations, and class assignments. Student ID auto-populates when a Student Name is entered.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,20), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,20), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  // Filter row
  vals.push({ range:`${S}!A3`, values:[['Filter by class or use the search function. Active = checkmark.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,20), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,20), cell:{userEnteredFormat:{
    backgroundColor:hex(C.ELA), textFormat:{fontSize:9,foregroundColor:hex(C.text),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:8},
  }}, fields:'userEnteredFormat' }});

  // Spacer row 4 — blank
  vals.push({ range:`${S}!A4`, values:[['']] });

  // Column headers row 5
  const hdrs = ['Student ID','Student Name','Preferred Name','Grade Level','Class ID','Class Name',
                'Student Number','Email','Parent / Guardian','Parent Email','Parent Phone',
                'Emergency Contact','Support / Accommodation?','Support Summary','Medical / Safety Note',
                'Transportation / Dismissal','Active?','Start Date','End Date','Notes'];
  vals.push({ range:`${S}!A5`, values:[hdrs] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,20), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5}, properties:{pixelSize:30}, fields:'pixelSize' }});

  // Data rows 6-1005: Student ID formula + data
  for (let i = 0; i < 1000; i++) {
    const r = i + 6;
    vals.push({ range:`${S}!A${r}`, values:[[`=IF(B${r}="","","STU-"&TEXT(ROW()-5,"0000"))`]] });
  }

  // Class Name lookup by Class ID
  for (let i = 0; i < 1000; i++) {
    const r = i + 6;
    vals.push({ range:`${S}!F${r}`, values:[[
      `=IFERROR(INDEX('School Year Setup'!B${29}:B${48},MATCH(E${r},'School Year Setup'!A${29}:A${48},0)),"")`
    ]]});
  }

  // Sample students
  STUDENTS.forEach((stu, i) => {
    const r = i + 6;
    // B:T columns (Student Name through Notes), skip A (formula) and F (lookup)
    const [name, pref, grade, clsId, stuNum, email, parent, parentEmail, phone, emg,
           support, supportSum, medical, transport, active, startDate, notes] = stu;
    vals.push({ range:`${S}!B${r}:E${r}`, values:[[name, pref, grade, clsId]] });
    vals.push({ range:`${S}!G${r}:L${r}`, values:[[stuNum, email, parent, parentEmail, phone, emg]] });
    vals.push({ range:`${S}!M${r}`, values:[[support]] });
    vals.push({ range:`${S}!N${r}:P${r}`, values:[[supportSum, medical, transport]] });
    vals.push({ range:`${S}!Q${r}`, values:[[active]] });
    vals.push({ range:`${S}!R${r}:T${r}`, values:[[startDate,'',notes]] });
  });

  // Checkboxes for Support (M) and Active (Q)
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1005,12,13), rule:{ condition:{type:'BOOLEAN'}, showCustomUi:true }}});
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1005,16,17), rule:{ condition:{type:'BOOLEAN'}, showCustomUi:true }}});

  // Grade Level dropdown
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1005,3,4), rule:{
    condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!N2:N16`}] }, showCustomUi:true, strict:false,
  }}});

  // Class ID dropdown
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1005,4,5), rule:{
    condition:{ type:'ONE_OF_LIST', values:[
      {userEnteredValue:'CLS-001'},{userEnteredValue:'CLS-002'},{userEnteredValue:'CLS-003'},
      {userEnteredValue:'CLS-004'},{userEnteredValue:'CLS-005'},{userEnteredValue:'CLS-006'},
      {userEnteredValue:'CLS-007'},{userEnteredValue:'CLS-008'},{userEnteredValue:'CLS-009'},{userEnteredValue:'CLS-010'},
    ]}, showCustomUi:true, strict:false,
  }}});

  // Date format col R (Start Date)
  fmt.push({ repeatCell:{ range: gridRange(SID,5,1005,17,18), cell:{userEnteredFormat:{
    numberFormat:{type:'DATE',pattern:'mmm d, yyyy'},
  }}, fields:'userEnteredFormat.numberFormat' }});

  // Alternating row colors
  for (let i = 0; i < 200; i++) {
    fmt.push({ repeatCell:{ range: gridRange(SID,5+i,6+i,0,20), cell:{userEnteredFormat:{
      backgroundColor: hex(i%2===0 ? C.panel : C.altRow),
    }}, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Formula cells (A, F) blue tint
  fmt.push({ repeatCell:{ range: gridRange(SID,5,1005,0,1), cell:{userEnteredFormat:{backgroundColor:hex(C.formula)}}, fields:'userEnteredFormat.backgroundColor' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,5,1005,5,6), cell:{userEnteredFormat:{backgroundColor:hex(C.formula)}}, fields:'userEnteredFormat.backgroundColor' }});

  // Support/Medical columns — neutral gray (never highlight with attention color)
  fmt.push({ repeatCell:{ range: gridRange(SID,5,1005,13,16), cell:{userEnteredFormat:{
    backgroundColor:hex('#F0EEE9'), textFormat:{fontSize:8,fontFamily:'Arial',italic:true,foregroundColor:hex(C.secText)},
  }}, fields:'userEnteredFormat' }});

  // Column widths
  const widths = [80,160,110,90,80,150,90,160,150,160,100,140,75,180,180,160,60,100,100,160];
  widths.forEach((w,i) => {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1}, properties:{pixelSize:w}, fields:'pixelSize' }});
  });

  // Row height for data rows
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:5,endIndex:1005}, properties:{pixelSize:21}, fields:'pixelSize' }});

  // Freeze rows 1:5 (no column freeze — merged title rows span all cols)
  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:5}}, fields:'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '04-roster values');
  await batchUpdate(id, fmt, '04-roster format');
  console.log('✅ Student Roster done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
