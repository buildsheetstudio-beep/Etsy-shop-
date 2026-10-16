'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Assignments & Assessments'];
const S = "'Assignments & Assessments'";
const REF = "'Reference Data'";

// 85 sample assignments across all classes, types, statuses
const ASSIGNMENTS = [
  // [ClassID, AssignName, Type, Category, Topic, Standards, DatePlanned, DateAssigned, DueDate, MaxPts, ScoreType, Weight, Status, Graded, Returned, Notes]
  ['CLS-001','Daily Reading Log Wk 1','Homework','Homework','Reading Fluency','CCSS.ELA-RF.4.4','09/01/2025','09/02/2025','09/05/2025',10,'Points',1,'Returned',true,true,''],
  ['CLS-001','Character Analysis — Charlotte\'s Web','Classwork','Classwork','Story Elements','CCSS.ELA-RL.4.3','09/08/2025','09/08/2025','09/10/2025',20,'Points',1,'Returned',true,true,''],
  ['CLS-001','Vocabulary Quiz — Unit 1','Quiz','Quizzes','Vocabulary','CCSS.ELA-L.4.4','09/15/2025','09/12/2025','09/15/2025',25,'Points',1,'Returned',true,true,''],
  ['CLS-001','Reading Comprehension Test 1','Test','Tests','Main Idea & Details','CCSS.ELA-RI.4.2','10/03/2025','09/29/2025','10/03/2025',50,'Points',1,'Returned',true,true,''],
  ['CLS-001','Personal Narrative Draft','Essay','Essays','Narrative Writing','CCSS.ELA-W.4.3','10/10/2025','09/22/2025','10/10/2025',40,'Points',1,'Returned',true,false,'Drafts returned; final due 10/20'],
  ['CLS-001','Book Report — Fiction','Project','Projects','Literary Analysis','CCSS.ELA-RL.4.9','10/24/2025','09/29/2025','10/24/2025',60,'Points',1,'Returned',true,true,''],
  ['CLS-001','Reading Conference Participation','Participation','Participation','Reading Engagement','CCSS.ELA-SL.4.1','11/01/2025','','11/01/2025',10,'Points',1,'Returned',true,true,''],
  ['CLS-001','Vocabulary Quiz — Unit 2','Quiz','Quizzes','Context Clues','CCSS.ELA-L.4.4','11/10/2025','11/06/2025','11/10/2025',25,'Points',1,'Returned',true,true,''],
  ['CLS-001','Reading Response Journal Wk 10','Homework','Homework','Inferencing','CCSS.ELA-RL.4.1','11/14/2025','11/10/2025','11/14/2025',10,'Points',1,'Returned',true,true,''],
  ['CLS-001','Mid-Year ELA Assessment','Test','Tests','Comprehensive','CCSS.ELA.4','01/16/2026','01/14/2026','01/16/2026',100,'Points',1,'Returned',true,true,''],
  ['CLS-001','Poetry Unit Exit Ticket','Exit Ticket','Classwork','Poetry Elements','CCSS.ELA-RL.4.5','02/06/2026','02/06/2026','02/06/2026',5,'Points',1,'Returned',true,true,''],
  ['CLS-001','Research Project — Animal Habitats','Project','Projects','Informational Writing','CCSS.ELA-W.4.2','03/06/2026','02/09/2026','03/06/2026',80,'Rubric',1,'Grading',true,false,'Presentations 03/09–03/11'],
  ['CLS-001','Spring Reading Assessment','Test','Tests','Reading Levels','CCSS.ELA-RF.4.4','04/20/2026','04/16/2026','04/20/2026',100,'Points',1,'Planned',false,false,'Benchmark assessment'],
  ['CLS-001','Author Study Presentation','Presentation','Projects','Author Craft','CCSS.ELA-RL.4.6','05/08/2026','04/27/2026','05/08/2026',50,'Rubric',1,'Assigned',false,false,'Pair work'],
  ['CLS-001','Year-End Vocabulary Practice','Practice','Classwork','Vocabulary Review','CCSS.ELA-L.4.6','05/22/2026','','05/22/2026',20,'Points',1,'Planned',false,false,''],
  // CLS-002 Reading
  ['CLS-002','Fluency Check — Passage 1','Practice','Classwork','Reading Fluency','CCSS.ELA-RF.4.4','09/03/2025','09/03/2025','09/03/2025',20,'Points',1,'Returned',true,true,''],
  ['CLS-002','Story Mapping Activity','Classwork','Classwork','Story Structure','CCSS.ELA-RL.4.3','09/10/2025','09/10/2025','09/12/2025',15,'Points',1,'Returned',true,true,''],
  ['CLS-002','Guided Reading Quiz G4','Quiz','Quizzes','Comprehension','CCSS.ELA-RI.4.1','09/22/2025','09/18/2025','09/22/2025',20,'Points',1,'Returned',true,true,''],
  ['CLS-002','Independent Reading Log','Homework','Homework','Self-Selected Reading','CCSS.ELA-RF.4.4','10/01/2025','09/29/2025','10/01/2025',10,'Complete / Incomplete',1,'Returned',true,true,''],
  ['CLS-002','Reading Strategy Exit Ticket','Exit Ticket','Classwork','Summarizing','CCSS.ELA-RI.4.2','10/15/2025','10/15/2025','10/15/2025',5,'Points',1,'Returned',true,true,''],
  ['CLS-002','Nonfiction Text Feature Quiz','Quiz','Quizzes','Text Features','CCSS.ELA-RI.4.7','11/05/2025','11/03/2025','11/05/2025',20,'Points',1,'Returned',true,true,''],
  ['CLS-002','Reading Assessment — Benchmark 1','Test','Tests','Reading Level','CCSS.ELA-RF.4.4','11/21/2025','11/17/2025','11/21/2025',50,'Points',1,'Returned',true,true,''],
  ['CLS-002','Winter Reading Project','Project','Projects','Genre Study','CCSS.ELA-RL.4.9','01/23/2026','12/15/2025','01/23/2026',60,'Rubric',1,'Returned',true,true,''],
  ['CLS-002','Comprehension Strategy Practice','Practice','Classwork','Making Connections','CCSS.ELA-RL.4.1','02/20/2026','02/18/2026','02/20/2026',15,'Points',1,'Collected',true,false,'Grading in progress'],
  ['CLS-002','Spring Benchmark Reading Test','Test','Tests','Reading Levels','CCSS.ELA-RF.4.4','04/17/2026','04/13/2026','04/17/2026',50,'Points',1,'Assigned',false,false,''],
  // CLS-003 Math
  ['CLS-003','Place Value Worksheet','Homework','Homework','Place Value','CCSS.MATH.4.NBT.A.2','09/02/2025','09/01/2025','09/02/2025',20,'Points',1,'Returned',true,true,''],
  ['CLS-003','Addition & Subtraction Classwork','Classwork','Classwork','Operations','CCSS.MATH.4.NBT.B.4','09/08/2025','09/08/2025','09/08/2025',15,'Points',1,'Returned',true,true,''],
  ['CLS-003','Multiplication Facts Quiz','Quiz','Quizzes','Multiplication','CCSS.MATH.4.NBT.B.5','09/15/2025','09/12/2025','09/15/2025',25,'Points',1,'Returned',true,true,''],
  ['CLS-003','Unit 1 Math Test — Operations','Test','Tests','Addition through Multiplication','CCSS.MATH.4.NBT','10/06/2025','10/01/2025','10/06/2025',50,'Points',1,'Returned',true,true,''],
  ['CLS-003','Division Practice Sheet','Homework','Homework','Division','CCSS.MATH.4.NBT.B.6','10/13/2025','10/10/2025','10/13/2025',20,'Points',1,'Returned',true,true,''],
  ['CLS-003','Fractions Exit Ticket','Exit Ticket','Classwork','Equivalent Fractions','CCSS.MATH.4.NF.A.1','11/03/2025','11/03/2025','11/03/2025',10,'Points',1,'Returned',true,true,''],
  ['CLS-003','Fractions Unit Quiz','Quiz','Quizzes','Fractions','CCSS.MATH.4.NF','11/17/2025','11/12/2025','11/17/2025',30,'Points',1,'Returned',true,true,''],
  ['CLS-003','Math Project — Fractions in Real Life','Project','Projects','Applied Fractions','CCSS.MATH.4.NF.B.3','12/05/2025','11/10/2025','12/05/2025',50,'Rubric',1,'Returned',true,true,''],
  ['CLS-003','Mid-Year Math Assessment','Test','Tests','All Domains','CCSS.MATH.4','01/16/2026','01/14/2026','01/16/2026',100,'Points',1,'Returned',true,true,''],
  ['CLS-003','Geometry Homework — Angles','Homework','Homework','Geometry','CCSS.MATH.4.G.A.1','02/09/2026','02/06/2026','02/09/2026',20,'Points',1,'Returned',true,true,''],
  ['CLS-003','Geometry Quiz','Quiz','Quizzes','Angles & Lines','CCSS.MATH.4.G.A.2','02/23/2026','02/19/2026','02/23/2026',25,'Points',1,'Grading',true,false,''],
  ['CLS-003','State Math Assessment Prep','Practice','Classwork','Test Prep','CCSS.MATH.4','03/02/2026','02/27/2026','03/02/2026',20,'Points',1,'Returned',true,true,''],
  ['CLS-003','Measurement Unit Test','Test','Tests','Measurement','CCSS.MATH.4.MD','03/27/2026','03/23/2026','03/27/2026',50,'Points',1,'Returned',true,true,''],
  ['CLS-003','Data & Graphs Practice','Practice','Classwork','Data Analysis','CCSS.MATH.4.MD.B.4','04/13/2026','04/09/2026','04/13/2026',15,'Points',1,'Assigned',false,false,''],
  ['CLS-003','Spring Math Benchmark','Test','Tests','Benchmark','CCSS.MATH.4','04/20/2026','04/16/2026','04/20/2026',100,'Points',1,'Planned',false,false,''],
  ['CLS-003','End-of-Year Math Review','Classwork','Classwork','Cumulative Review','CCSS.MATH.4','05/18/2026','','05/18/2026',20,'Points',1,'Planned',false,false,''],
  // CLS-004 Science
  ['CLS-004','Science Notebook — Lab Rules','Classwork','Classwork','Safety','NGSS.4-PS3','09/03/2025','09/03/2025','09/03/2025',10,'Complete / Incomplete',1,'Returned',true,true,''],
  ['CLS-004','Energy Forms Lab','Lab','Labs','Forms of Energy','NGSS.4-PS3-1','09/17/2025','09/15/2025','09/17/2025',30,'Rubric',1,'Returned',true,true,''],
  ['CLS-004','Energy Unit Quiz','Quiz','Quizzes','Energy','NGSS.4-PS3','09/29/2025','09/25/2025','09/29/2025',25,'Points',1,'Returned',true,true,''],
  ['CLS-004','Waves & Sound Lab','Lab','Labs','Sound & Light','NGSS.4-PS4-1','10/22/2025','10/20/2025','10/22/2025',30,'Rubric',1,'Returned',true,true,''],
  ['CLS-004','Science Test — Energy & Waves','Test','Tests','Energy & Waves','NGSS.4-PS3; 4-PS4','11/10/2025','11/05/2025','11/10/2025',60,'Points',1,'Returned',true,true,''],
  ['CLS-004','Plant Life Cycle Observation Journal','Homework','Homework','Life Science','NGSS.4-LS1-1','12/01/2025','11/17/2025','12/01/2025',20,'Points',1,'Returned',true,true,''],
  ['CLS-004','Ecosystems Project','Project','Projects','Ecosystems','NGSS.4-LS1-2','01/30/2026','01/05/2026','01/30/2026',60,'Rubric',1,'Returned',true,true,''],
  ['CLS-004','Earth Features Lab','Lab','Labs','Rocks & Landforms','NGSS.4-ESS2-2','02/25/2026','02/23/2026','02/25/2026',30,'Rubric',1,'Grading',true,false,''],
  ['CLS-004','Earth Science Quiz','Quiz','Quizzes','Earth Features','NGSS.4-ESS2','03/09/2026','03/05/2026','03/09/2026',25,'Points',1,'Assigned',false,false,''],
  ['CLS-004','Science Fair Project','Project','Projects','Student Choice','NGSS.4','04/24/2026','03/02/2026','04/24/2026',100,'Rubric',1,'Planned',false,false,'Major project'],
  // CLS-005 Social Studies
  ['CLS-005','Map Skills Worksheet','Classwork','Classwork','Geography','C3.D2.Geo.1','09/05/2025','09/05/2025','09/05/2025',15,'Points',1,'Returned',true,true,''],
  ['CLS-005','Community Helpers Research','Homework','Homework','Community','C3.D2.His.1','09/19/2025','09/15/2025','09/19/2025',20,'Points',1,'Returned',true,true,''],
  ['CLS-005','American History Quiz — Explorers','Quiz','Quizzes','Exploration','C3.D2.His.4','10/06/2025','10/02/2025','10/06/2025',25,'Points',1,'Returned',true,true,''],
  ['CLS-005','Native American Cultures Project','Project','Projects','Indigenous History','C3.D2.His.5','11/14/2025','10/06/2025','11/14/2025',50,'Rubric',1,'Returned',true,true,''],
  ['CLS-005','Social Studies Test — Early America','Test','Tests','Colonial America','C3.D2.His','12/08/2025','12/01/2025','12/08/2025',50,'Points',1,'Returned',true,true,''],
  ['CLS-005','Current Events Report','Homework','Homework','Civics','C3.D2.Civ.1','01/12/2026','01/09/2026','01/12/2026',15,'Points',1,'Returned',true,true,''],
  ['CLS-005','Government Structures Quiz','Quiz','Quizzes','Civics','C3.D2.Civ.4','02/02/2026','01/29/2026','02/02/2026',25,'Points',1,'Returned',true,true,''],
  ['CLS-005','Economics Exit Ticket','Exit Ticket','Classwork','Economics','C3.D2.Eco.1','02/23/2026','02/23/2026','02/23/2026',5,'Points',1,'Collected',true,false,''],
  ['CLS-005','State History Research Project','Project','Projects','State Studies','C3.D2.Geo.4','04/17/2026','03/09/2026','04/17/2026',80,'Rubric',1,'Planned',false,false,''],
  ['CLS-005','Year-End Social Studies Review','Classwork','Classwork','Cumulative','C3','05/18/2026','','05/18/2026',20,'Points',1,'Planned',false,false,''],
  // CLS-006 Writing
  ['CLS-006','Sentence Writing Practice','Practice','Classwork','Grammar','CCSS.ELA-W.4','09/04/2025','09/04/2025','09/04/2025',10,'Points',1,'Returned',true,true,''],
  ['CLS-006','Personal Narrative Brainstorm','Classwork','Classwork','Narrative Writing','CCSS.ELA-W.4.3','09/10/2025','09/10/2025','09/12/2025',10,'Complete / Incomplete',1,'Returned',true,true,''],
  ['CLS-006','Personal Narrative Final Draft','Essay','Essays','Narrative Writing','CCSS.ELA-W.4.3','10/03/2025','09/10/2025','10/03/2025',50,'Rubric',1,'Returned',true,true,''],
  ['CLS-006','Informational Writing — How-To','Essay','Essays','Informational Writing','CCSS.ELA-W.4.2','11/07/2025','10/13/2025','11/07/2025',50,'Rubric',1,'Returned',true,true,''],
  ['CLS-006','Persuasive Essay Draft','Essay','Essays','Opinion Writing','CCSS.ELA-W.4.1','12/05/2025','11/10/2025','12/05/2025',50,'Rubric',1,'Returned',true,true,''],
  ['CLS-006','Grammar Quiz — Punctuation','Quiz','Quizzes','Mechanics','CCSS.ELA-L.4.2','10/20/2025','10/16/2025','10/20/2025',20,'Points',1,'Returned',true,true,''],
  ['CLS-006','Research Report — Final','Essay','Essays','Research Writing','CCSS.ELA-W.4.7','03/20/2026','02/02/2026','03/20/2026',80,'Rubric',1,'Grading',true,false,'Long-term project'],
  ['CLS-006','Poetry Writing Unit','Project','Projects','Creative Writing','CCSS.ELA-W.4.3','04/10/2026','03/30/2026','04/10/2026',40,'Rubric',1,'Assigned',false,false,''],
  ['CLS-006','End-of-Year Writing Portfolio','Project','Projects','Portfolio','CCSS.ELA-W.4','05/29/2026','04/27/2026','05/29/2026',100,'Rubric',1,'Planned',false,false,''],
  // Edge cases
  ['CLS-003','Cancelled Extra Credit Quiz','Quiz','Quizzes','Review','CCSS.MATH.4','11/24/2025','','11/24/2025',10,'Points',1,'Cancelled',false,false,'Cancelled — holiday week'],
  ['CLS-001','Missing Max Points Assignment','Homework','Homework','Practice','','02/16/2026','02/13/2026','02/16/2026',0,'Points',1,'Returned',true,true,'Max pts not set'],
  ['CLS-004','Future Lab — Earth Sci II','Lab','Labs','Weathering','NGSS.4-ESS2','05/04/2026','','05/04/2026',30,'Rubric',1,'Planned',false,false,'Needs supplies order'],
  ['CLS-001','Overdue Reading Quiz','Quiz','Quizzes','Comprehension','CCSS.ELA-RL.4','11/03/2025','10/30/2025','11/03/2025',20,'Points',1,'Returned',true,false,'Returned but not graded — data error example'],
];

(async () => {
  const vals = [];
  const fmt  = [];

  // BG
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1510,0,22), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});

  // Title
  vals.push({ range:`${S}!A1`, values:[['ASSIGNMENTS & ASSESSMENTS']] });
  vals.push({ range:`${S}!A2`, values:[['Track every assignment and assessment. Class Name auto-populates from Class ID. Missing Count and Average Score pull from the Gradebook tab.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,22), mergeType:'MERGE_ALL' }});
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,22), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,22), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,22), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});

  // Alert row
  vals.push({ range:`${S}!A3`, values:[['⚑ Flags: Due before assigned | Returned but not graded | Due in 7 days and Planned | Graded with missing scores | Missing Max Points']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,2,3,0,22), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,2,3,0,22), cell:{userEnteredFormat:{
    backgroundColor:hex(C.warning), textFormat:{fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
    horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE', padding:{left:8},
  }}, fields:'userEnteredFormat' }});

  vals.push({ range:`${S}!A4`, values:[['']] });
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:2}, properties:{pixelSize:36}, fields:'pixelSize' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:2,endIndex:3}, properties:{pixelSize:24}, fields:'pixelSize' }});

  // Column headers row 5
  const hdrs = ['Assignment ID','Class ID','Class Name','Assignment / Assessment','Assignment Type',
                'Grade Category','Subject / Topic','Standards','Date Planned','Date Assigned',
                'Due Date','Max Points','Score Type','Weight','Status','Graded?','Returned?',
                'Missing Count','Average Score','Notes'];
  vals.push({ range:`${S}!A5`, values:[hdrs] });
  fmt.push({ repeatCell:{ range: gridRange(SID,4,5,0,20), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:9,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:4,endIndex:5}, properties:{pixelSize:30}, fields:'pixelSize' }});

  // Assignment ID formula (rows 6-1505)
  for (let i = 0; i < 1500; i++) {
    const r = i + 6;
    vals.push({ range:`${S}!A${r}`, values:[[`=IF(D${r}="","","ASN-"&TEXT(ROW()-5,"0000"))`]] });
  }

  // Class Name lookup
  for (let i = 0; i < 1500; i++) {
    const r = i + 6;
    vals.push({ range:`${S}!C${r}`, values:[[
      `=IFERROR(INDEX('School Year Setup'!B29:B48,MATCH(B${r},'School Year Setup'!A29:A48,0)),"")`
    ]]});
  }

  // Missing Count from Gradebook
  for (let i = 0; i < 1500; i++) {
    const r = i + 6;
    vals.push({ range:`${S}!R${r}`, values:[[
      `=IFERROR(COUNTIFS('Gradebook'!F6:F5005,A${r},'Gradebook'!N6:N5005,TRUE),"")`
    ]]});
  }

  // Average Score from Gradebook
  for (let i = 0; i < 1500; i++) {
    const r = i + 6;
    vals.push({ range:`${S}!S${r}`, values:[[
      `=IFERROR(AVERAGEIFS('Gradebook'!L6:L5005,'Gradebook'!F6:F5005,A${r},'Gradebook'!Q6:Q5005,TRUE),"")`
    ]]});
  }

  // Sample data
  ASSIGNMENTS.forEach((asn, i) => {
    const r = i + 6;
    const [clsId, name, type, cat, topic, stds, datePlan, dateAssign, due, maxPts, scoreType, wt, status, graded, returned, notes] = asn;
    vals.push({ range:`${S}!B${r}`, values:[[clsId]] });
    vals.push({ range:`${S}!D${r}:N${r}`, values:[[name, type, cat, topic, stds, datePlan, dateAssign, due, maxPts, scoreType, wt]] });
    vals.push({ range:`${S}!O${r}:Q${r}`, values:[[status, graded, returned]] });
    vals.push({ range:`${S}!T${r}`, values:[[notes]] });
  });

  // Dropdowns
  ['E','F','M','O'].forEach((col, ci) => {
    const refRanges = [
      `=${REF}!E2:E13`, `=${REF}!J2:J10`, `=${REF}!K2:K6`, `=${REF}!F2:F8`
    ];
    fmt.push({ setDataValidation:{ range: gridRange(SID,5,1505,ci+4===4?4:ci+4===5?5:ci+4===12?12:14, ci+4===4?5:ci+4===5?6:ci+4===12?13:15), rule:{
      condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue: refRanges[ci]}] }, showCustomUi:true, strict:false,
    }}});
  });
  // Status dropdown explicitly
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1505,14,15), rule:{
    condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!F2:F8`}] }, showCustomUi:true, strict:false,
  }}});
  // Assignment Type
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1505,4,5), rule:{
    condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!E2:E13`}] }, showCustomUi:true, strict:false,
  }}});
  // Grade Category
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1505,5,6), rule:{
    condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!J2:J10`}] }, showCustomUi:true, strict:false,
  }}});
  // Score Type
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1505,12,13), rule:{
    condition:{ type:'ONE_OF_RANGE', values:[{userEnteredValue:`=${REF}!K2:K6`}] }, showCustomUi:true, strict:false,
  }}});
  // Graded / Returned checkboxes
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1505,15,16), rule:{ condition:{type:'BOOLEAN'}, showCustomUi:true }}});
  fmt.push({ setDataValidation:{ range: gridRange(SID,5,1505,16,17), rule:{ condition:{type:'BOOLEAN'}, showCustomUi:true }}});

  // Date formats cols I, J, K (8, 9, 10)
  [8,9,10].forEach(c => {
    fmt.push({ repeatCell:{ range: gridRange(SID,5,1505,c,c+1), cell:{userEnteredFormat:{
      numberFormat:{type:'DATE',pattern:'mmm d, yyyy'},
    }}, fields:'userEnteredFormat.numberFormat' }});
  });
  // Average Score — percentage
  fmt.push({ repeatCell:{ range: gridRange(SID,5,1505,18,19), cell:{userEnteredFormat:{
    numberFormat:{type:'PERCENT',pattern:'0%'},
  }}, fields:'userEnteredFormat.numberFormat' }});

  // Conditional formatting — overdue = muted attention
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,1505,0,20)],
    booleanRule:{
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=AND(O6<>"Cancelled",O6<>"Returned",K6<TODAY(),K6<>"")`}] },
      format:{ backgroundColor: hex('#F0D9D7') },
    },
  }, index:0 }});

  // Conditional — graded + missing scores (orange)
  fmt.push({ addConditionalFormatRule:{ rule:{
    ranges:[gridRange(SID,5,1505,0,20)],
    booleanRule:{
      condition:{ type:'CUSTOM_FORMULA', values:[{userEnteredValue:`=AND(P6=TRUE,R6>0)`}] },
      format:{ backgroundColor: hex('#FBF0D0') },
    },
  }, index:1 }});

  // Alternating rows
  for (let i = 0; i < 100; i++) {
    fmt.push({ repeatCell:{ range: gridRange(SID,5+i,6+i,0,20), cell:{userEnteredFormat:{
      backgroundColor: hex(i%2===0 ? C.panel : C.altRow),
    }}, fields:'userEnteredFormat.backgroundColor' }});
  }

  // Formula cells blue tint (A, C, R, S)
  [0,2,17,18].forEach(c => {
    fmt.push({ repeatCell:{ range: gridRange(SID,5,1505,c,c+1), cell:{userEnteredFormat:{backgroundColor:hex(C.formula)}}, fields:'userEnteredFormat.backgroundColor' }});
  });

  // Col widths
  const widths = [80,80,140,240,100,100,140,160,100,100,100,75,110,65,100,65,70,90,90,180];
  widths.forEach((w,i) => {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:i,endIndex:i+1}, properties:{pixelSize:w}, fields:'pixelSize' }});
  });

  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:5,endIndex:1505}, properties:{pixelSize:21}, fields:'pixelSize' }});
  fmt.push({ updateSheetProperties:{ properties:{sheetId:SID,gridProperties:{frozenRowCount:5}}, fields:'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '05-assignments values');
  await batchUpdate(id, fmt, '05-assignments format');
  console.log('✅ Assignments & Assessments done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
