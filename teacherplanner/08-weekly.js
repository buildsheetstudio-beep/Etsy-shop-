'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Weekly Planner'];
const S = "'Weekly Planner'";

// 6 weeks of lesson plans, each week = 14 rows (1 header + 5 class rows + 1 notes row + spacer)
// Layout: rows per week block = 14 (header 2 rows, 5 days × 2 rows each, 2-row notes)
// Columns: A=Class/Label  B=Monday  C=Tuesday  D=Wednesday  E=Thursday  F=Friday
// G=Standards  H=Materials  I=Differentiation  J=Homework Assigned

// 6 weeks starting 09/01/2025
const WEEKS = [
  { label:'Week 1',  dates:['Sep 1','Sep 2','Sep 3','Sep 4','Sep 5'],  monday:'09/01/2025' },
  { label:'Week 2',  dates:['Sep 8','Sep 9','Sep 10','Sep 11','Sep 12'], monday:'09/08/2025' },
  { label:'Week 3',  dates:['Sep 15','Sep 16','Sep 17','Sep 18','Sep 19'], monday:'09/15/2025' },
  { label:'Week 4',  dates:['Sep 22','Sep 23','Sep 24','Sep 25','Sep 26'], monday:'09/22/2025' },
  { label:'Week 5',  dates:['Sep 29','Sep 30','Oct 1','Oct 2','Oct 3'],  monday:'09/29/2025' },
  { label:'Week 6',  dates:['Oct 6','Oct 7','Oct 8','Oct 9','Oct 10'],  monday:'10/06/2025' },
  { label:'Week 7',  dates:['Oct 13','Oct 14','Oct 15','Oct 16','Oct 17'], monday:'10/13/2025' },
  { label:'Week 8',  dates:['Oct 20','Oct 21','Oct 22','Oct 23','Oct 24'], monday:'10/20/2025' },
];

// Classes taught each day
const CLASSES = [
  { id:'CLS-001', name:'ELA',          color: C.ELA },
  { id:'CLS-002', name:'Reading',      color: C.ELA },
  { id:'CLS-003', name:'Mathematics',  color: C.Math },
  { id:'CLS-004', name:'Science',      color: C.Science },
  { id:'CLS-005', name:'Social Studies', color: C.SocialStudies },
];

// Lesson plans per class per week (abbreviated)
const PLANS = [
  // ELA CLS-001
  [
    ['Intro to Reading Workshop; establish routines','Phonics: short vowels review','Reading aloud: Charlotte\'s Web Ch 1-2','Character traits discussion','Vocab journals: introduce 10 words'],
    ['Reading log modeling; partner reading','Phonics: long vowels','Charlotte\'s Web Ch 3-5; comprehension Q','Graphic organizer: character map','Vocab quiz review + fluency practice'],
    ['Guided reading groups begin','Making inferences mini-lesson','Charlotte\'s Web Ch 6-8','Author\'s purpose discussion','Unit 1 vocab quiz'],
    ['Text structure: sequence + compare','RL 4.1 main idea practice','Reading comprehension test review','Introduce narrative writing','Personal narrative brainstorm'],
    ['Author\'s craft: word choice','Inference practice with picture books','Reading comp test — Story Elements','Introduction to figurative language','Simile + metaphor practice'],
    ['Poetry unit launch','Imagery in poetry','Analyze a poem together','Students write own poems','Poetry share + feedback'],
    ['Introduce informational text','Nonfiction features scavenger hunt','Read informational article','Main idea + details practice','Exit ticket: main idea'],
    ['Book report project introduced','Research strategies lesson','Draft book report intro paragraph','Peer review of introductions','Revise + expand paragraphs'],
  ],
  // Reading CLS-002
  [
    ['Establish reading groups; running records','Fluency partner reading','Guided reading: Group A','Reading strategy: visualizing','Reading response journals'],
    ['Guided reading: Group B','Strategy: making connections','Read-aloud + discussion','Fluency timed reads','Reading log check'],
    ['Guided reading: Group A','Comprehension: story structure','Guided reading: Group B','Nonfiction text intro','Quiz prep: comprehension'],
    ['Guided reading quiz','Re-teach struggling skills','Small group strategy work','Reading stamina building','Independent reading + log'],
    ['Benchmark assessment prep','Benchmark reading assessment','Score + analyze results','Differentiated small groups','Reading celebration!'],
    ['Introduce new reading unit','Close reading: informational','Guided reading rotations','Text-based evidence practice','Exit ticket: text evidence'],
    ['Fluency check: all students','Guided reading groups','Comprehension strategy focus','Partner reading + discussion','Weekly reading log due'],
    ['Summarizing nonfiction text','Guided reading: Group A+B','Main idea + supporting details','Re-teach: key details','End-of-unit review'],
  ],
  // Math CLS-003
  [
    ['Place value: intro + anchor chart','Expanded form practice','Compare + order numbers','Round to nearest 10, 100','Place value quiz prep'],
    ['Place value quiz','Multi-digit addition: review','Add with regrouping','Subtraction with regrouping','Multi-step word problems'],
    ['Multiplication facts: 0-5','Arrays + area model','Multiplication facts: 6-9','Properties of multiplication','Multiplication quiz'],
    ['Division as sharing','Division fact families','Long division intro','Division practice + word problems','Unit test review'],
    ['Division unit test','Introduce fractions','Equivalent fractions model','Compare fractions','Fractions on a number line'],
    ['Fractions review + practice','Add fractions — same denominator','Subtract fractions','Mixed numbers intro','Fractions quiz'],
    ['Fractions in real life project introduced','Geometry: points, lines, angles','Measuring angles with protractor','Types of triangles','Geometry practice'],
    ['Symmetry + geometry vocabulary','Geometry quiz prep','Geometry quiz','Review measurement units','Convert measurement units'],
  ],
  // Science CLS-004
  [
    ['Science safety + notebooks','Forms of energy: intro','Energy transformations activity','Lab: rolling objects + energy','Energy vocabulary cards'],
    ['Thermal energy + heat','Light energy demonstrations','Sound energy: making instruments','Lab: sound vibrations','Energy unit quiz review'],
    ['Energy quiz','Intro to waves','Properties of waves lab','Sound + light wave compare','Waves observation journal'],
    ['Waves in nature discussion','Electromagnetic spectrum intro','Light: reflection + refraction','Lab: prism + light bending','Unit test review: energy + waves'],
    ['Energy + Waves unit test','Life science intro','Plant parts + functions','Photosynthesis anchor chart','Plant life cycle model'],
    ['Plant observation journals begin','Animal life cycles compare','Food webs + food chains','Ecosystems vocabulary','Lab: build a food web'],
    ['Ecosystems project introduced','Research ecosystems (rainforest, ocean, desert)','Project work time','Present ecosystems posters','Project debrief'],
    ['Introduce earth science','Types of rocks lab','Rock cycle diagram','Weathering + erosion demo','Earth science vocab quiz'],
  ],
  // Social Studies CLS-005
  [
    ['Map skills: compass rose + legend','Political vs physical maps','Latitude + longitude intro','Practice: find places on map','Map quiz prep'],
    ['Map quiz','Community + local government','Roles in the community research','Community helpers presentations','Reflection on community'],
    ['American history: Native Americans','Regional Native American cultures','Primary source: Native American artifact','Compare 2 cultures — Venn diagram','Culture project intro'],
    ['Culture project work time','Present culture projects','Explorers: reasons for exploration','Columbus + other explorers','Explorer quiz prep'],
    ['Explorer quiz','Colonial America intro','Life in colonial times','Compare colonial + modern life','Colonial America writing'],
    ['Colonial period review','Causes of the American Revolution','Boston Tea Party drama activity','Declaration of Independence','History test review'],
    ['American history test','Civics: types of government','Three branches of government','Checks + balances lesson','Government quiz prep'],
    ['Government quiz','Economics intro: needs vs wants','Producers + consumers','Supply + demand activity','Economics exit ticket'],
  ],
];

// Standards per class per week
const STANDARDS = [
  ['CCSS.ELA-RF.4.4','CCSS.ELA-RF.4.4','CCSS.ELA-RL.4.3','CCSS.ELA-RL.4.3','CCSS.ELA-L.4.4','CCSS.ELA-RL.4.5','CCSS.ELA-RI.4.1','CCSS.ELA-W.4.2'],
  ['CCSS.ELA-RF.4.4','CCSS.ELA-RL.4.1','CCSS.ELA-RL.4.3','CCSS.ELA-RI.4.1','CCSS.ELA-RF.4.4','CCSS.ELA-RI.4.2','CCSS.ELA-RF.4.4','CCSS.ELA-RI.4.2'],
  ['CCSS.MATH.4.NBT.A.2','CCSS.MATH.4.NBT.B.4','CCSS.MATH.4.NBT.B.5','CCSS.MATH.4.NBT.B.6','CCSS.MATH.4.NF.A.1','CCSS.MATH.4.NF.B.3','CCSS.MATH.4.G.A.1','CCSS.MATH.4.MD'],
  ['NGSS.4-PS3','NGSS.4-PS3','NGSS.4-PS4-1','NGSS.4-PS3;PS4','NGSS.4-LS1-1','NGSS.4-LS1-2','NGSS.4-LS1-2','NGSS.4-ESS2-2'],
  ['C3.D2.Geo.1','C3.D2.Civ.1','C3.D2.His.5','C3.D2.His.4','C3.D2.His','C3.D2.His','C3.D2.Civ.4','C3.D2.Eco.1'],
];

// Homework assigned per class per week
const HW = [
  ['Reading log 20 min/night','Vocab practice p.4-5','Charlotte\'s Web read Ch 1-5','Personal narrative brainstorm','Poetry journal'],
  ['Read 20 min + log','Guided reading at home','Reading response journal','Fluency practice passage','None — project'],
  ['Place value worksheet','Multiplication facts x5','Division fact family sheet','Fractions worksheet','Geometry shapes hunt'],
  ['Science notebook reflection','Energy fact cards','Waves journal entry','Study for energy test','Plant observation log'],
  ['Map worksheet','Community helper interview','Culture project research','Explorer facts card','Colonial life journal'],
];

// Materials per class per week
const MATS = [
  ['Charlotte\'s Web novels','Vocab journals, phonics cards','Comprehension graphic organizers','Poetry anthology','Research books'],
  ['Leveled readers, running record forms','Fluency passages','Guided reading books','Comprehension passages','Benchmark test materials'],
  ['Base-ten blocks, number lines','Grid paper, manipulatives','Fraction tiles, number lines','Protractors, rulers','Measurement tools'],
  ['Safety goggles, lab notebooks','Energy lab materials','Wave tank demo kit','Plant seeds, cups, soil','Rock samples, hand lenses'],
  ['Atlas, globe, map worksheets','Community resource posters','Culture project materials','Explorer timeline','Declaration of Independence copies'],
];

// Differentiation notes per week (same for all classes for simplicity)
const DIFF = [
  'IEP: extended time; ELL: visual supports; GT: extension reading',
  'Pull-out group Tue/Thu; sentence frames for ELL; challenge vocab for GT',
  'Small group reteach Wed; GT: open-ended problems; modified HW for IEP',
  'Lab partners strategically grouped; visual vocab for ELL; GT: research extension',
  'Primary sources scaffolded for struggling readers; GT: compare/contrast writing',
  'Re-teach mini-lesson Mon; extra fluency support; GT: independent research',
  'Check-in/check-out for behavior support students; differentiated readings',
  'Pull small group for foundational skills; enrichment choice board for GT',
];

// Week notes
const WEEK_NOTES = [
  'Back to school! Establish community agreements. Back to school night: Thursday 9/4.',
  'First full week. Set up class jobs. Monitor settling in. Collect supply lists.',
  'Running records complete by Friday. Inform groupings for guided reading.',
  'Conference reports due Monday. Personal narrative kickoff.',
  'Mid-unit check: monitor fluency data. Adjust groups as needed.',
  'Parent conferences Mon-Wed after school. Reduce homework load this week.',
  'Progress reports window opens. Check gradebook is current. PD on Friday.',
  'Begin benchmark prep. Check student engagement. Upcoming field trip planning.',
];

(async () => {
  const vals = [];
  const fmt  = [];

  // BG
  fmt.push({ repeatCell:{ range: gridRange(SID,0,500,0,18), cell:{userEnteredFormat:{backgroundColor:hex(C.bg)}}, fields:'userEnteredFormat.backgroundColor' }});

  // Title row 1
  vals.push({ range:`${S}!A1`, values:[['WEEKLY PLANNER']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,0,1,0,10), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,0,1,0,10), cell:{userEnteredFormat:{
    backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:16,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:0,endIndex:1}, properties:{pixelSize:44}, fields:'pixelSize' }});

  // Subtitle row 2
  vals.push({ range:`${S}!A2`, values:[['Plan lessons for the week. Each class gets its own row with objectives for Mon–Fri. Use the Notes row for reminders, events, or follow-ups.']] });
  fmt.push({ mergeCells:{ range: gridRange(SID,1,2,0,10), mergeType:'MERGE_ALL' }});
  fmt.push({ repeatCell:{ range: gridRange(SID,1,2,0,10), cell:{userEnteredFormat:{
    backgroundColor:hex('#4A5056'), textFormat:{fontSize:9,foregroundColor:hex('#D5D8DB'),italic:true,fontFamily:'Arial'},
    horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
  }}, fields:'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:1,endIndex:2}, properties:{pixelSize:28}, fields:'pixelSize' }});

  // Column widths
  const colWidths = [110,150,150,150,150,150,130,130,150,130];
  colWidths.forEach((px, c) => {
    fmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'COLUMNS',startIndex:c,endIndex:c+1}, properties:{pixelSize:px}, fields:'pixelSize' }});
  });

  await batchUpdate(id, fmt, '08-weekly format');

  // Each week block: rows (0-indexed from block start)
  // Row 0: Week label + dates header
  // Row 1: Column headers (Class | Mon | Tue | Wed | Thu | Fri | Standards | Materials | Differentiation | Homework)
  // Rows 2-6: 5 class rows
  // Row 7: Teacher Notes
  // Row 8: spacer
  // => 9 rows per week block, starting at row 3 (0-indexed) in the sheet

  const weekFmt = [];
  const weekVals = [];

  const startRow = 3; // 0-indexed sheet row where week 1 block starts

  WEEKS.forEach((week, wi) => {
    const blockStart = startRow + wi * 9; // sheet row (0-indexed)

    // Row 0 of block: week label + dates
    const headerRow = blockStart;
    weekVals.push({ range:`${S}!A${headerRow+1}`, values:[[
      `${week.label} — ${week.dates[0]}–${week.dates[4]}`,
      week.dates[0], week.dates[1], week.dates[2], week.dates[3], week.dates[4],
      'Standards','Materials / Resources','Differentiation','Homework Assigned'
    ]]});
    weekFmt.push({ mergeCells:{ range: gridRange(SID,headerRow,headerRow+1,0,1), mergeType:'MERGE_ALL' }});
    weekFmt.push({ repeatCell:{ range: gridRange(SID,headerRow,headerRow+1,0,10), cell:{userEnteredFormat:{
      backgroundColor:hex('#2F3437'), textFormat:{bold:true,fontSize:10,foregroundColor:hex('#FFFFFF'),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    weekFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:headerRow,endIndex:headerRow+1}, properties:{pixelSize:26}, fields:'pixelSize' }});

    // Row 1 of block: sub-headers (Class column labels)
    const subHdrRow = headerRow + 1;
    weekVals.push({ range:`${S}!A${subHdrRow+1}`, values:[['Class','Objective / Activity','Objective / Activity','Objective / Activity','Objective / Activity','Objective / Activity','Standards','Materials','Differentiation','Homework']]});
    weekFmt.push({ repeatCell:{ range: gridRange(SID,subHdrRow,subHdrRow+1,0,10), cell:{userEnteredFormat:{
      backgroundColor:hex(C.altRow), textFormat:{bold:true,fontSize:8,foregroundColor:hex(C.secText),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      borders:{ bottom:{style:'SOLID',color:hex(C.border)} },
    }}, fields:'userEnteredFormat' }});
    weekFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:subHdrRow,endIndex:subHdrRow+1}, properties:{pixelSize:20}, fields:'pixelSize' }});

    // Rows 2-6: 5 class rows
    CLASSES.forEach((cls, ci) => {
      const classRow = headerRow + 2 + ci;
      const dayPlans = PLANS[ci][wi] || Array(5).fill('');
      const std = STANDARDS[ci][wi] || '';
      const mat = MATS[ci][wi] || '';
      const diff = DIFF[wi] || '';
      const hw = HW[ci][wi] || '';
      weekVals.push({ range:`${S}!A${classRow+1}`, values:[[
        cls.name, ...dayPlans, std, mat, diff, hw
      ]]});
      weekFmt.push({ repeatCell:{ range: gridRange(SID,classRow,classRow+1,0,1), cell:{userEnteredFormat:{
        backgroundColor:hex(cls.color), textFormat:{bold:true,fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
      }}, fields:'userEnteredFormat' }});
      weekFmt.push({ repeatCell:{ range: gridRange(SID,classRow,classRow+1,1,10), cell:{userEnteredFormat:{
        backgroundColor:hex(C.panel), textFormat:{fontSize:8,foregroundColor:hex(C.text),fontFamily:'Arial'},
        verticalAlignment:'TOP', wrapStrategy:'WRAP',
        borders:{ bottom:{style:'SOLID',color:hex(C.border)}, right:{style:'SOLID',color:hex(C.border)} },
      }}, fields:'userEnteredFormat' }});
      weekFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:classRow,endIndex:classRow+1}, properties:{pixelSize:58}, fields:'pixelSize' }});
    });

    // Row 7: Notes row
    const notesRow = headerRow + 7;
    weekVals.push({ range:`${S}!A${notesRow+1}`, values:[['Notes',WEEK_NOTES[wi],'','','','','','','','']]});
    weekFmt.push({ mergeCells:{ range: gridRange(SID,notesRow,notesRow+1,1,10), mergeType:'MERGE_ALL' }});
    weekFmt.push({ repeatCell:{ range: gridRange(SID,notesRow,notesRow+1,0,1), cell:{userEnteredFormat:{
      backgroundColor:hex(C.warning), textFormat:{bold:true,fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    weekFmt.push({ repeatCell:{ range: gridRange(SID,notesRow,notesRow+1,1,10), cell:{userEnteredFormat:{
      backgroundColor:hex('#FFFBF0'), textFormat:{fontSize:9,foregroundColor:hex(C.text),fontFamily:'Arial'},
      verticalAlignment:'MIDDLE',
    }}, fields:'userEnteredFormat' }});
    weekFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:notesRow,endIndex:notesRow+1}, properties:{pixelSize:22}, fields:'pixelSize' }});

    // Row 8: spacer
    const spacerRow = headerRow + 8;
    weekFmt.push({ updateDimensionProperties:{ range:{sheetId:SID,dimension:'ROWS',startIndex:spacerRow,endIndex:spacerRow+1}, properties:{pixelSize:10}, fields:'pixelSize' }});
    weekFmt.push({ repeatCell:{ range: gridRange(SID,spacerRow,spacerRow+1,0,10), cell:{userEnteredFormat:{
      backgroundColor:hex(C.bg),
    }}, fields:'userEnteredFormat.backgroundColor' }});
  });

  await valuesBatchUpdate(id, weekVals, '08-weekly values');
  await batchUpdate(id, weekFmt, '08-weekly week fmt');

  // Freeze row 1 only (title — no frozen columns since merges span all)
  await batchUpdate(id, [{ updateSheetProperties:{ properties:{ sheetId:SID, gridProperties:{ frozenRowCount:2 }}, fields:'gridProperties.frozenRowCount' }}], '08-weekly freeze');

  console.log('✅ Weekly Planner done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
