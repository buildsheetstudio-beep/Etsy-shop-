'use strict';
const { batchUpdate, valuesBatchUpdate, hex, COLOR, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Assignments']; // 1

// 30 sample assignments: col order A=Row#, B=Title, C=Course, D=Type, E=Due Date, F=Priority, G=Credits/Weight, H=Status, I=Submission Link, J=Notes
const ASSIGNMENTS = [
  // Oct
  ['1','Reading Response 1','English 101','Reading','10/3/2025','Medium','5','Complete','','Chapter 1-3 summary'],
  ['2','Problem Set 1','Calculus II','Problem Set','10/3/2025','High','10','Complete','','Derivatives'],
  ['3','History Essay Outline','World History','Essay','10/5/2025','High','10','Complete','','WWI causes outline'],
  ['4','Bio Lab Report 1','Biology 201','Lab Report','10/7/2025','High','15','Complete','','Cell division lab'],
  ['5','Econ HW 1','Economics 101','Homework','10/8/2025','Low','5','Complete','','Supply & demand'],
  ['6','Reading Response 2','English 101','Reading','10/10/2025','Medium','5','Complete','','Chapter 4-6 summary'],
  ['7','Problem Set 2','Calculus II','Problem Set','10/10/2025','High','10','Complete','','Integration basics'],
  ['8','History Quiz 1','World History','Quiz','10/12/2025','Medium','10','Complete','','WWI quiz'],
  ['9','Econ HW 2','Economics 101','Homework','10/14/2025','Low','5','Complete','','Market structures'],
  ['10','Bio Lab Report 2','Biology 201','Lab Report','10/15/2025','High','15','Complete','','DNA extraction'],
  ['11','Essay Draft 1','English 101','Essay','10/17/2025','High','15','Complete','','Literary analysis'],
  ['12','Problem Set 3','Calculus II','Problem Set','10/17/2025','High','10','Submitted','','Riemann sums'],
  ['13','Econ Quiz 1','Economics 101','Quiz','10/19/2025','Medium','10','Complete','','Micro quiz'],
  ['14','History Essay Draft','World History','Essay','10/22/2025','High','20','Submitted','','WWII causes draft'],
  ['15','Bio Lab Report 3','Biology 201','Lab Report','10/25/2025','High','15','Submitted','','Photosynthesis'],
  // Nov
  ['16','Reading Response 3','English 101','Reading','11/3/2025','Medium','5','In Progress','','Chapter 7-9'],
  ['17','Problem Set 4','Calculus II','Problem Set','11/5/2025','High','10','In Progress','','Series & sequences'],
  ['18','History Research Notes','World History','Research Paper','11/7/2025','High','10','In Progress','','Source collection'],
  ['19','Econ HW 3','Economics 101','Homework','11/10/2025','Low','5','Not Started','','Macroeconomics intro'],
  ['20','Group Project Proposal','Biology 201','Project','11/12/2025','High','10','Not Started','','Ecosystem study'],
  ['21','Essay Final','English 101','Essay','11/14/2025','High','20','Not Started','','Literary analysis final'],
  ['22','Problem Set 5','Calculus II','Problem Set','11/14/2025','High','10','Not Started','','Differential equations'],
  ['23','Econ Presentation','Economics 101','Presentation','11/18/2025','Medium','15','Not Started','','Market analysis'],
  ['24','History Research Paper','World History','Research Paper','11/21/2025','High','30','Not Started','','WWII 10-page paper'],
  ['25','Bio Group Project','Biology 201','Project','11/25/2025','High','20','Not Started','','Ecosystem report'],
  // Dec
  ['26','Final Essay','English 101','Essay','12/5/2025','High','25','Not Started','','Semester portfolio'],
  ['27','Calculus Review Sheet','Calculus II','Homework','12/5/2025','High','5','Not Started','','Exam prep'],
  ['28','Econ Final Project','Economics 101','Project','12/8/2025','High','20','Not Started','','Market research report'],
  ['29','History Essay Revision','World History','Essay','12/10/2025','High','10','Not Started','','Revised WWII essay'],
  ['30','Bio Lab Final Report','Biology 201','Lab Report','12/12/2025','High','20','Not Started','','Cumulative lab'],
];

// Course → color mapping for CF
const COURSE_COLORS = {
  'English 101':    COLOR.paleSkyBlue,
  'Calculus II':    COLOR.paleMint,
  'World History':  COLOR.paleLavender,
  'Biology 201':    COLOR.palePeach,
  'Economics 101':  COLOR.paleGold,
};

(async () => {
  const reqs = [];

  // ── Column widths ──
  const widths = [
    [0,40],[1,220],[2,130],[3,120],[4,100],[5,80],[6,70],[7,110],[8,160],[9,220],
  ];
  widths.forEach(([ci,px]) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));

  // ── Row heights ──
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 50 }, fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 4 },
      properties: { pixelSize: 32 }, fields: 'pixelSize',
    },
  });

  // ── Title row merge (row 0, cols A-J) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });

  // ── Title row format ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 10),
      cell: {
        userEnteredValue: { stringValue: '📚 Assignment Tracker' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.skyBlue),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Subtitle row (row 1) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 10),
      cell: {
        userEnteredValue: { stringValue: 'Fall 2025 | Alex Rivera | State University' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleSkyBlue),
          textFormat: { foregroundColor: hex(COLOR.darkBlue), bold: false, fontSize: 11, italic: true },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Stats row (row 2): summary counts ──
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 2), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 2, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 4, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 6, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 8, 10), mergeType: 'MERGE_ALL' } });

  const statCells = [
    { col: 0, label: 'Total', formula: '=COUNTA(B5:B204)', bg: COLOR.subheaderBg },
    { col: 2, label: 'Complete', formula: '=COUNTIF(H5:H204,"Complete")+COUNTIF(H5:H204,"Submitted")', bg: COLOR.paleMint },
    { col: 4, label: 'In Progress', formula: '=COUNTIF(H5:H204,"In Progress")', bg: COLOR.paleSkyBlue },
    { col: 6, label: 'Not Started', formula: '=COUNTIF(H5:H204,"Not Started")', bg: COLOR.paleLavender },
    { col: 8, label: 'Overdue', formula: '=COUNTIFS(H5:H204,"Not Started",E5:E204,"<"&TODAY())+COUNTIFS(H5:H204,"In Progress",E5:E204,"<"&TODAY())', bg: COLOR.paleRed },
  ];

  statCells.forEach(({ col, label, formula, bg }) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 2, 3, col, col + 2),
        cell: {
          userEnteredValue: { stringValue: `${label}: ` },
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(COLOR.nearBlack), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  });

  // ── Column header row (row 3) ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 3, 4, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.darkBlue),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          borders: {
            bottom: { style: 'SOLID', color: hex(COLOR.lightBlue), width: 2 },
          },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)',
    },
  });

  // ── Alt row shading for data rows ──
  for (let r = 4; r < 204; r += 2) {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 10),
        cell: { userEnteredFormat: { backgroundColor: hex(COLOR.altRow) } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
  }

  // ── Data row text format ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 4, 204, 0, 10),
      cell: {
        userEnteredFormat: {
          textFormat: { fontSize: 10 },
          verticalAlignment: 'MIDDLE',
          wrapStrategy: 'CLIP',
        },
      },
      fields: 'userEnteredFormat(textFormat,verticalAlignment,wrapStrategy)',
    },
  });
  // Center #, Date, Priority, Weight, Status cols
  [0, 4, 5, 6, 7].forEach(ci => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 4, 204, ci, ci + 1),
        cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat.horizontalAlignment',
      },
    });
  });

  // ── Border on header row bottom ──
  reqs.push({
    updateBorders: {
      range: gridRange(SID, 3, 4, 0, 10),
      bottom: { style: 'SOLID_MEDIUM', color: hex(COLOR.skyBlue) },
    },
  });

  await batchUpdate(id, reqs, 'assignments-format');

  // ── Values ──
  const data = [
    {
      range: 'Assignments!A4:J4',
      values: [['#','Assignment Title','Course','Type','Due Date','Priority','Weight','Status','Link','Notes']],
    },
    {
      range: `Assignments!A5:J${5 + ASSIGNMENTS.length - 1}`,
      values: ASSIGNMENTS,
    },
  ];

  // Stat formulas
  statCells.forEach(({ col, formula }) => {
    const cellAddr = `Assignments!${String.fromCharCode(65 + col)}3`;
    // append formula as suffix with label — use separate write
    data.push({ range: cellAddr, values: [[`${statCells.find(s=>s.col===col)?.label??''}: ` + formula]] });
  });

  // Override stat cells with combined label+formula
  const statData = [
    { range: 'Assignments!A3', values: [['="""Total: """&COUNTA(B5:B204)']] },
    { range: 'Assignments!C3', values: [['="""Complete: """&(COUNTIF(H5:H204,"Complete")+COUNTIF(H5:H204,"Submitted"))']] },
    { range: 'Assignments!E3', values: [['="""In Progress: """&COUNTIF(H5:H204,"In Progress")']] },
    { range: 'Assignments!G3', values: [['="""Not Started: """&COUNTIF(H5:H204,"Not Started")']] },
    { range: 'Assignments!I3', values: [['="""Overdue: """&(COUNTIFS(H5:H204,"Not Started",E5:E204,"<"&TODAY())+COUNTIFS(H5:H204,"In Progress",E5:E204,"<"&TODAY()))']] },
  ];

  await valuesBatchUpdate(id, [
    { range: 'Assignments!A4:J4', values: [['#','Assignment Title','Course','Type','Due Date','Priority','Weight','Status','Link','Notes']] },
    { range: `Assignments!A5:J${5 + ASSIGNMENTS.length - 1}`, values: ASSIGNMENTS },
    ...statData,
  ], 'assignments-values');

  console.log('Assignments tab complete —', ASSIGNMENTS.length, 'entries');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
