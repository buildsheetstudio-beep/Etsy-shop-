'use strict';
const { batchUpdate, valuesBatchUpdate, hex, COLOR, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Dashboard']; // 0

(async () => {
  const reqs = [];

  // ── Column widths ──
  const widths = [
    [0,140],[1,130],[2,130],[3,130],[4,130],[5,130],[6,130],[7,130],[8,130],[9,130],
    [10,130],[11,130],[12,130],[13,50],
  ];
  widths.forEach(([ci,px]) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));

  // ── Row heights ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 5 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 6 }, properties: { pixelSize: 60 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 10 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 10, endIndex: 11 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 11, endIndex: 20 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 20, endIndex: 35 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── Title row (row 0) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 14),
      cell: {
        userEnteredValue: { stringValue: '🎓 Ultimate Student Planner — Academic Dashboard' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.skyBlue),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 20 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Subtitle row (row 1) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 8, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 8),
      cell: {
        userEnteredValue: { stringValue: 'Alex Rivera | State University | Fall 2025' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleSkyBlue),
          textFormat: { foregroundColor: hex(COLOR.darkBlue), bold: true, fontSize: 13 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          padding: { left: 12 },
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 8, 14),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleSkyBlue),
          textFormat: { foregroundColor: hex(COLOR.darkBlue), bold: false, italic: true, fontSize: 11 },
          horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
          padding: { right: 12 },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── KPI Cards Row (rows 3-5, 2 rows per card) ──
  // 6 KPI cards across 12 cols (2 cols each)
  const kpiCards = [
    { label: 'Semester GPA',          formula: `='Grades & GPA'!C32`,                                              fmt: '0.00',   bg: COLOR.paleLavender, fg: COLOR.violet,    icon: '🎓' },
    { label: 'Assignments Complete',   formula: `=COUNTIF(Assignments!H5:H204,"Complete")+COUNTIF(Assignments!H5:H204,"Submitted")`, fmt: '0', bg: COLOR.paleMint, fg: COLOR.emerald, icon: '✅' },
    { label: 'Due This Week',          formula: `=COUNTIFS(Assignments!E5:E204,">="&TODAY(),Assignments!E5:E204,"<="&TODAY()+7,Assignments!H5:H204,"<>"&"Complete",Assignments!H5:H204,"<>"&"Submitted")`, fmt: '0', bg: COLOR.paleGold, fg: COLOR.amber, icon: '📅' },
    { label: 'Study Hours This Week',  formula: `=COUNTIF('Weekly Planner'!B5:H37,"Study")*0.5`,                  fmt: '0.0',    bg: COLOR.paleSkyBlue,  fg: COLOR.darkBlue, icon: '⏱️' },
    { label: 'Exams Remaining',        formula: `=COUNTIF('Exam Prep'!F6:F10,">0")`,                              fmt: '0',      bg: COLOR.palePeach,    fg: COLOR.orange,   icon: '📝' },
    { label: 'Budget Balance',         formula: `='Student Budget'!D34`,                                           fmt: '"$"#,##0.00', bg: COLOR.paleMint, fg: COLOR.emerald, icon: '💰' },
  ];

  // Row 2: "KPI" label row
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 0, 14),
      cell: {
        userEnteredValue: { stringValue: '── KEY METRICS ──' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.darkBlue),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // KPI label row (row 3) and value row (row 4)
  kpiCards.forEach(({ label, bg, fg, icon }, ci) => {
    const c0 = ci * 2;
    const c1 = c0 + 2;
    reqs.push({ mergeCells: { range: gridRange(SID, 3, 4, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, 4, 5, c0, c1), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 3, 4, c0, c1),
        cell: {
          userEnteredValue: { stringValue: `${icon} ${label}` },
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(fg), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 4, 5, c0, c1),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(COLOR.white),
            textFormat: { foregroundColor: hex(fg), bold: true, fontSize: 20 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  });

  // ── Section labels row (row 5) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 5, 6, 0, 6), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 5, 6, 6, 14), mergeType: 'MERGE_ALL' } });

  const sectionLabels = [
    { c0: 0, c1: 6,  label: '📋 UPCOMING DEADLINES (Next 7 Days)',  bg: COLOR.skyBlue },
    { c0: 6, c1: 14, label: '📊 COURSE OVERVIEW & GRADES',           bg: COLOR.lavender },
  ];
  sectionLabels.forEach(({ c0, c1, label, bg }) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 5, 6, c0, c1),
        cell: {
          userEnteredValue: { stringValue: label },
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 11 },
            horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
            padding: { left: 8 },
          },
        },
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
  });

  // ── Upcoming Deadlines section (rows 6-15, cols 0-5) ──
  // Header row (row 6)
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 6, 7, 0, 6),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleSkyBlue),
          textFormat: { foregroundColor: hex(COLOR.darkBlue), bold: true, fontSize: 9 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Data rows 7-15
  for (let r = 7; r < 16; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 6),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 0 ? COLOR.altRow : COLOR.white),
            textFormat: { fontSize: 9 }, verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // ── Course Overview section (rows 6-15, cols 6-13) ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 6, 7, 6, 14),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleLavender),
          textFormat: { foregroundColor: hex(COLOR.violet), bold: true, fontSize: 9 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  for (let r = 7; r < 12; r++) {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 6, 14),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(r % 2 === 0 ? COLOR.paleLavender : COLOR.white),
            textFormat: { fontSize: 10 }, verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // ── Progress Bar section (rows 16-21, below deadlines) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 16, 17, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 16, 17, 0, 14),
      cell: {
        userEnteredValue: { stringValue: '📈 SEMESTER PROGRESS' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.mint),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
          padding: { left: 8 },
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 16, endIndex: 17 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // Progress rows 17-21 (5 courses)
  const courses = ['English 101','Calculus II','World History','Biology 201','Economics 101'];
  const courseBgs = [COLOR.paleSkyBlue, COLOR.paleMint, COLOR.paleLavender, COLOR.palePeach, COLOR.paleGold];
  courses.forEach((course, ci) => {
    const r = 17 + ci;
    reqs.push({ mergeCells: { range: gridRange(SID, r, r + 1, 0, 2), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, r, r + 1, 2, 8), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, r, r + 1, 8, 11), mergeType: 'MERGE_ALL' } });
    reqs.push({ mergeCells: { range: gridRange(SID, r, r + 1, 11, 14), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 14),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(courseBgs[ci]),
            textFormat: { fontSize: 10 }, verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  });

  // ── Chart anchor area (rows 22-38, cols N col 13+) ──
  // Chart helper data at N5:O9 (col 13-14, rows 4-8)
  reqs.push({ mergeCells: { range: gridRange(SID, 22, 23, 0, 14), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 22, 23, 0, 14),
      cell: {
        userEnteredValue: { stringValue: '── CHARTS ──' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.subheaderBg),
          textFormat: { foregroundColor: hex(COLOR.darkBlue), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Background for entire dashboard ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 60, 0, 14),
      cell: { userEnteredFormat: { backgroundColor: hex(COLOR.iceBlue) } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  });
  // Override specific sections above (they have their own bg)

  await batchUpdate(id, reqs, 'dashboard-format');

  // ── Values ──
  // KPI values
  const kpiValueData = kpiCards.map(({ formula, fmt }, ci) => ({
    range: `Dashboard!${String.fromCharCode(65 + ci * 2)}5`,
    values: [[formula]],
  }));

  // Subtitle date
  const subtitleDate = { range: 'Dashboard!I2', values: [[`="""Today: """&TEXT(TODAY(),"mmmm d, yyyy")`]] };

  // Deadline headers + QUERY formula
  const deadlineHeader = { range: 'Dashboard!A7:F7', values: [['Title','Course','Type','Due Date','Priority','Status']] };
  const deadlineQuery = {
    range: 'Dashboard!A8',
    values: [[`=IFERROR(QUERY(Assignments!A5:J204,"SELECT B,C,D,E,F,H WHERE E>=date '"&TEXT(TODAY(),"yyyy-mm-dd")&"' AND E<=date '"&TEXT(TODAY()+7,"yyyy-mm-dd")&"' AND H<>'Complete' AND H<>'Submitted' ORDER BY E ASC LIMIT 9",0),"")`]],
  };

  // Course overview headers + data
  const courseHeader = { range: 'Dashboard!G7:N7', values: [['Course','Credits','Current %','Letter','GPA Pts','Done','Pending','Sparkline']] };
  const courseRows = [
    `='Grades & GPA'!A7`,
    `='Grades & GPA'!A8`,
    `='Grades & GPA'!A9`,
    `='Grades & GPA'!A10`,
    `='Grades & GPA'!A11`,
  ].map((nameF, ci) => {
    const gr = ci + 7; // Grades row (1-indexed: 7-11)
    return {
      range: `Dashboard!G${8 + ci}:N${8 + ci}`,
      values: [[
        nameF,
        `='Grades & GPA'!B${gr}`,
        `=IFERROR('Grades & GPA'!O${gr},0)`,
        `='Grades & GPA'!Q${gr}`,
        `='Grades & GPA'!P${gr}`,
        `=COUNTIFS(Assignments!C$5:C$204,'Grades & GPA'!A${gr},Assignments!H$5:H$204,"Complete")+COUNTIFS(Assignments!C$5:C$204,'Grades & GPA'!A${gr},Assignments!H$5:H$204,"Submitted")`,
        `=COUNTIFS(Assignments!C$5:C$204,'Grades & GPA'!A${gr},Assignments!H$5:H$204,"<>"&"Complete",Assignments!H$5:H$204,"<>"&"Submitted")`,
        `=IFERROR(REPT("▰",ROUND('Grades & GPA'!O${gr}/10,0))&REPT("▱",10-ROUND('Grades & GPA'!O${gr}/10,0)),"")`,
      ]],
    };
  });

  // Progress bars (rows 18-22)
  const progressData = courses.map((course, ci) => {
    const r = 18 + ci;
    const grRow = ci + 7; // 1-indexed row in Grades
    return {
      range: `Dashboard!A${r}:N${r}`,
      values: [[
        course,
        '', '',
        `=IFERROR(REPT("▰",ROUND('Grades & GPA'!O${grRow}/10,0))&REPT("▱",10-ROUND('Grades & GPA'!O${grRow}/10,0)),"──────────")`,
        '', '', '', '', '',
        `=IFERROR(TEXT('Grades & GPA'!O${grRow}/100,"0%"),"")`,
        '', '',
        `='Grades & GPA'!Q${grRow}`,
        '',
      ]],
    };
  });

  // Chart helper data at col N (13), rows 5-9 (for assignment status donut)
  const chartHelper = {
    range: 'Dashboard!N5:O9',
    values: [
      ['Status', 'Count'],
      ['Complete', `=COUNTIF(Assignments!H5:H204,"Complete")+COUNTIF(Assignments!H5:H204,"Submitted")`],
      ['In Progress', `=COUNTIF(Assignments!H5:H204,"In Progress")`],
      ['Not Started', `=COUNTIF(Assignments!H5:H204,"Not Started")`],
      ['Late', `=COUNTIFS(Assignments!H5:H204,"Not Started",Assignments!E5:E204,"<"&TODAY())+COUNTIFS(Assignments!H5:H204,"In Progress",Assignments!E5:E204,"<"&TODAY())`],
    ],
  };

  await valuesBatchUpdate(id, [
    subtitleDate,
    ...kpiValueData,
    deadlineHeader,
    deadlineQuery,
    courseHeader,
    ...courseRows,
    ...progressData,
    chartHelper,
  ], 'dashboard-values');

  console.log('Dashboard tab complete');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
