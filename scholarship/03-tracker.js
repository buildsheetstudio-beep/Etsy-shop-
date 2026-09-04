'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TRK = sheetMap['📋 Application Tracker'];

// [name, provider, type, amount, frequency, deadline, status, priority, eligibility, requirements, urlRaw, amountAwarded, resultDate, notes]
const SCHOLARSHIPS = [
  // Awarded (5)
  ['Melbourne Global Scholars Award','University of Melbourne','University',8000,'Annual','2025-03-01','Awarded','High','GPA 3.7+, domestic student','Personal statement, academic reference','',8000,'2025-04-15','Flagship university award — reapply annually'],
  ['Commonwealth Bank Student Scholarship','Commonwealth Bank','Corporate',5000,'One-Off','2025-02-15','Awarded','High','Finance/business students','Application form, 500-word essay','',5000,'2025-03-20','Great for CV — mention in cover letters'],
  ['Rural Student Support Bursary','State Government','Government',2500,'Annual','2025-01-31','Awarded','Medium','Rural background, financial need','Income statement, 300-word statement','',2500,'2025-03-01','Auto-renewed if still eligible'],
  ['Women in STEM Excellence Award','STEM Foundation','Private Foundation',3000,'One-Off','2025-02-28','Awarded','High','Female-identifying, STEM faculty','Portfolio, essay, 2 references','',3000,'2025-04-01','Contact Dr. Lee re: next year cycle'],
  ['Faculty of Science Dean\'s Scholarship','University of Melbourne','University',1500,'Per Semester','2025-03-15','Awarded','Medium','GPA 3.5+, Science faculty','Auto-assessed at enrolment','',1500,'2025-03-30','Auto-applied — check enrolment records'],
  // Submitted (5)
  ['Westpac Future Leaders Scholarship','Westpac','Corporate',12000,'Annual','2025-04-30','Submitted','High','Leadership potential, community involvement','Application form, leadership essay, 2 references','','','','Result expected June 2025'],
  ['Australian Government Research Training Program','Dept. of Education','Government',28000,'Annual','2025-05-15','Submitted','High','Honours/Masters students','Research proposal, supervisor endorsement','','','','Full stipend scholarship — top priority'],
  ['Ian Potter Foundation Scholarship','Ian Potter Foundation','Private Foundation',6000,'One-Off','2025-04-01','Submitted','High','Creative arts students','Portfolio, artist statement, reference','','','','Portfolio submitted March 20'],
  ['Myer Foundation Young Leaders Grant','Myer Foundation','Private Foundation',4000,'One-Off','2025-04-15','Submitted','Medium','Community leadership','Essay, CV, 2 references','','','','Word count: 800 words submitted'],
  ['University Sports Excellence Award','Uni of Melbourne','University',2000,'Annual','2025-03-20','Submitted','Medium','Varsity athlete, GPA 3.0+','Coach reference, academic record','','','','Coach reference submitted on time'],
  // In Progress (5)
  ['NAB Village Scholarship','National Australia Bank','Corporate',10000,'Annual','2025-08-31','In Progress','High','Indigenous students','Application form, community statement, reference','','','','Draft statement in progress'],
  ['Google Generation Scholarship','Google','Corporate',7500,'One-Off','2025-07-15','In Progress','High','Computer science, diversity focus','Technical essay, transcript','','','','Tech essay 60% complete'],
  ['ANZ Emerging Leaders Program','ANZ','Corporate',5000,'Annual','2025-08-01','In Progress','High','Business/commerce, leadership','Video application, essay','','','','Need to film video component'],
  ['Melbourne School of Engineering Alumni Bursary','Eng. Alumni Assoc.','University',3500,'One-Off','2025-09-30','In Progress','Medium','Engineering students, financial need','Financial statement, personal essay','','','','Financial docs being prepared'],
  ['Rotary Foundation Ambassadorial Scholarship','Rotary International','Community',15000,'One-Off','2025-10-01','In Progress','Medium','Community service record','Multiple references, personal statement, interview','','','','3 references needed — 2 confirmed'],
  // Not Started (7)
  ['Atlassian Foundation Scholarship','Atlassian','Corporate',8000,'Annual','2025-09-15','Not Started','High','Tech students, first-gen college','Application form, coding challenge, interview','','','','Opens June 2025'],
  ['Queen\'s Commonwealth Trust Scholarship','Queen\'s Commonwealth Trust','Government',20000,'Annual','2025-11-01','Not Started','High','Commonwealth citizens, social change focus','Leadership essay, project proposal','','','','Need to gather project evidence'],
  ['Deloitte Next Generation Scholarship','Deloitte','Corporate',6000,'One-Off','2025-09-01','Not Started','Medium','Accounting/Finance, GPA 3.5+','Accounting aptitude test, interview','','','','Check eligibility — need to confirm GPA'],
  ['Frank Knox Memorial Fellowship','Harvard University','Private Foundation',40000,'Annual','2025-10-15','Not Started','High','Australian citizens for Harvard study','Research proposal, 3 academic references','','','','Highly competitive — allow 3 months prep'],
  ['Laby Research Scholarship','Uni of Melbourne','University',4000,'Annual','2025-08-15','Not Started','Medium','Physics PhD students','Research summary, supervisor letter','','','','Confirm supervisor endorsement first'],
  ['Endeavour Leadership Program','Dept. of Education','Government',25000,'Per Term','2025-09-30','Not Started','High','Short-term international study','Study plan, institutional endorsement','','','','Requires faculty coordinator approval'],
  ['AustralianSuper Scholarship','AustralianSuper','Corporate',5000,'One-Off','2025-10-31','Not Started','Medium','Finance/economics students','Application + essay','','','','Online portal opens Aug 2025'],
  // Rejected / Withdrawn (3)
  ['BHP Foundation Science and Engineering Award','BHP','Corporate',6000,'One-Off','2025-02-01','Rejected','High','Engineering/science, regional background','Technical essay, references','',0,'2025-03-15','Feedback: strengthen regional connection narrative'],
  ['Legal Aid Foundation Bursary','Legal Aid','Community',2000,'One-Off','2025-01-15','Rejected','Low','Law students, financial need','Financial declaration','',0,'2025-02-28','Didn\'t meet law faculty requirement'],
  ['University Choral Society Award','University of Melbourne','University',500,'One-Off','2025-03-01','Withdrawn','Low','Music students, choral experience','Audition + application','','','','Withdrew — ineligible after faculty transfer'],
];

const HDRS = ['SCHOLARSHIP / BURSARY NAME','PROVIDER / ORGANISATION','TYPE','AMOUNT','FREQUENCY','DEADLINE','STATUS','PRIORITY','ELIGIBILITY','REQUIREMENTS','APPLICATION URL','APPLY LINK','AMOUNT AWARDED','RESULT DATE','NOTES','DAYS LEFT'];
const WIDTHS = [240,200,130,110,110,120,120,90,160,200,180,100,120,120,220,60];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB  = "'📋 Application Tracker'";
  const NCOLS = 16; // A-P

  // Column widths
  WIDTHS.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: TRK, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner (merged A:P)
  reqs.push({ mergeCells: { range: gridRange(TRK,0,1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(TRK,0,1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TRK, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  📋  Application Tracker — Scholarship & Bursary Tracker']] });

  // Row 1: summary stats bar (merged)
  reqs.push({ mergeCells: { range: gridRange(TRK,1,2,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(TRK,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.royalBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TRK, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  // Summary formula in the merged bar
  vals.push({ range: `${TAB}!A2`, values: [['="  📊 Tracked: "&COUNTA(A4:A200)&"  |  Submitted: "&COUNTIF(G4:G200,"Submitted")&"  |  Awarded: "&COUNTIF(G4:G200,"Awarded")&"  |  Rejected: "&COUNTIF(G4:G200,"Rejected")&"  |  Awarded $: "&TEXT(SUMIF(G4:G200,\"Awarded\",M4:M200),\"$#,##0\")&"  |  Success Rate: "&TEXT(IFERROR(COUNTIF(G4:G200,\"Awarded\")/(COUNTIF(G4:G200,\"Awarded\")+COUNTIF(G4:G200,\"Rejected\")),0),\"0%\")']] });

  // Row 2: column headers
  reqs.push({ repeatCell: {
    range: gridRange(TRK,2,3,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TRK, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [HDRS] });

  // Data rows (0-indexed 3 to 3+25)
  for (let i = 0; i < SCHOLARSHIPS.length; i++) {
    const ri  = 3 + i;
    const row = 4 + i;
    const bg  = i % 2 === 0 ? C.white : C.paleBlue;
    reqs.push({ repeatCell: {
      range: gridRange(TRK, ri, ri+1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    // Amount col D (index 3): currency
    reqs.push({ repeatCell: {
      range: gridRange(TRK, ri, ri+1, 3, 4),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } }},
      fields: 'userEnteredFormat.numberFormat',
    }});
    // Deadline col F (index 5): date
    reqs.push({ repeatCell: {
      range: gridRange(TRK, ri, ri+1, 5, 6),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } }},
      fields: 'userEnteredFormat.numberFormat',
    }});
    // Amount Awarded col M (index 12): currency
    reqs.push({ repeatCell: {
      range: gridRange(TRK, ri, ri+1, 12, 13),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } }},
      fields: 'userEnteredFormat.numberFormat',
    }});
    // Result Date col N (index 13): date
    reqs.push({ repeatCell: {
      range: gridRange(TRK, ri, ri+1, 13, 14),
      cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } }},
      fields: 'userEnteredFormat.numberFormat',
    }});
    // Days left col P (index 15): formula cell bg
    reqs.push({ repeatCell: {
      range: gridRange(TRK, ri, ri+1, 15, 16),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
    }});

    const [name, provider, type, amount, freq, deadline, status, priority, elig, reqs2, url, awarded, resultDate, notes] = SCHOLARSHIPS[i];
    // Write data — HYPERLINK formula for col L if URL is present
    const hyperlink = url ? `=HYPERLINK("${url}","Apply →")` : '';
    const awardedVal = awarded === '' ? '' : awarded;
    const resultDateVal = resultDate === '' ? '' : resultDate;

    vals.push({ range: `${TAB}!A${row}`, values: [[
      name, provider, type, amount, freq, deadline, status, priority, elig, reqs2, url, hyperlink, awardedVal, resultDateVal, notes,
    ]] });
    // Days left formula
    vals.push({ range: `${TAB}!P${row}`, values: [[`=IFERROR(F${row}-TODAY(),"")`]] });
  }

  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TRK, dimension: 'ROWS', startIndex: 3, endIndex: 3+SCHOLARSHIPS.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(TRK,2,3+SCHOLARSHIPS.length,0,NCOLS-1), // exclude col P from main border
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // Col P (Days Left) — narrow, subtle
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TRK, dimension: 'COLUMNS', startIndex: 15, endIndex: 16 },
    properties: { pixelSize: 60 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'tracker-format');
  await valuesBatchUpdate(id, vals, 'tracker-values');
  console.log('Application Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
