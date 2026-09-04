'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Reference Data'];
const S = "'Reference Data'";

// Column layout (all starting row 1):
// A: School Levels      B: Term Types       C: Class Types
// D: Subjects           E: Assignment Types F: Assignment Statuses
// G: Task Categories    H: Task Statuses    I: Attendance Statuses
// J: Grade Categories   K: Score Types      L: Priority
// M: Yes/No             N: Grade Bands      O: Subject-Color Map (label)
// P: Subject-Color Map (hex)

const SCHOOL_LEVELS    = ['Elementary','Middle School','High School','K–12','Other'];
const TERM_TYPES       = ['Full Year','Semester','Trimester','Quarter','Custom'];
const CLASS_TYPES      = ['Homeroom','Subject','Period','Course Section','Intervention Group','Club / Activity','Other'];
const SUBJECTS         = ['English Language Arts','Reading','Writing','Mathematics','Science',
                          'Social Studies','History','Geography','Art','Music',
                          'Physical Education','Health','Technology','World Language',
                          'Special Education','Advisory','Other'];
const ASSIGN_TYPES     = ['Homework','Classwork','Quiz','Test','Project','Presentation',
                          'Lab','Essay','Participation','Practice','Exit Ticket','Other'];
const ASSIGN_STATUSES  = ['Planned','Assigned','In Progress','Collected','Grading','Returned','Cancelled'];
const TASK_CATS        = ['Lesson Planning','Grading','Parent Communication','Meetings',
                          'Classroom Management','Materials','Professional Development',
                          'Administrative','Student Support','Other'];
const TASK_STATUSES    = ['Not Started','In Progress','Waiting','Complete','Cancelled'];
const ATTEND_STATUSES  = ['Present','Absent','Tardy','Excused','Early Dismissal','Remote','School Activity','Not Scheduled'];
const GRADE_CATS       = ['Homework','Classwork','Quizzes','Tests','Projects','Participation','Labs','Essays','Other'];
const SCORE_TYPES      = ['Points','Percentage','Complete / Incomplete','Letter Grade','Rubric'];
const PRIORITY         = ['Low','Medium','High','Urgent'];
const YES_NO           = ['Yes','No'];
const GRADE_BANDS      = ['Kindergarten','Grade 1','Grade 2','Grade 3','Grade 4','Grade 5',
                          'Grade 6','Grade 7','Grade 8','Grade 9','Grade 10','Grade 11',
                          'Grade 12','Mixed Grade','Other'];
const SUBJECT_COLORS   = [
  ['English Language Arts', '#C9DDF2'],
  ['Reading',               '#C9DDF2'],
  ['Writing',               '#DED2EC'],
  ['Mathematics',           '#CDE5D5'],
  ['Science',               '#C5E1DD'],
  ['Social Studies',        '#E9D7B8'],
  ['History',               '#E9D7B8'],
  ['Geography',             '#E9D7B8'],
  ['Art',                   '#EBCFD9'],
  ['Music',                 '#D8D2EA'],
  ['Physical Education',    '#F2D0BA'],
  ['Health',                '#F2D0BA'],
  ['Technology',            '#D2E2F0'],
  ['World Language',        '#EADDAA'],
  ['Special Education',     '#D4E0C8'],
  ['Advisory',              '#E4E2DF'],
  ['Other',                 '#E4E2DF'],
];

function col(list, header, colLetter) {
  const rows = [[header], ...list.map(v => [v])];
  return { range: `${S}!${colLetter}1`, values: rows };
}

(async () => {
  // Expand Reference Data sheet to 20 columns
  await batchUpdate(id, [{
    updateSheetProperties: {
      properties: { sheetId: SID, gridProperties: { columnCount: 20 } },
      fields: 'gridProperties.columnCount',
    }
  }], '02-expand cols');

  const data = [
    col(SCHOOL_LEVELS,   'School Levels',       'A'),
    col(TERM_TYPES,      'Term Types',          'B'),
    col(CLASS_TYPES,     'Class Types',         'C'),
    col(SUBJECTS,        'Subjects',            'D'),
    col(ASSIGN_TYPES,    'Assignment Types',    'E'),
    col(ASSIGN_STATUSES, 'Assignment Statuses', 'F'),
    col(TASK_CATS,       'Task Categories',     'G'),
    col(TASK_STATUSES,   'Task Statuses',       'H'),
    col(ATTEND_STATUSES, 'Attendance Statuses', 'I'),
    col(GRADE_CATS,      'Grade Categories',    'J'),
    col(SCORE_TYPES,     'Score Types',         'K'),
    col(PRIORITY,        'Priority',            'L'),
    col(YES_NO,          'Yes / No',            'M'),
    col(GRADE_BANDS,     'Grade Bands',         'N'),
  ];

  // Subject-Color map in columns O:P
  data.push({
    range: `${S}!O1`, values: [
      ['Subject', 'Pastel Color'],
      ...SUBJECT_COLORS
    ]
  });

  await valuesBatchUpdate(id, data, '02-reference values');

  // Format header row bold + tinted
  const fmt = [];
  const cols = SCHOOL_LEVELS.length + TASK_CATS.length; // just color all 16 header cells
  for (let c = 0; c < 16; c++) {
    fmt.push({ repeatCell: {
      range: gridRange(SID, 0, 1, c, c+1),
      cell: { userEnteredFormat: {
        textFormat: { bold: true, foregroundColor: hex(C.text), fontFamily: 'Arial', fontSize: 9 },
        backgroundColor: hex(C.altRow),
        borders: { bottom: { style: 'SOLID', color: hex(C.border) } },
      }},
      fields: 'userEnteredFormat(textFormat,backgroundColor,borders)',
    }});
  }

  // Widen columns a bit
  for (let c = 0; c < 16; c++) {
    fmt.push({ updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 },
      properties: { pixelSize: c < 14 ? 160 : 140 },
      fields: 'pixelSize',
    }});
  }

  // Hide tab after build
  fmt.push({ updateSheetProperties: {
    properties: { sheetId: SID, hidden: true },
    fields: 'hidden',
  }});

  await batchUpdate(id, fmt, '02-reference format');
  console.log('✅ Reference Data done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
