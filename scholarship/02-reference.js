'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

const LISTS = [
  // [header, ...items]
  ['SCHOLARSHIP TYPE',   'University','Government','Private Foundation','Corporate','Community','Research Grant','Bursary','Fellowship','Award','Other'],
  ['FREQUENCY',          'One-Off','Annual','Per Semester','Per Term','Monthly','TBD'],
  ['STATUS',             'Not Started','Researching','In Progress','Submitted','Awarded','Rejected','Waitlisted','Withdrawn'],
  ['PRIORITY',           'High','Medium','Low'],
  ['RESULT STATUS',      'Awarded','Rejected','Waitlisted','Pending'],
  ['REQUIREMENTS',       'Personal statement','Academic transcript','References (2)','Portfolio','CV / Resume','Application form','Essay','Interview','Audition','Financial declaration'],
  ['ELIGIBILITY',        'GPA 3.5+','GPA 3.7+','Financial need','First-gen student','Indigenous student','Female-identifying','Domestic student','International student','Rural background','STEM faculty','Community service'],
  ['INSTITUTION TYPE',   'University','TAFE','Community College','Research Institute','Online Provider','Other'],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB  = "'📋 Reference Data'";
  const NCOLS = LISTS.length;

  // Col widths
  LISTS.forEach((_, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: 170 }, fields: 'pixelSize',
    }});
  });

  // Row 0 title
  reqs.push({ mergeCells: { range: gridRange(REF,0,1,0,NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(REF,0,1,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  📋  Reference Data — Scholarship & Bursary Tracker']] });

  // Row 1 column headers
  reqs.push({ repeatCell: {
    range: gridRange(REF,1,2,0,NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [LISTS.map(l => l[0])] });

  // Data rows
  const maxItems = Math.max(...LISTS.map(l => l.length - 1));
  for (let ri = 0; ri < maxItems; ri++) {
    const bg = ri % 2 === 0 ? C.white : C.paleBlue;
    reqs.push({ repeatCell: {
      range: gridRange(REF, 2+ri, 3+ri, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    const row = LISTS.map(l => l[1+ri] || '');
    vals.push({ range: `${TAB}!A${3+ri}`, values: [row] });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: REF, dimension: 'ROWS', startIndex: 2, endIndex: 2+maxItems },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  await batchUpdate(id, reqs, 'reference-format');
  await valuesBatchUpdate(id, vals, 'reference-values');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
