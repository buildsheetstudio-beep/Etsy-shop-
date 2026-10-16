'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const AP = sheetMap['🏥 Appointments Log'];

// 10 cols, 210 rows
// Row 0: Title, Row 1: Disclaimer, Row 2: Summary labels, Row 3: Summary values, Row 4: Headers, Row 5+: data

(async () => {
  const reqs = [];

  // Title row 0
  reqs.push({ mergeCells: { range: gridRange(AP, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(AP, 0, 1, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: AP, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Disclaimer row 1
  reqs.push({ mergeCells: { range: gridRange(AP, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(AP, 1, 2, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream), textFormat: { foregroundColor: hex(C.warmGrey), italic: true, fontSize: 8 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment,wrapStrategy)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: AP, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Summary rows 2-3
  reqs.push({ repeatCell: {
    range: gridRange(AP, 2, 3, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(AP, 3, 4, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: AP, dimension: 'ROWS', startIndex: 2, endIndex: 4 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Header row 4
  reqs.push({ repeatCell: {
    range: gridRange(AP, 4, 5, 0, 10),
    cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: AP, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // Alt rows
  for (let r = 5; r < 210; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(AP, r, r+1, 0, 10),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Input cell wrap
  reqs.push({ repeatCell: {
    range: gridRange(AP, 5, 210, 0, 10),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)',
  }});

  // Column widths
  const colW = [110, 150, 130, 100, 140, 100, 250, 100, 120, 120];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: AP, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  reqs.push({ updateDimensionProperties: { range: { sheetId: AP, dimension: 'ROWS', startIndex: 5, endIndex: 210 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(AP, 4, 210, 0, 10), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border } });

  await batchUpdate(id, reqs, 'ap-format');

  const vdata = [];
  vdata.push({ range: `'🏥 Appointments Log'!A1`, values: [['🏥 Appointments Log']] });
  vdata.push({ range: `'🏥 Appointments Log'!A2`, values: [['⚠️ Always attend all recommended prenatal appointments. Contact your provider if you need to reschedule.']] });

  vdata.push({ range: `'🏥 Appointments Log'!A3:J3`, values: [[
    'Total Appts','Completed','Upcoming','Next Appointment','Avg Weight Gain','','','','','',
  ]]});
  vdata.push({ range: `'🏥 Appointments Log'!A4:J4`, values: [[
    '=COUNTA(A6:A210)',
    '=COUNTIF(H6:H210,TRUE)',
    '=COUNTIF(H6:H210,FALSE)',
    '=IFERROR(TEXT(MINIFS(A6:A210,H6:H210,FALSE,A6:A210,">"&TODAY()),"DD MMM YYYY"),"None scheduled")',
    '=IFERROR(ROUND(AVERAGE(F6:F210),1)&" lbs","—")',
    '','','','','',
  ]]});

  vdata.push({ range: `'🏥 Appointments Log'!A5:J5`, values: [[
    'Date','Provider','Type','Week #','Location','Weight (lbs)','Notes / Questions','Completed','Follow-Up Needed','Follow-Up Date',
  ]]});

  const appts = [
    ['2025-06-01','Dr Sarah Collins','OB Checkup',8,'City OB Clinic',145,'Confirmed pregnancy! Due date set to Dec 15. Started prenatals.',true,false,''],
    ['2025-06-29','Dr Sarah Collins','OB Checkup',12,'City OB Clinic',146.5,'12-week bloodwork all clear. Nuchal translucency normal.',true,false,''],
    ['2025-07-06','Ultrasound Centre','Ultrasound',13,'Radiology Associates',147,'NT scan clear. Baby measuring on track!',true,false,''],
    ['2025-07-27','Dr Sarah Collins','OB Checkup',16,'City OB Clinic',149,'Heart rate 152bpm. All looking great!',true,false,''],
    ['2025-08-10','Ultrasound Centre','Anatomy Scan',20,'Radiology Associates',151,'Anatomy scan complete — everything looks wonderful!',true,false,''],
    ['2025-08-31','Dr Sarah Collins','OB Checkup',22,'City OB Clinic',153,'BP perfect. Measuring perfectly.',true,false,''],
    ['2025-09-28','Dr Sarah Collins','GD Screening',26,'City OB Clinic',155.5,'1-hour GD test. Results pending.',true,true,'2025-10-05'],
    ['2025-10-19','Dr Sarah Collins','OB Checkup',30,'City OB Clinic',158,'Growth scan ordered for week 32.',false,false,''],
    ['2025-11-02','Ultrasound Centre','Ultrasound',32,'Radiology Associates',160,'Growth scan — check baby size.',false,false,''],
    ['2025-11-23','Dr Sarah Collins','OB Checkup',36,'City OB Clinic',162,'GBS test. Final stretch!',false,false,''],
  ];
  vdata.push({ range: `'🏥 Appointments Log'!A6:J15`, values: appts });

  await valuesBatchUpdate(id, vdata, 'ap-values');
  console.log('Appointments Log complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
