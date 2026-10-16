'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const ST = sheetMap['🤒 Symptom & Wellbeing Tracker'];

// 12 cols, 210 rows
// Row 0: Title (full merge)
// Row 1: Disclaimer
// Row 2: Summary labels
// Row 3: Summary values (formulas)
// Row 4: Column headers
// Row 5+: Data

(async () => {
  const reqs = [];

  // Title row 0
  reqs.push({ mergeCells: { range: gridRange(ST, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(ST, 0, 1, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.deepLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: ST, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Disclaimer row 1
  reqs.push({ mergeCells: { range: gridRange(ST, 1, 2, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(ST, 1, 2, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream), textFormat: { foregroundColor: hex(C.warmGrey), italic: true, fontSize: 8 }, verticalAlignment: 'MIDDLE', horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment,wrapStrategy)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: ST, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Summary labels row 2
  reqs.push({ repeatCell: {
    range: gridRange(ST, 2, 3, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: ST, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 26 }, fields: 'pixelSize' } });

  // Summary values row 3
  reqs.push({ repeatCell: {
    range: gridRange(ST, 3, 4, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleLavender), textFormat: { foregroundColor: hex(C.deepPlum), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: ST, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Header row 4
  reqs.push({ repeatCell: {
    range: gridRange(ST, 4, 5, 0, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.softLavender), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: ST, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // Alt rows 5-209
  for (let r = 5; r < 210; r += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(ST, r, r+1, 0, 12),
      cell: { userEnteredFormat: { backgroundColor: hex(C.lavenderMist) } },
      fields: 'userEnteredFormat.backgroundColor',
    }});
  }

  // Input cells
  reqs.push({ repeatCell: {
    range: gridRange(ST, 5, 210, 0, 12),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } },
    fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)',
  }});

  // Column widths A-L
  const colW = [100, 90, 150, 90, 90, 90, 90, 90, 90, 160, 90, 150];
  colW.forEach((w, i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: ST, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  // Data row height
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ST, dimension: 'ROWS', startIndex: 5, endIndex: 210 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Borders
  const border = { style: 'SOLID', color: hex(C.lavenderBorder) };
  reqs.push({ updateBorders: { range: gridRange(ST, 4, 210, 0, 12), innerHorizontal: border, innerVertical: border, top: border, bottom: border, left: border, right: border } });

  await batchUpdate(id, reqs, 'st-format');

  // ── Values ──────────────────────────────────────────────────────────────────
  const vdata = [];

  vdata.push({ range: `'🤒 Symptom & Wellbeing Tracker'!A1`, values: [['🤒 Symptom & Wellbeing Tracker']] });
  vdata.push({ range: `'🤒 Symptom & Wellbeing Tracker'!A2`, values: [['⚠️ This tracker is for personal journaling only. Seek immediate medical attention for severe or unusual symptoms.']] });

  vdata.push({ range: `'🤒 Symptom & Wellbeing Tracker'!A3:L3`, values: [[
    'Total Entries','Severe Count','Most Common Symptom','Avg Mood Score',
    'Days Tracked','% Severe','Best Day','Worst Day','','','','',
  ]]});

  vdata.push({ range: `'🤒 Symptom & Wellbeing Tracker'!A4:L4`, values: [[
    '=COUNTA(A6:A210)',
    '=COUNTIF(E6:E210,"Severe")+COUNTIF(E6:E210,"Very Severe")',
    '=IFERROR(INDEX(C6:C210,MATCH(MAX(COUNTIF(C6:C210,C6:C210)),COUNTIF(C6:C210,C6:C210),0)),"—")',
    '=IFERROR(ROUND(AVERAGE(J6:J210),1),"—")',
    '=COUNTA(A6:A210)',
    '=IFERROR(TEXT(B4/A4,"0%"),"—")',
    '=IFERROR(TEXT(MAXIFS(A6:A210,J6:J210,MAX(J6:J210)),"DD MMM YYYY"),"—")',
    '=IFERROR(TEXT(MINIFS(A6:A210,J6:J210,MIN(J6:J210)),"DD MMM YYYY"),"—")',
    '','','','',
  ]]});

  vdata.push({ range: `'🤒 Symptom & Wellbeing Tracker'!A5:L5`, values: [[
    'Date','Week #','Symptom','Time of Day','Severity','Duration','Trigger','Relief Used','Mood (1-10)','Notes','','',
  ]]});

  // 15 sample entries (adjust dates to be before due date 2025-12-15)
  const entries = [
    ['2025-06-01',8,'Morning sickness','Morning','Moderate','2 hours','Empty stomach','Ginger tea',6,'Pretty rough morning'],
    ['2025-06-08',9,'Fatigue','All day','Mild','All day','Busy week','Rest',5,'Needed two naps'],
    ['2025-06-15',10,'Morning sickness','Morning','Severe','3 hours','Coffee smell','Crackers + rest',4,'Worst week so far'],
    ['2025-06-22',11,'Heartburn','Evening','Mild','1 hour','Spicy food','Antacids',7,'Getting used to this'],
    ['2025-06-29',12,'Round ligament pain','Afternoon','Moderate','30 mins','Standing too fast','Slow movements',6,'Startled me at first'],
    ['2025-07-06',13,'Fatigue','Afternoon','Moderate','3 hours','Work stress','Early night',5,'Need more rest'],
    ['2025-07-13',14,'Headache','Morning','Moderate','2 hours','Dehydration','Water + rest',6,'Drinking more water now'],
    ['2025-07-20',15,'Mood swings','All day','Mild','All day','Hormones','Talk to partner',7,'Cried at a commercial!'],
    ['2025-07-27',16,'Back pain','Evening','Mild','2 hours','Sitting at desk','Stretching',7,'Need better chair'],
    ['2025-08-03',17,'Swelling (edema)','Evening','Mild','Few hours','Hot weather','Elevate feet',7,'Feet look like sausages'],
    ['2025-08-10',18,'Insomnia','Night','Moderate','4 hours','Baby movements','Pillow fort',5,'Baby very active at night!'],
    ['2025-08-17',19,'Heartburn','After meals','Moderate','1 hour','Large meals','Smaller meals',6,'Eating 6 small meals now'],
    ['2025-08-24',20,'Braxton Hicks','Afternoon','Mild','20 mins','Exercise','Rest + water',8,'Completely normal apparently'],
    ['2025-08-31',21,'Leg cramps','Night','Severe','1 hour','Calcium deficiency','Stretching',4,'Woke me up at 2am'],
    ['2025-09-07',22,'Shortness of breath','Climbing stairs','Mild','5 mins','Physical activity','Slow down',7,'Uterus pressing on diaphragm'],
  ];
  vdata.push({ range: `'🤒 Symptom & Wellbeing Tracker'!A6:J20`, values: entries });

  await valuesBatchUpdate(id, vdata, 'st-values');
  console.log('Symptom Tracker complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
