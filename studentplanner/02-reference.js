'use strict';
const { sheets, batchUpdate, valuesBatchUpdate, hex, COLOR, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Reference']; // 6

(async () => {
  const reqs = [];

  // ── Column widths ──
  const colWidths = [
    [0, 160], [1, 150], [2, 140], [3, 120], [4, 130],
    [5, 100], [6, 140], [7, 150],
  ];
  colWidths.forEach(([ci, px]) => reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    },
  }));

  // ── Header row ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.skyBlue),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Data rows alt color ──
  for (let r = 1; r <= 20; r += 2) {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, r, r + 1, 0, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(COLOR.altRow) } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
  }

  await batchUpdate(id, reqs, 'reference-format');

  // ── Values ──
  const data = [
    // Row 1: headers
    { range: 'Reference!A1:H1', values: [['Courses', 'Assignment Types', 'Status', 'Priority', 'Confidence', 'Months', 'Activity Types', 'Budget Categories']] },
    // Col A: Courses
    { range: 'Reference!A2:A6', values: [['English 101'],['Calculus II'],['World History'],['Biology 201'],['Economics 101']] },
    // Col B: Assignment Types
    { range: 'Reference!B2:B11', values: [['Essay'],['Problem Set'],['Lab Report'],['Presentation'],['Reading'],['Quiz'],['Exam'],['Research Paper'],['Project'],['Homework']] },
    // Col C: Status
    { range: 'Reference!C2:C6', values: [['Not Started'],['In Progress'],['Complete'],['Submitted'],['Late']] },
    // Col D: Priority
    { range: 'Reference!D2:D4', values: [['High'],['Medium'],['Low']] },
    // Col E: Confidence
    { range: 'Reference!E2:E6', values: [['Very Low'],['Low'],['Medium'],['High'],['Very High']] },
    // Col F: Months
    { range: 'Reference!F2:F11', values: [['August'],['September'],['October'],['November'],['December'],['January'],['February'],['March'],['April'],['May']] },
    // Col G: Activity Types
    { range: 'Reference!G2:G11', values: [['Lecture'],['Study'],['Lab'],['Office Hours'],['Group Study'],['Review Session'],['Break'],['Personal'],['Exercise'],['Sleep']] },
    // Col H: Budget Categories
    { range: 'Reference!H2:H12', values: [['Housing'],['Food'],['Transportation'],['Books & Supplies'],['Entertainment'],['Health'],['Technology'],['Clothing'],['Personal Care'],['Savings'],['Other']] },
  ];

  await valuesBatchUpdate(id, data, 'reference-values');
  console.log('Reference tab complete');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
