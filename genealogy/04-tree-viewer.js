'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TV = sheetMap['🌳 Family Tree / Relationship Viewer'];
const FM_NAME = "'👨‍👩‍👧‍👦 Family Members'";

// Lookup helper: finds col X from Family Members by ID in col A
function fm(col, matchId) {
  return `IFERROR(INDEX(${FM_NAME}!${col}:${col},MATCH(${matchId},${FM_NAME}!A:A,0)),"")`;
}

// Returns "Name (Birth Date – Death Date)" string for a given ID
function personSummary(idRef) {
  const name = `IFERROR(INDEX(${FM_NAME}!B:B,MATCH(${idRef},${FM_NAME}!A:A,0)),"")`;
  const bd   = `IFERROR(TEXT(INDEX(${FM_NAME}!D:D,MATCH(${idRef},${FM_NAME}!A:A,0)),"YYYY-MM-DD"),"")`;
  const dd   = `IFERROR(TEXT(INDEX(${FM_NAME}!F:F,MATCH(${idRef},${FM_NAME}!A:A,0)),"YYYY-MM-DD"),"")`;
  return `=IFERROR(IF(${name}="","",${name}&" ("&${bd}&IF(${dd}<>"","–"&${dd},"")&")"),"")`;
}

(async () => {
  const data = [];

  // Row 1: title + input
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A1", values: [['🌳 FAMILY TREE / RELATIONSHIP VIEWER']] });
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A2", values: [['Enter Person ID →']] });
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!B2", values: [[31]] }); // default: Sarah (user)

  // Auto-populated fields (rows 4-16)
  const fields = [
    ['Full Name',      `=IFERROR(${fm('B','$B$2')},"—")`],
    ['Nickname',       `=IFERROR(${fm('C','$B$2')},"—")`],
    ['Birth Date',     `=IFERROR(TEXT(${fm('D','$B$2')},"YYYY-MM-DD"),"—")`],
    ['Birth Place',    `=IFERROR(${fm('E','$B$2')},"—")`],
    ['Death Date',     `=IFERROR(IF(${fm('F','$B$2')}="","Living",TEXT(${fm('F','$B$2')},"YYYY-MM-DD")),"—")`],
    ['Death Place',    `=IFERROR(IF(${fm('G','$B$2')}="","—",${fm('G','$B$2')}),"—")`],
    ['Cause of Death', `=IFERROR(IF(${fm('H','$B$2')}="","—",${fm('H','$B$2')}),"—")`],
    ['Nationality',    `=IFERROR(${fm('I','$B$2')},"—")`],
    ['Languages',      `=IFERROR(${fm('J','$B$2')},"—")`],
    ['Education',      `=IFERROR(${fm('K','$B$2')},"—")`],
    ['Notes',          `=IFERROR(${fm('O','$B$2')},"—")`],
  ];

  fields.forEach(([label, formula], i) => {
    const row = 4 + i;
    data.push({ range: `'🌳 Family Tree / Relationship Viewer'!A${row}:B${row}`, values: [[label, formula]] });
  });

  // Section: Parents (row 17)
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A17", values: [['👨‍👩‍👧 PARENTS']] });
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A18:B18", values: [['Parent 1', personSummary(`${fm('K2','$B$2')}`)]] });
  // Correction: Parent IDs are in columns L (idx 11) and M (idx 12)
  // In A1 notation that's columns L and M
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A18:B18", values: [[
    'Parent 1',
    `=IFERROR(IF(${fm('L','$B$2')}="","—",${fm('L','$B$2')}&" — "&IFERROR(INDEX(${FM_NAME}!B:B,MATCH(${fm('L','$B$2')},${FM_NAME}!A:A,0)),"")),"—")`
  ]] });
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A19:B19", values: [[
    'Parent 2',
    `=IFERROR(IF(${fm('M','$B$2')}="","—",${fm('M','$B$2')}&" — "&IFERROR(INDEX(${FM_NAME}!B:B,MATCH(${fm('M','$B$2')},${FM_NAME}!A:A,0)),"")),"—")`
  ]] });

  // Section: Spouse (row 21)
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A21", values: [['💍 SPOUSE']] });
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A22:B22", values: [[
    'Spouse',
    `=IFERROR(IF(${fm('N','$B$2')}="","—",${fm('N','$B$2')}&" — "&IFERROR(INDEX(${FM_NAME}!B:B,MATCH(${fm('N','$B$2')},${FM_NAME}!A:A,0)),"")),"—")`
  ]] });

  // Section: Children (row 24 onward)
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A24", values: [['👶 CHILDREN (from Family Members registry)']] });
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A25:C25", values: [['ID','Full Name','Birth Date']] });
  // FILTER: rows where Parent 1 ID (col L) or Parent 2 ID (col M) = B2
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!A26", values: [
    [`=IFERROR(FILTER(${FM_NAME}!A:A,(${FM_NAME}!L:L=$B$2)+(${FM_NAME}!M:M=$B$2)>0),"No children found in registry")`]
  ] });
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!B26", values: [
    [`=IFERROR(FILTER(${FM_NAME}!B:B,(${FM_NAME}!L:L=$B$2)+(${FM_NAME}!M:M=$B$2)>0),"")`]
  ] });
  data.push({ range: "'🌳 Family Tree / Relationship Viewer'!C26", values: [
    [`=IFERROR(FILTER(${FM_NAME}!D:D,(${FM_NAME}!L:L=$B$2)+(${FM_NAME}!M:M=$B$2)>0),"")`]
  ] });

  await valuesBatchUpdate(id, data, 'tv-values');

  const reqs = [];

  // Title banner row 1
  reqs.push({
    mergeCells: { range: gridRange(TV, 0, 1, 0, 5), mergeType: 'MERGE_ALL' },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 0, 1, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TV, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Input row 2
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 1, 2, 0, 1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.gold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 1, 2, 1, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { foregroundColor: hex(C.sepia), bold: true, fontSize: 16 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: TV, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Label col A (rows 4-14)
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 3, 15, 0, 1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { foregroundColor: hex(C.darkText), bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  // Value col B (rows 4-14)
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 3, 15, 1, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { foregroundColor: hex(C.darkText), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    },
  });

  // Section headers (rows 17, 21, 24)
  [16, 20, 23].forEach(ri => {
    reqs.push({
      repeatCell: {
        range: gridRange(TV, ri, ri+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(C.gold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({ mergeCells: { range: gridRange(TV, ri, ri+1, 0, 5), mergeType: 'MERGE_ALL' } });
    reqs.push({ updateDimensionProperties: { range: { sheetId: TV, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  });

  // Parent/spouse rows
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 17, 20, 0, 1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 17, 20, 1, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 21, 23, 0, 1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 21, 23, 1, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Children header row 25
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 24, 25, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 }, horizontalAlignment: 'CENTER' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  // Children data rows (formula spills)
  reqs.push({
    repeatCell: {
      range: gridRange(TV, 25, 36, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.parchment), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // Column widths
  [[0,180],[1,260],[2,120]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: TV, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'tv-format');
  console.log('Family Tree / Relationship Viewer complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
