'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SP = sheetMap['🛡️ Safety Plan & Crisis Resources'];

(async () => {
  const data = [];

  // Row 1: title
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A1", values: [['🛡️ MY SAFETY PLAN & CRISIS RESOURCES']] });
  // Row 2: italicized guidance note
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A2", values: [['Complete this page with your therapist or psychiatrist — it works best as a shared plan, not a solo one.']] });

  // Section 1: Warning Signs
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A4", values: [['⚠️ WARNING SIGNS I NOTICE (moods, behaviors, thoughts that tell me something is shifting)']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A5:B5", values: [['Warning Sign 1','← Type your warning signs here']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A6", values: [['Warning Sign 2']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A7", values: [['Warning Sign 3']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A8", values: [['Warning Sign 4']] });

  // Section 2: Coping Strategies
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A10", values: [['💡 COPING STRATEGIES THAT HELP ME (things I can do on my own to feel safer)']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A11:B11", values: [['Strategy 1','← Type your coping strategies here']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A12", values: [['Strategy 2']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A13", values: [['Strategy 3']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A14", values: [['Strategy 4']] });

  // Section 3: People to Contact
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A16", values: [['👥 PEOPLE I CAN CONTACT (friends or family I can reach out to)']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A17:C17", values: [['Name','Relationship','Phone']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A18:C18", values: [['','','']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A19:C19", values: [['','','']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A20:C20", values: [['','','']] });

  // Section 4: Care Team
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A22", values: [['🏥 MY CARE TEAM']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A23:C23", values: [['Provider Name','Role','Phone / Contact']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A24:C24", values: [['','Psychiatrist','']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A25:C25", values: [['','Therapist','']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A26:C26", values: [['','Primary Care','']] });

  // Section 5: Reasons to Stay Safe
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A28", values: [['💙 REASONS TO STAY SAFE (what matters to me — people, things, goals)']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A29:B29", values: [['Reason 1','← Type your reasons here']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A30", values: [['Reason 2']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A31", values: [['Reason 3']] });

  // Section 6: Environment Safety
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A33", values: [['🔒 MAKING MY ENVIRONMENT SAFER (steps I\'ve taken or can take)']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A34:B34", values: [['Step 1','← Type your environment safety steps here']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A35", values: [['Step 2']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A36", values: [['Step 3']] });

  // Section 7: CRISIS RESOURCES — permanently pre-filled
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A38", values: [['🆘 CRISIS RESOURCES']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A39:B39", values: [['988 Suicide & Crisis Lifeline','Call or text 988 — available 24/7 in the US. Free, confidential.']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A40:B40", values: [['Emergency','Call 911 or go to your nearest emergency room.']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A41:B41", values: [['Outside the US','Contact your local emergency number or crisis line.']] });
  data.push({ range: "'🛡️ Safety Plan & Crisis Resources'!A42:B42", values: [['Crisis Text Line (US)','Text HOME to 741741']] });

  await valuesBatchUpdate(id, data, 'sp-values');

  const reqs = [];

  // Title banner row 1
  reqs.push({ mergeCells: { range: gridRange(SP, 0, 1, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SP, 0, 1, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SP, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });

  // Row 2: guidance note — italic
  reqs.push({ mergeCells: { range: gridRange(SP, 1, 2, 0, 5), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SP, 1, 2, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.darkText) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SP, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 32 }, fields: 'pixelSize' } });

  // Section header rows (4,10,16,22,28,33,38) = idx 3,9,15,21,27,32,37
  const sectionRows = [3, 9, 15, 21, 27, 32, 37];
  sectionRows.forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(SP, ri, ri+1, 0, 5), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(SP, ri, ri+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(C.taupe), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({ updateDimensionProperties: { range: { sheetId: SP, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  });

  // Label col A for input sections
  const labelRanges = [
    [4, 8], [10, 14], [28, 31], [33, 36],
  ];
  labelRanges.forEach(([r1, r2]) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SP, r1, r2+1, 0, 1),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.darkText) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    // Input cells col B+
    reqs.push({
      repeatCell: {
        range: gridRange(SP, r1, r2+1, 1, 4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { fontSize: 10 }, wrapStrategy: 'WRAP' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)',
      },
    });
  });

  // Table header rows for People and Care Team
  [16, 22].forEach(ri => {
    reqs.push({
      repeatCell: {
        range: gridRange(SP, ri, ri+1, 0, 3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  });
  // People input rows 18-20 (idx 17-19)
  reqs.push({
    repeatCell: {
      range: gridRange(SP, 17, 20, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  // Care Team input rows 24-26 (idx 23-25)
  reqs.push({
    repeatCell: {
      range: gridRange(SP, 23, 26, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  // Provider role col B — pre-filled but visible
  reqs.push({
    repeatCell: {
      range: gridRange(SP, 23, 26, 1, 2),
      cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 9 } } },
      fields: 'userEnteredFormat(textFormat)',
    },
  });

  // Crisis Resources section rows 39-42 (idx 38-41) — bold, permanently styled
  reqs.push({
    repeatCell: {
      range: gridRange(SP, 38, 42, 0, 1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(SP, 38, 42, 1, 4),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { foregroundColor: hex(C.darkText), bold: true, fontSize: 10 }, wrapStrategy: 'WRAP' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SP, dimension: 'ROWS', startIndex: 38, endIndex: 42 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });

  // Column widths
  [[0,180],[1,380],[2,150],[3,150]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SP, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'sp-format');
  console.log('Safety Plan & Crisis Resources complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
