'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard — Annual Overview'];
const DSL_NAME = "'📅 Daily Symptom Log'";
const MED_NAME = "'💊 Medication & Treatment Tracker'";
const SLP_NAME = "'😴 Sleep Tracker'";

(async () => {
  const data = [];

  // Row 1: title
  data.push({ range: "'📊 Dashboard — Annual Overview'!A1", values: [['BIPOLAR SYMPTOM TRACKER — ANNUAL OVERVIEW']] });

  // Row 2: input label
  data.push({ range: "'📊 Dashboard — Annual Overview'!A2:B2", values: [['Tracking Year Start Date', '2025-01-01']] });

  // Row 3: section header
  data.push({ range: "'📊 Dashboard — Annual Overview'!A3", values: [['📊 KEY STATISTICS']] });

  // KPI rows 5-12
  data.push({ range: "'📊 Dashboard — Annual Overview'!A5:D5", values: [[
    'Avg Overall Severity',
    `=IFERROR(ROUND(AVERAGE(${DSL_NAME}!B2:B400),1)&"%","—")`,
    'Days Logged This Year',
    `=IFERROR(COUNTA(${DSL_NAME}!A2:A400),"—")`,
  ]] });
  data.push({ range: "'📊 Dashboard — Annual Overview'!A7:D7", values: [[
    'Medication Adherence %',
    `=IFERROR(TEXT(COUNTIF(${MED_NAME}!E2:E400,TRUE)/COUNTA(${MED_NAME}!A2:A400),"0%"),"—")`,
    'Avg Sleep Hours',
    `=IFERROR(ROUND(AVERAGE(${SLP_NAME}!B2:B400),1)&" hrs","—")`,
  ]] });
  data.push({ range: "'📊 Dashboard — Annual Overview'!A9:D9", values: [[
    'Stable/Euthymic Days',
    `=IFERROR(COUNTIF(${DSL_NAME}!C2:C400,"Stable/Euthymic"),"—")`,
    'Manic/Hypomanic Days',
    `=IFERROR(COUNTIF(${DSL_NAME}!C2:C400,"Manic")+COUNTIF(${DSL_NAME}!C2:C400,"Hypomanic"),"—")`,
  ]] });
  data.push({ range: "'📊 Dashboard — Annual Overview'!A11:D11", values: [[
    'Depressive Days',
    `=IFERROR(COUNTIF(${DSL_NAME}!C2:C400,"Depressive"),"—")`,
    'Therapy Appointments',
    `=IFERROR(COUNTA('🗓️ Therapy & Appointment Log'!A2:A400),"—")`,
  ]] });

  // Disclaimer box rows 13-15
  const disclaimer = 'This tracker is a personal organization and symptom-monitoring tool only. It is not a diagnostic tool and does not provide medical or psychiatric advice, and is not a substitute for professional care. Share this data with your psychiatrist, therapist, or care team as part of your treatment. If you are in crisis or having thoughts of self-harm, call or text 988 (Suicide & Crisis Lifeline, available 24/7 in the US) or your local emergency number immediately.';
  data.push({ range: "'📊 Dashboard — Annual Overview'!A13", values: [[disclaimer]] });

  // Monthly helper data for charts — rows 17+
  data.push({ range: "'📊 Dashboard — Annual Overview'!A16", values: [['📈 MONTHLY DATA — helper for charts']] });
  data.push({ range: "'📊 Dashboard — Annual Overview'!A17:E17", values: [['Month','Avg Severity (All)','Avg Severity (Manic)','Avg Severity (Depressive)','Avg Sleep Hrs']] });

  const months = ['Jan 2025','Feb 2025','Mar 2025','Apr 2025','May 2025','Jun 2025',
                  'Jul 2025','Aug 2025','Sep 2025','Oct 2025','Nov 2025','Dec 2025'];
  const monthStarts = [
    'DATE(2025,1,1)','DATE(2025,2,1)','DATE(2025,3,1)','DATE(2025,4,1)',
    'DATE(2025,5,1)','DATE(2025,6,1)','DATE(2025,7,1)','DATE(2025,8,1)',
    'DATE(2025,9,1)','DATE(2025,10,1)','DATE(2025,11,1)','DATE(2025,12,1)',
  ];
  const monthEnds = [
    'DATE(2025,2,1)','DATE(2025,3,1)','DATE(2025,4,1)','DATE(2025,5,1)',
    'DATE(2025,6,1)','DATE(2025,7,1)','DATE(2025,8,1)','DATE(2025,9,1)',
    'DATE(2025,10,1)','DATE(2025,11,1)','DATE(2025,12,1)','DATE(2026,1,1)',
  ];

  months.forEach((month, i) => {
    const r = 18 + i;
    const ms = monthStarts[i];
    const me = monthEnds[i];
    data.push({ range: `'📊 Dashboard — Annual Overview'!A${r}:E${r}`, values: [[
      month,
      `=IFERROR(AVERAGEIFS(${DSL_NAME}!$B:$B,${DSL_NAME}!$A:$A,">="&${ms},${DSL_NAME}!$A:$A,"<"&${me}),"")`,
      `=IFERROR(AVERAGEIFS(${DSL_NAME}!$B:$B,${DSL_NAME}!$A:$A,">="&${ms},${DSL_NAME}!$A:$A,"<"&${me},${DSL_NAME}!$C:$C,"Manic"),"")`,
      `=IFERROR(AVERAGEIFS(${DSL_NAME}!$B:$B,${DSL_NAME}!$A:$A,">="&${ms},${DSL_NAME}!$A:$A,"<"&${me},${DSL_NAME}!$C:$C,"Depressive"),"")`,
      `=IFERROR(AVERAGEIFS(${SLP_NAME}!$B:$B,${SLP_NAME}!$A:$A,">="&${ms},${SLP_NAME}!$A:$A,"<"&${me}),"")`,
    ]] });
  });

  // Medication adherence monthly helper rows 32+
  data.push({ range: "'📊 Dashboard — Annual Overview'!A32", values: [['💊 MONTHLY MEDICATION ADHERENCE — helper for Chart 3']] });
  data.push({ range: "'📊 Dashboard — Annual Overview'!A33:B33", values: [['Month','Adherence %']] });
  months.forEach((month, i) => {
    const r = 34 + i;
    const ms = monthStarts[i];
    const me = monthEnds[i];
    data.push({ range: `'📊 Dashboard — Annual Overview'!A${r}:B${r}`, values: [[
      month,
      `=IFERROR(COUNTIFS(${MED_NAME}!$A:$A,">="&${ms},${MED_NAME}!$A:$A,"<"&${me},${MED_NAME}!$E:$E,TRUE)/COUNTIFS(${MED_NAME}!$A:$A,">="&${ms},${MED_NAME}!$A:$A,"<"&${me}),"")`,
    ]] });
  });

  await valuesBatchUpdate(id, data, 'dash-values');

  const reqs = [];

  // Title banner
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 0, 1, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' } });

  // Row 2: input
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 1, 2, 0, 1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 1, 2, 1, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.periwinkle) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 1, 3), mergeType: 'MERGE_ALL' } });

  // Row 3: section header
  reqs.push({ mergeCells: { range: gridRange(DSH, 2, 3, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.taupe), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // KPI rows label + value pairs
  [4, 6, 8, 10].forEach(ri => {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 0, 1),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 1, 3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.periwinkle) }, verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 3, 4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 4, 6),
        cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { bold: true, fontSize: 13, foregroundColor: hex(C.periwinkle) }, verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  });

  // Disclaimer row 13 (idx 12)
  reqs.push({ mergeCells: { range: gridRange(DSH, 12, 14, 0, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 12, 14, 0, 7),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.darkText) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 12, endIndex: 14 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });

  // Helper section headers idx 15, 31
  [15, 31].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri+1, 0, 6), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 0, 6),
        cell: { userEnteredFormat: { backgroundColor: hex(C.periwinkle), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9, italic: true } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } });
  });

  // Helper header rows (idx 16, 32)
  [16, 32].forEach(ri => {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 0, 5),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  });

  // Helper data
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 17, 30, 0, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.fog), textFormat: { fontSize: 9 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 33, 46, 0, 2),
      cell: { userEnteredFormat: { backgroundColor: hex(C.fog), textFormat: { fontSize: 9 }, numberFormat: { type: 'PERCENT', pattern: '0%' } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)',
    },
  });

  [[0,180],[1,130],[2,120],[3,160],[4,130],[5,120]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'dash-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
