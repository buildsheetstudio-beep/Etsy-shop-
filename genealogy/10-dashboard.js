'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const DSH = sheetMap['📊 Dashboard'];
const FM_NAME = "'👨‍👩‍👧‍👦 Family Members'";
const DNA_NAME = "'🧬 DNA & Ethnicity Estimate'";
const MIG_NAME = "'🚢 Immigration & Migration Tracker'";

(async () => {
  const data = [];

  // ── Row 1: Title banner ───────────────────────────────────────────────
  data.push({ range: "'📊 Dashboard'!A1", values: [['🌳 FAMILY TREE DASHBOARD']] });

  // ── Row 2: spacer label
  data.push({ range: "'📊 Dashboard'!A2", values: [['Family Name']] });
  data.push({ range: "'📊 Dashboard'!B2", values: [['Sullivan-Kowalski-Nakamura']] });

  // ── Row 3: header separator
  data.push({ range: "'📊 Dashboard'!A3", values: [['📊 KEY STATISTICS']] });

  // ── KPI section rows 5-12 ─────────────────────────────────────────────
  // Col A = label, Col B = formula/value, Col C = label2, Col D = formula2
  data.push({ range: "'📊 Dashboard'!A5:D5", values: [['Total Ancestors', `=IFERROR(COUNTA(${FM_NAME}!A2:A300),"")`, 'Generations Tracked', '=IFERROR(4,"")']] });
  data.push({ range: "'📊 Dashboard'!A7:D7", values: [[
    'Earliest Ancestor',
    `=IFERROR(INDEX(${FM_NAME}!B:B,MATCH(MIN(${FM_NAME}!D2:D300),${FM_NAME}!D:D,0))&" ("&TEXT(MIN(${FM_NAME}!D2:D300),"YYYY")&")","")`,
    'Most Common Birth Country',
    `=IFERROR(INDEX(${FM_NAME}!E2:E300,MATCH(MAX(COUNTIF(${FM_NAME}!E2:E300,${FM_NAME}!E2:E300)),COUNTIF(${FM_NAME}!E2:E300,${FM_NAME}!E2:E300),0)),"")`,
  ]] });
  data.push({ range: "'📊 Dashboard'!A9:D9", values: [[
    'DNA Sources Tested',
    `=IFERROR(COUNTA(UNIQUE(${DNA_NAME}!D2:D100))-COUNTIF(${DNA_NAME}!D2:D100,""),"")`,
    'Immigration Events',
    `=IFERROR(COUNTA(${MIG_NAME}!A2:A100),"")`,
  ]] });
  data.push({ range: "'📊 Dashboard'!A11:D11", values: [[
    'Sources Cited (Confirmed)',
    `=IFERROR(COUNTIF('📜 Source & Citation Tracker'!E2:E100,"Confirmed"),"")`,
    'Sources (Uncertain — needs research)',
    `=IFERROR(COUNTIF('📜 Source & Citation Tracker'!E2:E100,"Uncertain"),"")`,
  ]] });

  // ── Physician-equivalent note row 13 ─────────────────────────────────
  data.push({ range: "'📊 Dashboard'!A13", values: [['This tracker is for personal genealogy research and record-keeping only. It does not constitute legal proof of identity, citizenship, or family relationship — consult official records and legal advisors for official documentation purposes.']] });

  // ── Section headers rows 15, 16 ─────────────────────────────────────
  data.push({ range: "'📊 Dashboard'!A15", values: [['📊 ANCESTORS BY ERA OF BIRTH — helper data for Chart 1']] });
  data.push({ range: "'📊 Dashboard'!A16:B16", values: [['Era (50-yr bracket)','Count']] });

  // Era buckets: <1800, 1800-49, 1850-99, 1900-49, 1950+
  const eraBuckets = [
    ['Before 1800', `=IFERROR(COUNTIFS(${FM_NAME}!D2:D300,"<"&DATE(1800,1,1)),"")`],
    ['1800–1849',   `=IFERROR(COUNTIFS(${FM_NAME}!D2:D300,">="&DATE(1800,1,1),${FM_NAME}!D2:D300,"<"&DATE(1850,1,1)),"")`],
    ['1850–1899',   `=IFERROR(COUNTIFS(${FM_NAME}!D2:D300,">="&DATE(1850,1,1),${FM_NAME}!D2:D300,"<"&DATE(1900,1,1)),"")`],
    ['1900–1949',   `=IFERROR(COUNTIFS(${FM_NAME}!D2:D300,">="&DATE(1900,1,1),${FM_NAME}!D2:D300,"<"&DATE(1950,1,1)),"")`],
    ['1950–1999',   `=IFERROR(COUNTIFS(${FM_NAME}!D2:D300,">="&DATE(1950,1,1),${FM_NAME}!D2:D300,"<"&DATE(2000,1,1)),"")`],
    ['2000+',       `=IFERROR(COUNTIFS(${FM_NAME}!D2:D300,">="&DATE(2000,1,1)),"")`],
  ];
  eraBuckets.forEach(([label, formula], i) => {
    data.push({ range: `'📊 Dashboard'!A${17+i}:B${17+i}`, values: [[label, formula]] });
  });

  // ── Birth Location helper rows 25-34 (for donut chart 2) ──────────────
  data.push({ range: "'📊 Dashboard'!A25", values: [['📊 ANCESTORS BY BIRTH COUNTRY — helper data for Chart 2']] });
  data.push({ range: "'📊 Dashboard'!A26:B26", values: [['Birth Country','Count']] });
  const countries = ['Ireland','Poland','Japan','Germany','USA','Other'];
  countries.forEach((country, i) => {
    let formula;
    if (country === 'Other') {
      const knownTotal = countries.slice(0,-1).map(c => `COUNTIF(${FM_NAME}!E2:E300,"*${c}*")`).join('+');
      formula = `=IFERROR(COUNTA(${FM_NAME}!A2:A300)-(${knownTotal}),"")`;
    } else {
      formula = `=IFERROR(COUNTIF(${FM_NAME}!E2:E300,"*${country}*"),"")`;
    }
    data.push({ range: `'📊 Dashboard'!A${27+i}:B${27+i}`, values: [[country, formula]] });
  });

  // ── DNA Ethnicity helper rows 35-46 (for donut chart 3) ───────────────
  data.push({ range: "'📊 Dashboard'!A35", values: [['🧬 DNA ETHNICITY BREAKDOWN — helper data for Chart 3']] });
  data.push({ range: "'📊 Dashboard'!A36:B36", values: [['Region','% of DNA']] });
  // Pull all distinct regions from DNA tab with their percentages
  const regions = ['British Isles','Western Europe','Eastern Europe','Scandinavia','East Asia','Southern Europe','Native American/Indigenous','Sub-Saharan Africa','Other'];
  regions.forEach((region, i) => {
    data.push({ range: `'📊 Dashboard'!A${37+i}:B${37+i}`, values: [
      [region, `=IFERROR(SUMIF(${DNA_NAME}!B2:B100,A${37+i},${DNA_NAME}!C2:C100),"")`]
    ] });
  });

  await valuesBatchUpdate(id, data, 'dash-values');

  const reqs = [];

  // Title banner A1 (merge cols A-H)
  reqs.push({ mergeCells: { range: gridRange(DSH, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 0, 1, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 18 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });

  // Row 2: Family Name input
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 1, 2, 0, 1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10 }, horizontalAlignment: 'RIGHT' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 1, 2, 1, 5),
      cell: { userEnteredFormat: { backgroundColor: hex(C.inputBg), textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.sepia) } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({ mergeCells: { range: gridRange(DSH, 1, 2, 1, 5), mergeType: 'MERGE_ALL' } });

  // Row 3: Section header
  reqs.push({ mergeCells: { range: gridRange(DSH, 2, 3, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 2, 3, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.gold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });

  // KPI label cells (odd rows 5,7,9,11 — 0-idx: 4,6,8,10) col A,C
  [4,6,8,10].forEach(ri => {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 0, 1),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.darkText) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 1, 3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { fontSize: 11, foregroundColor: hex(C.sepia), bold: true }, verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 3, 4),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.darkText) }, horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
      },
    });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 4, 6),
        cell: { userEnteredFormat: { backgroundColor: hex(C.formulaBg), textFormat: { fontSize: 11, foregroundColor: hex(C.sepia), bold: true }, verticalAlignment: 'MIDDLE' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  });

  // Row 13 disclaimer — merged, italic
  reqs.push({ mergeCells: { range: gridRange(DSH, 12, 13, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 12, 13, 0, 8),
      cell: { userEnteredFormat: { backgroundColor: hex(C.parchment), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.darkText) }, wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,wrapStrategy,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: 12, endIndex: 13 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } });

  // Helper section headers (rows 15, 25, 35) — idx 14, 24, 34
  [14, 24, 34].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(DSH, ri, ri+1, 0, 6), mergeType: 'MERGE_ALL' } });
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 0, 6),
        cell: { userEnteredFormat: { backgroundColor: hex(C.sepia), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9, italic: true }, horizontalAlignment: 'LEFT' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'ROWS', startIndex: ri, endIndex: ri+1 }, properties: { pixelSize: 22 }, fields: 'pixelSize' } });
  });

  // Helper table headers (rows 16, 26, 36) — idx 15, 25, 35
  [15, 25, 35].forEach(ri => {
    reqs.push({
      repeatCell: {
        range: gridRange(DSH, ri, ri+1, 0, 3),
        cell: { userEnteredFormat: { backgroundColor: hex(C.altRow), textFormat: { bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER' } },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
      },
    });
  });

  // Helper data cells
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 16, 24, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.parchment), textFormat: { fontSize: 9 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 26, 34, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.parchment), textFormat: { fontSize: 9 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(DSH, 36, 46, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: hex(C.parchment), textFormat: { fontSize: 9 }, numberFormat: { type: 'PERCENT', pattern: '0%' } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat)',
    },
  });

  // Column widths
  [[0,180],[1,160],[2,160],[3,200],[4,180],[5,160]].forEach(([ci,w]) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: DSH, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  await batchUpdate(id, reqs, 'dash-format');
  console.log('Dashboard complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
