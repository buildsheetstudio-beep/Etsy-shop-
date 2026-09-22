'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const UT = sheetMap['⚡ Utilities & Services'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  values.push({ range: "'⚡ Utilities & Services'!A1", values: [['⚡ All-in-One Moving Planner — Utilities & Services']] });

  // ── Summary block ─────────────────────────────────────────────────────────
  values.push({ range: "'⚡ Utilities & Services'!A2", values: [[
    'Old Home — Disconnected', 'New Home — Connected', 'Pending Disconnects', 'Pending Connections',
  ]]});
  values.push({ range: "'⚡ Utilities & Services'!A3", values: [[
    '=COUNTIF(G6:G25,"Disconnected")',
    '=COUNTIF(K6:K25,"Connected")',
    '=COUNTIF(G6:G25,"Active")',
    '=COUNTIF(K6:K25,"Not Started")+COUNTIF(K6:K25,"Applied")',
  ]]});

  // ── Column headers (row 5) ────────────────────────────────────────────────
  values.push({ range: "'⚡ Utilities & Services'!A5", values: [[
    '#', 'Utility Service', 'Provider (Old)', 'Account # (Old)', 'Disconnect Date',
    'Final Reading', 'Old Status',
    'Provider (New)', 'Account # (New)', 'Connect Date',
    'New Status', 'Notes',
  ]]});

  // ── 10 pre-populated utility rows ────────────────────────────────────────
  const utilities = [
    [1,  'Electricity',       'AGL',            'AGL-987654',   '2025-08-30', '',  'Active', 'Origin Energy',    '',              '2025-08-30', 'Not Started', 'Meter read on move day'],
    [2,  'Gas',               'AGL',            'AGL-987655',   '2025-08-30', '',  'Active', 'Origin Energy',    '',              '2025-08-30', 'Not Started', 'Meter read on move day'],
    [3,  'Water',             'SA Water',       'SAW-112233',   '2025-08-30', '',  'Active', 'South East Water',  '',             '2025-08-30', 'Not Started', 'Contact council'],
    [4,  'Internet',          'Aussie BB',      'AUSBB-445566', '2025-09-01', '',  'Active', 'Aussie Broadband', '',              '2025-08-30', 'Not Started', 'NBN port may need setup'],
    [5,  'Mobile',            'Telstra',        'TEL-778899',   '',           '',  'Active', 'Telstra',          '',              '',           'Active',      'No change — SIM only'],
    [6,  'Home Insurance',    'NRMA',           'NRMA-334455',  '2025-08-30', '',  'Active', 'NRMA',             '',              '2025-08-30', 'Not Started', 'Update address on policy'],
    [7,  'Contents Insurance','NRMA',           'NRMA-334456',  '2025-08-30', '',  'Active', 'NRMA',             '',              '2025-08-30', 'Not Started', 'Update cover level'],
    [8,  'Council Tax',       'Springfield CC', 'SCC-001122',   '2025-08-30', '',  'Active', 'Riverside CC',     '',             '2025-08-30', 'Not Started', 'Notify both councils'],
    [9,  'Waste Collection',  'Springfield CC', 'SCC-001123',   '2025-08-30', '',  'Active', 'Riverside CC',     '',             '2025-08-30', 'Not Started', 'Arrange bin swap'],
    [10, 'Solar',             'SolarEdge',      'SE-667788',    '2025-08-30', '',  'Active', 'N/A',              'N/A',           '',           'N/A',         'System stays with property'],
  ];

  const utilRows = utilities.map(u => [
    u[0], u[1], u[2], u[3], u[4], u[5], u[6], u[7], u[8], u[9], u[10], u[11],
  ]);
  values.push({ range: "'⚡ Utilities & Services'!A6", values: utilRows });

  await valuesBatchUpdate(id, values, 'utilities-values');

  // ── Merge title ───────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(UT, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(UT, 0, 1, 0, 12),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.tealGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Summary header row 2 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(UT, 1, 2, 0, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.antiqueBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Summary value row 3 ───────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(UT, 2, 3, 0, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.ivory),
          textFormat: { foregroundColor: hex(C.nearBlack), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Group headers: Old Home (cols C–G) / New Home (cols H–K) ─────────────
  reqs.push({ mergeCells: { range: gridRange(UT, 4, 5, 2, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(UT, 4, 5, 7, 11), mergeType: 'MERGE_ALL' } });

  // ── Column headers row 5 (sub-headers row 6) ─────────────────────────────
  // Actually, let's do headers on row 5, data starts row 6
  reqs.push({
    repeatCell: {
      range: gridRange(UT, 4, 5, 0, 12),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Zebra stripe data rows ────────────────────────────────────────────────
  for (let r = 0; r < 10; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(UT, 5 + r, 6 + r, 0, 12),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { foregroundColor: hex(C.nearBlack), fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [35, 110, 110, 110, 100, 100, 90, 110, 110, 100, 90, 180];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: UT, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: UT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 40 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: UT, dimension: 'ROWS', startIndex: 1, endIndex: 5 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: UT, dimension: 'ROWS', startIndex: 5, endIndex: 15 },
      properties: { pixelSize: 24 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze rows 1–5 ───────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: UT, gridProperties: { frozenRowCount: 5 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'utilities-format');
  console.log('Utilities & Services complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
