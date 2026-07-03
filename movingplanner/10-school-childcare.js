'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SC = sheetMap['🏫 School & Childcare'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  values.push({ range: "'🏫 School & Childcare'!A1", values: [['🏫 All-in-One Moving Planner — School & Childcare']] });

  // ── SCHOOLS section header (row 2) ────────────────────────────────────────
  values.push({ range: "'🏫 School & Childcare'!A2", values: [['SCHOOLS']] });

  // ── Schools column headers (row 3) ───────────────────────────────────────
  values.push({ range: "'🏫 School & Childcare'!A3", values: [[
    '#', 'School Name', 'Type', 'Year Level(s)', 'Address', 'Phone',
    'Status', 'Open Day', 'Application Deadline', 'Notes',
  ]]});

  // ── 8 sample schools ─────────────────────────────────────────────────────
  const schools = [
    [1, 'Riverside Primary School',     'Public',    'K–6',  '22 River Rd, Riverside IL 60546', '(708) 555-0101', 'Research',   '',           '',           'Zoned school for new address'],
    [2, 'Oakmont Elementary School',    'Public',    'K–5',  '14 Oak St, Riverside IL 60546',   '(708) 555-0102', 'Contacted',  '2025-07-20', '2025-08-01', 'Tour booked'],
    [3, 'St. Francis of Assisi School', 'Religious', 'K–8',  '7 Francis Ave, Riverside IL',     '(708) 555-0103', 'Research',   '',           '',           'Catholic school — waiting list expected'],
    [4, 'Riverside Montessori',         'Montessori','3–12', '55 Pine Blvd, Riverside IL',      '(708) 555-0104', 'Research',   '2025-08-05', '',           'Preschool through middle school'],
    [5, 'West Riverside Middle School',  'Public',   '6–8',  '100 West St, Riverside IL 60546', '(708) 555-0105', 'Not Started','',           '',           'For Yr 6+ child if applicable'],
    [6, 'Riverside High School',        'Public',    '9–12', '200 Main Dr, Riverside IL 60546', '(708) 555-0106', 'Not Started','',           '',           ''],
    [7, 'Lincoln Charter School',       'Charter',   'K–8',  '38 Lincoln Ave, Riverside IL',    '(708) 555-0107', 'Research',   '',           '',           'Application lottery in Feb'],
    [8, 'Heritage International School','International','K–12','90 Global Way, Riverside IL',   '(708) 555-0108', 'Research',   '',           '',           'IB program'],
  ];

  const schoolRows = schools.map(s => [s[0], s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8], s[9]]);
  values.push({ range: "'🏫 School & Childcare'!A4", values: schoolRows });

  // ── CHILDCARE section header (row 13) ────────────────────────────────────
  values.push({ range: "'🏫 School & Childcare'!A13", values: [['CHILDCARE & EARLY EDUCATION']] });

  // ── Childcare column headers (row 14) ────────────────────────────────────
  values.push({ range: "'🏫 School & Childcare'!A14", values: [[
    '#', 'Centre Name', 'Type', 'Ages', 'Address', 'Phone',
    'Status', 'Spots Available?', 'Weekly Cost ($)', 'Notes',
  ]]});

  // ── 6 sample childcare entries ───────────────────────────────────────────
  const childcare = [
    [1, 'Riverside Early Learning Centre', 'Long Day Care',   '6wk–5yr', '12 Elm St, Riverside IL', '(708) 555-0201', 'Contacted',  true,  320, 'Waitlist — 3 months est.'],
    [2, 'Little Acorns Family Day Care',   'Family Day Care', '6wk–5yr', 'Various, Riverside IL',   '(708) 555-0202', 'Research',   false, 220, 'Home-based care'],
    [3, 'Oakmont Preschool',               'Preschool',       '3–5yr',   '14 Oak St, Riverside IL', '(708) 555-0203', 'Contacted',  true,  160, '3 days/week program'],
    [4, 'St. Francis Early Childhood',     'Preschool',       '3–5yr',   '7 Francis Ave, Riverside',  '(708) 555-0204', 'Research', false, 180, 'Attached to St. Francis School'],
    [5, 'YMCA After School Care',          'After School Care','5–12yr', '55 Centre Dr, Riverside IL','(708) 555-0205', 'Not Started',false, 85,  'For school-age children'],
    [6, 'Occasional Care — Riverside CC',  'Occasional Care', '1–5yr',   '22 Civic Ave, Riverside IL','(708) 555-0206', 'Research', true,  35,  'Casual / drop-in basis'],
  ];

  const childcareRows = childcare.map(c => [c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7], c[8], c[9]]);
  values.push({ range: "'🏫 School & Childcare'!A15", values: childcareRows });

  await valuesBatchUpdate(id, values, 'school-values');

  // ── Merge title row 1 ────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SC, 0, 1, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SC, 0, 1, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.warmBrown),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Schools section header row 2 ─────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SC, 1, 2, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SC, 1, 2, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.antiqueBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Schools column headers row 3 ─────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(SC, 2, 3, 0, 10),
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

  // ── School data rows 4–11 ─────────────────────────────────────────────────
  for (let r = 0; r < 8; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(SC, 3 + r, 4 + r, 0, 10),
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

  // ── Childcare section header row 13 ──────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(SC, 12, 13, 0, 10), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SC, 12, 13, 0, 10),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.tealGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 12 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Childcare column headers row 14 ──────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(SC, 13, 14, 0, 10),
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

  // ── Childcare data rows 15–20 ─────────────────────────────────────────────
  for (let r = 0; r < 6; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(SC, 14 + r, 15 + r, 0, 10),
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

  // Currency for weekly cost column (I = index 8)
  reqs.push({
    repeatCell: {
      range: gridRange(SC, 14, 20, 8, 9),
      cell: {
        userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' } },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [35, 190, 100, 80, 200, 110, 90, 100, 90, 200];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: SC, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SC, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
      properties: { pixelSize: 36 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: SC, dimension: 'ROWS', startIndex: 2, endIndex: 25 },
      properties: { pixelSize: 24 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze rows 1–3 ───────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: SC, gridProperties: { frozenRowCount: 3 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'school-format');
  console.log('School & Childcare complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
