'use strict';
const { C, hex, batchUpdate, valuesBatchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GST = sheetMap['👥 Guest List & RSVP'];

// Cols: A=# B=Name C=Email D=Phone E=Side F=RSVP G=Table(formula) H=Dietary I=Meal J=+1 K=+1Name L=Gift M=TY N=Notes
// 14 cols (A:N, indices 0-13)
// Row 0: standalone A + merged B:N (title) — required for col A freeze
// Row 1: standalone A + stats cells
// Row 2: headers
// Rows 3-32: 30 guests

const GUESTS = [
  ['Emma Thompson',    'emma.t@email.com',    '555-0101','Partner 1',  'Attending',     'None',         'Chicken','TRUE', 'James Thompson',  'Given – Registry', ''],
  ['James Wilson',     'james.w@email.com',   '555-0102','Partner 2',  'Attending',     'None',         'Beef',   'FALSE','',                 'Not Given',         ''],
  ['Olivia Chen',      'olivia.c@email.com',  '555-0103','Both',        'Attending',     'Vegan',        'Vegan',  'FALSE','',                 'Not Given',         ''],
  ['Noah Martinez',    'noah.m@email.com',    '555-0104','Partner 1',  'Attending',     'Gluten-Free',  'Fish',   'TRUE', 'Lily Martinez',   'Given – Cash',      ''],
  ['Sophia Kim',       'sophia.k@email.com',  '555-0105','Partner 2',  'Attending',     'None',         'Chicken','FALSE','',                 'Not Given',         ''],
  ['Liam Johnson',     'liam.j@email.com',    '555-0106','Partner 1',  'Attending',     'None',         'Beef',   'TRUE', 'Sarah Johnson',   'Given – Registry',  ''],
  ['Isabella Williams','isabella.w@email.com','555-0107','Both',        'Attending',     'Vegetarian',   'Vegetarian','FALSE','',             'Given – Other',     ''],
  ['Mason Davis',      'mason.d@email.com',   '555-0108','Partner 2',  'Attending',     'None',         'Chicken','FALSE','',                 'Not Given',         ''],
  ['Ava Rodriguez',    'ava.r@email.com',     '555-0109','Friend of Both','Attending',  'Dairy-Free',   'Fish',   'FALSE','',                 'Given – Cash',      ''],
  ['Ethan Lee',        'ethan.l@email.com',   '555-0110','Partner 1',  'Attending',     'None',         'Beef',   'FALSE','',                 'Not Given',         ''],
  ['Mia Anderson',     'mia.a@email.com',     '555-0111','Partner 2',  'Attending',     'None',         'Chicken','TRUE', 'Tom Anderson',    'Given – Registry',  ''],
  ['Alexander Taylor', 'alex.t@email.com',    '555-0112','Friend of Both','Attending',  'Nut Allergy',  'Beef',   'FALSE','',                 'Not Given',         ''],
  ['Charlotte Brown',  'charlotte.b@email.com','555-0113','Partner 1', 'Attending',     'None',         'Chicken','FALSE','',                 'Given – Other',     ''],
  ['Benjamin Garcia',  'ben.g@email.com',     '555-0114','Partner 2',  'Maybe',         'None',         '',       'FALSE','',                 'Not Given',         'Tentative'],
  ['Amelia Miller',    'amelia.m@email.com',  '555-0115','Friend of Both','Attending',  'Vegan',        'Vegan',  'FALSE','',                 'Given – Cash',      ''],
  ['Daniel Moore',     'daniel.m@email.com',  '555-0116','Partner 1',  'Attending',     'None',         'Beef',   'TRUE', 'Kate Moore',      'Not Given',         ''],
  ['Harper Jackson',   'harper.j@email.com',  '555-0117','Partner 2',  'Attending',     'None',         'Fish',   'FALSE','',                 'Given – Registry',  ''],
  ['Lucas White',      'lucas.w@email.com',   '555-0118','Both',        'Attending',     'None',         'Chicken','FALSE','',                 'Not Given',         ''],
  ['Evelyn Harris',    'evelyn.h@email.com',  '555-0119','Partner 1',  'Awaiting Reply','None',         '',       'FALSE','',                 'Not Given',         'Sent reminder'],
  ['Henry Thomas',     'henry.t@email.com',   '555-0120','Partner 2',  'Attending',     'None',         'Beef',   'FALSE','',                 'Given – Other',     ''],
  ['Abigail Martin',   'abby.m@email.com',    '555-0121','Friend of Both','Attending',  'None',         'Chicken','FALSE','',                 'Given – Cash',      ''],
  ['Sebastian Thompson','seb.t@email.com',    '555-0122','Partner 1',  'Declined',      'None',         '',       'FALSE','',                 'Not Given',         'Out of town'],
  ['Emily Clark',      'emily.c@email.com',   '555-0123','Partner 2',  'Attending',     'Halal',        'Chicken','FALSE','',                 'Not Given',         ''],
  ['Jack Lewis',       'jack.l@email.com',    '555-0124','Friend of Both','Attending',  'None',         'Beef',   'TRUE', 'Amy Lewis',       'Given – Registry',  ''],
  ['Elizabeth Robinson','liz.r@email.com',    '555-0125','Partner 1',  'Attending',     'None',         'Chicken','FALSE','',                 'Given – Other',     ''],
  ['Owen Walker',      'owen.w@email.com',    '555-0126','Partner 2',  'Awaiting Reply','None',         '',       'FALSE','',                 'Not Given',         ''],
  ['Sofia Hall',       'sofia.h@email.com',   '555-0127','Both',        'Attending',     'None',         'Fish',   'FALSE','',                 'Given – Cash',      ''],
  ['Aiden Young',      'aiden.y@email.com',   '555-0128','Partner 1',  'Attending',     'None',         'Beef',   'FALSE','',                 'Not Given',         ''],
  ['Chloe Allen',      'chloe.a@email.com',   '555-0129','Friend of Both','Declined',   'None',         '',       'FALSE','',                 'Not Given',         ''],
  ['Michael Wright',   'michael.w@email.com', '555-0130','Partner 2',  'Attending',     'None',         'Chicken','FALSE','',                 'Given – Registry',  ''],
];

function fmtCell(bg, textColor, bold, size, align, italic) {
  return {
    backgroundColor: hex(bg),
    textFormat: { foregroundColor: hex(textColor), bold: !!bold, fontSize: size || 10, italic: !!italic },
    horizontalAlignment: align || 'LEFT',
    verticalAlignment: 'MIDDLE',
  };
}

(async () => {
  const reqs = [];

  // ── Row 0: Title banner ──
  // A0 standalone (emoji/icon cell)
  reqs.push({ repeatCell: {
    range: gridRange(GST, 0, 1, 0, 1),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 16, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // B0:N0 merged
  reqs.push({ mergeCells: { range: gridRange(GST, 0, 1, 1, 14), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GST, 0, 1, 1, 14),
    cell: { userEnteredFormat: fmtCell(C.deepPlum, C.white, true, 15, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Row 1: Stats bar (standalone A) ──
  reqs.push({ repeatCell: {
    range: gridRange(GST, 1, 2, 0, 1),
    cell: { userEnteredFormat: fmtCell(C.roseGold, C.white, true, 9, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  // Stats cells B-N (pairs: label + value merged)
  const statsFmt = { backgroundColor: hex(C.roseGold), textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' };
  reqs.push({ repeatCell: {
    range: gridRange(GST, 1, 2, 1, 14),
    cell: { userEnteredFormat: statsFmt },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Row 2: Column headers ──
  reqs.push({ repeatCell: {
    range: gridRange(GST, 2, 3, 0, 14),
    cell: { userEnteredFormat: fmtCell(C.mutedRust, C.white, true, 10, 'CENTER') },
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // ── Data rows 3-32 ──
  for (let r = 0; r < 30; r++) {
    const bg = r % 2 === 0 ? C.ivoryCream : C.altRow;
    reqs.push({ repeatCell: {
      range: gridRange(GST, r + 3, r + 4, 0, 14),
      cell: { userEnteredFormat: fmtCell(bg, C.bodyText, false, 10, 'LEFT') },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    }});
  }

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GST, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Column widths
  const colWidths = [40, 180, 180, 110, 100, 120, 80, 130, 100, 60, 150, 130, 80, 150];
  colWidths.forEach((w, ci) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GST, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Border for data range
  const borderStyle = { style: 'SOLID', color: hex(C.border) };
  reqs.push({ updateBorders: {
    range: gridRange(GST, 2, 33, 0, 14),
    top: borderStyle, bottom: borderStyle, left: borderStyle, right: borderStyle,
    innerHorizontal: borderStyle, innerVertical: borderStyle,
  }});

  await batchUpdate(id, reqs, 'guests-format');

  // ── Values ──
  const data = [];

  // Row 0 values
  data.push({ range: `'👥 Guest List & RSVP'!A1`, values: [['👥']] });
  data.push({ range: `'👥 Guest List & RSVP'!B1`, values: [['GUEST LIST & RSVP']] });

  // Row 1 stats
  data.push({ range: `'👥 Guest List & RSVP'!A2`, values: [['STATS']] });
  data.push({ range: `'👥 Guest List & RSVP'!B2`, values: [['Total Guests:']] });
  data.push({ range: `'👥 Guest List & RSVP'!C2`, values: [[`=COUNTA('👥 Guest List & RSVP'!B4:B200)`]] });
  data.push({ range: `'👥 Guest List & RSVP'!E2`, values: [['Attending:']] });
  data.push({ range: `'👥 Guest List & RSVP'!F2`, values: [[`=COUNTIF('👥 Guest List & RSVP'!F4:F200,"Attending")`]] });
  data.push({ range: `'👥 Guest List & RSVP'!H2`, values: [['Declined:']] });
  data.push({ range: `'👥 Guest List & RSVP'!I2`, values: [[`=COUNTIF('👥 Guest List & RSVP'!F4:F200,"Declined")`]] });
  data.push({ range: `'👥 Guest List & RSVP'!K2`, values: [['Awaiting:']] });
  data.push({ range: `'👥 Guest List & RSVP'!L2`, values: [[`=COUNTIF('👥 Guest List & RSVP'!F4:F200,"Awaiting Reply")`]] });
  data.push({ range: `'👥 Guest List & RSVP'!N2`, values: [[`=COUNTIF('👥 Guest List & RSVP'!J4:J200,TRUE)&" +1s"`]] });

  // Row 2 headers
  data.push({
    range: `'👥 Guest List & RSVP'!A3:N3`,
    values: [['#','Guest Name','Email','Phone','Side','RSVP Status','Table #','Dietary Needs','Meal Choice','+1?','+1 Name','Gift Status','TY Sent','Notes']],
  });

  // Data rows 3-32 (A4:N33 in A1)
  const rows = GUESTS.map((g, i) => {
    const row = i + 4; // 1-indexed row in sheet
    const tableFormula = `=IFERROR(INDEX('🪑 Seating Chart'!$A$4:$A$11, MATCH(1, MMULT(('🪑 Seating Chart'!$F$4:$O$11=B${row})*1, SEQUENCE(10,1,1,0)), 0)), "")`;
    return [i + 1, g[0], g[1], g[2], g[3], g[4], tableFormula, g[5], g[6], g[7], g[8], g[9], 'FALSE', g[10]];
  });

  data.push({ range: `'👥 Guest List & RSVP'!A4:N33`, values: rows });

  await valuesBatchUpdate(id, data, 'guests-values');
  console.log('Guest List & RSVP complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
