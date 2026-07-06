'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const IT = sheetMap['🗓️ Party Itinerary'];

// Layout:
// Row 1: Title (A1:I1 merged)
// Row 2: Summary / party date strip
// Row 3: Col headers
// Rows 4-17: 14 itinerary segments
// Cols: A=Start, B=End, C=Duration(formula), D=Activity, E=Activity Type, F=Host, G=Location/Notes, H=Overlap Check, I=Guests

const NCOLS = 9;

// [startTime, endTime, activity, type, host, locationNotes, guests]
// Times stored as text "HH:MM AM/PM" for USER_ENTERED
const SEGMENTS = [
  ['11:00 AM','11:30 AM','Guest Arrival & Welcome Drinks','Welcome','Emma Clarke','Front garden / welcome table','All'],
  ['11:30 AM','11:50 AM','Welcome & Introductions','Welcome','Emma Clarke & Lily Thompson','Living room','All'],
  ['11:50 AM','12:10 PM','Mum-to-Be Spotlight — memories & wishes','Activity','Lily Thompson','Living room','All'],
  ['12:10 PM','12:40 PM','Game 1 — Baby Bingo','Games','Emma Clarke','Living room','All'],
  ['12:40 PM','1:00 PM','Game 2 — Guess the Baby Food','Games','Mia Patel','Kitchen / living room','All'],
  ['1:00 PM','1:45 PM','Lunch — Afternoon Tea Served','Food & Drinks','Caterer / hosts','Dining area','All'],
  ['1:45 PM','2:00 PM','Cake Cutting & Desserts','Food & Drinks','Lily Thompson','Dining area','All'],
  ['2:00 PM','2:20 PM','Game 3 — Baby Trivia','Games','Zoe Roberts','Living room','All'],
  ['2:20 PM','2:40 PM','Game 4 — Baby Price is Right','Games','Amy Johnson','Living room','All'],
  ['2:40 PM','3:00 PM','Gift Opening — Part 1','Gift Opening','Sarah Mitchell','Living room','All'],
  ['3:00 PM','3:20 PM','Gift Opening — Part 2','Gift Opening','Sarah Mitchell','Living room','All'],
  ['3:20 PM','3:35 PM','Speeches & Toasts','Speeches','Emma Clarke','Living room','All'],
  ['3:35 PM','3:55 PM','Photo Time — group & candid shots','Photo','Lucy Chen','Garden / living room','All'],
  ['3:55 PM','4:30 PM','Farewell & Final Nibbles','Farewell','All hosts','Entrance / garden','All'],
];

(async () => {
  const reqs = [];

  // Column widths: A=80, B=80, C=80, D=240, E=120, F=140, G=200, H=80, I=100
  [80,80,80,240,120,140,200,80,100].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: IT, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: IT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: IT, dimension: 'ROWS', startIndex: 1, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: IT, dimension: 'ROWS', startIndex: 3, endIndex: 18 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Title row
  reqs.push({ mergeCells: { range: gridRange(IT, 0, 1, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(IT, 0, 1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Summary row 2
  reqs.push({ repeatCell: {
    range: gridRange(IT, 1, 2, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  [[0,2],[2,4],[4,6],[6,9]].forEach(([c1,c2]) => {
    reqs.push({ mergeCells: { range: gridRange(IT, 1, 2, c1, c2), mergeType: 'MERGE_ALL' }});
  });

  // Col headers row 3
  reqs.push({ repeatCell: {
    range: gridRange(IT, 2, 3, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows 4-17 (indices 3-16)
  for (let r = 3; r < 17; r++) {
    const bg = (r % 2 === 1) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(IT, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Time cols A, B — center, text
  [0, 1].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(IT, 3, 17, c, c + 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', numberFormat: { type: 'TEXT' } }},
      fields: 'userEnteredFormat(horizontalAlignment,numberFormat)',
    }});
  });

  // Duration col C — center
  reqs.push({ repeatCell: {
    range: gridRange(IT, 3, 17, 2, 3),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // Activity col D — wrap
  reqs.push({ repeatCell: {
    range: gridRange(IT, 3, 17, 3, 4),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' }},
    fields: 'userEnteredFormat.wrapStrategy',
  }});

  // Overlap check col H — center
  reqs.push({ repeatCell: {
    range: gridRange(IT, 3, 17, 7, 8),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // Notes col G — wrap
  reqs.push({ repeatCell: {
    range: gridRange(IT, 3, 17, 6, 7),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' }},
    fields: 'userEnteredFormat.wrapStrategy',
  }});

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(IT, 2, 17, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'itinerary-format');

  // Values
  const vals = [];
  vals.push({ range: "'🗓️ Party Itinerary'!A1", values: [['  🗓️  Party Itinerary']] });

  // Summary row 2
  vals.push({ range: "'🗓️ Party Itinerary'!A2", values: [['📅  Party Date: 20 July 2025']] });
  vals.push({ range: "'🗓️ Party Itinerary'!C2", values: [['⏰  11:00 AM – 4:30 PM']] });
  vals.push({ range: "'🗓️ Party Itinerary'!E2", values: [['📍  Home of Emma Clarke']] });
  vals.push({ range: "'🗓️ Party Itinerary'!G2", values: [['👥  20 guests · 14 segments · 5.5 hours']] });

  // Col headers
  vals.push({ range: "'🗓️ Party Itinerary'!A3", values: [['START','END','DURATION','ACTIVITY / SEGMENT','TYPE','HOST','LOCATION / NOTES','⚠️','GUESTS']] });

  // Segment rows — times as text, duration formula, overlap formula
  const rows = SEGMENTS.map(([start, end, activity, type, host, notes, guests], i) => {
    const row = 4 + i; // 1-indexed
    // Duration in mins: =(TIMEVALUE(B4)-TIMEVALUE(A4))*1440 — since times are text
    const dur = `=IFERROR(ROUND((TIMEVALUE(B${row})-TIMEVALUE(A${row}))*1440,0)&" min","")`;
    // Overlap check: flag if this segment's start < previous segment's end (row > 4 only)
    const overlap = row > 4
      ? `=IFERROR(IF(TIMEVALUE(A${row})<TIMEVALUE(B${row-1}),"⚠️ OVERLAP",""),"")`
      : '';
    return [start, end, dur, activity, type, host, notes, overlap, guests];
  });
  vals.push({ range: "'🗓️ Party Itinerary'!A4", values: rows });

  await valuesBatchUpdate(id, vals, 'itinerary-values');
  console.log('Party Itinerary complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
