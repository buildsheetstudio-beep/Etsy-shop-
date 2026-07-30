'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RP = sheetMap['Reception Planner'];
const S  = "'Reception Planner'";

const HEADERS = [
  '#', 'Start Time', 'End Time', 'Duration (auto)', 'Segment Name',
  'Description', 'Responsible Person', 'Vendor Involved',
  'Location / Zone', 'AV / Mic Needed', 'Music Cue', 'Status',
  'Announce to Guests', 'Time Conflict (auto)', 'Notes', 'Completed',
];

// 28 timeline segments: Ceremony → Cocktail Hour → Reception → Sendoff
const segments = [
  // Pre-ceremony prep
  ['12:00 PM', '2:00 PM',  'Venue Setup',           'Vendor setup: florist, DJ, rentals, lighting', 'Sophia Rivera (Coordinator)', 'All Vendors', 'Both Venues', false, '',                   'Confirmed', false, ''],
  ['2:00 PM',  '3:30 PM',  'Hair & Makeup — Final', 'Final touches for bride + bridesmaids', 'Simone Nwosu (Glam)', 'Luxe Bridal Beauty', 'Bridal Suite', false, '', 'Confirmed', false, ''],
  ['2:30 PM',  '3:00 PM',  'Bridal Party Photos',   'Pre-ceremony portraits at garden', 'Isaiah Torres (Photo)', 'Golden Hour Studio', 'Garden', false, 'Soft instrumental', 'Confirmed', false, ''],
  // Ceremony
  ['3:30 PM',  '3:50 PM',  'Guest Arrival — Ceremony', 'Ushers seat guests; music plays softly', 'Noel Parker (Best Man)', 'Vibe Events DJ', 'Garden Ceremony', false, 'Prelude — Boccherini', 'Confirmed', true, ''],
  ['3:50 PM',  '3:55 PM',  'Processional — Parents', 'Seating of parents of bride and groom', 'Usher Team', 'Vibe Events DJ', 'Garden Ceremony', false, 'Canon in D', 'Confirmed', true, ''],
  ['3:55 PM',  '4:00 PM',  'Processional — Wedding Party', 'Bridesmaids & groomsmen walk in', 'Wedding Party', 'Vibe Events DJ', 'Garden Ceremony', false, 'A Thousand Years', 'Confirmed', true, ''],
  ['4:00 PM',  '4:02 PM',  'Bridal Entrance',       'Bride walks with father', 'Kwame Osei (Father)', 'Vibe Events DJ', 'Garden Ceremony', false, 'Here Comes the Sun', 'Confirmed', true, ''],
  ['4:02 PM',  '4:30 PM',  'Ceremony — Vows & Rings', 'Readings, vows, ring exchange', 'Rev. David Nguyen', 'LifeFrame Films', 'Garden Ceremony', true, 'Silence during ceremony', 'Confirmed', true, ''],
  ['4:30 PM',  '4:32 PM',  'Ceremony Recessional',  'Mr. & Mrs. Thornton-Osei exit', 'Rev. David Nguyen', 'Vibe Events DJ', 'Garden Ceremony', false, 'Signed Sealed Delivered', 'Confirmed', true, ''],
  ['4:32 PM',  '4:50 PM',  'Formal Portraits',      'Couple + families + wedding party photos', 'Isaiah Torres', 'Golden Hour Studio', 'Garden', false, 'Quiet', 'Confirmed', false, 'Guests move to cocktail hour'],
  // Cocktail Hour
  ['5:00 PM',  '6:00 PM',  'Cocktail Hour',         'Guests enjoy appetizers, drinks, and lawn games', 'Sophia Rivera', 'Harvest Table Catering + DJ', 'Terrace', false, 'Neo-soul playlist', 'Confirmed', false, ''],
  ['5:30 PM',  '6:00 PM',  'Photo Booth Opens',     'Open-air booth with custom backdrop', 'ClickBox ATX', 'ClickBox ATX', 'Ballroom Foyer', false, '', 'Confirmed', false, ''],
  // Reception
  ['6:00 PM',  '6:05 PM',  'Doors Open to Ballroom','Guests find seats; music transitions', 'Sophia Rivera + DJ', 'Vibe Events DJ', 'Ballroom', false, 'Playlist: R&B & Afrobeats', 'Confirmed', false, ''],
  ['6:05 PM',  '6:10 PM',  'Wedding Party Entrance','Introduced by DJ; party walks in', 'MC / DJ Marcus', 'Vibe Events DJ', 'Ballroom', true, 'Upbeat entrance songs', 'Confirmed', true, ''],
  ['6:10 PM',  '6:14 PM',  'Couple Entrance',       'Grand introduction of newlyweds', 'MC / DJ Marcus', 'Vibe Events DJ', 'Ballroom', true, 'TBD — couple picks', 'Confirmed', true, ''],
  ['6:14 PM',  '6:20 PM',  'Welcome Remarks',       'Both families welcome guests', 'Robert Thornton + Kwame Osei', 'LifeFrame Films', 'Ballroom', true, 'Mic open', 'Confirmed', true, ''],
  ['6:20 PM',  '6:30 PM',  'First Dance',           'Couple first dance', 'Amara & Marcus', 'Vibe Events DJ', 'Dance Floor', false, 'All of Me — John Legend', 'Confirmed', true, ''],
  ['6:30 PM',  '6:36 PM',  'Parent Dances',         'Father-daughter then mother-son', 'Parents', 'Vibe Events DJ', 'Dance Floor', false, 'Isn\'t She Lovely / My Wish', 'Confirmed', true, ''],
  ['6:36 PM',  '6:42 PM',  'Wedding Party Dance',   'Wedding party joins the dance floor', 'Wedding Party', 'Vibe Events DJ', 'Dance Floor', false, 'Upbeat song TBD', 'Confirmed', true, ''],
  ['6:42 PM',  '7:30 PM',  'Dinner Service — First & Second Course', 'Salad and entrée service', 'Harvest Table Catering', 'Caterer', 'Ballroom', false, 'Dinner jazz playlist', 'Confirmed', false, ''],
  ['7:30 PM',  '7:35 PM',  'Toasts — Best Man',     'Best Man Noel Parker toast', 'Noel Parker', 'LifeFrame Films', 'Ballroom', true, 'Mic open', 'Confirmed', true, ''],
  ['7:35 PM',  '7:40 PM',  'Toast — Maid of Honor', 'Jordan Osei toast', 'Jordan Osei', 'LifeFrame Films', 'Ballroom', true, 'Mic open', 'Confirmed', true, ''],
  ['7:40 PM',  '7:42 PM',  'Toast — Couple',        'Bride & groom thank guests', 'Amara & Marcus', 'LifeFrame Films', 'Ballroom', true, 'Mic open', 'Confirmed', true, ''],
  ['7:45 PM',  '8:00 PM',  'Cake Cutting',          'Formal cake cutting + dessert bar opens', 'Amara & Marcus + Staff', 'Sugar & Spire + Catering', 'Dessert Station', false, 'How Sweet It Is', 'Confirmed', true, ''],
  ['8:00 PM',  '9:30 PM',  'Open Dancing',          'DJ opens dance floor — full energy set', 'Vibe Events DJ', 'Vibe Events DJ', 'Dance Floor', false, 'Mixed hits — DJ curated', 'Confirmed', false, ''],
  ['9:30 PM',  '9:35 PM',  'Bouquet Toss',          'Bride\'s bouquet toss to guests', 'MC / Amara', 'Vibe Events DJ', 'Dance Floor', true, 'TBD', 'Confirmed', true, ''],
  ['10:30 PM', '10:45 PM', 'Last Dance',             'Slow final song for all guests', 'MC / DJ Marcus', 'Vibe Events DJ', 'Dance Floor', false, 'At Last — Etta James', 'Confirmed', true, ''],
  ['10:45 PM', '11:00 PM', 'Grand Sendoff',          'Sparkler send-off at main entrance', 'All Guests', 'LifeFrame Films + Photo', 'Main Entrance', false, 'Upbeat outro', 'Confirmed', true, 'Sparklers distributed 10:40'],
];

(async () => {
  const data = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  data.push({ range: `${S}!A1`, values: [['RECEPTION RUN-OF-SHOW PLANNER']] });
  reqs.push({ mergeCells: { range: gridRange(RP, 0, 1, 0, 16), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: { range: gridRange(RP, 0, 1, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 14 },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: RP, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize' }});

  // ── Summary row 2 ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A2`, values: [[
    'Total Segments:','',
    `=IFERROR(COUNTA(E5:E200),"")`, '',
    'Confirmed:','',
    `=IFERROR(SUMPRODUCT((L5:L200="Confirmed")*1),"")`, '',
    'Completed:','',
    `=IFERROR(SUMPRODUCT((P5:P200=TRUE)*1),"")`, '',
    'Conflicts:','',
    `=IFERROR(SUMPRODUCT((N5:N200="YES")*1),"")`,
  ]] });
  reqs.push({ repeatCell: { range: gridRange(RP, 1, 2, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.altRow),
    textFormat: { bold: true, foregroundColor: hex(C.primary), fontSize: 9 },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Headers row 4 ─────────────────────────────────────────────────────────
  data.push({ range: `${S}!A4`, values: [HEADERS] });
  reqs.push({ repeatCell: { range: gridRange(RP, 3, 4, 0, 16), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary),
    textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 9 },
    horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)' }});
  reqs.push({ updateDimensionProperties: { range: { sheetId: RP, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 44 }, fields: 'pixelSize' }});

  // ── Timeline data rows ────────────────────────────────────────────────────
  // Phase section headers
  const phases = {
    0: 'PRE-CEREMONY PREPARATION',
    3: 'CEREMONY',
    10: 'COCKTAIL HOUR',
    12: 'RECEPTION',
  };

  segments.forEach((seg, i) => {
    const row = 5 + i;
    const [start, end, name, desc, responsible, vendor, location, avNeeds, music, status, announce, notes] = seg;

    data.push({ range: `${S}!A${row}`, values: [[i + 1]] });
    data.push({ range: `${S}!B${row}`, values: [[start]] });
    data.push({ range: `${S}!C${row}`, values: [[end]] });
    // D: Duration (auto) — uses TIME arithmetic
    data.push({ range: `${S}!D${row}`, values: [[`=IFERROR(TEXT(TIMEVALUE(C${row})-TIMEVALUE(B${row}),"h:mm"),"")` ]] });
    data.push({ range: `${S}!E${row}`, values: [[name]] });
    data.push({ range: `${S}!F${row}`, values: [[desc]] });
    data.push({ range: `${S}!G${row}`, values: [[responsible]] });
    data.push({ range: `${S}!H${row}`, values: [[vendor]] });
    data.push({ range: `${S}!I${row}`, values: [[location]] });
    data.push({ range: `${S}!J${row}`, values: [[avNeeds]] });
    data.push({ range: `${S}!K${row}`, values: [[music]] });
    data.push({ range: `${S}!L${row}`, values: [[status]] });
    data.push({ range: `${S}!M${row}`, values: [[announce]] });
    // N: Time Conflict — start time < previous end time
    if (i === 0) {
      data.push({ range: `${S}!N${row}`, values: [['']] });
    } else {
      data.push({ range: `${S}!N${row}`, values: [[
        `=IFERROR(IF(AND(B${row}<>"",C${row-1}<>"",TIMEVALUE(B${row})<TIMEVALUE(C${row-1})),"YES",""),"")`,
      ]]) });
    }
    data.push({ range: `${S}!O${row}`, values: [[notes]] });
    data.push({ range: `${S}!P${row}`, values: [[false]] });
  });

  // Checkboxes: J (AV), M (Announce), P (Completed)
  [9, 12, 15].forEach(col => {
    reqs.push({ setDataValidation: {
      range: gridRange(RP, 4, 4 + segments.length, col, col + 1),
      rule: { condition: { type: 'BOOLEAN' }, showCustomUi: true },
    }});
  });

  // Status dropdown
  reqs.push({ setDataValidation: {
    range: gridRange(RP, 4, 200, 11, 12),
    rule: { condition: { type: 'ONE_OF_LIST', values: [
      'Confirmed', 'Tentative', 'TBD', 'Cancelled',
    ].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: false },
  }});

  // Formula col D, N — blue-gray
  [3, 13].forEach(col => {
    reqs.push({ repeatCell: { range: gridRange(RP, 4, 4 + segments.length, col, col + 1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.formula),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });

  // Phase sub-headers (section dividers)
  Object.entries(phases).forEach(([segIdx, label]) => {
    const phaseRow = 5 + parseInt(segIdx) - 0.5; // insert before segment
    // We'll use conditional formatting instead — mark col E with section label styling
  });

  // Alternating rows
  for (let i = 0; i < segments.length; i++) {
    const bg = i % 2 === 0 ? C.panel : C.altRow;
    reqs.push({ repeatCell: { range: gridRange(RP, 4 + i, 5 + i, 0, 16), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, foregroundColor: hex(C.mainText) },
    }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});
  }

  // CF: Conflict → red
  reqs.push({ addConditionalFormatRule: {
    rule: { ranges: [gridRange(RP, 4, 200, 13, 14)],
      booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'YES' }] },
        format: { backgroundColor: hex(C.attention), textFormat: { bold: true, foregroundColor: hex(C.white) } } },
    }, index: 0,
  }});

  // CF: Status colors
  [{ text: 'Confirmed', bg: C.success }, { text: 'TBD', bg: C.warning }].forEach(({ text, bg }) => {
    reqs.push({ addConditionalFormatRule: {
      rule: { ranges: [gridRange(RP, 4, 200, 11, 12)],
        booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] },
          format: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.white) } } },
      }, index: 0,
    }});
  });

  // Column widths
  const widths = [35, 70, 70, 65, 160, 220, 140, 160, 120, 70, 160, 80, 70, 80, 180, 70];
  widths.forEach((px, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: RP, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 },
      properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Freeze row 4
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: RP, gridProperties: { frozenRowCount: 4 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  reqs.push({ updateSheetProperties: {
    properties: { sheetId: RP, tabColor: hex(C.accent) },
    fields: 'tabColor',
  }});

  await valuesBatchUpdate(id, data, 'reception-data');
  await batchUpdate(id, reqs, 'reception-format');

  console.log('✓ Reception Planner built with', segments.length, 'segments');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
