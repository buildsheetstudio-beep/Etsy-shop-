'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Habit Tracker'];
const S = "'Habit Tracker'";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function fmtDate(d) { return d.toISOString().split('T')[0]; }

// 10 habit definitions
const HABITS = [
  ['HB-001','Drink 8 glasses of water',     'Hydration',  '8 glasses',  'Daily',    true,  ''],
  ['HB-002','Morning stretch (10 min)',      'Movement',   '10 min',     'Daily',    true,  ''],
  ['HB-003','7+ hours of sleep',             'Sleep',      '7 hours',    'Daily',    true,  ''],
  ['HB-004','Log meals in Nutrition Log',    'Nutrition',  '3 meals',    'Daily',    true,  ''],
  ['HB-005','10-min mindfulness/meditation', 'Mindfulness','10 min',     'Daily',    true,  ''],
  ['HB-006','Read for 20 minutes',           'Self-Care',  '20 min',     'Daily',    true,  ''],
  ['HB-007','Log weight (when prompted)',    'Nutrition',  '1 entry',    'Weekdays', true,  'Weekly is fine too'],
  ['HB-008','Limit alcohol to 1 drink max',  'Nutrition',  '1 or less',  'Daily',    true,  ''],
  ['HB-009','No screens 30 min before bed',  'Sleep',      '30 min',     'Daily',    false, 'Working on this'],
  ['HB-010','Evening walk (15+ min)',        'Movement',   '15 min',     'Daily',    true,  ''],
];

// 153 habit log entries: 9 active habits × 17 days (Aug 26 - Sep 11, 2026)
// Each day log all 9 active habits
const LOG_START = new Date('2026-08-26');
const ACTIVE_HABITS = HABITS.filter(h => h[5] === true);
const habitLogRows = [];
for (let day = 0; day < 17; day++) {
  const date = fmtDate(addDays(LOG_START, day));
  ACTIVE_HABITS.forEach((habit) => {
    const [hid, hname] = habit;
    // Mostly done, occasional miss to simulate realistic tracking
    const done = !(day % 5 === 4 && hid === 'HB-009') && !(day % 7 === 6 && hid === 'HB-006');
    habitLogRows.push([date, hid, hname, done, '']);
  });
}

(async () => {
  const fmt = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,10000,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Row 1: Title
  vals.push({ range: `${S}!A1`, values: [['HABIT TRACKER']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 48 }, fields: 'pixelSize' }});

  // ── SECTION 1: Habit Definitions (rows 3-13) ─────────────────────────────
  // Section header row 2
  vals.push({ range: `${S}!A2`, values: [['HABIT DEFINITIONS — Add or edit habits below. Active habits appear in the daily log section.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Row 3: Definition headers
  vals.push({ range: `${S}!A3`, values: [['Habit ID','Habit Name','Category','Daily Target','Frequency','Active?','Done Today','Notes']] });
  fmt.push({ repeatCell: { range: gridRange(SID,2,3,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // Habit definition rows 4-13 (ri=3-12)
  // A=ID(formula), B=name, C=category, D=target, E=freq, F=active(checkbox), G=done_today(formula), H=notes
  fmt.push({ repeatCell: { range: gridRange(SID,3,13,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B4="","","HB-"&TEXT(ROW()-3,"000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.primary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // G: Done Today formula (counts today's done entries in log section below)
  fmt.push({ repeatCell: { range: gridRange(SID,3,13,6,7), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(COUNTIFS($B$20:$B$10000,A4,$C$20:$C$10000,TODAY(),$D$20:$D$10000,TRUE),0)` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial', bold: true }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push habit definition data (B=name, C=cat, D=target, E=freq, F=active, G=formula, H=notes)
  // HABITS format: [id, name, cat, target, freq, active, notes]
  vals.push({ range: `${S}!B4`, values: HABITS.map(h => [h[1], h[2], h[3], h[4], h[5], '', h[6]]) });

  // Style definition rows
  fmt.push({ repeatCell: { range: gridRange(SID,3,13,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.warmTint), textFormat: { fontSize: 9, fontFamily: 'Arial' },
    verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 13 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ── SECTION 2: Daily Habit Log (rows 17+) ────────────────────────────────
  // Spacer rows 14-16
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 13, endIndex: 16 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Section header row 17
  vals.push({ range: `${S}!A17`, values: [['DAILY HABIT LOG — Log each habit daily. Filter by date or habit to review your patterns.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,16,17,0,8), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,16,17,0,8), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondaryTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 16, endIndex: 17 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Row 18: quick stats for log
  const STATS = [
    ['Log Entries',    `=COUNTA(C20:C10000)`],
    ['Days Logged',    `=SUMPRODUCT((C20:C5000<>"")*IFERROR(1/COUNTIF(C20:C5000,C20:C5000),0))`],
    ['Completion Rate',`=IFERROR(TEXT(COUNTIF(D20:D10000,TRUE)/COUNTA(D20:D10000),"0%"),"")`],
  ];
  STATS.forEach(([label, formula], i) => {
    const r = 18;
    const c = i * 3;
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,c,c+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ mergeCells: { range: gridRange(SID,r-1,r,c+2,c+3), mergeType: 'MERGE_ALL' }});
    vals.push({ range: `${S}!${String.fromCharCode(65+c)}18`, values: [[label]] });
    vals.push({ range: `${S}!${String.fromCharCode(65+c+2)}18`, values: [[formula]] });
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,c,c+2), cell: { userEnteredFormat: {
      backgroundColor: hex(C.primaryTint), textFormat: { bold: true, fontSize: 9, fontFamily: 'Arial', foregroundColor: hex(C.primary) },
      horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ repeatCell: { range: gridRange(SID,r-1,r,c+2,c+3), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input), textFormat: { bold: true, fontSize: 10, fontFamily: 'Arial', foregroundColor: hex(C.secondary) },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 17, endIndex: 18 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Row 19: Log headers (ri=18)
  vals.push({ range: `${S}!A19`, values: [['Log ID','Habit ID','Date','Habit Name','Done?','Value / Notes','']] });
  fmt.push({ repeatCell: { range: gridRange(SID,18,19,0,7), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 18, endIndex: 19 }, properties: { pixelSize: 32 }, fields: 'pixelSize' }});

  // A: Log ID formula for log rows (ri=19+)
  fmt.push({ repeatCell: { range: gridRange(SID,19,10000,0,1), cell: {
    userEnteredValue: { formulaValue: `=IF(B20="","","HL-"&TEXT(ROW()-19,"00000"))` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { bold: true, fontFamily: 'Arial', foregroundColor: hex(C.secondary) }, horizontalAlignment: 'CENTER' },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // D: Habit Name VLOOKUP from definitions section
  fmt.push({ repeatCell: { range: gridRange(SID,19,10000,3,4), cell: {
    userEnteredValue: { formulaValue: `=IFERROR(VLOOKUP(B20,$A$4:$H$13,2,FALSE),"")` },
    userEnteredFormat: { backgroundColor: hex(C.formula), textFormat: { fontFamily: 'Arial' } },
  }, fields: 'userEnteredValue,userEnteredFormat' }});

  // Push log data (B=habitId, C=date, D=formula, E=done, F=notes)
  // habitLogRows: [date, hid, hname, done, notes]
  vals.push({ range: `${S}!B20`, values: habitLogRows.map(r => [r[1], r[0], '', r[3], r[4]]) });

  // Row banding for log section
  fmt.push({ addBanding: { bandedRange: {
    range: gridRange(SID,19,10000,0,7),
    rowProperties: { firstBandColor: hex(C.white), secondBandColor: hex(C.altRow) },
  }}});

  // Date format col C (log section)
  fmt.push({ repeatCell: { range: gridRange(SID,19,10000,2,3), cell: { userEnteredFormat: {
    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' },
  }}, fields: 'userEnteredFormat(numberFormat)' }});

  // Column widths
  const WIDTHS = [80, 80, 100, 180, 65, 220, 100];
  WIDTHS.forEach((w,ci) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: w }, fields: 'pixelSize' }}));

  // Freeze row 1 only (sections have their own sub-headers)
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 1 } }, fields: 'gridProperties.frozenRowCount' }});

  await valuesBatchUpdate(id, vals, '10-habits values');
  await batchUpdate(id, fmt, '10-habits format');
  console.log(`✅  Habit Tracker done — ${HABITS.length} habits, ${habitLogRows.length} log entries.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
