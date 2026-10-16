'use strict';
const { batchUpdate, valuesBatchUpdate, hex, COLOR, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = sheetMap['Weekly Planner']; // 2

// Time slots: 7:00 AM to 11:00 PM in 30-min increments = 33 slots
// Rows 4-36 (0-indexed: 3-35)
const TIMES = [];
for (let h = 7; h <= 22; h++) {
  TIMES.push(`${h > 12 ? h - 12 : h}:00 ${h >= 12 ? 'PM' : 'AM'}`);
  if (h < 22) TIMES.push(`${h > 12 ? h - 12 : h}:30 ${h >= 12 ? 'PM' : 'AM'}`);
}
TIMES.push('11:00 PM');
// 7:00 AM to 10:30 PM = 32 entries then 11:00 PM = 33 total
// Let me recount: 7am,7:30,8am,8:30,...,10pm,10:30pm,11pm
// h=7: 7:00 AM, 7:30 AM  (2)
// h=8-11: 4 pairs = 8
// h=12: 12:00 PM, 12:30 PM (2)
// h=13-21: 9 pairs = 18
// h=22: 10:00 PM, 10:30 PM (2)
// then we add 11:00 PM = 1
// Total = 2+8+2+18+2+1 = 33 ✓

// Pre-filled schedule for Oct 6-12, 2025 (Mon-Sun)
// Cols: B=Mon, C=Tue, D=Wed, E=Thu, F=Fri, G=Sat, H=Sun
// Row index 0 = 7:00 AM (row 4), row index 32 = 11:00 PM (row 36)
// Time slot index: 0=7am, 1=7:30, 2=8am, 3=8:30, 4=9am, 5=9:30, 6=10am, 7=10:30,
//   8=11am, 9=11:30, 10=12pm, 11=12:30, 12=1pm, 13=1:30, 14=2pm, 15=2:30,
//   16=3pm, 17=3:30, 18=4pm, 19=4:30, 20=5pm, 21=5:30, 22=6pm, 23=6:30,
//   24=7pm, 25=7:30, 26=8pm, 27=8:30, 28=9pm, 29=9:30, 30=10pm, 31=10:30, 32=11pm

// Build schedule grid (33 rows × 7 cols), initialize empty
const schedule = Array.from({ length: 33 }, () => Array(7).fill(''));

// Helper: set slot (timeIdx, dayIdx 0=Mon..6=Sun, text)
function set(t, d, text) { schedule[t][d] = text; }

// Monday Oct 6
set(2, 0, 'English 101');      // 8:00 AM
set(3, 0, 'English 101');      // 8:30 AM
set(4, 0, 'Calculus II');      // 9:00 AM
set(5, 0, 'Calculus II');      // 9:30 AM
set(8, 0, 'World History');    // 11:00 AM
set(9, 0, 'World History');    // 11:30 AM
set(16, 0, 'Study');           // 3:00 PM
set(17, 0, 'Study');           // 3:30 PM
set(18, 0, 'Study');           // 4:00 PM
set(26, 0, 'Reading');         // 8:00 PM
set(27, 0, 'Reading');         // 8:30 PM

// Tuesday Oct 7
set(4, 1, 'Biology 201');      // 9:00 AM
set(5, 1, 'Biology 201');      // 9:30 AM
set(6, 1, 'Biology 201');      // 10:00 AM (lab)
set(12, 1, 'Economics 101');   // 1:00 PM
set(13, 1, 'Economics 101');   // 1:30 PM
set(16, 1, 'Group Study');     // 3:00 PM
set(17, 1, 'Group Study');     // 3:30 PM
set(18, 1, 'Group Study');     // 4:00 PM
set(24, 1, 'Study');           // 7:00 PM
set(25, 1, 'Study');           // 7:30 PM

// Wednesday Oct 8
set(2, 2, 'English 101');      // 8:00 AM
set(3, 2, 'English 101');      // 8:30 AM
set(4, 2, 'Calculus II');      // 9:00 AM
set(5, 2, 'Calculus II');      // 9:30 AM
set(8, 2, 'World History');    // 11:00 AM
set(9, 2, 'World History');    // 11:30 AM
set(14, 2, 'Office Hours');    // 2:00 PM
set(16, 2, 'Study');           // 3:00 PM
set(17, 2, 'Study');           // 3:30 PM
set(26, 2, 'Reading');         // 8:00 PM
set(27, 2, 'Reading');         // 8:30 PM
set(28, 2, 'Reading');         // 9:00 PM

// Thursday Oct 9
set(4, 3, 'Biology 201');      // 9:00 AM
set(5, 3, 'Biology 201');      // 9:30 AM
set(12, 3, 'Economics 101');   // 1:00 PM
set(13, 3, 'Economics 101');   // 1:30 PM
set(16, 3, 'Study');           // 3:00 PM
set(17, 3, 'Study');           // 3:30 PM
set(18, 3, 'Study');           // 4:00 PM
set(22, 3, 'Exercise');        // 6:00 PM
set(23, 3, 'Exercise');        // 6:30 PM

// Friday Oct 10
set(2, 4, 'English 101');      // 8:00 AM
set(3, 4, 'English 101');      // 8:30 AM
set(4, 4, 'Calculus II');      // 9:00 AM
set(5, 4, 'Calculus II');      // 9:30 AM
set(8, 4, 'World History');    // 11:00 AM
set(9, 4, 'World History');    // 11:30 AM
set(16, 4, 'Review Session');  // 3:00 PM
set(17, 4, 'Review Session');  // 3:30 PM
set(20, 4, 'Personal');        // 5:00 PM
set(21, 4, 'Personal');        // 5:30 PM

// Saturday Oct 11
set(8, 5, 'Study');            // 11:00 AM
set(9, 5, 'Study');            // 11:30 AM
set(10, 5, 'Study');           // 12:00 PM
set(11, 5, 'Study');           // 12:30 PM
set(16, 5, 'Group Study');     // 3:00 PM
set(17, 5, 'Group Study');     // 3:30 PM
set(22, 5, 'Personal');        // 6:00 PM
set(23, 5, 'Personal');        // 6:30 PM

// Sunday Oct 12
set(10, 6, 'Study');           // 12:00 PM
set(11, 6, 'Study');           // 12:30 PM
set(12, 6, 'Study');           // 1:00 PM
set(14, 6, 'Reading');         // 2:00 PM
set(15, 6, 'Reading');         // 2:30 PM
set(24, 6, 'Review Session');  // 7:00 PM
set(25, 6, 'Review Session');  // 7:30 PM

(async () => {
  const reqs = [];

  // ── Column widths ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 90 }, fields: 'pixelSize' } });
  for (let c = 1; c <= 7; c++) {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 }, properties: { pixelSize: 115 }, fields: 'pixelSize' } });
  }

  // ── Row heights ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 30 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 36 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // ── Title row ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 8),
      cell: {
        userEnteredValue: { stringValue: '📅 Weekly Planner' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.mint),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Week label row (row 1) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 1, 2, 0, 8),
      cell: {
        userEnteredValue: { stringValue: 'Week of Oct 6 – Oct 12, 2025' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleMint),
          textFormat: { foregroundColor: hex(COLOR.emerald), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Stats row (row 2) ──
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 4), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(SID, 2, 3, 4, 8), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 0, 4),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.subheaderBg),
          textFormat: { foregroundColor: hex(COLOR.darkBlue), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 2, 3, 4, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.paleMint),
          textFormat: { foregroundColor: hex(COLOR.emerald), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Day header row (row 3) ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 3, 4, 0, 8),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.emerald),
          textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Time column (A) format for data rows ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 3, 36, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(COLOR.subheaderBg),
          textFormat: { foregroundColor: hex(COLOR.darkBlue), bold: true, fontSize: 9 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Alt hour shading (every 2 rows = 1 hour) ──
  for (let i = 0; i < 33; i += 4) {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 3 + i, Math.min(3 + i + 2, 36), 1, 8),
        cell: { userEnteredFormat: { backgroundColor: hex(COLOR.iceBlue) } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
  }

  // ── Data cell default format ──
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 3, 36, 1, 8),
      cell: {
        userEnteredFormat: {
          textFormat: { fontSize: 9 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
          wrapStrategy: 'CLIP',
        },
      },
      fields: 'userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });

  // ── Borders ──
  reqs.push({
    updateBorders: {
      range: gridRange(SID, 3, 36, 0, 8),
      innerHorizontal: { style: 'SOLID', color: hex(COLOR.softGrey) },
      innerVertical: { style: 'SOLID', color: hex(COLOR.midGrey) },
      bottom: { style: 'SOLID', color: hex(COLOR.midGrey) },
      right: { style: 'SOLID', color: hex(COLOR.midGrey) },
    },
  });

  await batchUpdate(id, reqs, 'weekly-format');

  // ── Values ──
  const days = ['Mon 10/6','Tue 10/7','Wed 10/8','Thu 10/9','Fri 10/10','Sat 10/11','Sun 10/12'];
  const headerRow = ['Time', ...days];
  const timeRows = TIMES.map((t, i) => [t, ...schedule[i]]);

  await valuesBatchUpdate(id, [
    { range: 'Weekly Planner!A3', values: [['="""Study Hours: """&COUNTIF(B4:H36,"Study")*0.5&""" hrs"""']] },
    { range: 'Weekly Planner!E3', values: [['="""Lecture Hours: """&(COUNTIF(B4:H36,"English 101")+COUNTIF(B4:H36,"Calculus II")+COUNTIF(B4:H36,"World History")+COUNTIF(B4:H36,"Biology 201")+COUNTIF(B4:H36,"Economics 101"))*0.5&""" hrs"""']] },
    { range: 'Weekly Planner!A4:H4', values: [headerRow] },
    { range: 'Weekly Planner!A5:H37', values: timeRows },
  ], 'weekly-values');

  console.log('Weekly Planner complete —', TIMES.length, 'time slots');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
