'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const TL = sheetMap['🗓️ Moving Timeline'];

(async () => {
  const values = [];
  const reqs = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  values.push({ range: "'🗓️ Moving Timeline'!A1", values: [['🗓️ All-in-One Moving Planner — Moving Timeline']] });

  // ── Column headers (row 2) ────────────────────────────────────────────────
  values.push({ range: "'🗓️ Moving Timeline'!A2", values: [[
    '#', 'Phase', 'Weeks Before Move', 'Target Date', 'Status', 'Tasks Completed', 'Total Tasks', 'Progress', 'Notes',
  ]]});

  // ── 8 phases ──────────────────────────────────────────────────────────────
  // Dashboard B5 = move date, B2 = weeks placeholder (actually Dashboard sets it)
  // Use hard-coded move date 2025-08-30 for target date calc
  // Phase target dates = MOVE_DATE - (weeks * 7)
  // Move date serial: 2025-08-30
  const phases = [
    [1, '8 Weeks Out — Research & Planning',      8, '=DATE(2025,8,30)-(B3*7)', 'Not Started', 0, 10, '', 'Research, book movers, declutter'],
    [2, '6 Weeks Out — Bookings & Notifications',  6, '=DATE(2025,8,30)-(B4*7)', 'Not Started', 0, 8,  '', 'Book movers, notify schools, utilities'],
    [3, '4 Weeks Out — Packing Begins',            4, '=DATE(2025,8,30)-(B5*7)', 'Not Started', 0, 8,  '', 'Start packing, notify financial inst.'],
    [4, '2 Weeks Out — Finishing Up',              2, '=DATE(2025,8,30)-(B6*7)', 'Not Started', 0, 6,  '', 'Final packing, confirm all bookings'],
    [5, 'Moving Week — Final Preparations',        1, '=DATE(2025,8,30)-(B7*7)', 'Not Started', 0, 5,  '', 'Pack essentials, disassemble furniture'],
    [6, 'Moving Day — 30 August 2025',             0, 'Moving Day',              'Not Started', 0, 3,  '', 'Supervise movers, meter readings'],
    [7, 'First Week in New Home',                 -1, '=DATE(2025,8,30)+7',      'Not Started', 0, 5,  '', 'Unpack essentials, update address'],
    [8, 'Post-Move — Wrap Up',                    -4, '=DATE(2025,8,30)+28',     'Not Started', 0, 5,  '', 'Final admin, confirm all changes done'],
  ];

  // Write weeks in col B (used in formulas)
  const weeksValues = phases.map(p => [p[2]]);
  values.push({ range: "'🗓️ Moving Timeline'!B3", values: weeksValues });

  const phaseRows = phases.map((p, i) => [
    p[0],
    p[1],
    p[2],
    p[3],
    p[4],
    p[5],
    p[6],
    `=IFERROR(SPARKLINE({F${3 + i},MAX(G${3 + i}-F${3 + i},0)},{"charttype","bar";"color1","#2C4A35";"color2","#FBF5E6"}),"-")`,
    p[8],
  ]);
  values.push({ range: "'🗓️ Moving Timeline'!A3", values: phaseRows });

  await valuesBatchUpdate(id, values, 'timeline-values');

  // ── Merge title ───────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(TL, 0, 1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(TL, 0, 1, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.deepForest),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // ── Column headers row 2 ──────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(TL, 1, 2, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.mossGreen),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Zebra stripe data rows ────────────────────────────────────────────────
  for (let r = 0; r < 8; r++) {
    const bg = r % 2 === 0 ? C.ivory : C.warmCream;
    reqs.push({
      repeatCell: {
        range: gridRange(TL, 2 + r, 3 + r, 0, 9),
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

  // Moving Day row — special highlight (row index 7 = row 8, phase 6)
  reqs.push({
    repeatCell: {
      range: gridRange(TL, 7, 8, 0, 9),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.antiqueBrass),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    },
  });

  // Date format for col D
  reqs.push({
    repeatCell: {
      range: gridRange(TL, 2, 7, 3, 4),
      cell: {
        userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(TL, 8, 10, 3, 4),
      cell: {
        userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'dd mmm yyyy' } },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [35, 240, 100, 120, 100, 100, 80, 150, 220];
  colWidths.forEach((w, i) => {
    reqs.push({
      updateDimensionProperties: {
        range: { sheetId: TL, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w },
        fields: 'pixelSize',
      },
    });
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: TL, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 40 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: TL, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
      properties: { pixelSize: 28 },
      fields: 'pixelSize',
    },
  });
  reqs.push({
    updateDimensionProperties: {
      range: { sheetId: TL, dimension: 'ROWS', startIndex: 2, endIndex: 10 },
      properties: { pixelSize: 32 },
      fields: 'pixelSize',
    },
  });

  // ── Freeze rows 1–2 ───────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: TL, gridProperties: { frozenRowCount: 2 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'timeline-format');
  console.log('Moving Timeline complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
