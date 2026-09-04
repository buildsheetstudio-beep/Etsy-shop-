'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Chore & Allowance Tracker'];

// Column index constants
// A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12 N=13 O=14 P=15 Q=16
const ALLOW_COLS = [2,4,6,8,10,12,14,16]; // C,E,G,I,K,M,O,Q (0-indexed)
const HOUR_COLS  = [1,3,5,7,9,11,13,15];  // B,D,F,H,J,L,N,P (0-indexed)

const MEM_BASE = m => 23 + m * 7;

(async () => {
  const reqs = [];

  // ── Merge helper ─────────────────────────────────────────────────────────
  function merge(r1, r2, c1, c2) {
    reqs.push({ mergeCells: { range: gridRange(TAB, r1, r2, c1, c2), mergeType: 'MERGE_ALL' } });
  }

  // ── Row 1 merges (title block + rate table block + family block) ──────────
  // A1: title spans A1:E1
  merge(0, 1, 0, 5);
  // F1: rate table header spans F1:H1
  merge(0, 1, 5, 8);
  // I1: family members header spans I1:K1
  merge(0, 1, 8, 11);

  // ── Row 2 merges ──────────────────────────────────────────────────────────
  // A2:B2 already a values range (no merge needed)

  // ── Row 13: Day header pairs (B:C, D:E, F:G, H:I, J:K, L:M, N:O, P:Q) ──
  for (let d = 0; d < 7; d++) {
    const c = 1 + d * 2; // B=1, D=3, F=5, H=7, J=9, L=11, N=13
    merge(12, 13, c, c + 2);
  }
  // P:Q merge for Weekly Achievement
  merge(12, 13, 15, 17);

  // ── Row 15: TOTAL label spans A15:Q15 header ─────────────────────────────
  // Just A15 label, no merge across col A needed

  // ── Member header rows: each spans A-Q within one row ────────────────────
  for (let m = 0; m < 10; m++) {
    const headerRow = 22 + m * 7; // 1-indexed
    merge(headerRow - 1, headerRow, 0, 17); // span A to Q (0-indexed: 0 to 17)
  }

  // ── Row 15 TOTAL block header spans A15 only (no merge) ─────────────────
  // No merges needed for TOTAL header

  // ── Format helpers ────────────────────────────────────────────────────────
  function fmt(r1, r2, c1, c2, cell) {
    reqs.push({ repeatCell: { range: gridRange(TAB, r1, r2, c1, c2), cell, fields: 'userEnteredFormat' } });
  }

  const bold = (bg, fg, sz, align) => ({
    userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(fg), bold: true, fontSize: sz || 10 },
      horizontalAlignment: align || 'CENTER',
      verticalAlignment: 'MIDDLE',
    }
  });
  const plain = (bg, fg, sz, align) => ({
    userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(fg), fontSize: sz || 9 },
      horizontalAlignment: align || 'CENTER',
      verticalAlignment: 'MIDDLE',
    }
  });
  const currency = (bg, fg) => ({
    userEnteredFormat: {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(fg), fontSize: 9 },
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
      horizontalAlignment: 'CENTER',
      verticalAlignment: 'MIDDLE',
    }
  });

  // ── Row 1: Block headers ──────────────────────────────────────────────────
  // Title block A1:E1
  fmt(0, 1, 0, 5, bold(C.charcoal, C.white, 13, 'CENTER'));
  // Rate table header F1:H1
  fmt(0, 1, 5, 8, bold(C.chrome, C.charcoal, 10, 'CENTER'));
  // Family members header I1:K1
  fmt(0, 1, 8, 11, bold(C.chrome, C.charcoal, 10, 'CENTER'));

  // ── Row 2: sub-headers ────────────────────────────────────────────────────
  fmt(1, 2, 0, 2, bold(C.chrome, C.charcoal, 9, 'LEFT'));   // A2:B2 Start Date
  fmt(1, 2, 5, 8, bold(C.chrome, C.charcoal, 9, 'CENTER')); // F2:H2 col headers
  fmt(1, 2, 8, 11, bold(C.chrome, C.charcoal, 9, 'CENTER')); // I2:K2 col headers

  // ── Rows 3-8: Rate table data (F:H) ──────────────────────────────────────
  for (let r = 2; r < 8; r++) {
    const bg = r % 2 === 0 ? C.white : C.chrome;
    fmt(r, r + 1, 5, 8, plain(bg, C.charcoal, 9, 'LEFT'));
  }

  // ── Rows 3-12: Family member block (I:K) ─────────────────────────────────
  for (let m = 0; m < 10; m++) {
    const r = m + 2; // 0-indexed rows 2-11
    const bg = C.slot[m];
    fmt(r, r + 1, 8, 9, bold(bg, C.charcoal, 9, 'LEFT'));   // I: name
    fmt(r, r + 1, 9, 11, plain(bg, C.charcoal, 9, 'CENTER')); // J:K: weekly totals
    // Currency format for K (weekly allowance, col index 10)
    reqs.push({ repeatCell: { range: gridRange(TAB, r, r + 1, 10, 11), cell: currency(bg, C.charcoal), fields: 'userEnteredFormat' } });
  }

  // ── Row 13: Day headers ───────────────────────────────────────────────────
  fmt(12, 13, 0, 17, bold(C.charcoal, C.white, 10, 'CENTER'));

  // ── Row 14: H/$ sub-headers ──────────────────────────────────────────────
  fmt(13, 14, 0, 17, bold(C.chrome, C.charcoal, 8, 'CENTER'));

  // ── Row 15: TOTAL block label ─────────────────────────────────────────────
  fmt(14, 15, 0, 17, bold(C.charcoal, C.white, 10, 'LEFT'));

  // ── Rows 16-21: TOTAL chore rows ──────────────────────────────────────────
  for (let c = 0; c < 6; c++) {
    const r = 15 + c; // 0-indexed rows 15-20
    const bg = c % 2 === 0 ? C.chrome : C.white;
    // Label col A
    fmt(r, r + 1, 0, 1, bold(bg, C.charcoal, 9, 'LEFT'));
    // Hour cols
    for (const ci of HOUR_COLS) {
      fmt(r, r + 1, ci, ci + 1, plain(bg, C.charcoal, 9, 'CENTER'));
    }
    // Allow cols — currency
    for (const ci of ALLOW_COLS) {
      reqs.push({ repeatCell: { range: gridRange(TAB, r, r + 1, ci, ci + 1), cell: currency(bg, C.charcoal), fields: 'userEnteredFormat' } });
    }
  }

  // ── Rows 22-91: 10 member blocks ─────────────────────────────────────────
  for (let m = 0; m < 10; m++) {
    const headerRow = 22 + m * 7; // 1-indexed
    const hr = headerRow - 1;     // 0-indexed
    const base = MEM_BASE(m);     // 1-indexed first chore row
    const slotBg = C.slot[m];
    const slotAlt = C.slotAlt[m];

    // Member header row (full width)
    fmt(hr, hr + 1, 0, 17, bold(slotBg, C.charcoal, 10, 'LEFT'));

    // 6 chore rows
    for (let c = 0; c < 6; c++) {
      const r = hr + 1 + c; // 0-indexed
      const bg = c % 2 === 0 ? slotBg : slotAlt;
      // Label
      fmt(r, r + 1, 0, 1, plain(bg, C.charcoal, 9, 'LEFT'));
      // Hour cols
      for (const ci of HOUR_COLS) {
        fmt(r, r + 1, ci, ci + 1, plain(bg, C.charcoal, 9, 'CENTER'));
      }
      // Allow cols — currency
      for (const ci of ALLOW_COLS) {
        reqs.push({ repeatCell: { range: gridRange(TAB, r, r + 1, ci, ci + 1), cell: currency(bg, C.charcoal), fields: 'userEnteredFormat' } });
      }
    }
  }

  // ── Borders ───────────────────────────────────────────────────────────────
  const thinBorder = { style: 'SOLID', color: hex('#BBBBBB') };
  const thickBorder = { style: 'SOLID_MEDIUM', color: hex(C.charcoal) };

  // Outer border around header block (rows 1-12)
  reqs.push({ updateBorders: { range: gridRange(TAB, 0, 12, 0, 11),
    top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder } });

  // Border around TOTAL block (rows 15-21)
  reqs.push({ updateBorders: { range: gridRange(TAB, 14, 21, 0, 17),
    top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder } });

  // Inner thin borders for grid rows 13-91
  reqs.push({ updateBorders: { range: gridRange(TAB, 12, 91, 0, 17),
    innerHorizontal: thinBorder, innerVertical: thinBorder } });

  // ── Column widths ─────────────────────────────────────────────────────────
  // A=150, B,D,F,H,J,L,N=55, C,E,G,I,K,M,O=70, P=75, Q=85
  const colWidths = [
    [0, 150],   // A
    [1, 55],    // B
    [2, 70],    // C
    [3, 55],    // D
    [4, 70],    // E
    [5, 55],    // F
    [6, 70],    // G
    [7, 55],    // H
    [8, 70],    // I
    [9, 55],    // J
    [10, 70],   // K
    [11, 55],   // L
    [12, 70],   // M
    [13, 55],   // N
    [14, 70],   // O
    [15, 75],   // P
    [16, 85],   // Q
  ];
  for (const [ci, w] of colWidths) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: TAB, dimension: 'COLUMNS', startIndex: ci, endIndex: ci + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    } });
  }

  // ── Row heights ───────────────────────────────────────────────────────────
  // Header rows 1-2: 30px, grid header rows 13-15: 28px, data rows: 22px
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize'
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 12, endIndex: 15 },
    properties: { pixelSize: 26 }, fields: 'pixelSize'
  } });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 15, endIndex: 91 },
    properties: { pixelSize: 22 }, fields: 'pixelSize'
  } });

  // ── Date format for B2 ────────────────────────────────────────────────────
  reqs.push({ repeatCell: {
    range: gridRange(TAB, 1, 2, 1, 2),
    cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' } } },
    fields: 'userEnteredFormat.numberFormat'
  } });

  await batchUpdate(id, reqs, 'tracker-format');
  console.log('Tracker format complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
