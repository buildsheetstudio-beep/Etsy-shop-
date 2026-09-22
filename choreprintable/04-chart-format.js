'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const TAB = sheetMap['Weekly Chore Chart'];

// Column indices (0-indexed)
// A=0, B=1,C=2,D=3, E=4,F=5,G=6, H=7,I=8,J=9, K=10,L=11,M=12,
// N=13,O=14,P=15, Q=16,R=17,S=18, T=19,U=20,V=21
const CHK = [1,4,7,10,13,16,19];   // B,E,H,K,N,Q,T (checkbox cols)
const DSC = [2,5,8,11,14,17,20];   // C,F,I,L,O,R,U (task left col)
const MEM_START = [14,18,22,26,30,34]; // 1-indexed first row per member

(async () => {
  const reqs = [];

  // ── Merge helper ─────────────────────────────────────────────────────────
  function merge(r1,r2,c1,c2) {
    reqs.push({ mergeCells: { range: gridRange(TAB,r1,r2,c1,c2), mergeType:'MERGE_ALL' } });
  }

  // ── Merges ────────────────────────────────────────────────────────────────
  // Title A1:V1 and A2:V2 (merge both rows together as one block via row merge)
  merge(0,2,0,22);   // A1:V2
  merge(2,3,0,22);   // A3:V3

  // Row 5-7 setup area
  merge(4,7,0,5);    // A5:F7 — chart label
  merge(4,5,6,10);   // G5:J5 — week start label
  merge(5,7,6,10);   // G6:J7 — week start input
  merge(4,7,10,13);  // K5:M7 — total chores card
  merge(4,7,13,16);  // N5:P7 — completed card
  merge(4,7,16,19);  // Q5:S7 — remaining card
  merge(4,7,19,22);  // T5:V7 — completion % card

  // Row 9-11 family member strip
  merge(8,11,0,3);   // A9:C11 — label
  merge(8,11,3,6);   // D9:F11 — member 1
  merge(8,11,6,9);   // G9:I11 — member 2
  merge(8,11,9,12);  // J9:L11 — member 3
  merge(8,11,12,15); // M9:O11 — member 4
  merge(8,11,15,18); // P9:R11 — member 5
  merge(8,11,18,22); // S9:V11 — member 6

  // Row 13: day header merges
  // A13 no merge
  merge(12,13,1,4);  // B13:D13 Monday
  merge(12,13,4,7);  // E13:G13 Tuesday
  merge(12,13,7,10); // H13:J13 Wednesday
  merge(12,13,10,13);// K13:M13 Thursday
  merge(12,13,13,16);// N13:P13 Friday
  merge(12,13,16,19);// Q13:S13 Saturday
  merge(12,13,19,22);// T13:V13 Sunday

  // Grid rows 14-37: column A per member, task desc pairs per day
  for (let m = 0; m < 6; m++) {
    const r1 = MEM_START[m] - 1;     // 0-indexed start
    const r2 = r1 + 4;               // 4 rows
    merge(r1, r2, 0, 1);             // A col merged per member (4 rows)
    // Task desc pairs: C:D, F:G, I:J, L:M, O:P, R:S, U:V for each chore row
    for (let r = r1; r < r2; r++) {
      for (let d = 0; d < 7; d++) {
        const c = DSC[d]; // left desc col
        merge(r, r+1, c, c+2); // merge desc+spacer col pair
      }
    }
  }

  // Notes area
  merge(38,39,0,22);  // A39:V39 notes label
  merge(39,44,0,22);  // A40:V44 notes area

  // Footer
  merge(45,47,0,22);  // A46:V47

  // ── Format helpers ────────────────────────────────────────────────────────
  function fmt(r1,r2,c1,c2,cell) {
    reqs.push({ repeatCell: { range: gridRange(TAB,r1,r2,c1,c2), cell, fields:'userEnteredFormat' } });
  }

  const center = 'CENTER';
  const left   = 'LEFT';
  const mid    = 'MIDDLE';

  // ── Title rows 1-2 ────────────────────────────────────────────────────────
  fmt(0,2,0,22, { userEnteredFormat: {
    backgroundColor: hex(C.bg),
    textFormat: { foregroundColor: hex(C.black), bold: true, fontSize: 24, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
  }});

  // ── Subtitle row 3 ────────────────────────────────────────────────────────
  fmt(2,3,0,22, { userEnteredFormat: {
    backgroundColor: hex(C.bg),
    textFormat: { foregroundColor: hex('#3D3D3D'), bold: true, fontSize: 11, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
  }});

  // ── Chore Chart label A5:F7 ───────────────────────────────────────────────
  fmt(4,7,0,5, { userEnteredFormat: {
    backgroundColor: hex(C.headerGray),
    textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 13, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
  }});

  // ── Week Start label G5:J5 ────────────────────────────────────────────────
  fmt(4,5,6,10, { userEnteredFormat: {
    backgroundColor: hex(C.headerGray),
    textFormat: { foregroundColor: hex(C.secondary), bold: true, fontSize: 9, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
  }});

  // ── Week Start input G6:J7 ────────────────────────────────────────────────
  fmt(5,7,6,10, { userEnteredFormat: {
    backgroundColor: hex(C.white),
    textFormat: { foregroundColor: hex(C.mainText), bold: false, fontSize: 13, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
    numberFormat: { type: 'DATE', pattern: 'MMM D, YYYY' },
  }});

  // ── Summary cards K5:V7 ───────────────────────────────────────────────────
  const cardLabel = (bg) => ({ userEnteredFormat: {
    backgroundColor: hex(bg),
    textFormat: { foregroundColor: hex(C.secondary), bold: true, fontSize: 8, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
  }});
  const cardValue = (bg, nfmt) => {
    const f = {
      backgroundColor: hex(bg),
      textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 22, fontFamily: 'Arial' },
      horizontalAlignment: center, verticalAlignment: mid,
    };
    if (nfmt) f.numberFormat = nfmt;
    return { userEnteredFormat: f };
  };

  // Each card is rows 4-6 (0-indexed), merged into one. We format the whole merged range uniformly.
  // The label is in the top portion, value in the lower — but since merged we use two separate ranges
  // with their own format applied to row 4 (label row) and rows 5-6 (value rows).
  // Actually since these are merged, the cell value shows in the entire merged region.
  // We'll format the whole card as one unit with small label font, then the value cell separately.

  // Cards at cols K5:M7 (10-12), N5:P7 (13-15), Q5:S7 (16-18), T5:V7 (19-21)
  // Label row 5 (idx 4), value cell row 6-7 (idx 5-6) — but all merged as single cell
  // Strategy: put label in top row text + value in merged cell; use two separate write ranges
  // Since the range is merged (r4-r7), we only have one effective cell.
  // We'll use font size for the combined cell at a mid size, then put label/value as two text lines
  // using newline in the value string. Actually we wrote them as separate cells.
  // Since we merged K5:M7 entirely, there's only one cell. We wrote 'TOTAL CHORES' to K5 and the formula to K6.
  // After merging K5:M7, K5 is the anchor and K6 value is lost. We need a different approach:
  // DON'T merge the full K5:M7 vertically — merge label (K5:M5) and value (K6:M7) separately.

  // Let me reconsider the merge strategy for summary cards.
  // The merges above merged K5:M7 as a single block. That means only K5's value shows.
  // Better: merge K5:M5 (label) and K6:M7 (value) as two separate merges.
  // But we already pushed the full K5:M7 merge. We need to redo this.
  // Unmerge K5:M7 and redo as two merges per card.

  // Unmerge cards first
  [[4,7,10,13],[4,7,13,16],[4,7,16,19],[4,7,19,22]].forEach(([r1,r2,c1,c2]) => {
    reqs.push({ unmergeCells: { range: gridRange(TAB,r1,r2,c1,c2) } });
  });
  // Re-merge as label row + value area
  [[10,13],[13,16],[16,19],[19,22]].forEach(([c1,c2]) => {
    merge(4,5,c1,c2);   // label row 5
    merge(5,7,c1,c2);   // value rows 6-7
  });

  const cardBg = C.headerGray;
  [[10,13],[13,16],[16,19],[19,22]].forEach(([c1,c2]) => {
    fmt(4,5,c1,c2, cardLabel(cardBg));
    fmt(5,7,c1,c2, cardValue(cardBg));
  });
  // Percent format for completion card
  reqs.push({ repeatCell: {
    range: gridRange(TAB,5,7,19,22),
    cell: { userEnteredFormat: {
      backgroundColor: hex(cardBg),
      textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 22, fontFamily: 'Arial' },
      horizontalAlignment: center, verticalAlignment: mid,
      numberFormat: { type: 'PERCENT', pattern: '0%' },
    }},
    fields: 'userEnteredFormat'
  }});

  // ── Family member header label A9:C11 ────────────────────────────────────
  fmt(8,11,0,3, { userEnteredFormat: {
    backgroundColor: hex(C.headerGray),
    textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 10, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
  }});

  // ── Member name cells D9:V11 (6 members) ─────────────────────────────────
  const memCols = [[3,6],[6,9],[9,12],[12,15],[15,18],[18,22]];
  for (let m = 0; m < 6; m++) {
    const [c1,c2] = memCols[m];
    fmt(8,11,c1,c2, { userEnteredFormat: {
      backgroundColor: hex(C.member[m]),
      textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 11, fontFamily: 'Arial' },
      horizontalAlignment: center, verticalAlignment: mid,
    }});
  }

  // ── Row 13: Day headers ───────────────────────────────────────────────────
  // A13
  fmt(12,13,0,1, { userEnteredFormat: {
    backgroundColor: hex(C.headerGray),
    textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 10, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
  }});
  // Mon-Fri: B:D, E:G, H:J, K:M, N:P
  [[1,4],[4,7],[7,10],[10,13],[13,16]].forEach(([c1,c2]) => {
    fmt(12,13,c1,c2, { userEnteredFormat: {
      backgroundColor: hex(C.headerGray),
      textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 10, fontFamily: 'Arial' },
      horizontalAlignment: center, verticalAlignment: mid,
    }});
  });
  // Sat-Sun: Q:S, T:V — weekend tint
  [[16,19],[19,22]].forEach(([c1,c2]) => {
    fmt(12,13,c1,c2, { userEnteredFormat: {
      backgroundColor: hex(C.weekend),
      textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 10, fontFamily: 'Arial' },
      horizontalAlignment: center, verticalAlignment: mid,
    }});
  });

  // ── Grid rows 14-37: member blocks ───────────────────────────────────────
  for (let m = 0; m < 6; m++) {
    const r1 = MEM_START[m] - 1; // 0-indexed
    const r2 = r1 + 4;

    // Col A member label
    fmt(r1,r2,0,1, { userEnteredFormat: {
      backgroundColor: hex(C.member[m]),
      textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 9, fontFamily: 'Arial' },
      horizontalAlignment: center, verticalAlignment: mid,
      wrapStrategy: 'WRAP',
    }});

    // Each row in this member block
    for (let r = r1; r < r2; r++) {
      // Checkbox cols (B,E,H,K,N,Q,T)
      for (const ci of CHK) {
        const isWeekend = ci >= 16; // Q or T
        const bg = isWeekend ? C.weekend : C.white;
        fmt(r,r+1,ci,ci+1, { userEnteredFormat: {
          backgroundColor: hex(bg),
          horizontalAlignment: center, verticalAlignment: mid,
        }});
      }
      // Task desc cols (merged C:D, F:G, I:J, L:M, O:P, R:S, U:V)
      for (let d = 0; d < 7; d++) {
        const ci = DSC[d];
        const isWeekend = d >= 5;
        const bg = isWeekend ? C.weekend : C.white;
        fmt(r,r+1,ci,ci+2, { userEnteredFormat: {
          backgroundColor: hex(bg),
          textFormat: { foregroundColor: hex(C.mainText), fontSize: 8, fontFamily: 'Arial' },
          horizontalAlignment: left, verticalAlignment: mid,
          wrapStrategy: 'WRAP',
        }});
      }
    }
  }

  // ── Notes area ────────────────────────────────────────────────────────────
  fmt(38,39,0,22, { userEnteredFormat: {
    backgroundColor: hex(C.headerGray),
    textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 10, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
  }});
  fmt(39,44,0,22, { userEnteredFormat: {
    backgroundColor: hex(C.notes),
    textFormat: { foregroundColor: hex(C.mainText), fontSize: 9, fontFamily: 'Arial' },
    horizontalAlignment: left, verticalAlignment: 'TOP',
    wrapStrategy: 'WRAP',
  }});

  // ── Footer ────────────────────────────────────────────────────────────────
  fmt(45,47,0,22, { userEnteredFormat: {
    backgroundColor: hex(C.footer),
    textFormat: { foregroundColor: hex(C.mainText), bold: true, fontSize: 11, fontFamily: 'Arial' },
    horizontalAlignment: center, verticalAlignment: mid,
  }});

  // ── Borders ───────────────────────────────────────────────────────────────
  const medium = { style:'SOLID_MEDIUM', color: hex(C.border) };
  const thin   = { style:'SOLID',        color: hex(C.border) };

  // Major section outer borders
  [
    [4,7,0,5],         // chart label
    [4,5,6,10],        // week start label
    [5,7,6,10],        // week start input
    [4,5,10,22],       // summary card labels
    [5,7,10,22],       // summary card values
    [8,11,0,22],       // member name strip
    [12,13,0,22],      // day headers
    [13,37,0,22],      // full grid
    [38,44,0,22],      // notes section
    [45,47,0,22],      // footer
  ].forEach(([r1,r2,c1,c2]) => {
    reqs.push({ updateBorders: { range: gridRange(TAB,r1,r2,c1,c2),
      top:medium, bottom:medium, left:medium, right:medium } });
  });

  // Inner thin borders for grid cells
  reqs.push({ updateBorders: { range: gridRange(TAB,13,37,0,22),
    innerHorizontal: thin, innerVertical: thin } });

  // ── Column widths ─────────────────────────────────────────────────────────
  // A=88, B,E,H,K,N,Q,T=28, C:D,F:G,I:J,L:M,O:P,R:S,U:V=62 each pair
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 88 }, fields: 'pixelSize'
  }});
  for (const ci of CHK) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: TAB, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 },
      properties: { pixelSize: 28 }, fields: 'pixelSize'
    }});
  }
  // Task col pairs: C:D (2-3), F:G (5-6), I:J (8-9), L:M (11-12), O:P (14-15), R:S (17-18), U:V (20-21)
  [[2,4],[5,7],[8,10],[11,13],[14,16],[17,19],[20,22]].forEach(([c1,c2]) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: TAB, dimension: 'COLUMNS', startIndex: c1, endIndex: c2 },
      properties: { pixelSize: 62 }, fields: 'pixelSize'
    }});
  });

  // ── Row heights ───────────────────────────────────────────────────────────
  // Title rows 1-2: 34px each
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
    properties: { pixelSize: 34 }, fields: 'pixelSize'
  }});
  // Row 3 (subtitle): 24px
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 24 }, fields: 'pixelSize'
  }});
  // Row 4 gap
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 3, endIndex: 4 },
    properties: { pixelSize: 8 }, fields: 'pixelSize'
  }});
  // Setup rows 5-7: ~34px each
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 4, endIndex: 7 },
    properties: { pixelSize: 34 }, fields: 'pixelSize'
  }});
  // Row 8 gap
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 7, endIndex: 8 },
    properties: { pixelSize: 8 }, fields: 'pixelSize'
  }});
  // Member name strip 9-11: ~34px each
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 8, endIndex: 11 },
    properties: { pixelSize: 34 }, fields: 'pixelSize'
  }});
  // Row 12 gap
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 11, endIndex: 12 },
    properties: { pixelSize: 8 }, fields: 'pixelSize'
  }});
  // Day header row 13: 26px
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 12, endIndex: 13 },
    properties: { pixelSize: 26 }, fields: 'pixelSize'
  }});
  // Chore rows 14-37: 30px
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 13, endIndex: 37 },
    properties: { pixelSize: 30 }, fields: 'pixelSize'
  }});
  // Row 38 gap
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 37, endIndex: 38 },
    properties: { pixelSize: 8 }, fields: 'pixelSize'
  }});
  // Notes label + content rows 39-44: 24px
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 38, endIndex: 44 },
    properties: { pixelSize: 24 }, fields: 'pixelSize'
  }});
  // Row 45 gap
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 44, endIndex: 45 },
    properties: { pixelSize: 8 }, fields: 'pixelSize'
  }});
  // Footer rows 46-47: 28px
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: TAB, dimension: 'ROWS', startIndex: 45, endIndex: 47 },
    properties: { pixelSize: 28 }, fields: 'pixelSize'
  }});

  await batchUpdate(id, reqs, 'chart-format');
  console.log('Chart format complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
