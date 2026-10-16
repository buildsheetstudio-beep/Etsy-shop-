'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const LAYOUT = sheetMap['Garden Layout'];

// Zone color definitions
const ZONE_COLORS = {
  A: C.paleSage,        // Herbs — light sage
  B: C.paleTerracotta,  // Tomatoes — light terracotta
  C: C.paleWheat,       // Root Veg — light wheat
  D: C.cream2,          // Greens — cream
  E: C.oliveLight,      // Cucurbits — olive light (will be toned down with #E8F0D8)
  F: C.lightGray,       // Flowers — light gray/lavender
  G: '#E8F4E8',         // Perennials — very light green
  P: '#F5F0E8',         // Path — parchment
  W: '#D6EEF8',         // Water — light blue
};

// 30-row × 30-col garden grid starting at row 5 (index 4), col 2 (B)
// Rows 0-29 (30 rows), Cols 0-29 (30 cols) relative to the data zone
// Layout: grid starts at spreadsheet row 5 (idx 4), col B (idx 1)
// We offset: grid row R = sheet row R+4, grid col C = sheet col C+1

const GRID_START_ROW = 4;  // row index 4 = row 5
const GRID_START_COL = 1;  // col index 1 = col B
const GRID_ROWS = 30;
const GRID_COLS = 30;

// Zone map: 30x30 grid — each cell is a zone letter or '.'
// Row 0 = top of garden
function buildZoneMap() {
  const m = [];
  for (let r = 0; r < GRID_ROWS; r++) {
    const row = [];
    for (let c = 0; c < GRID_COLS; c++) {
      // Paths (2 cells wide) at rows 5, 11, 17, 23 and cols 5, 11, 17, 23
      if (r === 5 || r === 11 || r === 17 || r === 23) { row.push('P'); continue; }
      if (c === 5 || c === 11 || c === 17 || c === 23) { row.push('P'); continue; }
      // Water feature at top right corner rows 0-4, cols 25-29
      if (r <= 3 && c >= 26) { row.push('W'); continue; }
      // Zone assignments based on quadrants
      if (r <= 4 && c <= 4)           row.push('A'); // Herbs — top left
      else if (r <= 4 && c >= 6 && c <= 10)   row.push('B'); // Tomatoes
      else if (r <= 4 && c >= 12 && c <= 16)  row.push('B'); // Tomatoes
      else if (r <= 4 && c >= 18 && c <= 22)  row.push('C'); // Root Veg
      else if (r <= 4 && c >= 24 && c <= 25)  row.push('C'); // Root Veg
      else if (r >= 6 && r <= 10 && c <= 4)   row.push('D'); // Greens
      else if (r >= 6 && r <= 10 && c >= 6 && c <= 10)  row.push('B'); // Tomatoes
      else if (r >= 6 && r <= 10 && c >= 12 && c <= 16) row.push('E'); // Cucurbits
      else if (r >= 6 && r <= 10 && c >= 18 && c <= 22) row.push('C'); // Root Veg
      else if (r >= 6 && r <= 10 && c >= 24 && c <= 29) row.push('F'); // Flowers
      else if (r >= 12 && r <= 16 && c <= 4)  row.push('D'); // Greens
      else if (r >= 12 && r <= 16 && c >= 6 && c <= 10)  row.push('A'); // Herbs
      else if (r >= 12 && r <= 16 && c >= 12 && c <= 16) row.push('E'); // Cucurbits
      else if (r >= 12 && r <= 16 && c >= 18 && c <= 22) row.push('F'); // Flowers
      else if (r >= 12 && r <= 16 && c >= 24 && c <= 29) row.push('G'); // Perennials
      else if (r >= 18 && r <= 22 && c <= 4)  row.push('G'); // Perennials
      else if (r >= 18 && r <= 22 && c >= 6 && c <= 10)  row.push('C'); // Root Veg
      else if (r >= 18 && r <= 22 && c >= 12 && c <= 16) row.push('D'); // Greens
      else if (r >= 18 && r <= 22 && c >= 18 && c <= 22) row.push('E'); // Cucurbits
      else if (r >= 18 && r <= 22 && c >= 24 && c <= 29) row.push('A'); // Herbs
      else if (r >= 24 && c <= 4)   row.push('F'); // Flowers — bottom left
      else if (r >= 24 && c >= 6 && c <= 10)   row.push('G'); // Perennials
      else if (r >= 24 && c >= 12 && c <= 16)  row.push('B'); // Tomatoes
      else if (r >= 24 && c >= 18 && c <= 22)  row.push('D'); // Greens
      else if (r >= 24 && c >= 24 && c <= 29)  row.push('E'); // Cucurbits
      else row.push('.');
    }
    m.push(row);
  }
  return m;
}

// Sample plant abbreviations in various zones
const PLANT_LABELS = {
  '1,1':  'BSL', '1,2': 'BSL', '2,1': 'THY', '2,2': 'THY', '3,1': 'RSM', '3,2': 'MNT', // Herbs
  '7,7':  'TOM', '7,8': 'TOM', '8,7': 'TOM', '8,8': 'TOM', '9,7': 'CHR', '9,8': 'CHR', // Tomatoes
  '7,9':  'PEP', '8,9': 'PEP', '9,9': 'JAL',
  '7,13': 'ZUC', '7,14': 'ZUC', '8,13': 'CUC', '8,14': 'CUC', '9,13': 'PMP',
  '1,18': 'CAR', '1,19': 'CAR', '2,18': 'BET', '2,19': 'BET', '3,18': 'RAD', '3,19': 'RAD', '4,18': 'TRN', '4,19': 'GAR',
  '6,25': 'MGD', '6,26': 'MGD', '7,25': 'NST', '7,26': 'SUN', '8,25': 'LAV', '8,26': 'LAV',
  '6,2':  'LET', '6,3': 'LET', '7,2': 'SPN', '7,3': 'SPN', '8,2': 'KAL', '9,2': 'ARC', '10,3': 'CRD',
  '18,25':'BSL', '19,25': 'CIL', '20,25': 'DIL', '21,25': 'CHV', '22,25': 'PTH',
  '18,2': 'STR', '19,2': 'STR', '20,2': 'LVD', '21,2': 'LVD',
  '24,2': 'MGD', '25,2': 'NST', '26,2': 'SUN', '27,2': 'SUN',
  '24,25':'CUC', '25,25':'ZUC', '26,25':'PMP', '27,25':'PMP',
};

(async () => {
  const reqs = [];
  const values = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(LAYOUT, 0, 1, 0, 35), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Garden Layout!A1', values: [['🗺️ GARDEN LAYOUT MAP — Season 2025']] });
  reqs.push({
    repeatCell: {
      range: gridRange(LAYOUT, 0, 1, 0, 35),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.soil),
          textFormat: { foregroundColor: hex(C.cream), bold: true, fontSize: 16 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: LAYOUT, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 44 }, fields: 'pixelSize',
  }});

  // ── Zone Legend rows 2-3 ──────────────────────────────────────────────────
  values.push({ range: 'Garden Layout!A2', values: [
    ['ZONE LEGEND:', 'A — Herbs', '', 'B — Tomatoes & Peppers', '', 'C — Root Vegetables', '', 'D — Leafy Greens', '', 'E — Cucurbits', '', 'F — Flowers', '', 'G — Perennials', '', 'P — Path', '', 'W — Water'],
  ]});
  values.push({ range: 'Garden Layout!A3', values: [
    ['ABBREVIATIONS:', 'BSL=Basil', 'THY=Thyme', 'RSM=Rosemary', 'MNT=Mint', 'TOM=Tomato', 'CHR=Cherry Tom', 'PEP=Bell Pepper', 'JAL=Jalapeño', 'ZUC=Zucchini', 'CUC=Cucumber', 'PMP=Pumpkin', 'CAR=Carrot', 'BET=Beet', 'RAD=Radish', 'LET=Lettuce', 'SPN=Spinach', 'KAL=Kale', 'MGD=Marigold', 'STR=Strawberry'],
  ]});
  reqs.push({
    repeatCell: {
      range: gridRange(LAYOUT, 1, 4, 0, 35),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.parchment),
          textFormat: { fontSize: 8, bold: false },
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // ── Row 4: Column numbers header ──────────────────────────────────────────
  const colNums = [['']];
  for (let c = 1; c <= 30; c++) colNums[0].push(String(c));
  values.push({ range: 'Garden Layout!A4', values: colNums });
  reqs.push({
    repeatCell: {
      range: gridRange(LAYOUT, 3, 4, 0, 31),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.charcoal),
          textFormat: { foregroundColor: hex(C.cream), bold: true, fontSize: 8 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Grid cells ────────────────────────────────────────────────────────────
  const zoneMap = buildZoneMap();

  // Set all cell sizes: 28px wide, 22px tall
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: LAYOUT, dimension: 'COLUMNS', startIndex: 1, endIndex: 31 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: LAYOUT, dimension: 'ROWS', startIndex: GRID_START_ROW, endIndex: GRID_START_ROW + GRID_ROWS },
    properties: { pixelSize: 22 }, fields: 'pixelSize',
  }});
  // Row labels col A — narrow
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: LAYOUT, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});

  // Row number labels
  const rowLabels = [];
  for (let r = 0; r < GRID_ROWS; r++) rowLabels.push([String(r + 1)]);
  values.push({ range: `Garden Layout!A5`, values: rowLabels });
  reqs.push({
    repeatCell: {
      range: gridRange(LAYOUT, GRID_START_ROW, GRID_START_ROW + GRID_ROWS, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.charcoal),
          textFormat: { foregroundColor: hex(C.cream), bold: true, fontSize: 8 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Zone colors and labels
  const gridValues = zoneMap.map((row, r) =>
    row.map((zone, c) => {
      const key = `${r},${c}`;
      return PLANT_LABELS[key] || zone;
    })
  );

  for (let r = 0; r < GRID_ROWS; r++) {
    values.push({
      range: `Garden Layout!B${GRID_START_ROW + r + 1}:AE${GRID_START_ROW + r + 1}`,
      values: [gridValues[r]],
    });
  }

  // Color each zone cell
  for (let r = 0; r < GRID_ROWS; r++) {
    for (let c = 0; c < GRID_COLS; c++) {
      const zone = zoneMap[r][c];
      let bg, fg, bold;
      switch (zone) {
        case 'A': bg = '#D8EDD4'; fg = C.olive; bold = false; break;
        case 'B': bg = '#F5D5C3'; fg = C.terracotta; bold = false; break;
        case 'C': bg = '#FBF0C8'; fg = C.soil; bold = false; break;
        case 'D': bg = '#EAF5EA'; fg = C.growingGreen; bold = false; break;
        case 'E': bg = '#DFF0D8'; fg = C.oliveLight; bold = false; break;
        case 'F': bg = '#F3E5F5'; fg = '#7B1FA2'; bold = false; break;
        case 'G': bg = '#E8F5E9'; fg = C.growingGreen; bold = false; break;
        case 'P': bg = C.parchment; fg = C.warmGray; bold = false; break;
        case 'W': bg = '#BBDEFB'; fg = '#1565C0'; bold = false; break;
        default:  bg = C.white; fg = C.charcoal; bold = false; break;
      }
      const hasLabel = PLANT_LABELS[`${r},${c}`];
      reqs.push({
        repeatCell: {
          range: gridRange(LAYOUT, GRID_START_ROW + r, GRID_START_ROW + r + 1, GRID_START_COL + c, GRID_START_COL + c + 1),
          cell: {
            userEnteredFormat: {
              backgroundColor: hex(bg),
              textFormat: { foregroundColor: hex(fg), bold: hasLabel ? true : bold, fontSize: 7 },
              horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
              borders: {
                top:    { style: 'SOLID', color: hex('#CCCCCC') },
                bottom: { style: 'SOLID', color: hex('#CCCCCC') },
                left:   { style: 'SOLID', color: hex('#CCCCCC') },
                right:  { style: 'SOLID', color: hex('#CCCCCC') },
              },
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)',
        },
      });
    }
  }

  // ── Zone color key below the grid ─────────────────────────────────────────
  const keyRow = GRID_START_ROW + GRID_ROWS + 1;
  reqs.push({ mergeCells: { range: gridRange(LAYOUT, keyRow, keyRow + 1, 0, 35), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Garden Layout!A${keyRow + 1}`, values: [['Zone Color Key — Edit cells above to plan your garden layout using zone letters (A-G), abbreviations, or plant names']] });
  reqs.push({
    repeatCell: {
      range: gridRange(LAYOUT, keyRow, keyRow + 1, 0, 35),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.parchment),
          textFormat: { foregroundColor: hex(C.soil), italic: true, fontSize: 9 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  await valuesBatchUpdate(id, values, 'layout-values');
  await batchUpdate(id, reqs, 'layout-format');
  console.log('Garden Layout complete — 30×30 grid with 7 zones');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
