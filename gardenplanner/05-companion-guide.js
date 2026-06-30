'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const COMP = sheetMap['Companion Planting Guide'];

// 20 plants for the companion matrix
const MATRIX_PLANTS = [
  'Tomato','Basil','Carrot','Marigold','Cucumber','Beans',
  'Lettuce','Spinach','Kale','Radish','Pepper','Zucchini',
  'Garlic','Rosemary','Mint','Dill','Nasturtium','Peas',
  'Sunflower','Chives',
];

// Companion matrix: Good(✅), Bad(❌), Neutral(—)
// Row plant + Col plant relationship
// [row][col] — below diagonal mirrors above (skipping diagonal = same plant)
const RELATIONSHIPS = {
  'Tomato-Basil': '✅', 'Tomato-Carrot': '✅', 'Tomato-Marigold': '✅',
  'Tomato-Cucumber': '—', 'Tomato-Beans': '—', 'Tomato-Lettuce': '✅',
  'Tomato-Spinach': '✅', 'Tomato-Kale': '—', 'Tomato-Radish': '—',
  'Tomato-Pepper': '✅', 'Tomato-Zucchini': '—', 'Tomato-Garlic': '✅',
  'Tomato-Rosemary': '—', 'Tomato-Mint': '✅', 'Tomato-Dill': '❌',
  'Tomato-Nasturtium': '✅', 'Tomato-Peas': '❌', 'Tomato-Sunflower': '—',
  'Tomato-Chives': '✅',
  'Basil-Carrot': '—', 'Basil-Marigold': '✅', 'Basil-Cucumber': '✅',
  'Basil-Beans': '—', 'Basil-Lettuce': '—', 'Basil-Spinach': '—',
  'Basil-Kale': '—', 'Basil-Radish': '—', 'Basil-Pepper': '✅',
  'Basil-Zucchini': '—', 'Basil-Garlic': '—', 'Basil-Rosemary': '—',
  'Basil-Mint': '—', 'Basil-Dill': '❌', 'Basil-Nasturtium': '✅',
  'Basil-Peas': '—', 'Basil-Sunflower': '—', 'Basil-Chives': '—',
  'Carrot-Marigold': '✅', 'Carrot-Cucumber': '—', 'Carrot-Beans': '✅',
  'Carrot-Lettuce': '✅', 'Carrot-Spinach': '—', 'Carrot-Kale': '—',
  'Carrot-Radish': '✅', 'Carrot-Pepper': '✅', 'Carrot-Zucchini': '—',
  'Carrot-Garlic': '—', 'Carrot-Rosemary': '✅', 'Carrot-Mint': '—',
  'Carrot-Dill': '❌', 'Carrot-Nasturtium': '—', 'Carrot-Peas': '✅',
  'Carrot-Sunflower': '—', 'Carrot-Chives': '✅',
  'Marigold-Cucumber': '✅', 'Marigold-Beans': '—', 'Marigold-Lettuce': '✅',
  'Marigold-Spinach': '—', 'Marigold-Kale': '✅', 'Marigold-Radish': '—',
  'Marigold-Pepper': '✅', 'Marigold-Zucchini': '✅', 'Marigold-Garlic': '—',
  'Marigold-Rosemary': '—', 'Marigold-Mint': '—', 'Marigold-Dill': '—',
  'Marigold-Nasturtium': '✅', 'Marigold-Peas': '—', 'Marigold-Sunflower': '✅',
  'Marigold-Chives': '—',
  'Cucumber-Beans': '✅', 'Cucumber-Lettuce': '✅', 'Cucumber-Spinach': '—',
  'Cucumber-Kale': '—', 'Cucumber-Radish': '✅', 'Cucumber-Pepper': '—',
  'Cucumber-Zucchini': '❌', 'Cucumber-Garlic': '❌', 'Cucumber-Rosemary': '—',
  'Cucumber-Mint': '—', 'Cucumber-Dill': '✅', 'Cucumber-Nasturtium': '✅',
  'Cucumber-Peas': '✅', 'Cucumber-Sunflower': '✅', 'Cucumber-Chives': '—',
  'Beans-Lettuce': '✅', 'Beans-Spinach': '✅', 'Beans-Kale': '—',
  'Beans-Radish': '—', 'Beans-Pepper': '—', 'Beans-Zucchini': '✅',
  'Beans-Garlic': '❌', 'Beans-Rosemary': '✅', 'Beans-Mint': '—',
  'Beans-Dill': '—', 'Beans-Nasturtium': '—', 'Beans-Peas': '—',
  'Beans-Sunflower': '✅', 'Beans-Chives': '❌',
  'Lettuce-Spinach': '✅', 'Lettuce-Kale': '✅', 'Lettuce-Radish': '✅',
  'Lettuce-Pepper': '—', 'Lettuce-Zucchini': '—', 'Lettuce-Garlic': '✅',
  'Lettuce-Rosemary': '—', 'Lettuce-Mint': '✅', 'Lettuce-Dill': '✅',
  'Lettuce-Nasturtium': '✅', 'Lettuce-Peas': '✅', 'Lettuce-Sunflower': '—',
  'Lettuce-Chives': '✅',
  'Spinach-Kale': '✅', 'Spinach-Radish': '—', 'Spinach-Pepper': '—',
  'Spinach-Zucchini': '—', 'Spinach-Garlic': '✅', 'Spinach-Rosemary': '—',
  'Spinach-Mint': '—', 'Spinach-Dill': '—', 'Spinach-Nasturtium': '—',
  'Spinach-Peas': '✅', 'Spinach-Sunflower': '—', 'Spinach-Chives': '—',
  'Kale-Radish': '—', 'Kale-Pepper': '—', 'Kale-Zucchini': '—',
  'Kale-Garlic': '✅', 'Kale-Rosemary': '—', 'Kale-Mint': '—',
  'Kale-Dill': '✅', 'Kale-Nasturtium': '✅', 'Kale-Peas': '✅',
  'Kale-Sunflower': '—', 'Kale-Chives': '—',
  'Radish-Pepper': '—', 'Radish-Zucchini': '—', 'Radish-Garlic': '—',
  'Radish-Rosemary': '—', 'Radish-Mint': '—', 'Radish-Dill': '—',
  'Radish-Nasturtium': '✅', 'Radish-Peas': '—', 'Radish-Sunflower': '—',
  'Radish-Chives': '—',
  'Pepper-Zucchini': '—', 'Pepper-Garlic': '✅', 'Pepper-Rosemary': '—',
  'Pepper-Mint': '—', 'Pepper-Dill': '—', 'Pepper-Nasturtium': '—',
  'Pepper-Peas': '—', 'Pepper-Sunflower': '—', 'Pepper-Chives': '—',
  'Zucchini-Garlic': '—', 'Zucchini-Rosemary': '—', 'Zucchini-Mint': '—',
  'Zucchini-Dill': '✅', 'Zucchini-Nasturtium': '✅', 'Zucchini-Peas': '—',
  'Zucchini-Sunflower': '✅', 'Zucchini-Chives': '—',
  'Garlic-Rosemary': '—', 'Garlic-Mint': '—', 'Garlic-Dill': '—',
  'Garlic-Nasturtium': '—', 'Garlic-Peas': '❌', 'Garlic-Sunflower': '—',
  'Garlic-Chives': '✅',
  'Rosemary-Mint': '—', 'Rosemary-Dill': '—', 'Rosemary-Nasturtium': '—',
  'Rosemary-Peas': '—', 'Rosemary-Sunflower': '—', 'Rosemary-Chives': '—',
  'Mint-Dill': '—', 'Mint-Nasturtium': '—', 'Mint-Peas': '—',
  'Mint-Sunflower': '—', 'Mint-Chives': '—',
  'Dill-Nasturtium': '—', 'Dill-Peas': '—', 'Dill-Sunflower': '—',
  'Dill-Chives': '—',
  'Nasturtium-Peas': '—', 'Nasturtium-Sunflower': '✅', 'Nasturtium-Chives': '—',
  'Peas-Sunflower': '—', 'Peas-Chives': '—',
  'Sunflower-Chives': '—',
};

function getRelationship(p1, p2) {
  if (p1 === p2) return '◆';
  const key1 = `${p1}-${p2}`;
  const key2 = `${p2}-${p1}`;
  return RELATIONSHIPS[key1] || RELATIONSHIPS[key2] || '—';
}

(async () => {
  const reqs = [];
  const values = [];
  const N = MATRIX_PLANTS.length;

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(COMP, 0, 1, 0, N + 2), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Companion Planting Guide!A1', values: [['🌿 COMPANION PLANTING GUIDE — What grows well together']] });
  reqs.push({
    repeatCell: {
      range: gridRange(COMP, 0, 1, 0, N + 2),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.sage),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: COMP, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // ── Legend row 2 ──────────────────────────────────────────────────────────
  values.push({ range: 'Companion Planting Guide!A2', values: [
    ['✅ = Good companions (plant together)', '', '❌ = Bad companions (keep apart)', '', '— = Neutral', '', '◆ = Same plant'],
  ]});
  reqs.push({
    repeatCell: {
      range: gridRange(COMP, 1, 2, 0, N + 2),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.parchment),
          textFormat: { fontSize: 9 },
          horizontalAlignment: 'LEFT',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  // ── Matrix header row (row 3 = index 2) ───────────────────────────────────
  const headerRow = ['', ...MATRIX_PLANTS];
  values.push({ range: 'Companion Planting Guide!A3', values: [headerRow] });
  reqs.push({
    repeatCell: {
      range: gridRange(COMP, 2, 3, 0, N + 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.olive),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 8 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: COMP, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 60 }, fields: 'pixelSize',
  }});

  // ── Matrix data rows 4–23 (indices 3–22) ──────────────────────────────────
  const matrixRows = MATRIX_PLANTS.map((rowPlant, r) => {
    const row = [rowPlant];
    MATRIX_PLANTS.forEach((colPlant) => {
      row.push(getRelationship(rowPlant, colPlant));
    });
    return row;
  });
  values.push({ range: 'Companion Planting Guide!A4', values: matrixRows });

  // Row plant labels — col A
  reqs.push({
    repeatCell: {
      range: gridRange(COMP, 3, 3 + N, 0, 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.olive),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 8 },
          horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });

  // Matrix cells coloring
  for (let r = 0; r < N; r++) {
    for (let c = 0; c < N; c++) {
      const rel = getRelationship(MATRIX_PLANTS[r], MATRIX_PLANTS[c]);
      let bg, fg;
      if (rel === '✅') { bg = C.completeBg; fg = C.complete; }
      else if (rel === '❌') { bg = C.warningBg; fg = C.warningRed; }
      else if (rel === '◆') { bg = C.lightGray; fg = C.warmGray; }
      else { bg = C.white; fg = C.warmGray; }
      reqs.push({
        repeatCell: {
          range: gridRange(COMP, 3 + r, 4 + r, 1 + c, 2 + c),
          cell: {
            userEnteredFormat: {
              backgroundColor: hex(bg),
              textFormat: { foregroundColor: hex(fg), bold: rel === '✅' || rel === '❌', fontSize: 10 },
              horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
        },
      });
    }
  }

  // ── Column widths ─────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: COMP, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 100 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: COMP, dimension: 'COLUMNS', startIndex: 1, endIndex: N + 1 },
    properties: { pixelSize: 70 }, fields: 'pixelSize',
  }});

  // Row heights for matrix rows
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: COMP, dimension: 'ROWS', startIndex: 3, endIndex: 3 + N },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // ── Companion Lookup section (below matrix) ───────────────────────────────
  const lookupRow = 3 + N + 2; // row index
  reqs.push({ mergeCells: { range: gridRange(COMP, lookupRow - 1, lookupRow, 0, N + 1), mergeType: 'MERGE_ALL' } });
  values.push({ range: `Companion Planting Guide!A${lookupRow}`, values: [['QUICK LOOKUP — Enter a plant name to see its companions']] });
  reqs.push({
    repeatCell: {
      range: gridRange(COMP, lookupRow - 1, lookupRow, 0, N + 1),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.terracotta),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 11 },
          horizontalAlignment: 'CENTER',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
    },
  });

  values.push({ range: `Companion Planting Guide!A${lookupRow + 1}`, values: [['Plant Name:', '']] });
  values.push({ range: `Companion Planting Guide!B${lookupRow + 1}`, values: [['Tomato']] });

  values.push({ range: `Companion Planting Guide!A${lookupRow + 2}`, values: [['Good Companions:']] });
  values.push({ range: `Companion Planting Guide!B${lookupRow + 2}`, values: [
    [`=IFERROR(TEXTJOIN(", ",TRUE,FILTER(A3:A${3+N},OFFSET(A3:A${3+N},0,MATCH(B${lookupRow+1},B3:${String.fromCharCode(65+N)}3,0))="✅")),"None found")`]
  ]});

  values.push({ range: `Companion Planting Guide!A${lookupRow + 3}`, values: [['Bad Companions:']] });
  values.push({ range: `Companion Planting Guide!B${lookupRow + 3}`, values: [
    [`=IFERROR(TEXTJOIN(", ",TRUE,FILTER(A3:A${3+N},OFFSET(A3:A${3+N},0,MATCH(B${lookupRow+1},B3:${String.fromCharCode(65+N)}3,0))="❌")),"None found")`]
  ]});

  // Format lookup area
  reqs.push({
    repeatCell: {
      range: gridRange(COMP, lookupRow, lookupRow + 4, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 10 }, backgroundColor: hex(C.parchment) } },
      fields: 'userEnteredFormat(textFormat,backgroundColor)',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(COMP, lookupRow, lookupRow + 4, 1, N + 1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.cream2), textFormat: { fontSize: 10 } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    },
  });

  // ── Freeze header rows only (merged title prevents column freeze) ─────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: COMP, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
      fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
    },
  });

  await valuesBatchUpdate(id, values, 'companion-values');
  await batchUpdate(id, reqs, 'companion-format');
  console.log('Companion Planting Guide complete — 20×20 matrix');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
