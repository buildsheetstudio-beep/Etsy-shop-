'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const PLANTS = sheetMap['Plant Library'];

// 35 plants: Name, Type, Season, Sun, Water, Days to Germinate, Days to Harvest, Spacing (in), Zone, Companion Plants, Notes
const PLANT_DATA = [
  // Vegetables
  ['Tomato (Cherry)',      'Vegetable',     'Summer',     'Full Sun',     'High',     7,  65, 24, 'Zone B — Tomatoes',   'Basil, Marigold',       'Indeterminate; stake or cage'],
  ['Tomato (Beefsteak)',   'Vegetable',     'Summer',     'Full Sun',     'High',     8,  80, 36, 'Zone B — Tomatoes',   'Basil, Carrot',         'Needs heavy support'],
  ['Cucumber',             'Vine',          'Summer',     'Full Sun',     'High',     5,  55, 12, 'Zone E — Cucurbits',  'Radish, Nasturtium',    'Train on trellis'],
  ['Zucchini',             'Vegetable',     'Summer',     'Full Sun',     'High',     5,  50, 36, 'Zone E — Cucurbits',  'Nasturtium, Dill',      'Very prolific producer'],
  ['Pumpkin',              'Vine',          'Summer',     'Full Sun',     'Moderate', 7,  90, 60, 'Zone E — Cucurbits',  'Corn, Beans',           'Needs lots of space'],
  ['Bell Pepper',          'Vegetable',     'Summer',     'Full Sun',     'Moderate', 14, 70, 18, 'Zone B — Tomatoes',   'Basil, Carrot',         'Start indoors 8 weeks early'],
  ['Jalapeño',             'Vegetable',     'Summer',     'Full Sun',     'Moderate', 14, 75, 18, 'Zone B — Tomatoes',   'Basil, Marigold',       'Hot pepper; handle with care'],
  ['Eggplant',             'Vegetable',     'Summer',     'Full Sun',     'Moderate', 10, 70, 18, 'Zone B — Tomatoes',   'Basil, Tarragon',       'Loves heat'],
  ['Corn',                 'Vegetable',     'Summer',     'Full Sun',     'High',     7,  75, 12, 'Zone E — Cucurbits',  'Beans, Squash',         'Plant in blocks for pollination'],
  ['Beans (Bush)',         'Vegetable',     'Summer',     'Full Sun',     'Moderate', 8,  55, 6,  'Zone D — Greens',     'Carrot, Squash',        'Fix nitrogen in soil'],
  // Root Vegetables
  ['Carrot',               'Root Vegetable','Spring',     'Full Sun',     'Moderate', 14, 70, 3,  'Zone C — Root Veg',   'Rosemary, Sage',        'Thin to 2" apart'],
  ['Beet',                 'Root Vegetable','Spring',     'Full Sun',     'Moderate', 10, 60, 4,  'Zone C — Root Veg',   'Garlic, Onion',         'Eat greens too'],
  ['Radish',               'Root Vegetable','Spring',     'Full Sun',     'Moderate', 4,  25, 2,  'Zone C — Root Veg',   'Carrot, Cucumber',      'Fast-growing trap crop'],
  ['Turnip',               'Root Vegetable','Fall',       'Full Sun',     'Moderate', 7,  40, 4,  'Zone C — Root Veg',   'Peas, Lettuce',         'Harvest young for best flavor'],
  ['Garlic',               'Root Vegetable','All Season', 'Full Sun',     'Low',      14, 240,6,  'Zone C — Root Veg',   'Roses, Tomato',         'Plant in fall; harvest summer'],
  // Leafy Greens
  ['Lettuce (Romaine)',    'Leafy Green',   'Spring',     'Partial Sun',  'Moderate', 7,  70, 8,  'Zone D — Greens',     'Carrot, Radish',        'Bolt-resistant variety'],
  ['Spinach',              'Leafy Green',   'Spring',     'Partial Shade','High',     7,  45, 6,  'Zone D — Greens',     'Strawberry, Pea',       'Bolt in heat; grow in cool seasons'],
  ['Kale',                 'Leafy Green',   'Fall',       'Partial Sun',  'Moderate', 5,  60, 12, 'Zone D — Greens',     'Dill, Sage',            'Frost makes it sweeter'],
  ['Swiss Chard',          'Leafy Green',   'All Season', 'Partial Sun',  'Moderate', 7,  55, 9,  'Zone D — Greens',     'Beans, Onion',          'Cut-and-come-again'],
  ['Arugula',              'Leafy Green',   'Spring',     'Partial Shade','Moderate', 5,  40, 6,  'Zone D — Greens',     'Lettuce, Spinach',      'Peppery flavor; bolt-prone in heat'],
  // Herbs
  ['Basil',                'Herb',          'Summer',     'Full Sun',     'Moderate', 7,  30, 12, 'Zone A — Herbs',      'Tomato, Pepper',        'Pinch flowers to extend harvest'],
  ['Parsley',              'Herb',          'All Season', 'Partial Sun',  'Moderate', 21, 70, 6,  'Zone A — Herbs',      'Tomato, Asparagus',     'Slow to germinate; biennial'],
  ['Cilantro',             'Herb',          'Spring',     'Partial Sun',  'Moderate', 10, 50, 6,  'Zone A — Herbs',      'Spinach, Dill',         'Bolts in heat; succession sow'],
  ['Dill',                 'Herb',          'Summer',     'Full Sun',     'Low',      7,  40, 6,  'Zone A — Herbs',      'Cucumber, Lettuce',     'Attracts beneficial insects'],
  ['Rosemary',             'Herb',          'Perennial',  'Full Sun',     'Low',      14, 90, 18, 'Zone A — Herbs',      'Beans, Brassicas',      'Drought-tolerant once established'],
  ['Thyme',                'Herb',          'Perennial',  'Full Sun',     'Low',      14, 90, 9,  'Zone A — Herbs',      'Cabbage, Strawberry',   'Ground cover; repels pests'],
  ['Mint',                 'Herb',          'Perennial',  'Partial Shade','High',     10, 60, 12, 'Zone A — Herbs',      'Cabbage, Tomato',       'Contain in pots; very invasive'],
  ['Chives',               'Herb',          'Perennial',  'Full Sun',     'Moderate', 14, 60, 6,  'Zone A — Herbs',      'Carrot, Tomato',        'Repels aphids'],
  // Flowers
  ['Marigold',             'Flower',        'Summer',     'Full Sun',     'Low',      7,  50, 10, 'Zone F — Flowers',    'Tomato, Pepper',        'Nematode deterrent; edible'],
  ['Nasturtium',           'Flower',        'Summer',     'Full Sun',     'Low',      7,  35, 12, 'Zone F — Flowers',    'Squash, Cucumber',      'Trap crop for aphids; edible'],
  ['Sunflower',            'Flower',        'Summer',     'Full Sun',     'Moderate', 7,  80, 18, 'Zone F — Flowers',    'Cucumber, Corn',        'Pollinator magnet'],
  ['Lavender',             'Flower',        'Perennial',  'Full Sun',     'Low',      21, 90, 18, 'Zone G — Perennials', 'Chamomile, Cabbage',    'Drought tolerant; repels moths'],
  // Fruits
  ['Strawberry',           'Fruit',         'Spring',     'Full Sun',     'Moderate', 14, 90, 12, 'Zone G — Perennials', 'Spinach, Lettuce',      'Everbearing variety recommended'],
  ['Peas (Sugar Snap)',    'Vegetable',     'Spring',     'Full Sun',     'Moderate', 7,  60, 2,  'Zone D — Greens',     'Carrot, Turnip',        'Fix nitrogen; support needed'],
  ['Sweet Potato',         'Root Vegetable','Summer',     'Full Sun',     'Moderate', 14, 100,12, 'Zone C — Root Veg',   'Beans, Thyme',          'Needs 4 months; heat-loving'],
];

(async () => {
  const reqs = [];
  const values = [];

  // ── Title row ─────────────────────────────────────────────────────────────
  reqs.push({
    mergeCells: { range: gridRange(PLANTS, 0, 1, 0, 14), mergeType: 'MERGE_ALL' },
  });
  values.push({ range: 'Plant Library!A1', values: [['🌱 PLANT LIBRARY — All-Season Garden Planner 2025']] });

  // ── Header row ────────────────────────────────────────────────────────────
  const headers = [
    'Plant Name', 'Type', 'Season', 'Sun', 'Water',
    'Days to Germinate', 'Days to Harvest', 'Spacing (in)', 'Zone',
    'Companion Plants', 'Notes', 'Status', 'Date Planted', 'Qty Planted',
  ];
  values.push({ range: 'Plant Library!A2', values: [headers] });

  // ── Plant data rows ───────────────────────────────────────────────────────
  const plantRows = PLANT_DATA.map(p => [...p, 'Planned', '', '']);
  values.push({ range: 'Plant Library!A3', values: plantRows });

  await valuesBatchUpdate(id, values, 'plant-values');

  // ── Format title ──────────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(PLANTS, 0, 1, 0, 14),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.olive),
          textFormat: { foregroundColor: hex(C.cream), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PLANTS, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // ── Format header row ─────────────────────────────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(PLANTS, 1, 2, 0, 14),
      cell: {
        userEnteredFormat: {
          backgroundColor: hex(C.terracotta),
          textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
    },
  });
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PLANTS, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 32 }, fields: 'pixelSize',
  }});

  // ── Zebra striping for plant rows ─────────────────────────────────────────
  for (let i = 0; i < PLANT_DATA.length; i++) {
    const bg = i % 2 === 0 ? C.cream2 : C.cream;
    reqs.push({
      repeatCell: {
        range: gridRange(PLANTS, 2 + i, 3 + i, 0, 14),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(bg),
            textFormat: { fontSize: 10 },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
      },
    });
  }

  // ── Row heights for data ──────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: PLANTS, dimension: 'ROWS', startIndex: 2, endIndex: 2 + PLANT_DATA.length },
    properties: { pixelSize: 24 }, fields: 'pixelSize',
  }});

  // ── Column widths ─────────────────────────────────────────────────────────
  const colWidths = [160, 110, 80, 90, 80, 90, 90, 80, 140, 160, 200, 100, 90, 80];
  colWidths.forEach((px, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PLANTS, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: px }, fields: 'pixelSize',
    }});
  });

  // ── Bold plant name col A, left-align notes ───────────────────────────────
  reqs.push({
    repeatCell: {
      range: gridRange(PLANTS, 2, 2 + PLANT_DATA.length, 0, 1),
      cell: { userEnteredFormat: { textFormat: { bold: true } } },
      fields: 'userEnteredFormat.textFormat.bold',
    },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(PLANTS, 2, 2 + PLANT_DATA.length, 10, 11),
      cell: { userEnteredFormat: { wrapStrategy: 'WRAP', horizontalAlignment: 'LEFT' } },
      fields: 'userEnteredFormat(wrapStrategy,horizontalAlignment)',
    },
  });

  // ── Borders ───────────────────────────────────────────────────────────────
  reqs.push({
    updateBorders: {
      range: gridRange(PLANTS, 1, 2 + PLANT_DATA.length, 0, 14),
      innerHorizontal: { style: 'SOLID', color: hex(C.lightGray) },
      innerVertical:   { style: 'SOLID', color: hex(C.lightGray) },
      bottom: { style: 'SOLID_MEDIUM', color: hex(C.olive) },
      top:    { style: 'SOLID_MEDIUM', color: hex(C.olive) },
      left:   { style: 'SOLID_MEDIUM', color: hex(C.olive) },
      right:  { style: 'SOLID_MEDIUM', color: hex(C.olive) },
    },
  });

  // ── Freeze header rows ────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: PLANTS, gridProperties: { frozenRowCount: 2 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  await batchUpdate(id, reqs, 'plant-format');
  console.log('Plant Library complete — 35 plants loaded');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
