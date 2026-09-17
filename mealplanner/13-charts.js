'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DSH = sheetMap['📊 Dashboard'];
const GBD = sheetMap['💰 Grocery Budget'];
const NGO = sheetMap['🎯 Nutrition Goals'];

// Chart summary data will be written to Dashboard rows 35-65 (hidden below main content)
// Row 35 (idx 34): Chart data label
// Rows 36-40 (idx 35-39): Meal Type distribution (5 types)
// Rows 42-51 (idx 41-50): Cuisine distribution (10 types)
// Rows 53-62 (idx 52-61): Grocery Category spend (10 categories)

const mealTypes = ['Breakfast','Lunch','Dinner','Snack','Dessert'];
const cuisines  = ['American','Italian','Mexican','Asian','Mediterranean','Indian','French','Greek','Thai','Japanese'];
const groceryCats = ['Produce','Dairy','Meat','Seafood','Grains','Pantry','Frozen','Bakery','Beverages','Deli'];

function srcRange(sheetId, startRow, endRow, startCol, endCol) {
  return { sheetId, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: startCol, endColumnIndex: endCol };
}

function overlayPos(row, col, widthPx, heightPx) {
  return {
    overlayPosition: {
      anchorCell: { sheetId: DSH, rowIndex: row, columnIndex: col },
      offsetXPixels: 4, offsetYPixels: 4,
      widthPixels: widthPx, heightPixels: heightPx,
    },
  };
}

(async () => {
  // ── Write chart summary tables to dashboard rows 35-65 ────────────────────
  const mealTypeRows = mealTypes.map(t => [
    t, `=COUNTIF('📖 Recipe Book'!$D:$D,"${t}")`,
  ]);
  const cuisineRows = cuisines.map(c => [
    c, `=COUNTIF('📖 Recipe Book'!$C:$C,"${c}")`,
  ]);
  const categoryRows = groceryCats.map(cat => [
    cat, `=IFERROR(SUMIF('🛒 Grocery List'!$B:$B,"${cat}",'🛒 Grocery List'!$E:$E),0)`,
  ]);

  const chartData = [
    { range: "'📊 Dashboard'!A35", values: [['— Chart Data (Internal) —']] },
    { range: "'📊 Dashboard'!A36", values: [['Meal Type','Recipe Count']] },
    { range: "'📊 Dashboard'!A37", values: mealTypeRows },
    { range: "'📊 Dashboard'!A42", values: [['Cuisine','Recipe Count']] },
    { range: "'📊 Dashboard'!A43", values: cuisineRows },
    { range: "'📊 Dashboard'!A53", values: [['Category','Est. Spend ($)']] },
    { range: "'📊 Dashboard'!A54", values: categoryRows },
  ];
  await valuesBatchUpdate(id, chartData, 'chart-data');

  // ── Charts ────────────────────────────────────────────────────────────────
  const reqs = [];

  // Chart 1: LINE — Grocery Budget Trend (Budget vs Actual, 6 weeks from GBD)
  reqs.push({ addChart: { chart: {
    spec: {
      title: '💰 Grocery Budget Trend',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
      basicChart: {
        chartType: 'LINE',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Week Of' },
          { position: 'LEFT_AXIS', title: 'Amount ($)' },
        ],
        domains: [{ domain: { sourceRange: { sources: [srcRange(GBD, 2, 8, 0, 1)] } } }],
        series: [
          { series: { sourceRange: { sources: [srcRange(GBD, 2, 8, 1, 2)] } }, targetAxis: 'LEFT_AXIS', color: hex(C.deepBasil) },
          { series: { sourceRange: { sources: [srcRange(GBD, 2, 8, 2, 3)] } }, targetAxis: 'LEFT_AXIS', color: hex(C.warmOrange) },
        ],
        headerCount: 0,
      },
      backgroundColor: hex(C.warmCream),
    },
    position: overlayPos(29, 0, 500, 280),
  }}});

  // Chart 2: DONUT — Recipes by Meal Type (from dashboard summary rows 37-41)
  reqs.push({ addChart: { chart: {
    spec: {
      title: '🍽 Recipes by Meal Type',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0.4,
        domain: { sourceRange: { sources: [srcRange(DSH, 36, 41, 0, 1)] } },
        series: { sourceRange: { sources: [srcRange(DSH, 36, 41, 1, 2)] } },
      },
      backgroundColor: hex(C.warmCream),
    },
    position: overlayPos(29, 4, 450, 280),
  }}});

  // Chart 3: BAR — Calorie Goals vs Actual by Family Member (from NGO rows 5-7, current week)
  reqs.push({ addChart: { chart: {
    spec: {
      title: '🎯 Calorie Goals vs Actual',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
      basicChart: {
        chartType: 'BAR',
        legendPosition: 'BOTTOM_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Calories' },
          { position: 'LEFT_AXIS', title: 'Family Member' },
        ],
        domains: [{ domain: { sourceRange: { sources: [srcRange(NGO, 4, 7, 0, 1)] } } }],
        series: [
          { series: { sourceRange: { sources: [srcRange(NGO, 4, 7, 2, 3)] } }, targetAxis: 'BOTTOM_AXIS', color: hex(C.deepBasil) },
          { series: { sourceRange: { sources: [srcRange(NGO, 4, 7, 3, 4)] } }, targetAxis: 'BOTTOM_AXIS', color: hex(C.warmOrange) },
        ],
        headerCount: 0,
      },
      backgroundColor: hex(C.warmCream),
    },
    position: overlayPos(43, 0, 500, 280),
  }}});

  // Chart 4: COLUMN — Recipes by Cuisine (from dashboard summary rows 43-52)
  reqs.push({ addChart: { chart: {
    spec: {
      title: '🌍 Recipes by Cuisine',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
      basicChart: {
        chartType: 'COLUMN',
        legendPosition: 'NO_LEGEND',
        axis: [
          { position: 'BOTTOM_AXIS', title: 'Cuisine' },
          { position: 'LEFT_AXIS', title: 'Count' },
        ],
        domains: [{ domain: { sourceRange: { sources: [srcRange(DSH, 42, 52, 0, 1)] } } }],
        series: [{
          series: { sourceRange: { sources: [srcRange(DSH, 42, 52, 1, 2)] } },
          targetAxis: 'LEFT_AXIS',
          color: hex(C.goldenMustard),
        }],
        stackedType: 'NOT_STACKED',
      },
      backgroundColor: hex(C.warmCream),
    },
    position: overlayPos(43, 4, 450, 280),
  }}});

  // Chart 5: PIE — Grocery Spend by Category (from dashboard summary rows 54-63)
  reqs.push({ addChart: { chart: {
    spec: {
      title: '🛒 Grocery Spend by Category',
      titleTextFormat: { bold: true, fontSize: 11, foregroundColor: hex(C.darkText) },
      pieChart: {
        legendPosition: 'RIGHT_LEGEND',
        pieHole: 0,
        domain: { sourceRange: { sources: [srcRange(DSH, 53, 63, 0, 1)] } },
        series: { sourceRange: { sources: [srcRange(DSH, 53, 63, 1, 2)] } },
      },
      backgroundColor: hex(C.warmCream),
    },
    position: overlayPos(57, 0, 500, 280),
  }}});

  await batchUpdate(id, reqs, 'charts');
  console.log('Charts complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
