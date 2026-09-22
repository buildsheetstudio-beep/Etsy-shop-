'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const JRN = sheetMap['Garden Journal'];

// 12 months of garden journal entries for 2025
const JOURNAL_ENTRIES = [
  {
    month: 'January 2025', weather: 'Cold, overcast, avg 28°F', tasks: 'Reviewed seed catalogs; ordered seeds online; cleaned and sharpened tools; planned garden layout for the new season.',
    observations: 'Ground frozen solid. Garlic planted in October mulched with straw — looks healthy under the cover.',
    gratitude: 'Grateful for the garden rest period — planning season is so exciting!',
    nextMonth: 'Order remaining seeds; start pepper and tomato planning; set up grow light station.',
  },
  {
    month: 'February 2025', weather: 'Mild, some snow, avg 32°F', tasks: 'Set up indoor grow light station; started peppers and tomatoes indoors; organized seed packets alphabetically; prepped raised bed amendments.',
    observations: 'Garlic showing small green tips through mulch — encouraging sign! Seed germination room at 72°F.',
    gratitude: 'Love watching the first tiny seedlings emerge — magic every time.',
    nextMonth: 'Continue indoor seedling care; begin hardening off schedule planning; prep compost.',
  },
  {
    month: 'March 2025', weather: 'Variable, 35-55°F range', tasks: 'Direct sowed lettuce, spinach, radish outdoors under row covers; transplanted tomato and pepper seedlings to bigger pots; amended raised beds with compost.',
    observations: 'Soil temps at 4" = 45°F. Lettuce and spinach sprouted in 7 days. Impressive! Row covers keeping frost off greens.',
    gratitude: 'First outdoor seeds in the ground — season has truly begun!',
    nextMonth: 'Watch for last frost date (April 15); continue succession lettuce sowing; start cucumbers indoors.',
  },
  {
    month: 'April 2025', weather: 'Spring-like, 45-65°F, some rain', tasks: 'Harvested first radishes (Apr 14!); second succession of lettuce sown; started hardening off tomatoes; installed drip irrigation; fertilized raised beds.',
    observations: 'FIRST HARVEST of the season — radishes at 25 days. Beautiful! Tomatoes need more light; added reflective panels. Aphids spotted on kale — hand removed.',
    gratitude: 'That first crunch of a fresh radish pulled straight from the soil is pure joy.',
    nextMonth: 'Last frost past — transplant tomatoes, peppers, basil; direct sow cucumbers, zucchini, beans.',
  },
  {
    month: 'May 2025', weather: 'Warm, 60-78°F, perfect growing conditions', tasks: 'Transplanted all warm-season crops after last frost; direct sowed cucumbers, zucchini, beans, pumpkins; installed tomato cages; mulched all beds with wood chips.',
    observations: 'Garden transformation! 14 tomato plants in. Zucchini germinated in 4 days. Lettuce bolting in the heat — succession saved us. Carrot seedlings thinned to 2".',
    gratitude: 'Watching the garden fill up is the best feeling. All that planning coming to life.',
    nextMonth: 'Monitor tomatoes for pests; first basil harvest; continue cucurbit training on trellis.',
  },
  {
    month: 'June 2025', weather: 'Hot, 75-92°F, dry spell mid-month', tasks: 'First tomato harvest (June 20); cucumber, zucchini production ramping up; treated powdery mildew on zucchini; garlic scapes harvested; set up automatic drip timer for vacation.',
    observations: 'Blossom end rot on 3 beefsteak tomatoes — added calcium spray and adjusted watering consistency. Cucumber mosaic virus found — removed 2 plants. Pumpkins vining aggressively!',
    gratitude: 'Nothing beats a caprese salad with tomatoes from my own garden. Worth every effort.',
    nextMonth: 'Peak tomato season; start planning fall crops (kale, spinach, beets); harvest garlic when tops fall.',
  },
  {
    month: 'July 2025', weather: 'Hot, 80-98°F, drought conditions', tasks: 'Harvesting tomatoes, cucumbers, zucchini daily; direct sowed kale, spinach for fall; garlic harvested and cured; deep-watered trees and perennials.',
    observations: 'Zucchini producing 2-3 per day — more than we can eat! Sharing with neighbors. Drip irrigation paying off in the heat. Kale seedlings struggling in heat — shaded with row cover.',
    gratitude: 'The abundance is overwhelming — what a gift to have too much food to eat!',
    nextMonth: 'Transition to fall crops; remove spent summer plants; amend soil for fall planting.',
  },
  {
    month: 'August 2025', weather: 'Cooling, 65-85°F, some rain', tasks: 'Planted fall brassicas and root vegetables; removed spent cucumber and bean plants; composted pulled plants; planted garlic for next year in one bed.',
    observations: 'Garden shifting to fall mode. Kale loving the cooler temps — growth taking off. Tomatoes slowing — blight spotted on lower leaves. Peppers prolific now.',
    gratitude: 'Grateful for the rhythm of the seasons — summer winding down but autumn abundance ahead.',
    nextMonth: 'First fall harvest; prep beds for winter; plant more garlic; enjoy peak pepper harvest.',
  },
  {
    month: 'September 2025', weather: 'Cool, 50-72°F, frost possible late month', tasks: 'Harvested all remaining tomatoes before frost; pickled cucumbers; canned tomato sauce (8 quarts!); transplanted kale to protected spot; covered pepper plants on cold nights.',
    observations: 'End-of-season tomato bounty — processed 40 lbs in one weekend. Kale thriving after first light frost. Peppers still producing through September!',
    gratitude: 'A kitchen full of preserved garden goods is one of the greatest satisfactions in life.',
    nextMonth: 'Final cleanup; heavy mulch on perennial beds; test and amend soil for winter; reflect on season.',
  },
  {
    month: 'October 2025', weather: 'Cold, 35-58°F, first hard frost Oct 15', tasks: 'Final harvest of peppers, eggplant; cut back perennials; heavy leaf mulch on all beds; planted garlic bulbs for 2026 season; composted annuals.',
    observations: 'Garlic in for 2026! Excited for next season already. Kale surviving light frosts beautifully. Perennial herbs (thyme, rosemary) putting on new growth.',
    gratitude: 'Every season teaches something new. This year: water consistency matters more than fertilizer.',
    nextMonth: 'Season officially closed; plan 2026 improvements; review what worked and what didnt.',
  },
  {
    month: 'November 2025', weather: 'Cold, 28-45°F, snow possible', tasks: 'Final garden walk-through and photos; updated plant records; ordered next year\'s seed catalogs; cleaned and oiled all tools for storage; turned compost pile.',
    observations: 'Kale still producing under frost cloth — amazing cold tolerance! Garlic sprouted tiny green tips — on track. Tools all cleaned and hung for winter.',
    gratitude: 'A full season of growing — so grateful for every meal from this garden.',
    nextMonth: 'Rest, reflect, and dream about next year\'s garden. Seed catalog reading season!',
  },
  {
    month: 'December 2025', weather: 'Winter, 20-38°F, frozen ground', tasks: 'Season review and journaling; compared harvest records year-over-year; planned 2026 garden improvements; bought seeds for next year; covered raised beds with burlap.',
    observations: 'Total 2025 harvest value: ~$57+ from just 10 plants/varieties. ROI on $414 spend = $57 in produce. Plus the joy, health benefits, and satisfaction are priceless.',
    gratitude: '2025 garden season: my most successful yet. 35 varieties attempted, 12 types harvested, 12 months of memories.',
    nextMonth: 'Order seeds by Jan 15; start planning 2026 layout; start peppers indoors Feb 1.',
  },
];

(async () => {
  const reqs = [];
  const values = [];

  // ── Title ─────────────────────────────────────────────────────────────────
  reqs.push({ mergeCells: { range: gridRange(JRN, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  values.push({ range: 'Garden Journal!A1', values: [['📔 GARDEN JOURNAL — 2025 Season Record']] });
  reqs.push({
    repeatCell: {
      range: gridRange(JRN, 0, 1, 0, 12),
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
    range: { sheetId: JRN, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 50 }, fields: 'pixelSize',
  }});

  // ── Column widths ─────────────────────────────────────────────────────────
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: JRN, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 160 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: JRN, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 500 }, fields: 'pixelSize',
  }});

  // ── Each month block ──────────────────────────────────────────────────────
  let currentRow = 1; // 0-indexed
  const MONTH_COLORS = [
    '#E8F4F8', // Jan — icy blue
    '#EBF5FB', // Feb
    '#E8F8F5', // Mar — spring green
    '#EAF7EA', // Apr
    '#F0FAE8', // May
    '#FFFDE7', // Jun — summer gold
    '#FFF9E6', // Jul
    '#FFF3E0', // Aug
    '#FBE9E7', // Sep — autumn
    '#FCE4EC', // Oct
    '#F3E5F5', // Nov
    '#E3F2FD', // Dec — winter blue
  ];

  for (let m = 0; m < JOURNAL_ENTRIES.length; m++) {
    const entry = JOURNAL_ENTRIES[m];
    const bg = MONTH_COLORS[m];
    const headerBg = m < 3 ? C.soil : m < 6 ? C.oliveLight : m < 9 ? C.wheat : C.soilLight;

    // Month header
    reqs.push({ mergeCells: { range: gridRange(JRN, currentRow, currentRow + 1, 0, 2), mergeType: 'MERGE_ALL' } });
    values.push({ range: `Garden Journal!A${currentRow + 1}`, values: [[`🌿 ${entry.month}`]] });
    reqs.push({
      repeatCell: {
        range: gridRange(JRN, currentRow, currentRow + 1, 0, 2),
        cell: {
          userEnteredFormat: {
            backgroundColor: hex(headerBg),
            textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 13 },
            horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
            padding: { left: 12 },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,padding)',
      },
    });
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: JRN, dimension: 'ROWS', startIndex: currentRow, endIndex: currentRow + 1 },
      properties: { pixelSize: 36 }, fields: 'pixelSize',
    }});
    currentRow++;

    // Journal fields
    const fields = [
      { label: '🌤️ Weather', content: entry.weather },
      { label: '✅ Tasks Completed', content: entry.tasks },
      { label: '👁️ Observations', content: entry.observations },
      { label: '🙏 Gratitude', content: entry.gratitude },
      { label: '📌 Focus for Next Month', content: entry.nextMonth },
    ];

    fields.forEach(field => {
      values.push({ range: `Garden Journal!A${currentRow + 1}`, values: [[field.label, field.content]] });
      reqs.push({
        repeatCell: {
          range: gridRange(JRN, currentRow, currentRow + 1, 0, 1),
          cell: {
            userEnteredFormat: {
              backgroundColor: hex(bg),
              textFormat: { bold: true, fontSize: 10, foregroundColor: hex(C.soil) },
              verticalAlignment: 'TOP',
              padding: { left: 8 },
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,padding)',
        },
      });
      reqs.push({
        repeatCell: {
          range: gridRange(JRN, currentRow, currentRow + 1, 1, 2),
          cell: {
            userEnteredFormat: {
              backgroundColor: hex(bg),
              textFormat: { fontSize: 10, foregroundColor: hex(C.charcoal) },
              verticalAlignment: 'TOP',
              wrapStrategy: 'WRAP',
              padding: { left: 8, right: 8 },
            },
          },
          fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment,wrapStrategy,padding)',
        },
      });
      reqs.push({ updateDimensionProperties: {
        range: { sheetId: JRN, dimension: 'ROWS', startIndex: currentRow, endIndex: currentRow + 1 },
        properties: { pixelSize: 72 }, fields: 'pixelSize',
      }});
      currentRow++;
    });

    // Spacer row between months
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: JRN, dimension: 'ROWS', startIndex: currentRow, endIndex: currentRow + 1 },
      properties: { pixelSize: 16 }, fields: 'pixelSize',
    }});
    reqs.push({
      repeatCell: {
        range: gridRange(JRN, currentRow, currentRow + 1, 0, 2),
        cell: { userEnteredFormat: { backgroundColor: hex(C.white) } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
    currentRow++;
  }

  // ── Freeze title row ──────────────────────────────────────────────────────
  reqs.push({
    updateSheetProperties: {
      properties: { sheetId: JRN, gridProperties: { frozenRowCount: 1 } },
      fields: 'gridProperties.frozenRowCount',
    },
  });

  // Expand grid first, then write values, then apply formatting
  await batchUpdate(id, [{
    updateSheetProperties: {
      properties: { sheetId: JRN, gridProperties: { rowCount: currentRow + 10 } },
      fields: 'gridProperties.rowCount',
    },
  }], 'journal-expand');
  await valuesBatchUpdate(id, values, 'journal-values');
  await batchUpdate(id, reqs, 'journal-format');
  console.log('Garden Journal complete — 12 months (Jan–Dec 2025)');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
