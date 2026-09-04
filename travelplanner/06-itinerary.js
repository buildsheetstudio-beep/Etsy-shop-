const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 353766069; // Day-by-Day Itinerary

// 5 sample days (Jul 15-19, Paris)
const sampleDays = [
  {
    date: '2025-07-15', dest: 'Paris', overnight: 'Hotel Duquesne Eiffel',
    morning: 'Arrive Paris Gare du Nord via Eurostar. Check in, freshen up.',
    afternoon: 'Stroll along the Seine, visit Notre-Dame exterior (under restoration).',
    evening: 'Dinner at Café de Flore, Saint-Germain-des-Prés.',
    notes: 'Arrival day — relax and explore the neighbourhood',
    budget: 40, actual: 42,
  },
  {
    date: '2025-07-16', dest: 'Paris', overnight: 'Hotel Duquesne Eiffel',
    morning: 'Louvre Museum — book skip-the-line 9am slot. See Mona Lisa & Venus de Milo.',
    afternoon: 'Tuileries Garden picnic lunch. Walk to Musée d\'Orsay.',
    evening: 'Seine River Cruise at sunset. Dinner at Le Marais bistro.',
    notes: 'Art & culture day — wear comfortable shoes!',
    budget: 75, actual: 78,
  },
  {
    date: '2025-07-17', dest: 'Paris', overnight: 'Hotel Duquesne Eiffel',
    morning: 'Montmartre & Sacré-Cœur — explore the artists\' quarter.',
    afternoon: 'Lunch at a local brasserie. Visit Centre Pompidou modern art.',
    evening: 'Eiffel Tower summit visit (booked 8pm slot). Champagne on the terrace!',
    notes: 'Eiffel Tower evening slots: book well in advance',
    budget: 85, actual: 87,
  },
  {
    date: '2025-07-18', dest: 'Paris', overnight: 'Hotel Duquesne Eiffel',
    morning: 'Versailles day trip — Palace & Hall of Mirrors, gardens.',
    afternoon: 'Return to Paris. Shopping in Le Marais for vintage & boutiques.',
    evening: 'Farewell Paris dinner — splurge at a classic French restaurant.',
    notes: 'Versailles is best on weekdays to avoid crowds',
    budget: 110, actual: 115,
  },
  {
    date: '2025-07-19', dest: 'Paris', overnight: 'Hotel Duquesne Eiffel',
    morning: 'Morning at the Rodin Museum, then Rue Cler market for fresh produce.',
    afternoon: 'Afternoon free — Luxembourg Gardens, Panthéon, or St-Germain shops.',
    evening: 'Early dinner near hotel, pack bags for Barcelona travel tomorrow.',
    notes: 'Pack the night before — early train Jul 20',
    budget: 65, actual: 68,
  },
];

(async () => {
  const reqs = [];

  // ── Row 1: Title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 12),
      cell: {
        userEnteredValue: { stringValue: '🗓 DAY-BY-DAY ITINERARY' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 18 },
          horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // ── Row 2: Column headers ──
  const headers = [
    'Day #','Date','Day','Destination','Overnight','Morning Activities','Afternoon Activities','Evening Activities','Notes / Tips','Budget (£)','Actual (£)','Difference (£)'
  ];
  headers.forEach((h, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 1, 2, c, c+1),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.midnightBlue),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // ── Data rows (rows 2 onward, 0-indexed = rows 2-…) ──
  // Total trip = 22 days (Jul 15 – Aug 5)
  const totalDays = 22;

  for (let i = 0; i < totalDays; i++) {
    const row = 2 + i;
    const sample = sampleDays[i] || null;
    const bg = i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow);

    // # (col A = 0)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 1),
        cell: {
          userEnteredValue: { formulaValue: `=IF(B${row+1}="","",ROW()-1)` },
          userEnteredFormat: { backgroundColor: hex(COLOR.champagne), textFormat: { bold: true, foregroundColor: hex(COLOR.deepNavy) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    // Day of week (col C = 2) — formula
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 2, 3),
        cell: {
          userEnteredValue: { formulaValue: `=IF(B${row+1}="","",TEXT(B${row+1},"ddd"))` },
          userEnteredFormat: { backgroundColor: hex(COLOR.formulaBg), horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', textFormat: { foregroundColor: hex(COLOR.midnightBlue), bold: true } },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    // Difference (col L = 11)
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 11, 12),
        cell: {
          userEnteredValue: { formulaValue: `=IF(OR(J${row+1}="",K${row+1}=""),"",K${row+1}-J${row+1})` },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.formulaBg),
            numberFormat: { type: 'NUMBER', pattern: '£#,##0.00' },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    // Input cell styling
    const inputCols = [1,3,4,5,6,7,8,9,10]; // Date,Dest,Overnight,Morning,Afternoon,Evening,Notes,Budget,Actual
    inputCols.forEach(c => {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, row, row+1, c, c+1),
          cell: {
            userEnteredFormat: {
              backgroundColor: bg,
              numberFormat: c === 1 ? { type: 'DATE', pattern: 'dd mmm yyyy' } :
                            (c === 9 || c === 10) ? { type: 'NUMBER', pattern: '£#,##0.00' } : undefined,
              verticalAlignment: 'MIDDLE',
              wrapStrategy: (c >= 5 && c <= 8) ? 'WRAP' : 'CLIP',
            },
          },
          fields: 'userEnteredFormat',
        },
      });
    });
  }

  // ── Totals row ──
  const totRow = 2 + totalDays;
  reqs.push({ mergeCells: { range: gridRange(SID, totRow, totRow+1, 0, 9), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, totRow, totRow+1, 0, 9),
      cell: {
        userEnteredValue: { stringValue: 'TOTALS' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true },
          horizontalAlignment: 'RIGHT', verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });
  [9,10,11].forEach(c => {
    const letter = ['J','K','L'][c-9];
    reqs.push({
      repeatCell: {
        range: gridRange(SID, totRow, totRow+1, c, c+1),
        cell: {
          userEnteredValue: { formulaValue: `=SUM(${letter}3:${letter}${2+totalDays})` },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.antiqueGold),
            textFormat: { foregroundColor: hex(COLOR.deepNavy), bold: true, fontSize: 11 },
            numberFormat: { type: 'NUMBER', pattern: '£#,##0.00' },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // ── Column widths ──
  const widths = [50, 110, 50, 110, 160, 230, 230, 230, 200, 90, 90, 100];
  widths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // ── Row heights ──
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 38 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 2+totalDays }, properties: { pixelSize: 75 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'itinerary-structure');

  // ── Values ──
  const data = [];
  const rows = sampleDays.map(d => [
    d.date,
    '',          // Day col (formula)
    d.dest,
    d.overnight,
    d.morning,
    d.afternoon,
    d.evening,
    d.notes,
    d.budget,
    d.actual,
  ]);
  data.push(vr(`'Day-by-Day Itinerary'!B3`, rows));

  await valuesBatchUpdate(id, data, 'itinerary-values');

  console.log('Day-by-Day Itinerary built with', sampleDays.length, 'sample days (22-day trip total)');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
