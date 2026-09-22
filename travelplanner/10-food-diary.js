const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 1291844518; // Restaurant & Food Diary

// 8 sample entries
const entries = [
  // Restaurant, Destination, Date, MealType, Cuisine, Cost(£), Rating(1-5), MustReturn?, Dish, Notes
  ['Café de Flore','Paris','2025-07-15','Dinner','French',38,5,true,'Croque Monsieur & French onion soup','Historic literary café, great atmosphere. Book ahead for evenings.'],
  ['Le Relais de l\'Entrecôte','Paris','2025-07-16','Dinner','French',42,5,true,'Steak-frites (unlimited!)','No reservations needed — queue from 7pm. Worth the wait. Secret sauce!'],
  ['Boulangerie Utopie','Paris','2025-07-17','Breakfast','French',8,5,false,'Croissant & pain au chocolat','Best pastries in the 11th arr. Queue out the door by 8am.'],
  ['La Boqueria (market stalls)','Barcelona','2025-07-22','Lunch','Spanish',15,4,true,'Jamón ibérico & manchego board','Touristy but the produce is incredible. Go before noon.'],
  ['Bar Calders','Barcelona','2025-07-23','Dinner','Spanish',32,5,true,'Patatas bravas, calamares, pan con tomate','Neighbourhood gem in Sant Antoni. Outdoor terrace, local vibe.'],
  ['El Nacional','Barcelona','2025-07-25','Dinner','Mediterranean',55,4,false,'Seafood paella + vermouth','Stunning 1930s food hall. Pricey but worth it for the experience.'],
  ['Osteria dell\'Angelo','Rome','2025-07-30','Dinner','Italian',45,5,true,'Cacio e pepe & supplì','Classic Roman trattoria near Vatican. Cash only. Booking essential.'],
  ['Grué Coffee','Rome','2025-08-01','Breakfast','Italian',12,5,false,'Espresso, cornetto & granita','Best coffee in Rome. Standing at the bar is the Roman way!'],
];

(async () => {
  const reqs = [];

  // ── Row 1: Title ──
  reqs.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, 11), mergeType: 'MERGE_ALL' } });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 11),
      cell: {
        userEnteredValue: { stringValue: '🍽 RESTAURANT & FOOD DIARY' },
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
  const headers = ['#','Restaurant / Venue','Destination','Date','Meal Type','Cuisine','Cost (£)','Rating','Must Return?','Dish / Order','Notes & Tips'];
  headers.forEach((h, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 1, 2, c, c+1),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.deepCrimson),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
            wrapStrategy: 'WRAP',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // Chart helper cols (M-O): meal type breakdown
  ['Meal Type','Count','Total Spend'].forEach((h, ci) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 1, 2, 12+ci, 13+ci),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.deepCrimson),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 9 },
            horizontalAlignment: 'CENTER',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // ── Data rows ──
  const totalRows = 100;
  for (let i = 0; i < totalRows; i++) {
    const row = 2 + i;
    const bg = i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow);

    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 0, 1),
        cell: {
          userEnteredValue: { formulaValue: `=IF(B${row+1}="","",ROW()-1)` },
          userEnteredFormat: { backgroundColor: hex(COLOR.champagne), textFormat: { bold: true }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });

    [1,2,3,4,5,6,7,8,9,10].forEach(c => {
      reqs.push({
        repeatCell: {
          range: gridRange(SID, row, row+1, c, c+1),
          cell: {
            userEnteredFormat: {
              backgroundColor: bg,
              numberFormat: c === 3 ? { type: 'DATE', pattern: 'dd mmm yyyy' } :
                            c === 6 ? { type: 'NUMBER', pattern: '£#,##0.00' } : undefined,
              horizontalAlignment: (c === 6 || c === 7 || c === 8) ? 'CENTER' : 'LEFT',
              verticalAlignment: 'MIDDLE',
              wrapStrategy: (c === 1 || c === 9 || c === 10) ? 'WRAP' : 'CLIP',
            },
          },
          fields: 'userEnteredFormat',
        },
      });
    });
  }

  // ── Chart helper: meal type breakdown ──
  const mealTypes = ['Breakfast','Brunch','Lunch','Dinner','Snack / Cafe'];
  mealTypes.forEach((m, i) => {
    const row = 2 + i;
    reqs.push({
      repeatCell: {
        range: gridRange(SID, row, row+1, 12, 15),
        cell: {
          userEnteredFormat: {
            backgroundColor: i % 2 === 0 ? hex(COLOR.white) : hex(COLOR.altRow),
            numberFormat: { type: 'NUMBER', pattern: '£#,##0.00' },
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredFormat',
      },
    });
  });

  // Column widths
  const widths = [35,190,100,100,100,100,80,70,80,160,220];
  widths.forEach((w, c) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });
  [120,80,100].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 12+i, endIndex: 13+i }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 44 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 34 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 102 }, properties: { pixelSize: 55 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'food-diary-structure');

  // ── Values ──
  const data = [];
  const rows = entries.map(e => [e[0],e[1],e[2],e[3],e[4],e[5],e[6],e[7],e[8],e[9]]);
  data.push(vr(`'Restaurant & Food Diary'!B3`, rows));

  // Chart helper values
  const chartRows = mealTypes.map(m => [m, `=COUNTIF(E$3:E$102,"${m}")`, `=SUMIF(E$3:E$102,"${m}",G$3:G$102)`]);
  data.push(vr(`'Restaurant & Food Diary'!M3`, chartRows));

  await valuesBatchUpdate(id, data, 'food-diary-values');

  console.log('Restaurant & Food Diary built with', entries.length, 'entries');
})().catch(e => { console.error(e.errors || e); process.exit(1); });
