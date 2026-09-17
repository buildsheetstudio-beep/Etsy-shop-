const { hex, COLOR, gridRange, batchUpdate, valuesBatchUpdate, vr } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const SID = 2028483192; // Reference Data

const cols = {
  // A: Budget Categories
  budgetCats: ['Accommodation','Transport','Food & Drink','Activities','Shopping','Insurance','Visa & Fees','Tours','Miscellaneous'],
  // B: Destination Countries
  destinations: ['France','Spain','Italy','United Kingdom','Germany','Portugal','Greece','Croatia','Netherlands','Japan','Thailand','USA','Australia','UAE','Mexico','Singapore'],
  // C: Transport Types
  transport: ['Flight','Train','Bus','Car Hire','Ferry','Taxi / Rideshare','Metro / Subway','Walking','Cruise','Bicycle'],
  // D: Reservation Status
  resStatus: ['Confirmed','Pending','Cancelled','Waitlisted','Enquired'],
  // E: Activity Types
  activityTypes: ['Museum / Gallery','Guided Tour','Adventure / Outdoor','Beach / Water','Nature / Wildlife','Food & Drink','Shopping','Entertainment','Sport / Fitness','Spa / Wellness','Cultural Experience','Day Trip'],
  // F: Meal Types
  mealTypes: ['Breakfast','Brunch','Lunch','Dinner','Snack / Cafe','Street Food','Fine Dining','Drinks / Bar','Takeaway'],
  // G: Cuisine Types
  cuisines: ['French','Spanish','Italian','Japanese','Thai','Indian','American','Mediterranean','Mexican','Greek','Local / Regional','Fusion','Middle Eastern','Chinese'],
  // H: Visa Types
  visaTypes: ['No Visa Required','Visa on Arrival','E-Visa / eTA','Embassy Visa Required','Visa-Free (Passport)'],
  // I: Packing Categories
  packingCats: ['Clothing','Footwear','Toiletries & Health','Electronics & Tech','Documents & Finance','Entertainment','Food & Snacks','Miscellaneous'],
  // J: Priority
  priority: ['High','Medium','Low'],
  // K: Ratings
  ratings: ['★★★★★','★★★★☆','★★★☆☆','★★☆☆☆','★☆☆☆☆'],
  // L: Currencies
  currencies: ['GBP','USD','EUR','AUD','JPY','THB','AED','MXN','SGD'],
  // M: Accommodation Types
  accomTypes: ['Hotel','Airbnb / Vacation Rental','Hostel','Resort','Villa','Apartment','Camping / Glamping','Boutique Hotel'],
  // N: Traveller Types
  travellerTypes: ['Solo','Couple','Family','Group / Friends','Business'],
};

const headers = ['Budget Categories','Destinations','Transport','Res. Status','Activity Types','Meal Types','Cuisines','Visa Types','Packing Cats','Priority','Ratings','Currencies','Accommodation','Traveller Types'];

(async () => {
  const reqs = [];

  // Title row
  reqs.push({
    mergeCells: { range: gridRange(SID, 0, 1, 0, 14), mergeType: 'MERGE_ALL' },
  });
  reqs.push({
    repeatCell: {
      range: gridRange(SID, 0, 1, 0, 14),
      cell: {
        userEnteredValue: { stringValue: 'REFERENCE DATA' },
        userEnteredFormat: {
          backgroundColor: hex(COLOR.deepNavy),
          textFormat: { foregroundColor: hex(COLOR.antiqueGold), bold: true, fontSize: 14 },
          horizontalAlignment: 'CENTER',
          verticalAlignment: 'MIDDLE',
        },
      },
      fields: 'userEnteredValue,userEnteredFormat',
    },
  });

  // Header row
  headers.forEach((h, c) => {
    reqs.push({
      repeatCell: {
        range: gridRange(SID, 1, 2, c, c+1),
        cell: {
          userEnteredValue: { stringValue: h },
          userEnteredFormat: {
            backgroundColor: hex(COLOR.midnightBlue),
            textFormat: { foregroundColor: hex(COLOR.white), bold: true, fontSize: 10 },
            horizontalAlignment: 'CENTER',
            verticalAlignment: 'MIDDLE',
          },
        },
        fields: 'userEnteredValue,userEnteredFormat',
      },
    });
  });

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });

  // Column widths
  reqs.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: 0, endIndex: 14 }, properties: { pixelSize: 155 }, fields: 'pixelSize' } });

  await batchUpdate(id, reqs, 'reference-structure');

  // Values
  const data = [];
  const lists = Object.values(cols);
  const maxLen = Math.max(...lists.map(l => l.length));

  const rows = [];
  for (let r = 0; r < maxLen; r++) {
    rows.push(lists.map(l => l[r] || ''));
  }

  data.push(vr(`'Reference Data'!A3`, rows));
  await valuesBatchUpdate(id, data, 'reference-values');

  console.log('Reference Data populated:', Object.keys(cols).map((k,i) => `${headers[i]}(${cols[k].length})`).join(', '));
})().catch(e => { console.error(e.errors || e); process.exit(1); });
