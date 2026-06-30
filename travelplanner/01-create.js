const { sheets, hex } = require('./lib');
const fs = require('fs');

const tabColors = [
  { r: '#C9A84C' }, // 0 Trip Overview — gold
  { r: '#1A2F5A' }, // 1 Day-by-Day Itinerary — midnight blue
  { r: '#2D6A4F' }, // 2 Travel Budget — forest green
  { r: '#C9A84C' }, // 3 Cost Per Person Calculator — gold
  { r: '#B5830A' }, // 4 Currency Converter — amber
  { r: '#0D1B3E' }, // 5 Research & Reservations — deep navy
  { r: '#2D6A4F' }, // 6 Sightseeing & Activities — forest green
  { r: '#8B1A1A' }, // 7 Restaurant & Food Diary — crimson
  { r: '#1A2F5A' }, // 8 Visa & Entry Requirements — midnight blue
  { r: '#0D1B3E' }, // 9 Packing Checklist — deep navy
  { r: '#94A3B8' }, // 10 Reference Data — grey (hidden)
];

const sheetNames = [
  'Trip Overview',
  'Day-by-Day Itinerary',
  'Travel Budget',
  'Cost Per Person Calculator',
  'Currency Converter',
  'Research & Reservations',
  'Sightseeing & Activities',
  'Restaurant & Food Diary',
  'Visa & Entry Requirements',
  'Packing Checklist',
  'Reference Data',
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Travel Planner' },
      sheets: sheetNames.map((title, i) => ({
        properties: {
          title,
          index: i,
          tabColor: hex(tabColors[i].r),
          gridProperties: { rowCount: 1000, columnCount: 26 },
        },
      })),
    },
  });

  const id = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  fs.writeFileSync(__dirname + '/spreadsheet.json', JSON.stringify({ id, url }, null, 2));
  console.log('Created:', url);
  console.log('ID:', id);
  res.data.sheets.forEach(s => console.log(' ', s.properties.sheetId, s.properties.title));
})().catch(e => { console.error(e.errors || e); process.exit(1); });
