'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap, url } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🎉 Dashboard'];
const EL   = sheetMap['📋 Events Log'];
const CAL  = sheetMap['🗓️ Events Calendar'];
const TT   = sheetMap['✅ Task Tracker'];
const BUD  = sheetMap['💰 Event Budget'];
const GL   = sheetMap['👥 Guest List & RSVP'];
const VT   = sheetMap['🏪 Vendor & Supplier Tracker'];
const ROS  = sheetMap['⏱️ Run-of-Show'];
const AN   = sheetMap['📊 Event Analytics'];
const NC   = sheetMap['📝 Event Notes & Contacts'];
const REF  = sheetMap['📋 Reference Data'];

(async () => {
  const reqs = [];

  // Dashboard — freeze 2 rows, no col (title row has full-width merge)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DASH, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Events Log — freeze 2 rows, no col (full-width title merge in row 1)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: EL, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Events Calendar — freeze 4 rows (title, nav, day names, first week row)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: CAL, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Task Tracker — freeze 4 rows (title, summary labels, summary values, header)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: TT, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Budget — freeze 2 rows (title, alloc section header)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: BUD, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Guest List — freeze 4 rows (title, summary labels, summary values, header), no col (full-width merge issues)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: GL, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Vendor Tracker — freeze 4 rows (title, summary labels, summary values, header), no col
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: VT, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Run-of-Show — freeze 4 rows (title, event selector, sub-header, col header)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: ROS, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Analytics — freeze 4 rows (title, KPI labels, KPI values, header), no col
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: AN, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Notes — freeze 2 rows (title, contacts section header)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: NC, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Reference Data — freeze 1 row, hide tab
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true, gridProperties: { frozenRowCount: 1 } },
    fields: 'hidden,gridProperties.frozenRowCount',
  }});

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes applied + Reference Data hidden');
  console.log('\n✅ All-in-One Event Planner COMPLETE!');
  console.log('URL:', url);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
