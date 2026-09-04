'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap, url } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const AV  = sheetMap['🔐 Account Vault'];
const DB  = sheetMap['🛡️ Security Dashboard'];
const SUB = sheetMap['💳 Subscriptions & Billing'];
const REF = sheetMap['📋 Reference Data'];

(async () => {
  const reqs = [];

  // Account Vault — 9 frozen rows (title + notice + subtitle + header), 2 frozen cols (# + Service)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: AV, gridProperties: { frozenRowCount: 9, frozenColumnCount: 2 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Security Dashboard — 2 frozen rows (title + subtitle), no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DB, gridProperties: { frozenRowCount: 2, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Subscriptions — 9 frozen rows (title + notices + kpi rows + header), 2 frozen cols (# + Service)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: SUB, gridProperties: { frozenRowCount: 9, frozenColumnCount: 2 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Reference Data — 1 frozen row, already hidden
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true, gridProperties: { frozenRowCount: 1 } },
    fields: 'hidden,gridProperties.frozenRowCount',
  }});

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes applied + Reference Data hidden');
  console.log('\n✅ All-in-One Password & Account Vault COMPLETE!');
  console.log('URL:', url);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
