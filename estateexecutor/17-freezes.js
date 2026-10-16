'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap, url } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DB   = sheetMap['📋 Dashboard'];
const CK   = sheetMap['✅ Executor Checklist'];
const AD   = sheetMap['💼 Assets & Debts'];
const BEN  = sheetMap['👨‍👩‍👧 Beneficiaries'];
const DIST = sheetMap['🏦 Estate Distribution'];
const DA   = sheetMap['💻 Digital Assets & Online Accounts'];
const CL   = sheetMap['📞 Beneficiary Communication Log'];
const PS   = sheetMap['🏛️ Professional Services & Fees'];
const TL   = sheetMap['📅 Estate Timeline & Probate Progress'];
const DR   = sheetMap['📁 Documents Register'];
const NC   = sheetMap['📝 Notes & Contacts'];
const REF  = sheetMap['📋 Reference Data'];

(async () => {
  const reqs = [];

  // Dashboard — 3 frozen rows, no col freeze (full-width merges)
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DB, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Executor Checklist — 3 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: CK, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Assets & Debts — 4 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: AD, gridProperties: { frozenRowCount: 4, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Beneficiaries — 3 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: BEN, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Distribution — 3 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DIST, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Digital Assets — 3 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DA, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Communication Log — 3 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: CL, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Professional Services — 3 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: PS, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Timeline — 3 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: TL, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Documents Register — 3 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DR, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Notes & Contacts — 3 frozen rows, no col freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: NC, gridProperties: { frozenRowCount: 3, frozenColumnCount: 0 } },
    fields: 'gridProperties(frozenRowCount,frozenColumnCount)',
  }});

  // Reference Data — freeze 1 row, hide tab
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: REF, hidden: true, gridProperties: { frozenRowCount: 1 } },
    fields: 'hidden,gridProperties.frozenRowCount',
  }});

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes applied + Reference Data hidden');
  console.log('\n✅ Estate Executor Organiser COMPLETE!');
  console.log('URL:', url);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
