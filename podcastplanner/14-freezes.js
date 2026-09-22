'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const REF = sheetMap['📋 Reference Data'];
const DSH = sheetMap['📊 Dashboard'];
const CAL = sheetMap['🎙 Content Calendar'];
const GST = sheetMap['👥 Guest Outreach'];
const SPO = sheetMap['💰 Sponsorship Tracker'];
const ANA = sheetMap['📈 Episode Analytics'];
const SOC = sheetMap['📱 Social Promotion'];
const EQP = sheetMap['🎧 Equipment Checklist'];
const MCL = sheetMap['📅 Monthly Calendar'];

function freeze(sheetId, rows, cols) {
  return {
    updateSheetProperties: {
      properties: {
        sheetId,
        gridProperties: { frozenRowCount: rows, frozenColumnCount: cols },
      },
      fields: 'gridProperties.frozenRowCount,gridProperties.frozenColumnCount',
    },
  };
}

function hideSheet(sheetId) {
  return {
    updateSheetProperties: {
      properties: { sheetId, hidden: true },
      fields: 'hidden',
    },
  };
}

(async () => {
  const reqs = [
    // Dashboard: freeze rows 1-3 (title + subtitle + spacer, i.e. frozenRowCount=3)
    freeze(DSH, 3, 0),
    // Content Calendar: freeze row 1 (title) + row 2 (header) = 2 rows; freeze col A (standalone = safe)
    freeze(CAL, 2, 1),
    // Guest Outreach: freeze 2 rows + col A (standalone = safe)
    freeze(GST, 2, 1),
    // Sponsorship Tracker: freeze 2 rows + col A (standalone = safe)
    freeze(SPO, 2, 1),
    // Episode Analytics: freeze 2 rows + col A (standalone = safe)
    freeze(ANA, 2, 1),
    // Social Media Checklist: freeze row 1 only (full-width merge, no col freeze)
    freeze(SOC, 1, 0),
    // Equipment Checklist: freeze row 1 only (full-width merge, no col freeze)
    freeze(EQP, 1, 0),
    // Monthly Calendar: freeze rows 1-2 (title + year/month selector)
    freeze(MCL, 2, 0),
    // Reference Data: freeze header row, then hide
    freeze(REF, 1, 0),
    hideSheet(REF),
  ];

  await batchUpdate(id, reqs, 'freezes');
  console.log('Freezes & hide complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
