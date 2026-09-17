'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DSH = sheetMap['📊 Dashboard'];
const WMP = sheetMap['📅 Weekly Meal Plan'];
const FIN = sheetMap['🥗 Food Inventory'];
const MPR = sheetMap['🍳 Meal Prep Schedule'];
const GBD = sheetMap['💰 Grocery Budget'];
const NGO = sheetMap['🎯 Nutrition Goals'];

function cf(ranges, rule) {
  return { addConditionalFormatRule: { rule: { ranges, booleanRule: rule }, index: 0 } };
}

(async () => {
  const reqs = [];

  // ── 1. Food Inventory Status (col G = "Fresh" / "Use Soon" / "Expired") ──
  reqs.push(cf([gridRange(FIN, 2, 22, 0, 7)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G3="Fresh"' }] },
    format: { backgroundColor: hex(C.lightSage) },
  }));
  reqs.push(cf([gridRange(FIN, 2, 22, 0, 7)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G3="Use Soon"' }] },
    format: { backgroundColor: hex(C.lightOrange) },
  }));
  reqs.push(cf([gridRange(FIN, 2, 22, 0, 7)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G3="Expired"' }] },
    format: { backgroundColor: hex(C.lightRed) },
  }));

  // ── 2. Grocery Budget Variance (col E): positive=sage (under budget), negative=red (over) ──
  reqs.push(cf([gridRange(GBD, 2, 8, 4, 5)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E3>0' }] },
    format: { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepBasil) } },
  }));
  reqs.push(cf([gridRange(GBD, 2, 8, 4, 5)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E3<0' }] },
    format: { backgroundColor: hex(C.lightRed), textFormat: { foregroundColor: hex(C.deepBerry) } },
  }));

  // ── 3. Meal Prep Schedule: red if Use By Date is past and Portions Made > 0 ──
  reqs.push(cf([gridRange(MPR, 2, 12, 0, 7)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($E3<TODAY(),$F3>0)' }] },
    format: { backgroundColor: hex(C.lightRed), textFormat: { foregroundColor: hex(C.deepBerry), bold: true } },
  }));

  // ── 4. Nutrition Goals Variance% (col G) ──
  // ±10% = sage (on track)
  reqs.push(cf([gridRange(NGO, 2, 8, 6, 7)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($G3>=-0.1,$G3<=0.1)' }] },
    format: { backgroundColor: hex(C.lightSage), textFormat: { foregroundColor: hex(C.deepBasil) } },
  }));
  // Beyond ±10% but within ±25% = orange
  reqs.push(cf([gridRange(NGO, 2, 8, 6, 7)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($G3<-0.1,$G3>0.1)' }] },
    format: { backgroundColor: hex(C.lightOrange), textFormat: { foregroundColor: hex(C.warmOrange) } },
  }));
  // Beyond ±25% = red
  reqs.push(cf([gridRange(NGO, 2, 8, 6, 7)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($G3<-0.25,$G3>0.25)' }] },
    format: { backgroundColor: hex(C.lightRed), textFormat: { foregroundColor: hex(C.deepBerry) } },
  }));

  // ── 5. Weekly Meal Plan: sage highlight when Prepped in Advance = TRUE ──
  reqs.push(cf([gridRange(WMP, 2, 23, 0, 6)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F3=TRUE' }] },
    format: { backgroundColor: hex(C.lightSage) },
  }));

  // ── Dashboard: Inventory alerts — red for Expired, orange for Use Soon ──
  reqs.push(cf([gridRange(DSH, 17, 22, 0, 8)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B18="Expired"' }] },
    format: { backgroundColor: hex(C.lightRed) },
  }));
  reqs.push(cf([gridRange(DSH, 17, 22, 0, 8)], {
    condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$B18="Use Soon"' }] },
    format: { backgroundColor: hex(C.lightOrange) },
  }));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
