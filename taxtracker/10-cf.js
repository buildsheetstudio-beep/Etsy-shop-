'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const LOG = sheetMap['📝 Deduction Log'];
const QTR = sheetMap['📅 Quarterly Tax'];
const CHK = sheetMap['✅ Doc Checklist'];
const HOM = sheetMap['🏠 Home Office'];
const DSH = sheetMap['📊 Dashboard'];

function cfFormula(sheetId, r1, r2, c1, c2, formula, bg, textColor) {
  const req = {
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(sheetId, r1, r2, c1, c2)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
          format: {},
        },
      },
      index: 0,
    },
  };
  if (bg) req.addConditionalFormatRule.rule.booleanRule.format.backgroundColor = hex(bg);
  if (textColor) req.addConditionalFormatRule.rule.booleanRule.format.textFormat = { foregroundColor: hex(textColor), bold: true };
  return req;
}

(async () => {
  const reqs = [];

  // ── Deduction Log ──
  // Col E (Limit %): amber fill when <1 (50% limit applies — Meals & Entertainment)
  reqs.push(cfFormula(LOG, 2, 200, 4, 5, '=$E3<1', C.lightAmber, null));

  // Col F (Doc?): red text/fill when FALSE (undocumented expense)
  reqs.push(cfFormula(LOG, 2, 200, 5, 6, '=$F3=FALSE', C.lightRed, C.brickRed));

  // Entire row: highlight if undocumented
  reqs.push(cfFormula(LOG, 2, 200, 0, 8, '=$F3=FALSE', C.lightRed, null));

  // Col G (Net Deductible): green if positive
  reqs.push(cfFormula(LOG, 2, 200, 6, 7, '=$G3>0', C.lightSage, null));

  // ── Quarterly Tax ──
  // Rows 4-7: red fill if due date has passed and NOT paid
  reqs.push(cfFormula(QTR, 3, 7, 0, 10, '=AND($B4<TODAY(),$J4=FALSE)', C.lightRed, null));

  // Rows 4-7: green fill if paid
  reqs.push(cfFormula(QTR, 3, 7, 0, 10, '=$J4=TRUE', C.lightSage, null));

  // ── Doc Checklist ──
  // % Complete col H: sage ≥1 (100%), amber ≥0.5 (50-99%), red <0.5 (<50%)
  reqs.push(cfFormula(CHK, 3, 15, 7, 8, '=$H4>=1', C.mutedSage, C.white));
  reqs.push(cfFormula(CHK, 3, 15, 7, 8, '=AND($H4>=0.5,$H4<1)', C.amber, C.white));
  reqs.push(cfFormula(CHK, 3, 15, 7, 8, '=$H4<0.5', C.brickRed, C.white));

  // Entire row in checklist: green if all checked (H=1), amber if partial (H>=0.5), red if low
  reqs.push(cfFormula(CHK, 3, 15, 0, 8, '=$H4>=1', C.lightSage, null));
  reqs.push(cfFormula(CHK, 3, 15, 0, 8, '=AND($H4>=0.5,$H4<1)', C.lightAmber, null));
  reqs.push(cfFormula(CHK, 3, 15, 0, 8, '=$H4<0.5', C.lightRed, null));

  // ── Home Office ──
  // Simplified method block: gold if simplified > actual (rows 12-15, cols 0-4)
  reqs.push(cfFormula(HOM, 11, 16, 0, 5, '=$B14>$B23', C.lightAmber, null));

  // Actual method block: gold if actual > simplified (rows 17-23, cols 0-4)
  reqs.push(cfFormula(HOM, 16, 24, 0, 5, '=$B23>=$B14', C.lightSage, null));

  // Result row: gold bg always
  reqs.push(cfFormula(HOM, 24, 27, 0, 5, '=TRUE', C.paleGold, null));

  // ── Dashboard KPIs ──
  // Documentation rate KPI: green >=80%, amber >=60%, red <60%
  reqs.push(cfFormula(DSH, 4, 7, 8, 10, '=$I5>=0.8', C.lightSage, null));
  reqs.push(cfFormula(DSH, 4, 7, 8, 10, '=AND($I5>=0.6,$I5<0.8)', C.lightAmber, null));
  reqs.push(cfFormula(DSH, 4, 7, 8, 10, '=$I5<0.6', C.lightRed, null));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
