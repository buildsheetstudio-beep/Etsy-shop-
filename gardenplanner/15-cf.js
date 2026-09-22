'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const PLANTS = sheetMap['Plant Library'];
const PS     = sheetMap['Planting Schedule'];
const HARV   = sheetMap['Harvest & Yield Tracker'];
const PEST   = sheetMap['Pest & Disease Log'];
const BUD    = sheetMap['Garden Budget'];
const SEED   = sheetMap['Seed Inventory'];
const WAT    = sheetMap['Watering & Care Schedule'];

function cfRule(ranges, condition, format) {
  return { addConditionalFormatRule: { rule: { ranges, booleanRule: { condition, format } }, index: 0 } };
}

(async () => {
  const reqs = [];

  // ── Plant Library: Status column L (col 11) rows 3-37 (indices 2-36) ─────
  const plantStatusRanges = [gridRange(PLANTS, 2, 37, 11, 12)];

  reqs.push(cfRule(plantStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Harvested' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete), bold: true } }
  ));
  reqs.push(cfRule(plantStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Harvest Ready' }] },
    { backgroundColor: hex(C.harvestBg), textFormat: { foregroundColor: hex(C.harvestGold), bold: true } }
  ));
  reqs.push(cfRule(plantStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Growing' }] },
    { backgroundColor: hex(C.growingBg), textFormat: { foregroundColor: hex(C.growingGreen) } }
  ));
  reqs.push(cfRule(plantStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Transplanted' }] },
    { backgroundColor: hex(C.inProgressBg), textFormat: { foregroundColor: hex(C.inProgress) } }
  ));
  reqs.push(cfRule(plantStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Germinated' }] },
    { backgroundColor: hex(C.paleSage), textFormat: { foregroundColor: hex(C.olive) } }
  ));
  reqs.push(cfRule(plantStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Failed' }] },
    { backgroundColor: hex(C.warningBg), textFormat: { foregroundColor: hex(C.warningRed) } }
  ));
  reqs.push(cfRule(plantStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Planned' }] },
    { backgroundColor: hex(C.lightGray), textFormat: { foregroundColor: hex(C.warmGray) } }
  ));

  // ── Planting Schedule: Status column L (col 11) rows 4-18 (indices 3-17) ──
  const psStatusRanges = [gridRange(PS, 3, 18, 11, 12)];
  reqs.push(cfRule(psStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Harvested' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete), bold: true } }
  ));
  reqs.push(cfRule(psStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Growing' }] },
    { backgroundColor: hex(C.growingBg), textFormat: { foregroundColor: hex(C.growingGreen) } }
  ));
  reqs.push(cfRule(psStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Transplanted' }] },
    { backgroundColor: hex(C.inProgressBg), textFormat: { foregroundColor: hex(C.inProgress) } }
  ));
  reqs.push(cfRule(psStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Germinated' }] },
    { backgroundColor: hex(C.paleSage), textFormat: { foregroundColor: hex(C.olive) } }
  ));
  reqs.push(cfRule(psStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Planned' }] },
    { backgroundColor: hex(C.lightGray), textFormat: { foregroundColor: hex(C.warmGray) } }
  ));

  // ── Planting Schedule: Estimated Harvest Date — overdue = red ─────────────
  const psHarvestRanges = [gridRange(PS, 3, 18, 6, 7)];
  reqs.push(cfRule(psHarvestRanges,
    { type: 'DATE_BEFORE', values: [{ relativeDate: 'TODAY' }] },
    { backgroundColor: hex(C.warningBg), textFormat: { foregroundColor: hex(C.warningRed) } }
  ));

  // ── Pest Log: Severity column E (col 4) rows 3-8 (indices 2-7) ────────────
  const pestSevRanges = [gridRange(PEST, 2, 8, 4, 5)];
  reqs.push(cfRule(pestSevRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Critical' }] },
    { backgroundColor: hex(C.warningRed), textFormat: { foregroundColor: hex(C.white), bold: true } }
  ));
  reqs.push(cfRule(pestSevRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'High' }] },
    { backgroundColor: hex(C.warningBg), textFormat: { foregroundColor: hex(C.warningRed), bold: true } }
  ));
  reqs.push(cfRule(pestSevRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Medium' }] },
    { backgroundColor: hex(C.harvestBg), textFormat: { foregroundColor: hex(C.harvestGold) } }
  ));
  reqs.push(cfRule(pestSevRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Low' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete) } }
  ));

  // ── Pest Log: Status column H (col 7) ─────────────────────────────────────
  const pestStatusRanges = [gridRange(PEST, 2, 8, 7, 8)];
  reqs.push(cfRule(pestStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Resolved' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete), bold: true } }
  ));
  reqs.push(cfRule(pestStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Treating' }] },
    { backgroundColor: hex(C.harvestBg), textFormat: { foregroundColor: hex(C.harvestGold) } }
  ));
  reqs.push(cfRule(pestStatusRanges,
    { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Monitoring' }] },
    { backgroundColor: hex(C.inProgressBg), textFormat: { foregroundColor: hex(C.inProgress) } }
  ));

  // ── Budget: Difference column G (col 6) — negative = over budget (red) ────
  const budDiffRanges = [gridRange(BUD, 3, 13, 6, 7)];
  reqs.push(cfRule(budDiffRanges,
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.warningBg), textFormat: { foregroundColor: hex(C.warningRed), bold: true } }
  ));
  reqs.push(cfRule(budDiffRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete) } }
  ));

  // ── Seed Inventory: Remaining column H (col 7) — low stock warning ─────────
  const seedRemRanges = [gridRange(SEED, 2, 12, 7, 8)];
  reqs.push(cfRule(seedRemRanges,
    { type: 'NUMBER_LESS', values: [{ userEnteredValue: '20' }] },
    { backgroundColor: hex(C.warningBg), textFormat: { foregroundColor: hex(C.warningRed), bold: true } }
  ));
  reqs.push(cfRule(seedRemRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '100' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete) } }
  ));

  // ── Watering: Days Until Due column G (col 6) — urgent = red, soon = yellow
  const watDueRanges = [gridRange(WAT, 2, 14, 6, 7)];
  reqs.push(cfRule(watDueRanges,
    { type: 'NUMBER_LESS_THAN_EQ', values: [{ userEnteredValue: '0' }] },
    { backgroundColor: hex(C.warningBg), textFormat: { foregroundColor: hex(C.warningRed), bold: true } }
  ));
  reqs.push(cfRule(watDueRanges,
    { type: 'NUMBER_BETWEEN', values: [{ userEnteredValue: '1' }, { userEnteredValue: '2' }] },
    { backgroundColor: hex(C.harvestBg), textFormat: { foregroundColor: hex(C.harvestGold) } }
  ));
  reqs.push(cfRule(watDueRanges,
    { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '3' }] },
    { backgroundColor: hex(C.completeBg), textFormat: { foregroundColor: hex(C.complete) } }
  ));

  await batchUpdate(id, reqs, 'cf-rules');
  console.log(`Conditional formatting complete — ${reqs.length} rules`);
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
