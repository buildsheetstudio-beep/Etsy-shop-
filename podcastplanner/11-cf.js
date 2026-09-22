'use strict';
const { batchUpdate, gridRange, C, hex } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const CAL = sheetMap['🎙 Content Calendar'];
const GST = sheetMap['👥 Guest Outreach'];
const SPO = sheetMap['💰 Sponsorship Tracker'];
const SOC = sheetMap['📱 Social Promotion'];
const EQP = sheetMap['🎧 Equipment Checklist'];

function cfRule(ranges, formula, bg, textColor) {
  const rule = {
    addConditionalFormatRule: {
      rule: {
        ranges,
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
          format: { backgroundColor: hex(bg) },
        },
      },
      index: 0,
    },
  };
  if (textColor) {
    rule.addConditionalFormatRule.rule.booleanRule.format.textFormat = { foregroundColor: hex(textColor) };
  }
  return rule;
}

(async () => {
  const reqs = [];

  // ── 1. Content Calendar — Status col G (index 6) on full row ────────────
  // Published → sage
  reqs.push(cfRule([gridRange(CAL, 2, 500, 0, 12)], '=$G3="Published"', C.lightSage));
  // Scheduled → amber
  reqs.push(cfRule([gridRange(CAL, 2, 500, 0, 12)], '=$G3="Scheduled"', C.lightAmber));
  // Ready to Record → amber
  reqs.push(cfRule([gridRange(CAL, 2, 500, 0, 12)], '=$G3="Ready to Record"', C.lightAmber));
  // Recording → light blue-gray
  reqs.push(cfRule([gridRange(CAL, 2, 500, 0, 12)], '=$G3="Recording"', C.lightBlueGray));
  // Editing → light blue-gray
  reqs.push(cfRule([gridRange(CAL, 2, 500, 0, 12)], '=$G3="Editing"', C.lightBlueGray));
  // On Hold → light red
  reqs.push(cfRule([gridRange(CAL, 2, 500, 0, 12)], '=$G3="On Hold"', C.lightRed, C.rustRed));

  // ── 2. Guest Outreach — Status col F (index 5) on full row ──────────────
  // Confirmed/Recorded → sage
  reqs.push(cfRule([gridRange(GST, 2, 500, 0, 8)], '=$F3="Confirmed"', C.lightSage));
  reqs.push(cfRule([gridRange(GST, 2, 500, 0, 8)], '=$F3="Recorded"', C.lightSage));
  // Contacted → amber
  reqs.push(cfRule([gridRange(GST, 2, 500, 0, 8)], '=$F3="Contacted"', C.lightAmber));
  // Negotiating → amber
  reqs.push(cfRule([gridRange(GST, 2, 500, 0, 8)], '=$F3="Negotiating"', C.lightAmber));
  // Declined → light red
  reqs.push(cfRule([gridRange(GST, 2, 500, 0, 8)], '=$F3="Declined"', C.lightRed, C.rustRed));
  // Not Contacted → light gray
  reqs.push(cfRule([gridRange(GST, 2, 500, 0, 8)], '=$F3="Not Contacted"', C.lightGray));

  // ── 3. Sponsorship Tracker — Payment Status col E (index 4) on full row ─
  // Paid → sage
  reqs.push(cfRule([gridRange(SPO, 2, 500, 0, 8)], '=$E3="Paid"', C.lightSage));
  // Invoiced → amber
  reqs.push(cfRule([gridRange(SPO, 2, 500, 0, 8)], '=$E3="Invoiced"', C.lightAmber));
  // Pending → amber
  reqs.push(cfRule([gridRange(SPO, 2, 500, 0, 8)], '=$E3="Pending"', C.lightAmber));
  // Overdue → light red
  reqs.push(cfRule([gridRange(SPO, 2, 500, 0, 8)], '=$E3="Overdue"', C.lightRed, C.rustRed));

  // ── 4. Social Media — Posted (col E, index 4) → sage fill on full row ──
  reqs.push(cfRule([gridRange(SOC, 2, 500, 0, 6)], '=$E3=TRUE', C.lightSage));
  // Upcoming & unposted: Scheduled Date within 3 days and not posted → amber
  reqs.push(cfRule(
    [gridRange(SOC, 2, 500, 0, 6)],
    '=AND($E3=FALSE,$D3<>"",TODAY()-$D3>=-3,$D3>=TODAY())',
    C.lightAmber,
  ));

  // ── 5. Equipment Checklist — Status col C (index 2) on full row ─────────
  // Owned → sage
  reqs.push(cfRule([gridRange(EQP, 2, 200, 0, 5)], '=$C3="Owned"', C.lightSage));
  // Rented → amber
  reqs.push(cfRule([gridRange(EQP, 2, 200, 0, 5)], '=$C3="Rented"', C.lightAmber));
  // Need to Buy → light red
  reqs.push(cfRule([gridRange(EQP, 2, 200, 0, 5)], '=$C3="Need to Buy"', C.lightRed));

  await batchUpdate(id, reqs, 'cf-rules');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
