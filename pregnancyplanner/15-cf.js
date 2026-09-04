'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DASH = sheetMap['🌸 Dashboard'];
const WJ   = sheetMap['📅 Weekly Pregnancy Journal'];
const ST   = sheetMap['🤒 Symptom & Wellbeing Tracker'];
const AP   = sheetMap['🏥 Appointments Log'];
const BB   = sheetMap['💰 Baby Budget'];
const NP   = sheetMap['🛏️ Nursery Planner'];
const HB   = sheetMap['🎒 Hospital Bag Checklist'];
const BN   = sheetMap['👶 Baby Names'];
const FM   = sheetMap['🍱 Freezer Meal Planner'];
const PP   = sheetMap['🌙 Postpartum Recovery Tracker'];

(async () => {
  const reqs = [];

  // ── Symptom Tracker: severity CF ─────────────────────────────────────────
  // Severe — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(ST, 5, 210, 0, 12)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E6="Severe"' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex('#9B2335') } },
    },
  }, index: 0 }});
  // Very Severe — deep red bold
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(ST, 5, 210, 0, 12)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E6="Very Severe"' }] },
      format: { backgroundColor: hex('#9B2335'), textFormat: { foregroundColor: hex(C.white), bold: true } },
    },
  }, index: 1 }});
  // Mild — light green tint
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(ST, 5, 210, 0, 12)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E6="Mild"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 2 }});
  // Low mood (1-4) — pink
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(ST, 5, 210, 8, 9)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS_THAN_EQ', values: [{ userEnteredValue: '4' }] },
      format: { backgroundColor: hex(C.moodLowBg), textFormat: { foregroundColor: hex(C.dustyRose), bold: true } },
    },
  }, index: 3 }});
  // High mood (8-10) — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(ST, 5, 210, 8, 9)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '8' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), bold: true } },
    },
  }, index: 4 }});

  // ── Appointments: completed green, upcoming amber ─────────────────────────
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AP, 5, 210, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H6=TRUE' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});
  // Upcoming soon (within 7 days) — warm amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AP, 5, 210, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($H6=FALSE,$A6>TODAY(),$A6<=TODAY()+7)' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg), bold: true } },
    },
  }, index: 1 }});
  // Follow-up needed — lavender highlight
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(AP, 5, 210, 8, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I6=TRUE' }] },
      format: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepLavender), bold: true } },
    },
  }, index: 2 }});

  // ── Baby Budget: over-budget (variance negative) — red ────────────────────
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BB, 5, 17, 5, 6), gridRange(BB, 19, 65, 5, 6)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
      format: { backgroundColor: hex(C.errorBg), textFormat: { foregroundColor: hex('#9B2335'), bold: true } },
    },
  }, index: 0 }});
  // Under budget — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BB, 5, 17, 5, 6)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 1 }});
  // High priority — bold
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BB, 5, 17, 2, 3)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'High' }] },
      format: { textFormat: { foregroundColor: hex('#9B2335'), bold: true } },
    },
  }, index: 2 }});

  // ── Nursery Planner: purchased — green, not purchased high priority — amber ─
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(NP, 5, 37, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G6=TRUE' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});
  // High priority not purchased — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(NP, 5, 37, 0, 10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($E6="High",$G6=FALSE)' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 1 }});

  // ── Hospital Bag: packed — green strikethrough ────────────────────────────
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(HB, 4, 75, 0, 6)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$C5=TRUE' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg), strikethrough: true } },
    },
  }, index: 0 }});

  // ── Baby Names: top score — gold highlight ────────────────────────────────
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BN, 4, 55, 0, 11)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$K5=MAX($K$5:$K$54)' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.taupeGold), bold: true } },
    },
  }, index: 0 }});
  // High score (8+) — lavender
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BN, 4, 55, 6, 9)],
    booleanRule: {
      condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '8' }] },
      format: { backgroundColor: hex(C.whisperLav), textFormat: { foregroundColor: hex(C.deepLavender), bold: true } },
    },
  }, index: 1 }});

  // ── Freezer Meals: status CF ──────────────────────────────────────────────
  // Frozen — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(FM, 4, 35, 0, 11)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G5="Frozen"' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 0 }});
  // Used — grey strikethrough
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(FM, 4, 35, 0, 11)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G5="Used"' }] },
      format: { backgroundColor: hex(C.lavenderMist), textFormat: { foregroundColor: hex(C.warmGrey), strikethrough: true } },
    },
  }, index: 1 }});

  // ── Postpartum: wellbeing CF ──────────────────────────────────────────────
  // Struggling / Needs Support / Crisis — red
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(PP, 5, 17, 2, 3)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($C6="Struggling",$C6="Needs Support",$C6="Crisis — Get Help")' }] },
      format: { backgroundColor: hex('#9B2335'), textFormat: { foregroundColor: hex(C.white), bold: true } },
    },
  }, index: 0 }});
  // Good / Great — green
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(PP, 5, 17, 2, 3)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($C6="Good",$C6="Great")' }] },
      format: { backgroundColor: hex(C.doneBg), textFormat: { foregroundColor: hex(C.doneFg) } },
    },
  }, index: 1 }});
  // Fair — amber
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(PP, 5, 17, 2, 3)],
    booleanRule: {
      condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Fair' }] },
      format: { backgroundColor: hex(C.warnBg), textFormat: { foregroundColor: hex(C.warnFg) } },
    },
  }, index: 2 }});
  // Low energy (1-4) — pink
  reqs.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(PP, 5, 17, 4, 5)],
    booleanRule: {
      condition: { type: 'NUMBER_LESS_THAN_EQ', values: [{ userEnteredValue: '4' }] },
      format: { backgroundColor: hex(C.moodLowBg), textFormat: { foregroundColor: hex(C.dustyRose), bold: true } },
    },
  }, index: 3 }});

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
