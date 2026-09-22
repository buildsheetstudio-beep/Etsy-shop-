'use strict';
const { batchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const TT  = sheetMap['✅ Task Tracker'];
const PT  = sheetMap['📦 Packing Tracker'];
const BUD = sheetMap['💰 Moving Budget'];
const QT  = sheetMap['🚛 Moving Company Quotes'];
const COA = sheetMap['📋 Change of Address'];
const UT  = sheetMap['⚡ Utilities & Services'];
const TL  = sheetMap['🗓️ Moving Timeline'];
const INV = sheetMap['📊 Inventory List'];

(async () => {
  const reqs = [];

  // ── Task Tracker CF ────────────────────────────────────────────────────────

  // Done rows — grey-out entire row
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TT, 5, 45, 0, 10)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E6="Done"' }] },
          format: {
            backgroundColor: hex(C.doneBg),
            textFormat: { foregroundColor: hex(C.doneFg), strikethrough: true },
          },
        },
      },
      index: 0,
    },
  });

  // Blocked rows — red highlight
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TT, 5, 45, 0, 10)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E6="Blocked"' }] },
          format: {
            backgroundColor: hex(C.errorBg),
            textFormat: { foregroundColor: hex(C.errorFg) },
          },
        },
      },
      index: 1,
    },
  });

  // In Progress rows — amber highlight
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TT, 5, 45, 0, 10)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E6="In Progress"' }] },
          format: {
            backgroundColor: hex(C.warnBg),
            textFormat: { foregroundColor: hex(C.warnFg) },
          },
        },
      },
      index: 2,
    },
  });

  // Overdue tasks — due date in past, not done — red date cell
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TT, 5, 45, 5, 6)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=AND($F6<TODAY(),$E6<>"Done")' }] },
          format: {
            backgroundColor: hex(C.deepRust),
            textFormat: { foregroundColor: hex(C.white), bold: true },
          },
        },
      },
      index: 3,
    },
  });

  // High priority — bold
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TT, 5, 45, 3, 4)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'High' }] },
          format: { textFormat: { foregroundColor: hex(C.deepRust), bold: true } },
        },
      },
      index: 4,
    },
  });

  // ── Packing Tracker CF ─────────────────────────────────────────────────────

  // Packed status — green
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(PT, 2, 22, 0, 9)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E3="Packed"' }] },
          format: {
            backgroundColor: hex(C.doneBg),
            textFormat: { foregroundColor: hex(C.doneFg) },
          },
        },
      },
      index: 0,
    },
  });

  // Loaded/Arrived/Unpacked — deep green
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(PT, 2, 22, 0, 9)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($E3="Loaded",$E3="Arrived",$E3="Unpacked")' }] },
          format: {
            backgroundColor: hex(C.forestGreen),
            textFormat: { foregroundColor: hex(C.white) },
          },
        },
      },
      index: 1,
    },
  });

  // Fragile items — highlight fragile col
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(PT, 2, 22, 5, 6)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F3=TRUE' }] },
          format: {
            backgroundColor: hex(C.amber),
            textFormat: { foregroundColor: hex(C.white), bold: true },
          },
        },
      },
      index: 2,
    },
  });

  // ── Moving Budget CF ───────────────────────────────────────────────────────

  // Over budget variance — negative — red
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(BUD, 5, 15, 5, 6)],
        booleanRule: {
          condition: { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] },
          format: {
            backgroundColor: hex(C.errorBg),
            textFormat: { foregroundColor: hex(C.errorFg), bold: true },
          },
        },
      },
      index: 0,
    },
  });

  // Under budget — positive variance — green
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(BUD, 5, 15, 5, 6)],
        booleanRule: {
          condition: { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0' }] },
          format: {
            backgroundColor: hex(C.doneBg),
            textFormat: { foregroundColor: hex(C.doneFg) },
          },
        },
      },
      index: 1,
    },
  });

  // Budget total cell over budget warning
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(BUD, 2, 3, 3, 4)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$D3="YES ⚠️"' }] },
          format: {
            backgroundColor: hex(C.errorBg),
            textFormat: { foregroundColor: hex(C.errorFg), bold: true },
          },
        },
      },
      index: 2,
    },
  });

  // ── Moving Company Quotes CF ───────────────────────────────────────────────

  // Lowest quote — green
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(QT, 5, 8, 4, 5)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E6=MIN($E$6:$E$8)' }] },
          format: {
            backgroundColor: hex(C.doneBg),
            textFormat: { foregroundColor: hex(C.doneFg), bold: true },
          },
        },
      },
      index: 0,
    },
  });

  // Highest quote — red
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(QT, 5, 8, 4, 5)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E6=MAX($E$6:$E$8)' }] },
          format: {
            backgroundColor: hex(C.errorBg),
            textFormat: { foregroundColor: hex(C.errorFg) },
          },
        },
      },
      index: 1,
    },
  });

  // No insurance — yellow warning
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(QT, 5, 8, 5, 6)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$F6=FALSE' }] },
          format: {
            backgroundColor: hex(C.warnBg),
            textFormat: { foregroundColor: hex(C.warnFg), bold: true },
          },
        },
      },
      index: 2,
    },
  });

  // ── Change of Address CF ───────────────────────────────────────────────────

  // Done items (date notified filled) — green
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(COA, 5, 50, 0, 8)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G6<>""' }] },
          format: {
            backgroundColor: hex(C.doneBg),
            textFormat: { foregroundColor: hex(C.doneFg) },
          },
        },
      },
      index: 0,
    },
  });

  // ── Utilities CF ──────────────────────────────────────────────────────────

  // Old home — Disconnected — green
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(UT, 5, 15, 6, 7)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Disconnected' }] },
          format: {
            backgroundColor: hex(C.doneBg),
            textFormat: { foregroundColor: hex(C.doneFg) },
          },
        },
      },
      index: 0,
    },
  });

  // Old home — Cancellation Requested — amber
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(UT, 5, 15, 6, 7)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Cancellation Requested' }] },
          format: {
            backgroundColor: hex(C.warnBg),
            textFormat: { foregroundColor: hex(C.warnFg) },
          },
        },
      },
      index: 1,
    },
  });

  // New home — Connected/Active — green
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(UT, 5, 15, 10, 11)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=OR($K6="Connected",$K6="Active")' }] },
          format: {
            backgroundColor: hex(C.doneBg),
            textFormat: { foregroundColor: hex(C.doneFg) },
          },
        },
      },
      index: 2,
    },
  });

  // New home — Not Started — red
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(UT, 5, 15, 10, 11)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Not Started' }] },
          format: {
            backgroundColor: hex(C.errorBg),
            textFormat: { foregroundColor: hex(C.errorFg) },
          },
        },
      },
      index: 3,
    },
  });

  // ── Moving Timeline CF ────────────────────────────────────────────────────

  // Done phases — green
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TL, 2, 10, 0, 9)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E3="Done"' }] },
          format: {
            backgroundColor: hex(C.doneBg),
            textFormat: { foregroundColor: hex(C.doneFg) },
          },
        },
      },
      index: 0,
    },
  });

  // In Progress phases — amber
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(TL, 2, 10, 0, 9)],
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E3="In Progress"' }] },
          format: {
            backgroundColor: hex(C.warnBg),
            textFormat: { foregroundColor: hex(C.warnFg) },
          },
        },
      },
      index: 1,
    },
  });

  // ── Inventory CF ──────────────────────────────────────────────────────────

  // Poor condition — red
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(INV, 5, 15, 8, 9)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Poor' }] },
          format: {
            backgroundColor: hex(C.errorBg),
            textFormat: { foregroundColor: hex(C.errorFg), bold: true },
          },
        },
      },
      index: 0,
    },
  });

  // Excellent condition — green
  reqs.push({
    addConditionalFormatRule: {
      rule: {
        ranges: [gridRange(INV, 5, 15, 8, 9)],
        booleanRule: {
          condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: 'Excellent' }] },
          format: {
            backgroundColor: hex(C.doneBg),
            textFormat: { foregroundColor: hex(C.doneFg) },
          },
        },
      },
      index: 1,
    },
  });

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
