'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

(async () => {
  const fmt = [];

  const BD  = sheetMap['Brain Dump'];
  const MT  = sheetMap['Master Tasks'];
  const TB  = sheetMap['Task Breakdown'];
  const DF  = sheetMap['Daily Focus'];
  const WP  = sheetMap['Weekly Planner'];

  // ===== BRAIN DUMP: Ready? column G (col 6) =====
  // Green when TRUE (ready), neutral otherwise
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BD,7,2007,0,10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$G8=TRUE' }] },
      format: { backgroundColor: hex(C.tealTint) },
    },
  }, index: 0 }});

  // Brain Dump: Moved? column H (col 7) — dim when moved
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(BD,7,2007,0,10)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H8=TRUE' }] },
      format: { backgroundColor: hex(C.completedBg), textFormat: { foregroundColor: hex(C.medGray) } },
    },
  }, index: 1 }});

  // ===== MASTER TASKS: Focus Stage color coding (col B = index 1) =====
  const STAGE_CF = [
    { val: 'Now',    bg: C.coralTint,  text: C.coral    },
    { val: 'Next',   bg: C.blueTint,   text: C.blue     },
    { val: 'Later',  bg: C.lavTint,    text: C.lavender },
    { val: 'Event',  bg: C.tealTint,   text: C.teal     },
    { val: 'Revisit',bg: C.sandTint,   text: C.sand     },
  ];
  STAGE_CF.forEach(({ val, bg, text }, i) => {
    fmt.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(MT,7,5007,0,14)],
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$B8="${val}"` }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: i + 2 }});
  });

  // Master Tasks: Status-based row styling
  const STATUS_CF = [
    { val: 'Completed',   bg: C.completedBg,   text: C.medGray  },
    { val: 'Archived',    bg: C.archivedBg,     text: C.medGray  },
    { val: 'Waiting',     bg: C.waitingBg,      text: C.text     },
    { val: 'Rescheduled', bg: C.rescheduledBg,  text: C.text     },
    { val: 'In Progress', bg: C.coralTint,      text: C.coral    },
  ];
  STATUS_CF.forEach(({ val, bg, text }, i) => {
    fmt.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(MT,7,5007,0,14)],
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$E8="${val}"` }] },
        format: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(text) } },
      },
    }, index: i + 7 }});
  });

  // Master Tasks: Overdue highlighting (red bg) — only non-completed
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(MT,7,5007,5,6)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=AND($F8<TODAY(),$E8<>"Completed",$E8<>"Archived",$C8<>"")` }] },
      format: { backgroundColor: hex('#FFE0E0'), textFormat: { foregroundColor: hex('#CC0000'), bold: true } },
    },
  }, index: 12 }});

  // Master Tasks: Due today
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(MT,7,5007,5,6)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=AND($F8=TODAY(),$E8<>"Completed",$E8<>"Archived")` }] },
      format: { backgroundColor: hex(C.coralTint), textFormat: { bold: true, foregroundColor: hex(C.coral) } },
    },
  }, index: 13 }});

  // ===== TASK BREAKDOWN: Step status =====
  const STEP_STATUS_CF = [
    { val: 'Completed',  bg: C.completedBg, text: C.medGray },
    { val: 'In Progress',bg: C.blueTint,    text: C.blue    },
    { val: 'Blocked',    bg: '#FFE0E0',     text: '#CC0000' },
  ];
  STEP_STATUS_CF.forEach(({ val, bg, text }, i) => {
    fmt.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(TB,25,600,0,11)],
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$H26="${val}"` }] },
        format: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(text) } },
      },
    }, index: i + 14 }});
  });

  // Task Breakdown: Done? column I checkbox = TRUE → strikethrough
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(TB,25,600,0,11)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I26=TRUE' }] },
      format: { textFormat: { strikethrough: true, foregroundColor: hex(C.medGray) }, backgroundColor: hex(C.completedBg) },
    },
  }, index: 17 }});

  // ===== DAILY FOCUS: Done? col E for priorities, col I for task queue =====
  // Priorities done (row 13-15, col E = index 4)
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DF,12,15,0,12)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$E13=TRUE' }] },
      format: { backgroundColor: hex(C.completedBg), textFormat: { strikethrough: true, foregroundColor: hex(C.medGray) } },
    },
  }, index: 18 }});

  // Task queue done (col I = index 8, rows 19+)
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(DF,18,49,0,12)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$I19=TRUE' }] },
      format: { backgroundColor: hex(C.completedBg), textFormat: { strikethrough: true, foregroundColor: hex(C.medGray) } },
    },
  }, index: 19 }});

  // ===== WEEKLY PLANNER: Done? column H (index 7), task rows =====
  fmt.push({ addConditionalFormatRule: { rule: {
    ranges: [gridRange(WP,13,200,0,14)],
    booleanRule: {
      condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=$H14=TRUE' }] },
      format: { backgroundColor: hex(C.completedBg), textFormat: { strikethrough: true, foregroundColor: hex(C.medGray) } },
    },
  }, index: 20 }});

  // Weekly Planner: Focus stage coloring
  STAGE_CF.forEach(({ val, bg }, i) => {
    fmt.push({ addConditionalFormatRule: { rule: {
      ranges: [gridRange(WP,13,200,0,14)],
      booleanRule: {
        condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: `=$B14="${val}"` }] },
        format: { backgroundColor: hex(bg) },
      },
    }, index: i + 21 }});
  });

  await batchUpdate(id, fmt, '11-cf');
  console.log('✅  Conditional formatting done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
