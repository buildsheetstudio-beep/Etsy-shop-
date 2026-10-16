'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const ML = sheetMap['Master Task Log'];

// Master Task Log columns (0-indexed):
// A=0 TaskID, B=1 #, C=2 TaskName, D=3 Project, E=4 Dept, F=5 AssignedTo
// G=6 Priority, H=7 TaskType, I=8 StartDate, J=9 DueDate, K=10 Status
// L=11 CompletedDate, M=12 DaysRemaining, N=13 EstHrs, O=14 HrsLogged
// P=15 Overdue?, Q=16 RecID, R=17 Tags, S=18 Effort, T=19 Recurrence
// U=20 BlockedBy, V=21 Notes

function cfRule(ranges, condition, format, index) {
  return { addConditionalFormatRule:{ rule:{ ranges, booleanRule:{ condition, format } }, index } };
}
function textEq(val) { return { type:'TEXT_EQ', values:[{ userEnteredValue:val }] }; }
function textContains(val) { return { type:'TEXT_CONTAINS', values:[{ userEnteredValue:val }] }; }
function customFormula(formula) { return { type:'CUSTOM_FORMULA', values:[{ userEnteredValue:formula }] }; }

(async () => {
  const reqs = [];

  // ── Status column K (index 10), rows 5-504 ──
  const statusRange = [gridRange(ML,5,505,10,11)];
  // Full row tints — apply to entire row range
  const fullRowRange = [gridRange(ML,5,505,0,22)];

  // Status: Completed — sage green
  reqs.push(cfRule(statusRange, textEq('Completed'), { backgroundColor:hex(C.success), textFormat:{ foregroundColor:hex(C.white), bold:true } }, 0));
  // Status: In Progress — muted blue
  reqs.push(cfRule(statusRange, textEq('In Progress'), { backgroundColor:hex(C.mutedBlue), textFormat:{ foregroundColor:hex(C.mainText), bold:false } }, 0));
  // Status: Waiting — lavender
  reqs.push(cfRule(statusRange, textEq('Waiting'), { backgroundColor:hex(C.lavender), textFormat:{ foregroundColor:hex(C.mainText) } }, 0));
  // Status: Blocked — attention rust
  reqs.push(cfRule(statusRange, textEq('Blocked'), { backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true } }, 0));
  // Status: In Review — warning amber
  reqs.push(cfRule(statusRange, textEq('In Review'), { backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText) } }, 0));
  // Status: Not Started — neutral grey
  reqs.push(cfRule(statusRange, textEq('Not Started'), { backgroundColor:hex(C.altRow), textFormat:{ foregroundColor:hex(C.secText) } }, 0));
  // Status: Cancelled — dim
  reqs.push(cfRule(statusRange, textEq('Cancelled'), { backgroundColor:hex(C.border), textFormat:{ foregroundColor:hex(C.secText), strikethrough:true } }, 0));

  // ── Priority column G (index 6) ──
  const priorityRange = [gridRange(ML,5,505,6,7)];
  reqs.push(cfRule(priorityRange, textEq('Urgent'), { backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true } }, 0));
  reqs.push(cfRule(priorityRange, textEq('High'), { backgroundColor:hex(C.warning), textFormat:{ foregroundColor:hex(C.mainText), bold:true } }, 0));
  reqs.push(cfRule(priorityRange, textEq('Medium'), { backgroundColor:hex(C.mutedBlue), textFormat:{ foregroundColor:hex(C.mainText) } }, 0));
  reqs.push(cfRule(priorityRange, textEq('Low'), { backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText) } }, 0));

  // ── Overdue flag P (index 15) ──
  const overdueRange = [gridRange(ML,5,505,15,16)];
  reqs.push(cfRule(overdueRange, textEq('Yes'), { backgroundColor:hex(C.attention), textFormat:{ foregroundColor:hex(C.white), bold:true } }, 0));
  reqs.push(cfRule(overdueRange, textEq('No'), { backgroundColor:hex(C.bg), textFormat:{ foregroundColor:hex(C.secText) } }, 0));

  // ── Full row tints based on status (applied to full row range) ──
  // Overdue row tint — custom formula: due date < today and not completed/cancelled
  reqs.push(cfRule(fullRowRange,
    customFormula(`=AND($J6<>"",$J6<TODAY(),$K6<>"Completed",$K6<>"Cancelled")`),
    { backgroundColor:{ red:0.725, green:0.451, blue:0.451, alpha:0.15 } },
    0
  ));
  // Completed row tint — light sage
  reqs.push(cfRule(fullRowRange,
    customFormula(`=$K6="Completed"`),
    { backgroundColor:{ red:0.561, green:0.686, blue:0.573, alpha:0.2 } },
    0
  ));
  // Blocked row tint
  reqs.push(cfRule(fullRowRange,
    customFormula(`=$K6="Blocked"`),
    { backgroundColor:{ red:0.725, green:0.451, blue:0.451, alpha:0.12 } },
    0
  ));
  // Urgent row accent — subtle
  reqs.push(cfRule(fullRowRange,
    customFormula(`=AND($G6="Urgent",$K6<>"Completed",$K6<>"Cancelled")`),
    { backgroundColor:{ red:0.855, green:0.447, blue:0.447, alpha:0.08 } },
    0
  ));

  // ── Days Remaining col M (index 12) — color-code ──
  const daysRange = [gridRange(ML,5,505,12,13)];
  // Overdue (negative/zero and not completed)
  reqs.push(cfRule(daysRange,
    customFormula(`=AND($M6<1,$M6<>"",$K6<>"Completed",$K6<>"Cancelled")`),
    { textFormat:{ foregroundColor:hex(C.attention), bold:true } },
    0
  ));
  // Due soon (1-3 days)
  reqs.push(cfRule(daysRange,
    customFormula(`=AND($M6>=1,$M6<=3,$K6<>"Completed",$K6<>"Cancelled")`),
    { textFormat:{ foregroundColor:{ red:0.8, green:0.5, blue:0.1, alpha:1 }, bold:true } },
    0
  ));

  await batchUpdate(id, reqs, 'cf');
  console.log('Conditional formatting complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
