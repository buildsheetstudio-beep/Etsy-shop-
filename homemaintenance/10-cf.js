'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const TASK_SID  = sheetMap['Maintenance Task Log'];
const REP_SID   = sheetMap['Home Repair Log'];
const AST_SID   = sheetMap['Appliances & Systems'];
const REC_SID   = sheetMap['Recurring Maintenance'];
const COST_SID  = sheetMap['Maintenance Costs'];

function cfRule(ranges, formula, bg, textColor) {
  return {
    addConditionalFormatRule: {
      rule: {
        ranges,
        booleanRule: {
          condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: formula }] },
          format: {
            backgroundColor: hex(bg),
            textFormat: textColor ? { foregroundColor: hex(textColor) } : undefined,
          },
        },
      },
      index: 0,
    },
  };
}

(async () => {
  const fmt = [];

  const DS = 6; // 0-indexed first data row (row 7 in sheet = index 6)
  const DE = 500;

  // ────────────────────────────────────────────────────────────────────────────
  // MAINTENANCE TASK LOG — conditional formatting
  // ────────────────────────────────────────────────────────────────────────────
  const TASK_RANGE = [gridRange(TASK_SID, DS, DE, 0, 24)];

  // Status: Complete
  fmt.push(cfRule(TASK_RANGE, `=$H7="Complete"`, C.success + '40', null));
  // Status: In Progress
  fmt.push(cfRule(TASK_RANGE, `=$H7="In Progress"`, C.info + '50', null));
  // Status: Cancelled / Skipped
  fmt.push(cfRule(TASK_RANGE, `=OR($H7="Cancelled",$H7="Skipped")`, C.gray + '40', C.secText));
  // Status: Waiting
  fmt.push(cfRule(TASK_RANGE, `=$H7="Waiting"`, '#FFF9E6', null));
  // Overdue (col T = index 19, 1-indexed = col T)
  fmt.push(cfRule(TASK_RANGE, `=$T7=TRUE`, '#FFE0DC', C.attention));
  // Priority: Urgent
  fmt.push(cfRule([gridRange(TASK_SID, DS, DE, 6, 7)], `=$G7="Urgent"`, C.attention, C.white));
  // Priority: High
  fmt.push(cfRule([gridRange(TASK_SID, DS, DE, 6, 7)], `=$G7="High"`, C.warning, C.text));

  // ────────────────────────────────────────────────────────────────────────────
  // HOME REPAIR LOG — conditional formatting
  // ────────────────────────────────────────────────────────────────────────────
  const REP_RANGE = [gridRange(REP_SID, DS, DE, 0, 25)];

  // Status: Complete
  fmt.push(cfRule(REP_RANGE, `=$I7="Complete"`, C.success + '40', null));
  // Status: In Progress
  fmt.push(cfRule(REP_RANGE, `=$I7="In Progress"`, C.info + '50', null));
  // Status: Waiting for Parts
  fmt.push(cfRule(REP_RANGE, `=$I7="Waiting for Parts"`, '#FFF9E6', null));
  // Status: Cancelled
  fmt.push(cfRule(REP_RANGE, `=$I7="Cancelled"`, C.gray + '40', C.secText));
  // Priority: Urgent
  fmt.push(cfRule([gridRange(REP_SID, DS, DE, 7, 8)], `=$H7="Urgent"`, C.attention, C.white));
  // Priority: High
  fmt.push(cfRule([gridRange(REP_SID, DS, DE, 7, 8)], `=$H7="High"`, C.warning, C.text));
  // Warranty covered: highlight row in info color
  fmt.push(cfRule(REP_RANGE, `=$S7=TRUE`, C.info + '30', null));

  // ────────────────────────────────────────────────────────────────────────────
  // APPLIANCES & SYSTEMS — conditional formatting
  // ────────────────────────────────────────────────────────────────────────────
  const AST_RANGE = [gridRange(AST_SID, 7, 200, 0, 25)]; // data starts row 8 (0-indexed 7)

  // Condition: Excellent → success tint
  fmt.push(cfRule([gridRange(AST_SID, 7, 200, 15, 16)], `=$P8="Excellent"`, C.success, C.white));
  // Condition: Good → olive
  fmt.push(cfRule([gridRange(AST_SID, 7, 200, 15, 16)], `=$P8="Good"`, C.olive, C.text));
  // Condition: Fair → wheat
  fmt.push(cfRule([gridRange(AST_SID, 7, 200, 15, 16)], `=$P8="Fair"`, C.wheat, C.text));
  // Condition: Monitor → warning
  fmt.push(cfRule([gridRange(AST_SID, 7, 200, 15, 16)], `=$P8="Monitor"`, C.warning, C.text));
  // Condition: Service Recommended → attention
  fmt.push(cfRule([gridRange(AST_SID, 7, 200, 15, 16)], `=$P8="Service Recommended"`, C.attention, C.white));
  // Condition: Replacement Planning → dim red
  fmt.push(cfRule([gridRange(AST_SID, 7, 200, 15, 16)], `=$P8="Replacement Planning"`, C.attention, C.white));
  // Warranty status: Active → success
  fmt.push(cfRule([gridRange(AST_SID, 7, 200, 12, 13)], `=$M8="Active"`, C.success, C.white));
  // Warranty status: Expiring Soon → warning
  fmt.push(cfRule([gridRange(AST_SID, 7, 200, 12, 13)], `=$M8="Expiring Soon"`, C.warning, C.text));
  // Warranty status: Expired → gray
  fmt.push(cfRule([gridRange(AST_SID, 7, 200, 12, 13)], `=$M8="Expired"`, C.gray, C.white));
  // Replacement Planning checkbox highlighted
  fmt.push(cfRule(AST_RANGE, `=$W8=TRUE`, C.clay, null));

  // ────────────────────────────────────────────────────────────────────────────
  // RECURRING MAINTENANCE — conditional formatting
  // ────────────────────────────────────────────────────────────────────────────
  const REC_RANGE = [gridRange(REC_SID, DS, 300, 0, 25)];

  // Inactive (Active = FALSE)
  fmt.push(cfRule(REC_RANGE, `=$R7=FALSE`, C.gray + '30', C.secText));
  // Priority: Urgent
  fmt.push(cfRule([gridRange(REC_SID, DS, 300, 6, 7)], `=$G7="Urgent"`, C.attention, C.white));
  // Priority: High
  fmt.push(cfRule([gridRange(REC_SID, DS, 300, 6, 7)], `=$G7="High"`, C.warning, C.text));
  // Next due overdue (past TODAY)
  fmt.push(cfRule([gridRange(REC_SID, DS, 300, 15, 16)], `=AND($P7<>"",ISNUMBER($P7),$P7<TODAY())`, '#FFE0DC', C.attention));

  // ────────────────────────────────────────────────────────────────────────────
  // MAINTENANCE COSTS — conditional formatting
  // ────────────────────────────────────────────────────────────────────────────
  const COST_RANGE = [gridRange(COST_SID, DS, 1000, 0, 20)];

  // Payment status: Paid → light success
  fmt.push(cfRule(COST_RANGE, `=$O7="Paid"`, C.success + '25', null));
  // Payment status: Planned → pale olive
  fmt.push(cfRule(COST_RANGE, `=$O7="Planned"`, C.olive, null));
  // Payment status: Due → attention tint
  fmt.push(cfRule(COST_RANGE, `=$O7="Due"`, '#FFE0DC', null));
  // Reimbursable = TRUE → info tint
  fmt.push(cfRule(COST_RANGE, `=$S7=TRUE`, C.info + '30', null));

  await batchUpdate(id, fmt, 'cf');

  console.log('✓ Conditional formatting complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
