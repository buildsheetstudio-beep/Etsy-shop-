'use strict';
const { sheets, hex, batchUpdate, gridRange, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

// Shared conditional formatting helpers
const cfRule = (sheetId, range, condition, bgHex, textHex) => ({
  addConditionalFormatRule: {
    rule: {
      ranges: [range],
      booleanRule: {
        condition,
        format: {
          backgroundColor: bgHex ? { red: parseInt(bgHex.slice(1,3),16)/255, green: parseInt(bgHex.slice(3,5),16)/255, blue: parseInt(bgHex.slice(5,7),16)/255 } : undefined,
          textFormat: textHex ? { bold: true, foregroundColor: { red: parseInt(textHex.slice(1,3),16)/255, green: parseInt(textHex.slice(3,5),16)/255, blue: parseInt(textHex.slice(5,7),16)/255 } } : undefined,
        },
      },
    },
    index: 0,
  },
});

(async () => {
  const reqs = [];

  // ── Fund Setup & Goals — Status column (V = col 21) ──────────────────────────
  const SETUP_SID = sheetMap['Fund Setup & Goals'];
  const statusRng = gridRange(SETUP_SID, 7, 33, 21, 22); // col V

  const STATUS_CF = [
    { text: 'Goal Reached', bg: '#C8E6C9', fg: '#2E7D32' },
    { text: 'Ahead of Plan', bg: '#DCEDC8', fg: '#558B2F' },
    { text: 'On Track', bg: '#E8F5E9', fg: '#388E3C' },
    { text: 'Behind Plan', bg: '#FFEBEE', fg: '#C62828' },
    { text: 'Paused', bg: '#F5F5F5', fg: '#757575' },
    { text: 'Active', bg: '#E3F2FD', fg: '#1565C0' },
    { text: 'Not Started', bg: '#FFF8E1', fg: '#F57F17' },
  ];
  STATUS_CF.forEach(({ text, bg, fg }) => {
    reqs.push(cfRule(SETUP_SID, statusRng, { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] }, bg, fg));
  });

  // Funding Order column (W = col 22)
  const orderRng = gridRange(SETUP_SID, 7, 33, 22, 23);
  const ORDER_CF = [
    { text: 'Fund First', bg: '#B71C1C', fg: '#FFFFFF' },
    { text: 'High Planning Priority', bg: '#FF6F00', fg: '#FFFFFF' },
    { text: 'Standard Planning Priority', bg: '#E8F5E9', fg: '#2E7D32' },
    { text: 'Lower Planning Priority', bg: '#F5F5F5', fg: '#616161' },
    { text: 'Goal Reached', bg: '#C8E6C9', fg: '#1B5E20' },
    { text: 'Paused', bg: '#ECEFF1', fg: '#90A4AE' },
    { text: 'Insufficient Data', bg: '#FFF3E0', fg: '#E65100' },
  ];
  ORDER_CF.forEach(({ text, bg, fg }) => {
    reqs.push(cfRule(SETUP_SID, orderRng, { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] }, bg, fg));
  });

  // Priority column (E = col 4)
  const priRng = gridRange(SETUP_SID, 7, 33, 4, 5);
  [
    { text: 'Critical', bg: '#FFCDD2', fg: '#B71C1C' },
    { text: 'High', bg: '#FFE0B2', fg: '#E65100' },
    { text: 'Medium', bg: '#FFF9C4', fg: '#F57F17' },
    { text: 'Low', bg: '#F1F8E9', fg: '#558B2F' },
  ].forEach(({ text, bg, fg }) => {
    reqs.push(cfRule(SETUP_SID, priRng, { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] }, bg, fg));
  });

  // ── Contribution Log — Transaction Type column (H = col 7) ───────────────────
  const LOG_SID = sheetMap['Contribution Log'];
  const txTypeRng = gridRange(LOG_SID, 7, 5007, 7, 8);
  const TX_CF = [
    { text: 'Contribution', bg: '#E8F5E9', fg: '#2E7D32' },
    { text: 'Withdrawal', bg: '#FFEBEE', fg: '#C62828' },
    { text: 'Transfer In', bg: '#E3F2FD', fg: '#1565C0' },
    { text: 'Transfer Out', bg: '#EDE7F6', fg: '#4527A0' },
    { text: 'Adjustment Increase', bg: '#F3E5F5', fg: '#6A1B9A' },
    { text: 'Adjustment Decrease', bg: '#FBE9E7', fg: '#BF360C' },
    { text: 'Refund', bg: '#E0F7FA', fg: '#006064' },
  ];
  TX_CF.forEach(({ text, bg, fg }) => {
    reqs.push(cfRule(LOG_SID, txTypeRng, { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] }, bg, fg));
  });

  // Balance Effect column (N = col 13) — positive green, negative red
  const beRng = gridRange(LOG_SID, 7, 5007, 13, 14);
  reqs.push(cfRule(LOG_SID, beRng, { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '0' }] }, '#E8F5E9', '#2E7D32'));
  reqs.push(cfRule(LOG_SID, beRng, { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] }, '#FFEBEE', '#C62828'));

  // ── Goals & Milestones — Status column (K = col 10) ─────────────────────────
  const MILE_SID = sheetMap['Goals & Milestones'];
  const mileStatusRng = gridRange(MILE_SID, 5, 5+69, 10, 11);
  const MILE_CF = [
    { text: 'Achieved', bg: '#C8E6C9', fg: '#2E7D32' },
    { text: 'In Progress', bg: '#E3F2FD', fg: '#1565C0' },
    { text: 'Delayed', bg: '#FFEBEE', fg: '#C62828' },
    { text: 'Not Started', bg: '#FFF8E1', fg: '#F57F17' },
    { text: 'Reassess', bg: '#F3E5F5', fg: '#6A1B9A' },
  ];
  MILE_CF.forEach(({ text, bg, fg }) => {
    reqs.push(cfRule(MILE_SID, mileStatusRng, { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] }, bg, fg));
  });

  // Auto-check column (N = col 13) — TRUE = green
  const autoCheckRng = gridRange(MILE_SID, 5, 5+69, 13, 14);
  reqs.push(cfRule(MILE_SID, autoCheckRng, { type: 'TEXT_EQ', values: [{ userEnteredValue: 'TRUE' }] }, '#C8E6C9', '#2E7D32'));

  // ── Savings Forecast — Status column (Q = col 16) ───────────────────────────
  const FORE_SID = sheetMap['Savings Forecast'];
  const foreStatusRng = gridRange(FORE_SID, 9, 35, 16, 17);
  STATUS_CF.forEach(({ text, bg, fg }) => {
    reqs.push(cfRule(FORE_SID, foreStatusRng, { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] }, bg, fg));
  });

  // % Funded column (P = col 15) — gradient via thresholds
  const pctFundedRng = gridRange(FORE_SID, 9, 35, 15, 16);
  reqs.push(cfRule(FORE_SID, pctFundedRng, { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1' }] }, '#C8E6C9', '#2E7D32'));
  reqs.push(cfRule(FORE_SID, pctFundedRng, { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.75' }] }, '#DCEDC8', '#388E3C'));
  reqs.push(cfRule(FORE_SID, pctFundedRng, { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.5' }] }, '#FFF9C4', '#F9A825'));
  reqs.push(cfRule(FORE_SID, pctFundedRng, { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0.5' }] }, '#FFEBEE', '#C62828'));

  // ── Annual Summary — Status column (I = col 8 in fund section) ───────────────
  const ANN_SID = sheetMap['Annual Summary'];
  const annStatusRng = gridRange(ANN_SID, 32, 58, 8, 9);
  STATUS_CF.forEach(({ text, bg, fg }) => {
    reqs.push(cfRule(ANN_SID, annStatusRng, { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] }, bg, fg));
  });

  // Monthly breakdown net column (D = col 3): positive/negative
  const annNetRng = gridRange(ANN_SID, 13, 26, 3, 4);
  reqs.push(cfRule(ANN_SID, annNetRng, { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '0' }] }, '#E8F5E9', '#2E7D32'));
  reqs.push(cfRule(ANN_SID, annNetRng, { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0' }] }, '#FFEBEE', '#C62828'));

  // Monthly Planner — Plan vs Actual status (col P = col 15)
  const PLAN_SID = sheetMap['Monthly Funding Planner'];
  const planStatusRng = gridRange(PLAN_SID, 11, 511, 15, 16);
  [
    { text: 'Over Funded', bg: '#C8E6C9', fg: '#2E7D32' },
    { text: 'Fully Funded', bg: '#DCEDC8', fg: '#33691E' },
    { text: 'Under Funded', bg: '#FFF9C4', fg: '#F57F17' },
    { text: 'Unfunded', bg: '#FFEBEE', fg: '#C62828' },
    { text: 'Goal Reached', bg: '#E0F7FA', fg: '#006064' },
  ].forEach(({ text, bg, fg }) => {
    reqs.push(cfRule(PLAN_SID, planStatusRng, { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] }, bg, fg));
  });

  // Dashboard — % funded cells: color by value
  const DASH_SID = sheetMap['Sinking Funds Dashboard'];
  const dashPctRng = gridRange(DASH_SID, PRI_ROW_IDX(), PRI_ROW_IDX()+9, 6, 7);
  reqs.push(cfRule(DASH_SID, dashPctRng, { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '1' }] }, '#C8E6C9', '#2E7D32'));
  reqs.push(cfRule(DASH_SID, dashPctRng, { type: 'NUMBER_GREATER_THAN_EQ', values: [{ userEnteredValue: '0.5' }] }, '#FFF9C4', '#F9A825'));
  reqs.push(cfRule(DASH_SID, dashPctRng, { type: 'NUMBER_LESS', values: [{ userEnteredValue: '0.5' }] }, '#FFEBEE', '#C62828'));

  function PRI_ROW_IDX() { return 10; } // row 11 (0-indexed = 10)

  await batchUpdate(id, reqs, 'cf');

  console.log('✓ Conditional formatting complete');
})().catch(e => { console.error(e.message || e); process.exit(1); });
