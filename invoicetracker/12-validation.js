'use strict';
const { batchUpdate, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const CRM = sheetMap['Client CRM'];
const PS  = sheetMap['Products & Services'];
const IG  = sheetMap['Invoice Generator'];
const IL  = sheetMap['Invoice Log'];
const PT  = sheetMap['Payment Tracker'];
const BS  = sheetMap['Business Setup'];
const DB  = sheetMap['Dashboard'];

const RD  = "'Reference Data'";

function list(values, sheetId, r1, r2, c1, c2) {
  return { setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: { condition: { type: 'ONE_OF_LIST', values: values.map(v => ({ userEnteredValue: String(v) })) }, strict: true, showCustomUi: true },
  }};
}

function rangeVal(formula, sheetId, r1, r2, c1, c2) {
  return { setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: { condition: { type: 'ONE_OF_RANGE', values: [{ userEnteredValue: formula }] }, strict: false, showCustomUi: true },
  }};
}

function checkbox(sheetId, r1, r2, c1, c2) {
  return { setDataValidation: {
    range: gridRange(sheetId, r1, r2, c1, c2),
    rule: { condition: { type: 'BOOLEAN' } },
  }};
}

const invoiceStatuses  = ['Draft','Ready to Send','Sent','Viewed','Partially Paid','Paid','Overdue','Cancelled','Refunded'];
const paymentStatuses  = ['Unpaid','Deposit Paid','Partially Paid','Paid','Overpaid','Refunded','Cancelled'];
const paymentMethods   = ['Cash','Check','Bank Transfer','ACH','Credit Card','Debit Card','PayPal','Stripe','Venmo','Zelle','Other'];
const paymentTerms     = ['Due on Receipt','Net 7','Net 14','Net 15','Net 30','Net 45','Net 60','Custom'];
const deliveryMethods  = ['Email','Printed','Client Portal','In Person','Postal Mail','Other'];
const itemTypes        = ['Product','Service','Hourly Work','Retainer','Deposit','Reimbursement','Shipping','Discount','Other'];
const units            = ['Each','Hour','Day','Week','Month','Project','Session','Package','Mile','Custom'];
const taxTreatments    = ['Taxable','Non-Taxable','Exempt','Tax Included'];
const discountTypes    = ['None','Percentage','Fixed Amount'];
const currencies       = ['USD','CAD','GBP','EUR','AUD','NZD','Other'];
const followUpStatus   = ['Not Needed','Upcoming','Due Today','Overdue','Completed'];
const paymentTypes     = ['Deposit','Partial Payment','Final Payment','Overpayment','Refund','Credit','Adjustment'];
const yesNo            = ['Yes','No'];
const ownPeriods       = ['1','3','5','7'];
const layouts          = ['Classic','Modern','Elegant','Compact'];

(async () => {
  const reqs = [];

  // ── Client CRM ────────────────────────────────────────────────────────────
  reqs.push(list(paymentTerms,    CRM, 5, 305, 12, 13)); // Default Payment Terms col M
  reqs.push(list(currencies,      CRM, 5, 305, 14, 15)); // Currency col O
  reqs.push(list(deliveryMethods, CRM, 5, 305, 15, 16)); // Delivery Method col P
  reqs.push(checkbox(CRM, 5, 305, 16, 17));               // Active? col Q

  // ── Products & Services ──────────────────────────────────────────────────
  reqs.push(list(itemTypes,       PS,  5, 305,  2,  3)); // Item Type col C
  reqs.push(list(units,           PS,  5, 305,  4,  5)); // Unit col E
  reqs.push(list(taxTreatments,   PS,  5, 305,  6,  7)); // Tax Treatment col G
  reqs.push(checkbox(PS, 5, 305,  8,  9));                // Discount Eligible col I
  reqs.push(checkbox(PS, 5, 305,  9, 10));                // Active col J

  // ── Invoice Generator ────────────────────────────────────────────────────
  reqs.push(list(paymentTerms,    IG, 12, 13, 1, 2)); // Payment Terms B13
  reqs.push(list(currencies,      IG, 13, 14, 1, 2)); // Currency B14
  reqs.push(list(deliveryMethods, IG, 16, 17, 1, 2)); // Delivery Method B17
  reqs.push(list(invoiceStatuses, IG, 17, 18, 1, 2)); // Invoice Status B18
  reqs.push(list(discountTypes,   IG, 20, 21, 1, 2)); // Discount Type B21
  // Line item Taxable? col I rows 29-48
  reqs.push(list(yesNo,           IG, 28, 48, 8, 9));
  // Line item Unit col F rows 29-48
  reqs.push(list(units,           IG, 28, 48, 5, 6));

  // ── Invoice Log ─────────────────────────────────────────────────────────
  reqs.push(list(invoiceStatuses,  IL, 5, 1005, 19, 20)); // Invoice Status col T
  reqs.push(list(currencies,       IL, 5, 1005,  9, 10)); // Currency col J
  reqs.push(list(deliveryMethods,  IL, 5, 1005, 24, 25)); // Delivery Method col Y
  reqs.push(list(followUpStatus,   IL, 5, 1005, 23, 24)); // Follow-Up Status col X

  // ── Payment Tracker ─────────────────────────────────────────────────────
  reqs.push(list(paymentTypes,    PT, 5, 1505, 6, 7));   // Payment Type col G
  reqs.push(list(paymentMethods,  PT, 5, 1505, 8, 9));   // Payment Method col I
  reqs.push(checkbox(PT, 5, 1505, 10, 11));               // Deposit? col K
  reqs.push(checkbox(PT, 5, 1505, 11, 12));               // Refund? col L

  // ── Business Setup ───────────────────────────────────────────────────────
  reqs.push(list(currencies,      BS, 11, 12, 1, 2)); // Currency B12
  reqs.push(list(paymentTerms,    BS, 13, 14, 1, 2)); // Default Payment Terms B14
  reqs.push(list(layouts,         BS, 30, 31, 1, 2)); // Selected Layout B31

  // ── Dashboard ────────────────────────────────────────────────────────────
  // Month dropdown: Full Year + months
  const monthOptions = ['Full Year','January','February','March','April','May','June','July','August','September','October','November','December'];
  reqs.push(list(monthOptions, DB, 4, 5, 3, 4)); // D5
  // Status dropdown
  const allStatuses = ['All Statuses', ...invoiceStatuses];
  reqs.push(list(allStatuses, DB, 4, 5, 7, 8)); // H5

  await batchUpdate(id, reqs, 'validations');
  console.log('✓ Data Validation');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
