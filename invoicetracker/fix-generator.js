'use strict';
// Patches six formula bugs in the Invoice Generator tab:
// 1. B11 Due Date — was self-referential and used wrong Payment Terms ref
// 2. N54 Invoice-Level Discount — wrong column refs (B→N for subtotal)
// 3. N55 Shipping — referenced B23 (Notes) instead of B21 (Shipping)
// 4. N57 Grand Total — referenced B column instead of N column
// 5. N58 Deposit Required — referenced B24 (T&C) instead of B22
// 6. N60/N61 Balance Due / Overpaid — referenced B column instead of N column
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const S = "'Invoice Generator'";

(async () => {
  const fixes = [
    // B11: Due Date (Payment Terms = B12, Invoice Date = B10)
    {
      range: `${S}!B11`,
      values: [['=IFERROR(IF(B12="Due on Receipt",B10,IF(B12="Net 7",B10+7,IF(B12="Net 14",B10+14,IF(B12="Net 15",B10+15,IF(B12="Net 30",B10+30,IF(B12="Net 45",B10+45,IF(B12="Net 60",B10+60,B10))))))),"")']]
    },
    // N54: Invoice-Level Discount (Discount Type = B19, Subtotal = N52, Line Discounts = N53, Discount Value = B20)
    {
      range: `${S}!N54`,
      values: [['=IFERROR(IF(B19="None",0,IF(B19="Percentage",(N52-N53)*B20,IF(B19="Fixed Amount",B20,0))),0)']]
    },
    // N55: Shipping (Shipping input = B21)
    {
      range: `${S}!N55`,
      values: [['=IFERROR(B21,0)']]
    },
    // N57: Grand Total (all totals are in column N)
    {
      range: `${S}!N57`,
      values: [['=IFERROR(N52-N53-N54+N55+N56,0)']]
    },
    // N58: Deposit Required (Deposit input = B22)
    {
      range: `${S}!N58`,
      values: [['=IFERROR(B22,0)']]
    },
    // N60: Balance Due (Grand Total = N57, Payments Received = N59)
    {
      range: `${S}!N60`,
      values: [['=IFERROR(MAX(0,N57-N59),0)']]
    },
    // N61: Amount Overpaid
    {
      range: `${S}!N61`,
      values: [['=IFERROR(MAX(0,N59-N57),0)']]
    },
  ];

  await valuesBatchUpdate(id, fixes, 'fix-generator');
  console.log('✓ Invoice Generator formula fixes applied');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
