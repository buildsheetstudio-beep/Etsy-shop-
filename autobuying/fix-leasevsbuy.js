'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const S  = "'Lease vs Buy'";
const VQ = "'Vehicle Details & Quotes'";
const BS = "'Buyer Setup'";

// Buy column D — mirrored to match lease label rows C6:C24
//   D6  = Vehicle ID
//   D7  = MSRP / Listing Price      (VQ!$M)
//   D8  = Out-the-Door Price        (VQ!$W)      ← mirrors "Negotiated Cap Cost"
//   D9  = Down Payment              (BS!B18)     ← mirrors "Cap Cost Reduction"
//   D10 = Loan Term (months)        (VQ!$Y)      ← mirrors "Lease Term"
//   D11 = APR                       (VQ!$X)      ← mirrors "Money Factor"
//   D12 = ""                                     ← mirrors "Residual %"
//   D13 = Annual Mileage            (BS!B21)     ← mirrors "Annual Mileage"
//   D14 = ""                                     ← mirrors "Acquisition Fee"
//   D15 = Documentation Fee         (VQ!$P, ref) ← mirrors "Doc Fee"
//   D16 = Registration              (VQ!$Q, ref) ← mirrors "Registration"
//   D17 = ""                                     ← mirrors "Sales Tax Rate" (in OTD)
//   D18 = Monthly Payment           (PMT formula)← mirrors "Monthly Payment"
//   D19-D22 = ""                                 ← lease-only fee rows
//   D23 = Insurance (monthly)       (VQ!$AA)     ← mirrors "Insurance (Monthly $)"
//   D24 = Maintenance (monthly)     (VQ!$AB)     ← mirrors "Maintenance (Monthly $)"
//
// Comparison metrics (compStart=30, rows 31-39):
//   D31 Upfront Cash Required  = D9
//   D32 Monthly Payment        = D18
//   D33 Total Payments         = D18 * D10
//   D34 End-of-Term Fees       = 0
//   D35 Equity / Asset Value   = D8 * (1-0.12)^(D10/12)
//   D36 Total Cash Outflow     = D31 + D33
//   D37 Estimated Net Cost     = D36 - D35
//   D38 Cost per Month         = D37 / D10
//   D39 Cost per Mile          = D37 / (D13 * D10 / 12)

(async () => {
  const data = [];
  const compStart = 30;

  // ── Lease input cells — INDEX/MATCH so they update when C6 changes ──
  data.push({ range: `${S}!C7`, values: [[
    `=IFERROR(INDEX(${VQ}!$M$6:$M$105,MATCH(C6,${VQ}!$A$6:$A$105,0)),36500)`
  ]]});
  data.push({ range: `${S}!C8`, values: [[
    `=IFERROR(INDEX(${VQ}!$N$6:$N$105,MATCH(C6,${VQ}!$A$6:$A$105,0)),34800)`
  ]]});
  data.push({ range: `${S}!C15`, values: [[
    `=IFERROR(INDEX(${VQ}!$P$6:$P$105,MATCH(C6,${VQ}!$A$6:$A$105,0)),495)`
  ]]});
  data.push({ range: `${S}!C16`, values: [[
    `=IFERROR(INDEX(${VQ}!$Q$6:$Q$105,MATCH(C6,${VQ}!$A$6:$A$105,0)),285)`
  ]]});
  data.push({ range: `${S}!C23`, values: [[
    `=IFERROR(INDEX(${VQ}!$AA$6:$AA$105,MATCH(C6,${VQ}!$A$6:$A$105,0)),150)`
  ]]});
  data.push({ range: `${S}!C24`, values: [[
    `=IFERROR(INDEX(${VQ}!$AB$6:$AB$105,MATCH(C6,${VQ}!$A$6:$A$105,0)),55)`
  ]]});

  // ── C18 — Monthly Lease Payment ──
  data.push({ range: `${S}!C18`, values: [[
    `=IFERROR(((C8-C9-C7*C12)/C10+(C8-C9+C7*C12)*C11)*(1+C17),0)`
  ]]});

  // ── Buy column D — structured to match label rows ──

  // D7: MSRP / Listing Price
  data.push({ range: `${S}!D7`, values: [[
    `=IFERROR(INDEX(${VQ}!$M$6:$M$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),0)`
  ]]});
  // D8: Out-the-Door Price
  data.push({ range: `${S}!D8`, values: [[
    `=IFERROR(INDEX(${VQ}!$W$6:$W$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),0)`
  ]]});
  // D9: Down Payment
  data.push({ range: `${S}!D9`, values: [[
    `=IFERROR(${BS}!B18,0)`
  ]]});
  // D10: Loan Term (months)
  data.push({ range: `${S}!D10`, values: [[
    `=IFERROR(INDEX(${VQ}!$Y$6:$Y$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),60)`
  ]]});
  // D11: APR
  data.push({ range: `${S}!D11`, values: [[
    `=IFERROR(INDEX(${VQ}!$X$6:$X$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),0)`
  ]]});
  // D12: empty — no residual % for buy
  data.push({ range: `${S}!D12`, values: [['']]});
  // D13: Annual Mileage
  data.push({ range: `${S}!D13`, values: [[
    `=IFERROR(${BS}!B21,12000)`
  ]]});
  // D14: empty — no acquisition fee for buy
  data.push({ range: `${S}!D14`, values: [['']]});
  // D15: Documentation Fee (reference — already included in OTD)
  data.push({ range: `${S}!D15`, values: [[
    `=IFERROR(INDEX(${VQ}!$P$6:$P$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),0)`
  ]]});
  // D16: Registration (reference — already included in OTD)
  data.push({ range: `${S}!D16`, values: [[
    `=IFERROR(INDEX(${VQ}!$Q$6:$Q$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),0)`
  ]]});
  // D17: empty — sales tax is included in OTD
  data.push({ range: `${S}!D17`, values: [['']]});
  // D18: Monthly Payment — principal = OTD - DownPayment - TradeEquity + TradeRollover
  data.push({ range: `${S}!D18`, values: [[
    `=IFERROR(-PMT(D11/12,D10,MAX(0,D8-D9-IFERROR(${BS}!B19,0)+IFERROR(${BS}!B20,0))),0)`
  ]]});
  // D19-D22: empty — lease-only fee rows
  data.push({ range: `${S}!D19:D22`, values: [[''],[''],[''],['']].map(v=>[v[0]])});
  // D23: Insurance (monthly)
  data.push({ range: `${S}!D23`, values: [[
    `=IFERROR(INDEX(${VQ}!$AA$6:$AA$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),120)`
  ]]});
  // D24: Maintenance (monthly)
  data.push({ range: `${S}!D24`, values: [[
    `=IFERROR(INDEX(${VQ}!$AB$6:$AB$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),60)`
  ]]});

  // ── Comparison metrics — LEASE column (C31–C39) ──
  // C31: Upfront = Cap Cost Reduction + Acquisition Fee + 1st month payment
  data.push({ range: `${S}!C${compStart+1}`, values: [[`=IFERROR(C9+C14+C18,0)`]]});
  // C32: Monthly Payment
  data.push({ range: `${S}!C${compStart+2}`, values: [[`=IFERROR(C18,0)`]]});
  // C33: Total Payments = monthly × term
  data.push({ range: `${S}!C${compStart+3}`, values: [[`=IFERROR(C18*C10,0)`]]});
  // C34: End-of-Term Fees = Disposition + ExcessMiles×Rate + WearCharges
  data.push({ range: `${S}!C${compStart+4}`, values: [[`=IFERROR(C19+C20*C21+C22,0)`]]});
  // C35: Equity = 0
  data.push({ range: `${S}!C${compStart+5}`, values: [[0]]});
  // C36: Total Cash Outflow
  data.push({ range: `${S}!C${compStart+6}`, values: [[`=IFERROR(C${compStart+1}+C${compStart+3}+C${compStart+4},0)`]]});
  // C37: Estimated Net Cost
  data.push({ range: `${S}!C${compStart+7}`, values: [[`=IFERROR(C${compStart+6}-C${compStart+5},0)`]]});
  // C38: Cost per Month
  data.push({ range: `${S}!C${compStart+8}`, values: [[`=IFERROR(C${compStart+7}/C10,0)`]]});
  // C39: Cost per Mile
  data.push({ range: `${S}!C${compStart+9}`, values: [[`=IFERROR(C${compStart+7}/(C13*C10/12),0)`]]});

  // ── Comparison metrics — BUY column (D31–D39) ──
  // D31: Upfront = Down Payment
  data.push({ range: `${S}!D${compStart+1}`, values: [[`=IFERROR(D9,0)`]]});
  // D32: Monthly Payment
  data.push({ range: `${S}!D${compStart+2}`, values: [[`=IFERROR(D18,0)`]]});
  // D33: Total Payments = payment × term
  data.push({ range: `${S}!D${compStart+3}`, values: [[`=IFERROR(D18*D10,0)`]]});
  // D34: End-of-Term Fees = 0
  data.push({ range: `${S}!D${compStart+4}`, values: [[0]]});
  // D35: Equity = Estimated Resale Value (OTD × depreciation over term)
  data.push({ range: `${S}!D${compStart+5}`, values: [[`=IFERROR(D8*(1-0.12)^(D10/12),0)`]]});
  // D36: Total Cash Outflow = upfront + totalPmts
  data.push({ range: `${S}!D${compStart+6}`, values: [[`=IFERROR(D${compStart+1}+D${compStart+3},0)`]]});
  // D37: Estimated Net Cost = cashOut − equity
  data.push({ range: `${S}!D${compStart+7}`, values: [[`=IFERROR(D${compStart+6}-D${compStart+5},0)`]]});
  // D38: Cost per Month = netCost / term
  data.push({ range: `${S}!D${compStart+8}`, values: [[`=IFERROR(D${compStart+7}/D10,0)`]]});
  // D39: Cost per Mile = netCost / (AnnualMiles × term÷12)
  data.push({ range: `${S}!D${compStart+9}`, values: [[`=IFERROR(D${compStart+7}/(D13*D10/12),0)`]]});

  await valuesBatchUpdate(id, data, 'fix-leasevsbuy');
  console.log('✓ Lease vs Buy — buy column restructured and formulas fixed');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
