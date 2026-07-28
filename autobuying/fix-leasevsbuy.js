'use strict';
const { valuesBatchUpdate } = require('./lib');
const fs = require('fs');
const { id } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const S  = "'Lease vs Buy'";
const VQ = "'Vehicle Details & Quotes'";
const BS = "'Buyer Setup'";

// Row layout (buy column D):
//   D6  = Vehicle ID (V-001)
//   D7  = Out-the-Door Price   (lookup from VQ!$W, MATCH on VQ!$A = Vehicle ID)
//   D8  = Down Payment         (=Buyer Setup!B18)
//   D9  = APR                  (lookup from VQ!$X)
//   D10 = Loan Term (months)   (lookup from VQ!$Y)
//   D11 = Monthly Payment      (PMT formula)
//   D12 = Total Interest       (payment×term − principal)
//   D13 = Est. Resale Value    (OTD × (1−0.12)^(term/12))
//   D14 = Insurance (monthly)  (lookup from VQ!$AA)
//   D15 = Maintenance (monthly)(lookup from VQ!$AB)
//
// Comparison metrics start at row 31 (compStart+1, compStart=30):
//   [0] Upfront Cash Required
//   [1] Monthly Payment
//   [2] Total Payments
//   [3] End-of-Term Fees
//   [4] Equity / Asset Value
//   [5] Total Cash Outflow
//   [6] Estimated Net Cost
//   [7] Cost per Month
//   [8] Cost per Mile

(async () => {
  const data = [];
  const compStart = 30;

  // ── Fix 1: C18 — Monthly Lease Payment (was circular self-reference) ────────
  // Standard lease formula: ((AdjCapCost - Residual) / Term + (AdjCapCost + Residual) × MoneyFactor) × (1+TaxRate)
  // C8=Cap Cost, C9=CapReduction, C7=MSRP, C12=Residual%, C10=Term, C11=MoneyFactor, C17=TaxRate
  data.push({ range: `${S}!C18`, values: [[
    `=IFERROR(((C8-C9-C7*C12)/C10+(C8-C9+C7*C12)*C11)*(1+C17),0)`
  ]]});

  // ── Fix 2: D7 — OTD Price (was MATCH(D7,...) — self-reference) ─────────────
  data.push({ range: `${S}!D7`, values: [[
    `=IFERROR(INDEX(${VQ}!$W$6:$W$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),0)`
  ]]});

  // ── Fix 3–4: D9/D10 — APR and Loan Term (MATCH used D7=OTD, not D6=VehicleID) ──
  data.push({ range: `${S}!D9`, values: [[
    `=IFERROR(INDEX(${VQ}!$X$6:$X$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),0)`
  ]]});
  data.push({ range: `${S}!D10`, values: [[
    `=IFERROR(INDEX(${VQ}!$Y$6:$Y$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),60)`
  ]]});

  // ── Fix 5: D11 — Monthly Payment (was circular self-reference + wrong refs) ─
  // D7=OTD, D8=DownPmt, D9=APR, D10=LoanTerm; BS!B19=TradeInEquity, BS!B20=TradeInLoanBal
  data.push({ range: `${S}!D11`, values: [[
    `=IFERROR(-PMT(D9/12,D10,MAX(0,D7-D8-IFERROR(${BS}!B19,0)+IFERROR(${BS}!B20,0))),0)`
  ]]});

  // ── Fix 6: D12 — Total Interest (was D12*D11 — circular) ───────────────────
  data.push({ range: `${S}!D12`, values: [[
    `=IFERROR(D11*D10-MAX(0,D7-D8-IFERROR(${BS}!B19,0)+IFERROR(${BS}!B20,0)),0)`
  ]]});

  // ── Fix 7: D13 — Resale Value (was D8=DownPmt and C11=MoneyFactor) ─────────
  data.push({ range: `${S}!D13`, values: [[
    `=IFERROR(D7*(1-0.12)^(D10/12),0)`
  ]]});

  // ── Fix 8–9: D14/D15 — Insurance/Maintenance (MATCH used D7 not D6) ────────
  data.push({ range: `${S}!D14`, values: [[
    `=IFERROR(INDEX(${VQ}!$AA$6:$AA$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),120)`
  ]]});
  data.push({ range: `${S}!D15`, values: [[
    `=IFERROR(INDEX(${VQ}!$AB$6:$AB$105,MATCH(D6,${VQ}!$A$6:$A$105,0)),60)`
  ]]});

  // ── Fix comparison metrics — LEASE column (C31–C39) ────────────────────────
  // C31: Upfront = Cap Cost Reduction + Acquisition Fee + 1st month payment
  data.push({ range: `${S}!C${compStart+1}`, values: [[`=IFERROR(C9+C14+C18,0)`]]});
  // C32: Monthly Payment = lease monthly payment (already includes tax)
  data.push({ range: `${S}!C${compStart+2}`, values: [[`=IFERROR(C18,0)`]]});
  // C33: Total Payments = monthly × term
  data.push({ range: `${S}!C${compStart+3}`, values: [[`=IFERROR(C18*C10,0)`]]});
  // C34: End-of-Term Fees = Disposition + ExcessMiles×Rate + WearCharges
  data.push({ range: `${S}!C${compStart+4}`, values: [[`=IFERROR(C19+C20*C21+C22,0)`]]});
  // C35: Equity = 0 (no asset at lease end)
  data.push({ range: `${S}!C${compStart+5}`, values: [[0]]});
  // C36: Total Cash Outflow = upfront + totalPmts + endFees
  data.push({ range: `${S}!C${compStart+6}`, values: [[`=IFERROR(C${compStart+1}+C${compStart+3}+C${compStart+4},0)`]]});
  // C37: Estimated Net Cost = cashOut − equity
  data.push({ range: `${S}!C${compStart+7}`, values: [[`=IFERROR(C${compStart+6}-C${compStart+5},0)`]]});
  // C38: Cost per Month = netCost / term
  data.push({ range: `${S}!C${compStart+8}`, values: [[`=IFERROR(C${compStart+7}/C10,0)`]]});
  // C39: Cost per Mile = netCost / (AnnualMiles × term÷12)
  data.push({ range: `${S}!C${compStart+9}`, values: [[`=IFERROR(C${compStart+7}/(C13*C10/12),0)`]]});

  // ── Fix comparison metrics — BUY column (D31–D39) ──────────────────────────
  // D31: Upfront = Down Payment
  data.push({ range: `${S}!D${compStart+1}`, values: [[`=IFERROR(D8,0)`]]});
  // D32: Monthly Payment
  data.push({ range: `${S}!D${compStart+2}`, values: [[`=IFERROR(D11,0)`]]});
  // D33: Total Payments = payment × term
  data.push({ range: `${S}!D${compStart+3}`, values: [[`=IFERROR(D11*D10,0)`]]});
  // D34: End-of-Term Fees = 0
  data.push({ range: `${S}!D${compStart+4}`, values: [[0]]});
  // D35: Equity = Estimated Resale Value
  data.push({ range: `${S}!D${compStart+5}`, values: [[`=IFERROR(D13,0)`]]});
  // D36: Total Cash Outflow = upfront + totalPmts
  data.push({ range: `${S}!D${compStart+6}`, values: [[`=IFERROR(D${compStart+1}+D${compStart+3},0)`]]});
  // D37: Estimated Net Cost = cashOut − equity
  data.push({ range: `${S}!D${compStart+7}`, values: [[`=IFERROR(D${compStart+6}-D${compStart+5},0)`]]});
  // D38: Cost per Month = netCost / term
  data.push({ range: `${S}!D${compStart+8}`, values: [[`=IFERROR(D${compStart+7}/D10,0)`]]});
  // D39: Cost per Mile = netCost / (AnnualMiles × term÷12)
  data.push({ range: `${S}!D${compStart+9}`, values: [[`=IFERROR(D${compStart+7}/(IFERROR(${BS}!B21,12000)*D10/12),0)`]]});

  await valuesBatchUpdate(id, data, 'fix-leasevsbuy');
  console.log('✓ Lease vs Buy — formula errors fixed');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
