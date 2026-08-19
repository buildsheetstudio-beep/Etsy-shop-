'use strict';
const { sheets, hex, batchUpdate, valuesBatchUpdate, gridRange, colL, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Contribution Log'];
const S = "'Contribution Log'";
const FS = "'Fund Setup & Goals'";
const NC = 16; // A-P

// Col mapping:
// A(0)=Transaction ID  B(1)=Date  C(2)=Year(fml)  D(3)=Month(fml)
// E(4)=Fund ID  F(5)=Fund Name(fml)  G(6)=Owner(fml)  H(7)=Transaction Type
// I(8)=Contribution Source  J(9)=Amount  K(10)=From Fund ID  L(11)=To Fund ID
// M(12)=Transfer Pair ID  N(13)=Balance Effect(fml)  O(14)=Running Fund Balance(fml)
// P(15)=Notes

// [date, fundId, txType, source, amount, fromFundId, toFundId, pairId, notes]
const TXNS = [
  // ── 2024 January ───────────────────────────────────────────────────────────
  ['1/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Jan wedding fund — first contribution.'],
  ['1/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Jan property tax reserve.'],
  ['1/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Jan emergency buffer deposit.'],
  ['1/15/2024','FUND-009','Contribution','Monthly Budget',200,'','','','Jan insurance premium reserve.'],
  ['1/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Jan Q1 estimated tax contribution.'],
  ['1/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Jan medical deductible start.'],
  ['1/20/2024','FUND-013','Contribution','Monthly Budget',60,'','','','Jan birthday fund.'],
  ['1/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Jan home maintenance reserve.'],
  ['1/20/2024','FUND-025','Contribution','Monthly Budget',60,'','','','Jan streaming subscriptions fund.'],
  // ── 2024 February ──────────────────────────────────────────────────────────
  ['2/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Feb wedding fund.'],
  ['2/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Feb property tax reserve.'],
  ['2/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Feb emergency buffer.'],
  ['2/15/2024','FUND-009','Contribution','Monthly Budget',200,'','','','Feb insurance reserve.'],
  ['2/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Feb tax reserve.'],
  ['2/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Feb medical deductible.'],
  ['2/20/2024','FUND-013','Contribution','Monthly Budget',60,'','','','Feb birthday fund.'],
  ['2/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Feb home maintenance.'],
  ['2/20/2024','FUND-003','Contribution','Monthly Budget',100,'','','','Feb car maintenance start.'],
  ['2/20/2024','FUND-025','Contribution','Monthly Budget',60,'','','','Feb subscriptions fund.'],
  // ── 2024 March ─────────────────────────────────────────────────────────────
  ['3/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Mar wedding fund.'],
  ['3/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Mar property tax.'],
  ['3/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Mar emergency buffer.'],
  ['3/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Mar tax reserve — Q1 filing month.'],
  ['3/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Mar medical deductible.'],
  ['3/20/2024','FUND-013','Contribution','Monthly Budget',60,'','','','Mar birthday fund.'],
  ['3/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Mar home maintenance.'],
  ['3/20/2024','FUND-003','Contribution','Monthly Budget',100,'','','','Mar car maintenance.'],
  ['3/25/2024','FUND-010','Contribution','Cash Windfall',100,'','','','Pet fund — birthday gift cash from relatives.'],
  ['3/28/2024','FUND-021','Contribution','Monthly Budget',100,'','','','New smartphone — Marisol, month 1.'],
  // ── 2024 April ─────────────────────────────────────────────────────────────
  ['4/1/2024','FUND-025','Contribution','Monthly Budget',60,'','','','Apr subscriptions fund.'],
  ['4/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Apr wedding fund.'],
  ['4/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Apr property tax.'],
  ['4/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Apr tax reserve.'],
  ['4/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Apr emergency buffer.'],
  ['4/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Apr medical deductible.'],
  ['4/20/2024','FUND-013','Contribution','Monthly Budget',60,'','','','Apr birthday fund.'],
  ['4/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Apr home maintenance.'],
  ['4/20/2024','FUND-003','Contribution','Monthly Budget',100,'','','','Apr car maintenance.'],
  ['4/28/2024','FUND-003','Withdrawal','',250,'','','','Tire rotation + new wiper blades.'],
  ['4/28/2024','FUND-021','Contribution','Monthly Budget',100,'','','','New smartphone month 2.'],
  // ── 2024 May ───────────────────────────────────────────────────────────────
  ['5/1/2024','FUND-025','Contribution','Monthly Budget',60,'','','','May subscriptions fund.'],
  ['5/15/2024','FUND-008','Contribution','Paycheck',800,'','','','May wedding fund.'],
  ['5/15/2024','FUND-024','Contribution','Paycheck',700,'','','','May property tax.'],
  ['5/15/2024','FUND-016','Contribution','Paycheck',500,'','','','May tax reserve.'],
  ['5/15/2024','FUND-004','Contribution','Paycheck',300,'','','','May emergency buffer.'],
  ['5/20/2024','FUND-023','Contribution','Paycheck',300,'','','','May medical deductible.'],
  ['5/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','May home maintenance.'],
  ['5/20/2024','FUND-003','Contribution','Monthly Budget',100,'','','','May car maintenance.'],
  ['5/28/2024','FUND-021','Contribution','Monthly Budget',100,'','','','New smartphone month 3.'],
  // ── 2024 June ──────────────────────────────────────────────────────────────
  ['6/1/2024','FUND-025','Contribution','Monthly Budget',60,'','','','Jun subscriptions fund.'],
  ['6/5/2024','FUND-010','Contribution','Monthly Budget',100,'','','','Pet medical fund contribution.'],
  ['6/10/2024','FUND-010','Withdrawal','',120,'','','','Emergency vet visit — Biscuit ate something.'],
  ['6/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Jun wedding fund.'],
  ['6/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Jun property tax.'],
  ['6/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Jun tax reserve.'],
  ['6/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Jun emergency buffer.'],
  ['6/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Jun medical deductible.'],
  ['6/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Jun home maintenance.'],
  ['6/20/2024','FUND-003','Contribution','Monthly Budget',100,'','','','Jun car maintenance.'],
  ['6/28/2024','FUND-021','Contribution','Monthly Budget',100,'','','','New smartphone month 4.'],
  // ── 2024 July ──────────────────────────────────────────────────────────────
  ['7/1/2024','FUND-007','Contribution','Bonus',1500,'','','','Kitchen reno — work bonus applied.'],
  ['7/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Jul wedding fund.'],
  ['7/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Jul property tax.'],
  ['7/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Jul tax reserve.'],
  ['7/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Jul emergency buffer.'],
  ['7/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Jul medical deductible.'],
  ['7/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Jul home maintenance.'],
  ['7/20/2024','FUND-003','Contribution','Monthly Budget',100,'','','','Jul car maintenance.'],
  ['7/20/2024','FUND-011','Contribution','Monthly Budget',200,'','','','Car down payment — first contribution.'],
  ['7/28/2024','FUND-021','Contribution','Monthly Budget',100,'','','','New smartphone month 5.'],
  // ── 2024 August ────────────────────────────────────────────────────────────
  ['8/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Aug wedding fund.'],
  ['8/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Aug property tax.'],
  ['8/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Aug tax reserve.'],
  ['8/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Aug emergency buffer.'],
  ['8/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Aug medical deductible.'],
  ['8/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Aug home maintenance.'],
  ['8/20/2024','FUND-003','Contribution','Monthly Budget',100,'','','','Aug car maintenance.'],
  ['8/20/2024','FUND-011','Contribution','Monthly Budget',200,'','','','Car down payment month 2.'],
  ['8/20/2024','FUND-012','Contribution','Monthly Budget',300,'','','','Kids school expenses — first contribution.'],
  ['8/28/2024','FUND-021','Contribution','Monthly Budget',100,'','','','New smartphone month 6.'],
  // ── 2024 September ─────────────────────────────────────────────────────────
  ['9/5/2024','FUND-012','Withdrawal','',600,'','','','Back-to-school supplies, backpack, and activity fees.'],
  ['9/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Sep wedding fund.'],
  ['9/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Sep property tax.'],
  ['9/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Sep tax reserve.'],
  ['9/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Sep emergency buffer.'],
  ['9/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Sep medical deductible.'],
  ['9/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Sep home maintenance.'],
  ['9/20/2024','FUND-011','Contribution','Monthly Budget',200,'','','','Car down payment month 3.'],
  ['9/20/2024','FUND-012','Contribution','Monthly Budget',300,'','','','School expenses month 2.'],
  ['9/28/2024','FUND-021','Contribution','Monthly Budget',100,'','','','New smartphone month 7.'],
  // ── 2024 October ───────────────────────────────────────────────────────────
  ['10/1/2024','FUND-013','Contribution','Monthly Budget',60,'','','','Oct birthday fund — almost funded.'],
  ['10/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Oct wedding fund.'],
  ['10/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Oct property tax.'],
  ['10/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Oct tax reserve.'],
  ['10/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Oct emergency buffer — transfer needed for car.'],
  ['10/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Oct medical deductible.'],
  ['10/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Oct home maintenance.'],
  ['10/20/2024','FUND-011','Contribution','Monthly Budget',200,'','','','Car down payment month 4.'],
  ['10/20/2024','FUND-012','Contribution','Monthly Budget',300,'','','','School expenses month 3.'],
  ['10/28/2024','FUND-021','Contribution','Monthly Budget',100,'','','','New smartphone month 8.'],
  ['10/30/2024','FUND-004','Transfer Out','',500,'FUND-004','FUND-003','TXP-001','Transfer to car maintenance — emergency brake job.'],
  ['10/30/2024','FUND-003','Transfer In','',500,'FUND-004','FUND-003','TXP-001','Transfer from emergency buffer — brake job.'],
  ['10/31/2024','FUND-003','Withdrawal','',480,'','','','Brake job — front pads, rotors, and fluid flush.'],
  // ── 2024 November ──────────────────────────────────────────────────────────
  ['11/1/2024','FUND-013','Contribution','Monthly Budget',60,'','','','Nov birthday fund — goal almost reached.'],
  ['11/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Nov wedding fund.'],
  ['11/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Nov property tax.'],
  ['11/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Nov tax reserve.'],
  ['11/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Nov emergency buffer rebuild.'],
  ['11/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Nov medical deductible.'],
  ['11/20/2024','FUND-017','Contribution','Monthly Budget',250,'','','','Nov home maintenance.'],
  ['11/20/2024','FUND-003','Contribution','Monthly Budget',100,'','','','Nov car maintenance.'],
  ['11/20/2024','FUND-011','Contribution','Monthly Budget',200,'','','','Car down payment month 5.'],
  ['11/20/2024','FUND-012','Contribution','Monthly Budget',300,'','','','School expenses month 4.'],
  ['11/25/2024','FUND-009','Withdrawal','',1800,'','','','Home and auto insurance annual premium paid.'],
  ['11/28/2024','FUND-021','Contribution','Monthly Budget',100,'','','','New smartphone month 9 — goal reached!'],
  // ── 2024 December ──────────────────────────────────────────────────────────
  ['12/1/2024','FUND-013','Contribution','Monthly Budget',60,'','','','Dec birthday fund — goal reached!'],
  ['12/10/2024','FUND-017','Withdrawal','',400,'','','','Furnace filter and annual HVAC tune-up.'],
  ['12/15/2024','FUND-008','Contribution','Paycheck',800,'','','','Dec wedding fund — year-end.'],
  ['12/15/2024','FUND-024','Contribution','Paycheck',700,'','','','Dec property tax — fully funded for 2024.'],
  ['12/15/2024','FUND-016','Contribution','Paycheck',500,'','','','Dec tax reserve.'],
  ['12/15/2024','FUND-004','Contribution','Paycheck',300,'','','','Dec emergency buffer.'],
  ['12/20/2024','FUND-023','Contribution','Paycheck',300,'','','','Dec medical — fully funded!'],
  ['12/20/2024','FUND-022','Contribution','Side Income',200,'','','','Business tools — Terrence freelance income.'],
  ['12/20/2024','FUND-018','Contribution','Monthly Budget',200,'','','','Living room furniture start.'],
  // ── 2024 December Finale ────────────────────────────────────────────────────
  ['12/28/2024','FUND-024','Withdrawal','',8400,'','','','Annual property tax bill — paid in full.'],
  ['12/31/2024','FUND-009','Refund','',200,'','','','Insurance premium overpayment refund received.'],
  // ── 2025 January ───────────────────────────────────────────────────────────
  ['1/5/2025','FUND-001','Contribution','Monthly Budget',400,'','','','Hawaii vacation — first deposit.'],
  ['1/5/2025','FUND-005','Contribution','Monthly Budget',80,'','','','New laptop — Marisol, month 1.'],
  ['1/6/2025','FUND-020','Contribution','Monthly Budget',300,'','','','Anniversary trip — Paris, month 1.'],
  ['1/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Jan 2025 wedding fund.'],
  ['1/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Jan property tax restart for 2025.'],
  ['1/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Jan 2025 tax reserve restart.'],
  ['1/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Jan emergency buffer.'],
  ['1/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Jan insurance reserve restart.'],
  ['1/20/2025','FUND-017','Contribution','Monthly Budget',250,'','','','Jan home maintenance.'],
  ['1/20/2025','FUND-003','Contribution','Monthly Budget',100,'','','','Jan car maintenance.'],
  ['1/20/2025','FUND-011','Contribution','Monthly Budget',200,'','','','Car down payment month 6.'],
  ['1/20/2025','FUND-012','Contribution','Monthly Budget',300,'','','','School expenses month 5.'],
  ['1/20/2025','FUND-014','Contribution','Monthly Budget',40,'','','','Annual memberships — first contribution.'],
  // ── 2025 February ──────────────────────────────────────────────────────────
  ['2/5/2025','FUND-001','Contribution','Monthly Budget',400,'','','','Hawaii vacation month 2.'],
  ['2/5/2025','FUND-005','Contribution','Monthly Budget',80,'','','','New laptop month 2.'],
  ['2/6/2025','FUND-006','Contribution','Paycheck',150,'','','','Dental — Terrence, paycheck 1.'],
  ['2/6/2025','FUND-020','Contribution','Monthly Budget',300,'','','','Anniversary trip month 2.'],
  ['2/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Feb wedding fund.'],
  ['2/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Feb property tax.'],
  ['2/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Feb tax reserve.'],
  ['2/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Feb emergency buffer.'],
  ['2/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Feb insurance.'],
  ['2/20/2025','FUND-017','Contribution','Monthly Budget',250,'','','','Feb home maintenance.'],
  ['2/20/2025','FUND-003','Contribution','Monthly Budget',100,'','','','Feb car maintenance.'],
  ['2/20/2025','FUND-011','Contribution','Monthly Budget',200,'','','','Car down payment month 7.'],
  ['2/20/2025','FUND-012','Contribution','Monthly Budget',300,'','','','School expenses month 6.'],
  ['2/20/2025','FUND-014','Contribution','Monthly Budget',40,'','','','Memberships month 2.'],
  ['2/20/2025','FUND-022','Contribution','Side Income',200,'','','','Business tools month 2.'],
  ['2/20/2025','FUND-018','Contribution','Monthly Budget',200,'','','','Furniture fund month 2 — then pausing.'],
  // ── 2025 March ─────────────────────────────────────────────────────────────
  ['3/1/2025','FUND-019','Contribution','Monthly Budget',350,'','','','Baby fund — first contribution.'],
  ['3/5/2025','FUND-001','Contribution','Monthly Budget',400,'','','','Hawaii vacation month 3.'],
  ['3/5/2025','FUND-005','Contribution','Monthly Budget',80,'','','','New laptop month 3.'],
  ['3/6/2025','FUND-006','Contribution','Paycheck',150,'','','','Dental month 2 paycheck.'],
  ['3/6/2025','FUND-020','Contribution','Monthly Budget',300,'','','','Anniversary trip month 3.'],
  ['3/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Mar wedding fund.'],
  ['3/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Mar property tax.'],
  ['3/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Mar tax reserve.'],
  ['3/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Mar emergency buffer.'],
  ['3/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Mar insurance.'],
  ['3/20/2025','FUND-017','Contribution','Monthly Budget',250,'','','','Mar home maintenance.'],
  ['3/20/2025','FUND-003','Contribution','Monthly Budget',100,'','','','Mar car maintenance.'],
  ['3/20/2025','FUND-011','Contribution','Monthly Budget',200,'','','','Car down payment month 8.'],
  ['3/20/2025','FUND-012','Contribution','Monthly Budget',300,'','','','School expenses month 7.'],
  ['3/20/2025','FUND-014','Contribution','Monthly Budget',40,'','','','Memberships month 3.'],
  ['3/25/2025','FUND-003','Withdrawal','',175,'','','','Oil change, air filter, and wiper blades.'],
  // ── 2025 April ─────────────────────────────────────────────────────────────
  ['4/1/2025','FUND-019','Contribution','Monthly Budget',350,'','','','Baby fund month 2.'],
  ['4/5/2025','FUND-001','Contribution','Monthly Budget',400,'','','','Hawaii vacation month 4.'],
  ['4/5/2025','FUND-005','Contribution','Monthly Budget',80,'','','','New laptop month 4.'],
  ['4/6/2025','FUND-006','Contribution','Paycheck',150,'','','','Dental month 3 paycheck.'],
  ['4/6/2025','FUND-020','Contribution','Monthly Budget',300,'','','','Anniversary trip month 4.'],
  ['4/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Apr wedding fund.'],
  ['4/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Apr property tax.'],
  ['4/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Apr tax reserve.'],
  ['4/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Apr emergency buffer.'],
  ['4/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Apr insurance.'],
  ['4/20/2025','FUND-017','Contribution','Monthly Budget',250,'','','','Apr home maintenance.'],
  ['4/20/2025','FUND-014','Contribution','Monthly Budget',40,'','','','Memberships month 4.'],
  ['4/28/2025','FUND-007','Contribution','Tax Refund',1500,'','','','Kitchen reno — tax refund applied.'],
  // ── 2025 May ───────────────────────────────────────────────────────────────
  ['5/1/2025','FUND-019','Contribution','Monthly Budget',350,'','','','Baby fund month 3.'],
  ['5/5/2025','FUND-001','Contribution','Monthly Budget',400,'','','','Hawaii vacation month 5.'],
  ['5/5/2025','FUND-005','Contribution','Monthly Budget',80,'','','','New laptop month 5.'],
  ['5/6/2025','FUND-006','Contribution','Paycheck',150,'','','','Dental month 4 paycheck.'],
  ['5/6/2025','FUND-020','Contribution','Monthly Budget',300,'','','','Anniversary trip month 5.'],
  ['5/15/2025','FUND-008','Contribution','Paycheck',800,'','','','May wedding fund.'],
  ['5/15/2025','FUND-024','Contribution','Paycheck',700,'','','','May property tax.'],
  ['5/15/2025','FUND-016','Contribution','Paycheck',500,'','','','May tax reserve.'],
  ['5/15/2025','FUND-004','Contribution','Paycheck',300,'','','','May emergency buffer.'],
  ['5/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','May insurance.'],
  ['5/20/2025','FUND-017','Contribution','Monthly Budget',250,'','','','May home maintenance.'],
  ['5/20/2025','FUND-014','Contribution','Monthly Budget',40,'','','','Memberships month 5.'],
  ['5/30/2025','FUND-006','Withdrawal','',300,'','','','First dental appointment — filling #1.'],
  // ── 2025 June ──────────────────────────────────────────────────────────────
  ['6/1/2025','FUND-019','Contribution','Monthly Budget',350,'','','','Baby fund month 4.'],
  ['6/5/2025','FUND-001','Contribution','Monthly Budget',400,'','','','Hawaii vacation month 6.'],
  ['6/5/2025','FUND-005','Contribution','Monthly Budget',80,'','','','New laptop month 6.'],
  ['6/6/2025','FUND-006','Contribution','Paycheck',150,'','','','Dental month 5 paycheck.'],
  ['6/6/2025','FUND-020','Contribution','Monthly Budget',300,'','','','Anniversary trip month 6 — transfer leftover soon.'],
  ['6/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Jun wedding fund.'],
  ['6/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Jun property tax.'],
  ['6/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Jun tax reserve.'],
  ['6/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Jun emergency buffer.'],
  ['6/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Jun insurance.'],
  ['6/20/2025','FUND-017','Contribution','Monthly Budget',250,'','','','Jun home maintenance.'],
  ['6/20/2025','FUND-014','Contribution','Monthly Budget',40,'','','','Memberships month 6 — goal reached!'],
  ['6/28/2025','FUND-010','Contribution','Monthly Budget',100,'','','','Pet medical refill.'],
  ['6/28/2025','FUND-010','Withdrawal','',80,'','','','Annual rabies vaccination and wellness check.'],
  // ── 2025 July ──────────────────────────────────────────────────────────────
  ['7/1/2025','FUND-019','Contribution','Monthly Budget',350,'','','','Baby fund month 5.'],
  ['7/1/2025','FUND-015','Contribution','Monthly Budget',200,'','','','Moving costs — Marisol, month 1.'],
  ['7/5/2025','FUND-001','Contribution','Monthly Budget',400,'','','','Hawaii vacation month 7 — transfer from Christmas.'],
  ['7/5/2025','FUND-005','Contribution','Monthly Budget',80,'','','','New laptop month 7.'],
  ['7/6/2025','FUND-006','Contribution','Paycheck',150,'','','','Dental month 6 paycheck.'],
  ['7/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Jul wedding fund.'],
  ['7/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Jul property tax.'],
  ['7/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Jul tax reserve.'],
  ['7/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Jul emergency buffer.'],
  ['7/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Jul insurance.'],
  ['7/20/2025','FUND-017','Contribution','Monthly Budget',250,'','','','Jul home maintenance.'],
  ['7/20/2025','FUND-017','Withdrawal','',300,'','','','Plumber for leaky bathroom valve.'],
  ['7/25/2025','FUND-002','Transfer Out','',200,'FUND-002','FUND-001','TXP-002','Transfer Christmas surplus to Hawaii fund.'],
  ['7/25/2025','FUND-001','Transfer In','',200,'FUND-002','FUND-001','TXP-002','Transfer in from Christmas surplus — Hawaii fund boost.'],
  // ── 2025 August ────────────────────────────────────────────────────────────
  ['8/1/2025','FUND-019','Contribution','Monthly Budget',350,'','','','Baby fund month 6.'],
  ['8/1/2025','FUND-015','Contribution','Monthly Budget',200,'','','','Moving costs month 2.'],
  ['8/6/2025','FUND-006','Contribution','Paycheck',150,'','','','Dental month 7 paycheck.'],
  ['8/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Aug wedding fund.'],
  ['8/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Aug property tax.'],
  ['8/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Aug tax reserve.'],
  ['8/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Aug emergency buffer.'],
  ['8/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Aug insurance.'],
  ['8/20/2025','FUND-022','Contribution','Side Income',200,'','','','Business software fund boost.'],
  // ── 2025 September ─────────────────────────────────────────────────────────
  ['9/1/2025','FUND-019','Contribution','Monthly Budget',350,'','','','Baby fund month 7.'],
  ['9/1/2025','FUND-015','Contribution','Monthly Budget',200,'','','','Moving costs month 3.'],
  ['9/5/2025','FUND-002','Contribution','Monthly Budget',300,'','','','Christmas 2026 reload — month 1.'],
  ['9/6/2025','FUND-006','Contribution','Paycheck',150,'','','','Dental month 8 paycheck.'],
  ['9/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Sep wedding fund.'],
  ['9/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Sep property tax.'],
  ['9/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Sep tax reserve.'],
  ['9/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Sep emergency buffer.'],
  ['9/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Sep insurance.'],
  // ── 2025 October ───────────────────────────────────────────────────────────
  ['10/1/2025','FUND-019','Contribution','Monthly Budget',350,'','','','Baby fund month 8.'],
  ['10/1/2025','FUND-015','Contribution','Monthly Budget',200,'','','','Moving costs month 4.'],
  ['10/5/2025','FUND-002','Contribution','Monthly Budget',300,'','','','Christmas 2026 reload — month 2.'],
  ['10/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Oct wedding fund.'],
  ['10/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Oct property tax.'],
  ['10/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Oct tax reserve.'],
  ['10/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Oct emergency buffer.'],
  ['10/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Oct insurance.'],
  ['10/22/2025','FUND-022','Contribution','Side Income',200,'','','','Business tools — Oct freelance income.'],
  // ── 2025 November ──────────────────────────────────────────────────────────
  ['11/1/2025','FUND-015','Contribution','Monthly Budget',200,'','','','Moving costs month 5.'],
  ['11/5/2025','FUND-002','Contribution','Monthly Budget',300,'','','','Christmas 2026 reload — month 3.'],
  ['11/10/2025','FUND-009','Withdrawal','',1800,'','','','Home and auto insurance renewal — paid annually.'],
  ['11/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Nov wedding fund.'],
  ['11/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Nov property tax.'],
  ['11/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Nov tax reserve.'],
  ['11/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Nov emergency buffer.'],
  ['11/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Nov insurance restart after payment.'],
  ['11/20/2025','FUND-007','Contribution','Monthly Budget',500,'','','','Kitchen reno — holiday saving push.'],
  // ── 2025 December ──────────────────────────────────────────────────────────
  ['12/1/2025','FUND-015','Contribution','Monthly Budget',200,'','','','Moving costs month 6.'],
  ['12/5/2025','FUND-002','Contribution','Monthly Budget',300,'','','','Christmas 2026 reload — month 4.'],
  ['12/15/2025','FUND-008','Contribution','Paycheck',800,'','','','Dec 2025 wedding fund.'],
  ['12/15/2025','FUND-024','Contribution','Paycheck',700,'','','','Dec property tax.'],
  ['12/15/2025','FUND-016','Contribution','Paycheck',500,'','','','Dec tax reserve.'],
  ['12/15/2025','FUND-004','Contribution','Paycheck',300,'','','','Dec emergency buffer.'],
  ['12/20/2025','FUND-009','Contribution','Monthly Budget',200,'','','','Dec insurance reserve.'],
  ['12/20/2025','FUND-026','Contribution','Monthly Budget',150,'','','','Appliance replacement fund — first contribution.'],
  ['12/31/2025','FUND-024','Withdrawal','',8400,'','','','Annual property tax bill — 2025 payment.'],
  // ── 2026 January ───────────────────────────────────────────────────────────
  ['1/5/2026','FUND-002','Contribution','Monthly Budget',300,'','','','Christmas 2026 reload — month 5 — goal reached!'],
  ['1/15/2026','FUND-008','Contribution','Paycheck',800,'','','','Jan 2026 wedding fund.'],
  ['1/15/2026','FUND-024','Contribution','Paycheck',700,'','','','Jan property tax restart 2026.'],
  ['1/15/2026','FUND-016','Contribution','Paycheck',500,'','','','Jan 2026 tax reserve.'],
  ['1/15/2026','FUND-004','Contribution','Paycheck',300,'','','','Jan emergency buffer.'],
  ['1/20/2026','FUND-009','Contribution','Monthly Budget',200,'','','','Jan insurance reserve.'],
  ['1/20/2026','FUND-010','Contribution','Monthly Budget',100,'','','','Pet medical refill 2026.'],
  ['1/20/2026','FUND-026','Contribution','Monthly Budget',150,'','','','Appliance fund month 2.'],
  // ── 2026 February ──────────────────────────────────────────────────────────
  ['2/1/2026','FUND-001','Contribution','Monthly Budget',400,'','','','Hawaii vacation month 8 — final push!'],
  ['2/6/2026','FUND-016','Contribution','Paycheck',500,'','','','Feb 2026 tax reserve.'],
  ['2/15/2026','FUND-008','Contribution','Paycheck',800,'','','','Feb 2026 wedding fund.'],
  ['2/15/2026','FUND-024','Contribution','Paycheck',700,'','','','Feb property tax.'],
  ['2/15/2026','FUND-004','Contribution','Paycheck',300,'','','','Feb emergency buffer.'],
  ['2/20/2026','FUND-009','Contribution','Monthly Budget',200,'','','','Feb insurance.'],
  ['2/20/2026','FUND-026','Contribution','Monthly Budget',150,'','','','Appliance fund month 3 — pausing.'],
  ['2/28/2026','FUND-010','Withdrawal','',100,'','','','Annual wellness exam + heartworm test.'],
];

(async () => {
  const fmt = [];
  const vals = [];
  const NDATA = 5000;

  // ── Title ──────────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 0, 1, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 0, 1, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 22, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
  vals.push({ range: `${S}!A1`, values: [['CONTRIBUTION LOG']] });

  // ── Subtitle ───────────────────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 1, 2, 0, NC), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 1, 2, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.dustyBlue), textFormat: { italic: true, fontSize: 10, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)' } });
  vals.push({ range: `${S}!A2`, values: [['Record every contribution, withdrawal, and transfer. Balance Effect column is auto-calculated from Transaction Type. All amounts entered as positive values.']] });

  // ── Transfer Helper Block ──────────────────────────────────────────────────
  fmt.push({ mergeCells: { range: gridRange(SID, 2, 3, 0, 7), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 2, 3, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.info), textFormat: { bold: true, fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  vals.push({ range: `${S}!A3`, values: [['TRANSFER GUIDE: A transfer requires TWO rows — one Transfer Out (source fund) and one Transfer In (destination fund) — both with the same Transfer Pair ID (e.g. TXP-001).']] });

  fmt.push({ mergeCells: { range: gridRange(SID, 3, 4, 0, 7), mergeType: 'MERGE_ALL' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 3, 4, 0, 7), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 9 } } }, fields: 'userEnteredFormat(backgroundColor,textFormat)' } });
  vals.push({ range: `${S}!A4`, values: [['Balance Effect signs: Contribution / Transfer In / Adjustment Increase / Refund → +Amount   |   Withdrawal / Transfer Out / Adjustment Decrease → −Amount']] });

  // Summary KPIs row 5-6
  const kpiDefs = [
    { label: 'Total Contributions',   col: 8,  fml: `=SUMIF($H$8:$H$5007,"Contribution",$J$8:$J$5007)` },
    { label: 'Total Withdrawals',     col: 11, fml: `=IFERROR(ABS(SUMPRODUCT(($H$8:$H$5007="Withdrawal")*$N$8:$N$5007)),0)` },
    { label: 'Net Balance Effect',    col: 14, fml: `=IFERROR(SUM($N$8:$N$5007),0)` },
  ];
  kpiDefs.forEach(k => {
    fmt.push({ mergeCells: { range: gridRange(SID, 4, 5, k.col, k.col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ mergeCells: { range: gridRange(SID, 5, 7, k.col, k.col+3), mergeType: 'MERGE_ALL' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 4, 5, k.col, k.col+3), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 8, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    fmt.push({ repeatCell: { range: gridRange(SID, 5, 7, k.col, k.col+3), cell: { userEnteredFormat: { backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 18, foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)' } });
    vals.push({ range: `${S}!${colL(k.col)}5`, values: [[k.label]] });
    vals.push({ range: `${S}!${colL(k.col)}6`, values: [[k.fml]] });
  });

  // ── Column headers (row 7, 0-indexed 6) ────────────────────────────────────
  const HDRS = ['Transaction ID','Date','Year','Month','Fund ID','Fund Name','Owner','Transaction Type','Contribution Source','Amount','From Fund ID','To Fund ID','Transfer Pair ID','Balance Effect','Running Fund Balance','Notes'];
  fmt.push({ repeatCell: { range: gridRange(SID, 6, 7, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.white) }, horizontalAlignment: 'CENTER', wrapStrategy: 'WRAP', verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy,verticalAlignment)' } });
  vals.push({ range: `${S}!A7`, values: [HDRS] });

  // ── Write sample transaction data (B,E,H,I,J,K,L,M,P) first ───────────────
  TXNS.forEach((tx, i) => {
    const r = 8 + i;
    const [date, fundId, txType, source, amount, fromFund, toFund, pairId, notes] = tx;
    vals.push({ range: `${S}!B${r}:P${r}`, values: [[date, '', '', fundId, '', '', txType, source, amount, fromFund, toFund, pairId, '', '', notes]] });
  });

  // ── Formula columns: A, C, D, F, G, N, O ──────────────────────────────────
  const txnIdFmls  = Array.from({ length: NDATA }, (_, i) => [`=IF(B${8+i}="","","SFT-"&TEXT(ROW()-7,"00000"))`]);
  const yearFmls   = Array.from({ length: NDATA }, (_, i) => [`=IFERROR(IF(B${8+i}="","",YEAR(B${8+i})),"")`]);
  const monthFmls  = Array.from({ length: NDATA }, (_, i) => [`=IFERROR(IF(B${8+i}="","",TEXT(B${8+i},"mmmm")),"")`]);
  const nameFmls   = Array.from({ length: NDATA }, (_, i) => [`=IFERROR(IF(E${8+i}="","",VLOOKUP(E${8+i},${FS}!$A$8:$B$507,2,FALSE)),"")`]);
  const ownerFmls  = Array.from({ length: NDATA }, (_, i) => [`=IFERROR(IF(E${8+i}="","",VLOOKUP(E${8+i},${FS}!$A$8:$D$507,4,FALSE)),"")`]);
  const beFmls     = Array.from({ length: NDATA }, (_, i) => {
    const r = 8+i;
    return [`=IF(J${r}="","",IF(OR(H${r}="Contribution",H${r}="Transfer In",H${r}="Adjustment Increase",H${r}="Refund"),J${r},IF(OR(H${r}="Withdrawal",H${r}="Transfer Out",H${r}="Adjustment Decrease"),-J${r},J${r})))`];
  });

  // Running balance for sample rows only (growing range — performant for real usage)
  const rbFmls = TXNS.map((_, i) => {
    const r = 8 + i;
    return [`=IF(E${r}="","",IFERROR(SUMPRODUCT(($E$8:$E${r}=E${r})*($N$8:$N${r})),0))`];
  });

  vals.push({ range: `${S}!A8:A${7+NDATA}`, values: txnIdFmls });
  vals.push({ range: `${S}!C8:C${7+NDATA}`, values: yearFmls });
  vals.push({ range: `${S}!D8:D${7+NDATA}`, values: monthFmls });
  vals.push({ range: `${S}!F8:F${7+NDATA}`, values: nameFmls });
  vals.push({ range: `${S}!G8:G${7+NDATA}`, values: ownerFmls });
  vals.push({ range: `${S}!N8:N${7+NDATA}`, values: beFmls });
  vals.push({ range: `${S}!O8:O${7+TXNS.length}`, values: rbFmls });

  // ── Row styling ─────────────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.panel), textFormat: { fontSize: 8, fontFamily: 'Arial', foregroundColor: hex(C.text) }, verticalAlignment: 'MIDDLE' } }, fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)' } });
  for (let r = 0; r < NDATA; r++) {
    if (r % 2 !== 0) fmt.push({ repeatCell: { range: gridRange(SID, 7+r, 8+r, 0, NC), cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) } }, fields: 'userEnteredFormat.backgroundColor' } });
  }
  [0,2,3,5,6,13,14].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.formula) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });
  [1,4,7,8,9,10,11,12,15].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), cell: { userEnteredFormat: { backgroundColor: hex(C.input) } }, fields: 'userEnteredFormat.backgroundColor' } });
  });

  // Number formats
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 1, 2),   cell: { userEnteredFormat: { numberFormat: { type: 'DATE', pattern: 'mmm d, yyyy' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 9, 10),  cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmt.push({ repeatCell: { range: gridRange(SID, 7, 7+NDATA, 13, 15), cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00;[Red]-$#,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Dropdowns
  const dv = (ci, items) => ({ setDataValidation: { range: gridRange(SID, 7, 7+NDATA, ci, ci+1), rule: { condition: { type: 'ONE_OF_LIST', values: items.map(v=>({userEnteredValue:v})) }, showCustomUi: true } } });
  fmt.push(dv(7, ['Contribution','Withdrawal','Transfer In','Transfer Out','Adjustment Increase','Adjustment Decrease','Refund','Other']));
  fmt.push(dv(8, ['Paycheck','Monthly Budget','Bonus','Side Income','Tax Refund','Gift','Cash Windfall','Transfer','Other']));

  // ── Freeze rows 1-7, cols A-D ─────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' } });

  // Row heights
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 36 }, fields: 'pixelSize' } });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 7+NDATA }, properties: { pixelSize: 20 }, fields: 'pixelSize' } });

  // Column widths
  const WIDTHS = [70,90,50,70,70,160,90,100,110,80,80,80,80,80,90,200];
  WIDTHS.forEach((w,i) => fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } }));

  await batchUpdate(id, fmt, 'cl-fmt');
  await valuesBatchUpdate(id, vals, 'cl-vals');
  console.log(`✓ Contribution Log complete — ${TXNS.length} transactions written`);
})().catch(e => { console.error(e.message || e); process.exit(1); });
