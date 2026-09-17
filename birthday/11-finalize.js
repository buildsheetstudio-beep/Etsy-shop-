'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const RD = sheetMap['Reference Data'];
const PS = sheetMap['Party Setup'];
const GL = sheetMap['Guest List & RSVP'];
const IT = sheetMap['Invitation Tracker'];
const BE = sheetMap['Budget & Expenses'];
const FD = sheetMap['Food, Drinks & Cake'];
const SL = sheetMap['Shopping List'];
const PC = sheetMap['Prep, Setup & Cleanup'];
const DB = sheetMap['Dashboard'];

(async () => {
  const reqs = [];

  // ── Party Setup column widths ──────────────────────────────────────────────
  const psWidths = [200, 260, 20, 140, 20, 160, 20, 140];
  psWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PS, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });

  // ── Guest List column widths ───────────────────────────────────────────────
  // A(id)  B(first) C(last) D(type) E(household) F(age) G(email) H(phone) I(RSVP) J(adults) K(children) L(dietary) M(gift) N(thankyou) O(notes) P(invited by) Q(table) R(meal)
  const glWidths = [80, 100, 100, 100, 130, 60, 180, 120, 110, 65, 65, 120, 70, 80, 160, 100, 60, 80];
  glWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GL, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });

  // ── Invitation Tracker column widths ──────────────────────────────────────
  // A(guest id) B(name) C(household) D(method) E(sent?) F(sent date) G(msg preview) H(RSVP) I(follow up) J(notes) K(days since) L(reminder) M(contact) N(link)
  const itWidths = [80, 130, 130, 120, 65, 100, 160, 110, 90, 140, 80, 100, 180, 120];
  itWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: IT, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });

  // ── Budget & Expenses column widths ───────────────────────────────────────
  // A(expense id) B(date) C(category) D(subcategory) E(item) F(qty) G(unit cost) H(est cost) I(actual) J(payment) K(vendor) L(notes) M(receipt) N(approved)
  const beWidths = [80, 90, 120, 120, 160, 55, 80, 80, 80, 100, 120, 160, 70, 70];
  beWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: BE, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });

  // ── Food, Drinks & Cake column widths ─────────────────────────────────────
  // A(id) B(type) C(item) D(description) E(serves) F(servings/guest) G(qty) H(unit) I(cost) J(total) K(assigned) L(veg) M(vegan) N(gf) O(nut-free) P(status)
  const fdWidths = [70, 90, 160, 180, 65, 90, 60, 70, 70, 70, 120, 65, 65, 65, 65, 80];
  fdWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: FD, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });

  // ── Shopping List column widths ────────────────────────────────────────────
  // A(id) B(cat) C(item) D(description) E(status) F(qty needed) G(qty bought) H(unit) I(est price) J(actual price) K(purchased) L(store) M(packed) N(notes) O(link)
  const slWidths = [70, 100, 160, 160, 80, 75, 75, 60, 80, 80, 75, 120, 65, 140, 120];
  slWidths.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: SL, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });

  // ── Prep, Setup & Cleanup column widths ───────────────────────────────────
  // A(id) B(phase) C(task) D(description) E(assignee) F(priority) G(due date) H(priority cf) I(status) J(done) K(completed date) L(notes) M(overdue) N(days rem) O(effort) P(dependency)
  const pcWidths = [75, 120, 180, 180, 100, 80, 90, 0, 100, 60, 100, 160, 70, 80, 70, 140];
  pcWidths.forEach((w, i) => {
    if (w === 0) return; // skip hidden-like col H (priority cf display)
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: PC, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });

  // ── Dashboard column widths ────────────────────────────────────────────────
  // A-I content, J-L pivot data (narrow)
  [140, 180, 120, 50, 120, 180, 120, 50, 120].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DB, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });
  // Pivot cols J-L: narrow
  [100, 120, 120].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: DB, dimension: 'COLUMNS', startIndex: 9 + i, endIndex: 10 + i },
      properties: { pixelSize: w }, fields: 'pixelSize'
    }});
  });

  // ── Row heights: title rows taller ────────────────────────────────────────
  [PS, GL, IT, BE, FD, SL, PC, DB].forEach(sid => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: sid, dimension: 'ROWS', startIndex: 0, endIndex: 2 },
      properties: { pixelSize: 36 }, fields: 'pixelSize'
    }});
  });

  // ── Rename sheet tab color ─────────────────────────────────────────────────
  // Primary tabs: primary color; Dashboard: secondary
  [PS, GL, IT, BE, FD, SL, PC].forEach(sid => {
    reqs.push({ updateSheetProperties: {
      properties: { sheetId: sid, tabColorStyle: { rgbColor: hex(C.primary) } },
      fields: 'tabColorStyle'
    }});
  });
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: DB, tabColorStyle: { rgbColor: hex(C.secondary) } },
    fields: 'tabColorStyle'
  }});

  await batchUpdate(id, reqs, 'finalize');
  console.log('Finalize complete');
})();
