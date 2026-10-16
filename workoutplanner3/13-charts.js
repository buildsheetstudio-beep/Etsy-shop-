'use strict';
const { batchUpdate } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const YEARLY_SID = sheetMap['Yearly Fitness Planner'];
const DASH_SID   = sheetMap['Fitness Dashboard'];

// gridRange inline (avoid importing extra)
const gr = (sheetId, r1, r2, c1, c2) => ({ sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 });

// src: { sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex }
const srcRange = (sheetId, r1, r2, c1, c2) => ({ sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 });

function columnChart({ title, sheetId, anchorRow, anchorCol, width, height, domainC, seriesC, r1, r2, headerCount }) {
  return {
    addChart: {
      chart: {
        spec: {
          title,
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: '' },
              { position: 'LEFT_AXIS', title: '' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(sheetId, r1, r2, domainC, domainC+1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(sheetId, r1, r2, seriesC, seriesC+1)] } },
              targetAxis: 'LEFT_AXIS',
            }],
            headerCount,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId, rowIndex: anchorRow, columnIndex: anchorCol },
            widthPixels: width,
            heightPixels: height,
          },
        },
      },
    },
  };
}

function barChart({ title, sheetId, anchorRow, anchorCol, width, height, domainC, seriesC, r1, r2, headerCount }) {
  return {
    addChart: {
      chart: {
        spec: {
          title,
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'NO_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Sessions' },
              { position: 'LEFT_AXIS',   title: '' },
            ],
            domains: [{ domain: { sourceRange: { sources: [srcRange(sheetId, r1, r2, domainC, domainC+1)] } } }],
            series: [{
              series: { sourceRange: { sources: [srcRange(sheetId, r1, r2, seriesC, seriesC+1)] } },
              targetAxis: 'BOTTOM_AXIS',
            }],
            headerCount,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId, rowIndex: anchorRow, columnIndex: anchorCol },
            widthPixels: width,
            heightPixels: height,
          },
        },
      },
    },
  };
}

(async () => {
  const fmt = [];

  // ── Yearly Fitness Planner ────────────────────────────────────────────────
  // Monthly data: ri=2 (header row 3) through ri=14 (December, row 15)
  // Col A (0)=Month, Col B (1)=Workout Days, Col C (2)=Total Sets, Col D (3)=Est. Volume

  // Chart 1: Monthly workout sessions (column)
  // Anchor below annual goals section (ri=24 = row 25)
  fmt.push(columnChart({
    title: 'Monthly Workout Sessions',
    sheetId: YEARLY_SID,
    anchorRow: 24, anchorCol: 0,
    width: 480, height: 280,
    domainC: 0, seriesC: 1,
    r1: 2, r2: 15,
    headerCount: 1,
  }));

  // Chart 2: Monthly estimated volume (column)
  // Anchor at same row, offset to right (col 7)
  fmt.push(columnChart({
    title: 'Monthly Est. Volume (lb)',
    sheetId: YEARLY_SID,
    anchorRow: 24, anchorCol: 7,
    width: 480, height: 280,
    domainC: 0, seriesC: 3,
    r1: 2, r2: 15,
    headerCount: 1,
  }));

  // ── Fitness Dashboard ─────────────────────────────────────────────────────
  // Workout type breakdown table: ri=14 (sub-header) through ri=19 (HIIT)
  // Col A (0) = Type, Col B (1) = Sessions This Year (SUMPRODUCT formula)
  // Chart anchor below recent log entries (ri=34 = row 35)

  fmt.push(barChart({
    title: 'Sessions by Workout Type (This Year)',
    sheetId: DASH_SID,
    anchorRow: 34, anchorCol: 0,
    width: 450, height: 260,
    domainC: 0, seriesC: 1,
    r1: 14, r2: 20,
    headerCount: 1,
  }));

  await batchUpdate(id, fmt, '13-charts');
  console.log('✅  Charts done.');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
