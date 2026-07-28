'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));

const DB  = sheetMap['Dashboard'];
const MB  = sheetMap['Monthly Budget Impact'];
const TCO = sheetMap['Total Cost of Ownership'];
const VC  = sheetMap['Vehicle Comparison'];
const LC  = sheetMap['Auto Loan Calculator'];
const LVB = sheetMap['Lease vs Buy'];

function colChart(title, sheetId, headerRow, dataRows, seriesCol, domainCol, anchor, width, height) {
  return {
    addChart: {
      chart: {
        spec: {
          title,
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: '' },
              { position: 'LEFT_AXIS', title: '' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [{ sheetId, startRowIndex: headerRow, endRowIndex: dataRows, startColumnIndex: domainCol, endColumnIndex: domainCol+1 }] }
              }
            }],
            series: [{
              series: {
                sourceRange: { sources: [{ sheetId, startRowIndex: headerRow, endRowIndex: dataRows, startColumnIndex: seriesCol, endColumnIndex: seriesCol+1 }] }
              },
              type: 'COLUMN',
              color: hex(C.primary),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: anchor.sheetId || sheetId, rowIndex: anchor.row, columnIndex: anchor.col },
            widthPixels: width,
            heightPixels: height,
          }
        }
      }
    }
  };
}

function barChart(title, sheetId, headerRow, dataRows, domainCol, seriesCols, colors, anchor, width, height) {
  return {
    addChart: {
      chart: {
        spec: {
          title,
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: '' },
              { position: 'LEFT_AXIS', title: '' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [{ sheetId, startRowIndex: headerRow, endRowIndex: dataRows, startColumnIndex: domainCol, endColumnIndex: domainCol+1 }] }
              }
            }],
            series: seriesCols.map((col, i) => ({
              series: {
                sourceRange: { sources: [{ sheetId, startRowIndex: headerRow, endRowIndex: dataRows, startColumnIndex: col, endColumnIndex: col+1 }] }
              },
              type: 'COLUMN',
              color: hex(colors[i] || C.primary),
            })),
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: anchor.sheetId || sheetId, rowIndex: anchor.row, columnIndex: anchor.col },
            widthPixels: width,
            heightPixels: height,
          }
        }
      }
    }
  };
}

function lineChart(title, sheetId, headerRow, dataRows, domainCol, seriesCols, colors, anchor, width, height) {
  return {
    addChart: {
      chart: {
        spec: {
          title,
          basicChart: {
            chartType: 'LINE',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Payment #' },
              { position: 'LEFT_AXIS', title: 'Balance ($)' },
            ],
            domains: [{
              domain: {
                sourceRange: { sources: [{ sheetId, startRowIndex: headerRow, endRowIndex: dataRows, startColumnIndex: domainCol, endColumnIndex: domainCol+1 }] }
              }
            }],
            series: seriesCols.map((col, i) => ({
              series: {
                sourceRange: { sources: [{ sheetId, startRowIndex: headerRow, endRowIndex: dataRows, startColumnIndex: col, endColumnIndex: col+1 }] }
              },
              color: hex(colors[i] || C.primary),
            })),
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: anchor.sheetId || sheetId, rowIndex: anchor.row, columnIndex: anchor.col },
            widthPixels: width,
            heightPixels: height,
          }
        }
      }
    }
  };
}

(async () => {
  const reqs = [];

  // ── Dashboard: Monthly Cost Comparison (anchor at row 40, col 0) ──────────
  // Uses Vehicle Comparison tab cols A (Vehicle ID, col 0) and H (Total Monthly, col 7)
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Monthly Vehicle Cost Comparison',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS' }, { position: 'LEFT_AXIS', title: 'Monthly Cost ($)' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: VC, startRowIndex: 15, endRowIndex: 26, startColumnIndex: 0, endColumnIndex: 1 }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: VC, startRowIndex: 15, endRowIndex: 26, startColumnIndex: 7, endColumnIndex: 8 }] } },
              type: 'COLUMN', color: hex(C.primary),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DB, rowIndex: 40, columnIndex: 0 },
            widthPixels: 480, heightPixels: 280,
          }
        }
      }
    }
  });

  // ── Dashboard: 5-Year Ownership Cost Comparison ───────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: '5-Year Ownership Cost Comparison',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS' }, { position: 'LEFT_AXIS', title: '5-Year Cost ($)' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: VC, startRowIndex: 15, endRowIndex: 26, startColumnIndex: 0, endColumnIndex: 1 }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: VC, startRowIndex: 15, endRowIndex: 26, startColumnIndex: 8, endColumnIndex: 9 }] } },
              type: 'COLUMN', color: hex(C.secondary),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DB, rowIndex: 40, columnIndex: 5 },
            widthPixels: 480, heightPixels: 280,
          }
        }
      }
    }
  });

  // ── Dashboard: Loan Balance Over Time (from Loan Calc amortization) ───────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Loan Balance Over Time',
          basicChart: {
            chartType: 'LINE',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS', title: 'Payment #' }, { position: 'LEFT_AXIS', title: 'Balance ($)' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: LC, startRowIndex: 26, endRowIndex: 111, startColumnIndex: 0, endColumnIndex: 1 }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: LC, startRowIndex: 26, endRowIndex: 111, startColumnIndex: 7, endColumnIndex: 8 }] } },
              color: hex(C.primary),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DB, rowIndex: 57, columnIndex: 0 },
            widthPixels: 480, heightPixels: 280,
          }
        }
      }
    }
  });

  // ── Dashboard: Cost per Mile by Vehicle ───────────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Cost per Mile by Vehicle',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS', title: '$/mile' }, { position: 'LEFT_AXIS' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: VC, startRowIndex: 15, endRowIndex: 26, startColumnIndex: 0, endColumnIndex: 1 }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: VC, startRowIndex: 15, endRowIndex: 26, startColumnIndex: 9, endColumnIndex: 10 }] } },
              type: 'COLUMN', color: hex(C.attention),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: DB, rowIndex: 57, columnIndex: 5 },
            widthPixels: 480, heightPixels: 280,
          }
        }
      }
    }
  });

  // ── Monthly Budget Impact: Monthly Cost by Vehicle (stacked) ─────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Monthly Cost by Category',
          basicChart: {
            chartType: 'COLUMN',
            stackedType: 'STACKED',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS' }, { position: 'LEFT_AXIS', title: 'Monthly Cost ($)' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: MB, startRowIndex: 11, endRowIndex: 12, startColumnIndex: 3, endColumnIndex: 11 }] } }
            }],
            series: [15,16,17,18].map((ri, i) => ({
              series: { sourceRange: { sources: [{ sheetId: MB, startRowIndex: ri-1, endRowIndex: ri, startColumnIndex: 3, endColumnIndex: 11 }] } },
              type: 'COLUMN',
              color: hex([C.primary, C.secondary, C.attention, C.warning][i]),
            })),
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: MB, rowIndex: 36, columnIndex: 0 },
            widthPixels: 560, heightPixels: 300,
          }
        }
      }
    }
  });

  // ── Monthly Budget Impact: Remaining Discretionary ────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Remaining Discretionary Income by Vehicle',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS' }, { position: 'LEFT_AXIS', title: '$' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: MB, startRowIndex: 11, endRowIndex: 12, startColumnIndex: 3, endColumnIndex: 11 }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: MB, startRowIndex: 29, endRowIndex: 30, startColumnIndex: 3, endColumnIndex: 11 }] } },
              type: 'COLUMN', color: hex(C.success),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: MB, rowIndex: 36, columnIndex: 7 },
            widthPixels: 480, heightPixels: 300,
          }
        }
      }
    }
  });

  // ── TCO: Total Cost of Ownership Comparison ───────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Total Cost of Ownership Comparison',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS' }, { position: 'LEFT_AXIS', title: 'Total Cost ($)' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: TCO, startRowIndex: 14, endRowIndex: 15, startColumnIndex: 3, endColumnIndex: 11 }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: TCO, startRowIndex: 29, endRowIndex: 30, startColumnIndex: 3, endColumnIndex: 11 }] } },
              type: 'COLUMN', color: hex(C.primary),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: TCO, rowIndex: 44, columnIndex: 0 },
            widthPixels: 560, heightPixels: 300,
          }
        }
      }
    }
  });

  // ── TCO: Estimated Resale Value ───────────────────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Estimated Resale Value by Vehicle',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS' }, { position: 'LEFT_AXIS', title: 'Resale Value ($)' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: TCO, startRowIndex: 14, endRowIndex: 15, startColumnIndex: 3, endColumnIndex: 11 }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: TCO, startRowIndex: 28, endRowIndex: 29, startColumnIndex: 3, endColumnIndex: 11 }] } },
              type: 'COLUMN', color: hex(C.success),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: TCO, rowIndex: 44, columnIndex: 7 },
            widthPixels: 480, heightPixels: 300,
          }
        }
      }
    }
  });

  // ── Vehicle Comparison: Weighted Score ────────────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Weighted Score by Vehicle',
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS', title: 'Score' }, { position: 'LEFT_AXIS' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: VC, startRowIndex: 15, endRowIndex: 26, startColumnIndex: 1, endColumnIndex: 2 }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: VC, startRowIndex: 15, endRowIndex: 26, startColumnIndex: 20, endColumnIndex: 21 }] } },
              type: 'COLUMN', color: hex(C.primary),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: VC, rowIndex: 36, columnIndex: 0 },
            widthPixels: 480, heightPixels: 300,
          }
        }
      }
    }
  });

  // ── Loan Calculator: Principal vs Interest ────────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Principal vs Interest Paid',
          basicChart: {
            chartType: 'COLUMN',
            stackedType: 'STACKED',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS', title: 'Payment #' }, { position: 'LEFT_AXIS', title: '$' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: LC, startRowIndex: 26, endRowIndex: 111, startColumnIndex: 0, endColumnIndex: 1 }] } }
            }],
            series: [
              { series: { sourceRange: { sources: [{ sheetId: LC, startRowIndex: 26, endRowIndex: 111, startColumnIndex: 5, endColumnIndex: 6 }] } }, type: 'COLUMN', color: hex(C.primary) },
              { series: { sourceRange: { sources: [{ sheetId: LC, startRowIndex: 26, endRowIndex: 111, startColumnIndex: 6, endColumnIndex: 7 }] } }, type: 'COLUMN', color: hex(C.rust) },
            ],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: LC, rowIndex: 2, columnIndex: 7 },
            widthPixels: 480, heightPixels: 280,
          }
        }
      }
    }
  });

  // ── Lease vs Buy: Total Cost Comparison ───────────────────────────────────
  reqs.push({
    addChart: {
      chart: {
        spec: {
          title: 'Lease vs Buy — Total Cash Outflow',
          basicChart: {
            chartType: 'COLUMN',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [{ position: 'BOTTOM_AXIS' }, { position: 'LEFT_AXIS', title: 'Total Cost ($)' }],
            domains: [{
              domain: { sourceRange: { sources: [{ sheetId: LVB, startRowIndex: 29, endRowIndex: 30, startColumnIndex: 2, endColumnIndex: 4 }] } }
            }],
            series: [{
              series: { sourceRange: { sources: [{ sheetId: LVB, startRowIndex: 35, endRowIndex: 36, startColumnIndex: 2, endColumnIndex: 4 }] } },
              type: 'COLUMN', color: hex(C.secondary),
            }],
            headerCount: 1,
          }
        },
        position: {
          overlayPosition: {
            anchorCell: { sheetId: LVB, rowIndex: 48, columnIndex: 0 },
            widthPixels: 480, heightPixels: 280,
          }
        }
      }
    }
  });

  await batchUpdate(id, reqs, 'charts');
  console.log('✓ Charts');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
