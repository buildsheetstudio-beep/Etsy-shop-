'use strict';
const { batchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const BE = sheetMap['Budget & Expenses'];
const R = "'Reference Data'";

(async () => {
  const reqs = [];

  // Background
  reqs.push({ repeatCell: {
    range: gridRange(BE,0,400,0,14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});

  // Freeze
  reqs.push({ updateSheetProperties: {
    properties: { sheetId:BE, gridProperties:{ frozenRowCount:5 } },
    fields:'gridProperties.frozenRowCount'
  }});

  // Title row 1 (index 0)
  reqs.push({ mergeCells: { range: gridRange(BE,0,1,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BE,0,1,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:16, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:0, endIndex:1 },
    properties: { pixelSize:44 }, fields:'pixelSize'
  }});

  // Subtitle row 2 (index 1)
  reqs.push({ mergeCells: { range: gridRange(BE,1,2,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BE,1,2,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Info rows 3-4 (index 2-3)
  [2,3].forEach(ri => {
    reqs.push({ mergeCells: { range: gridRange(BE,ri,ri+1,0,14), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,0,14),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.bg),
        textFormat: { italic:true, fontSize:9, foregroundColor:hex(C.secText), fontFamily:'Arial' },
        horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:BE, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:22 }, fields:'pixelSize'
    }});
  });

  // Summary cards — two groups of 4 across rows 4-5 and 7-8 (indices)
  const cardPositions = [
    [4, 0, 3, 3, 7],
    [4, 7, 10, 10, 14],
    [5, 0, 3, 3, 7],
    [5, 7, 10, 10, 14],
    [7, 0, 3, 3, 7],
    [7, 7, 10, 10, 14],
    [8, 0, 3, 3, 7],
    [8, 7, 10, 10, 14],
  ];

  cardPositions.forEach(([ri, lc1, lc2, vc1, vc2]) => {
    reqs.push({ mergeCells: { range: gridRange(BE,ri,ri+1,lc1,lc2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,lc1,lc2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.primary),
        textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ mergeCells: { range: gridRange(BE,ri,ri+1,vc1,vc2), mergeType:'MERGE_ALL' } });
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,vc1,vc2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.panel),
        textFormat: { bold:true, fontSize:13, foregroundColor:hex(C.secondary), fontFamily:'Arial' },
        horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
    }});
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:BE, dimension:'ROWS', startIndex:ri, endIndex:ri+1 },
      properties: { pixelSize:38 }, fields:'pixelSize'
    }});
  });

  // Separator row 6 (index 6)
  reqs.push({ repeatCell: {
    range: gridRange(BE,6,7,0,14),
    cell: { userEnteredFormat: { backgroundColor: hex(C.bg) }},
    fields:'userEnteredFormat(backgroundColor)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:6, endIndex:7 },
    properties: { pixelSize:8 }, fields:'pixelSize'
  }});

  // Section header row 10 (index 9)
  reqs.push({ mergeCells: { range: gridRange(BE,9,10,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BE,9,10,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:9, endIndex:10 },
    properties: { pixelSize:28 }, fields:'pixelSize'
  }});

  // Category table column headers row 11 (index 10)
  reqs.push({ repeatCell: {
    range: gridRange(BE,10,11,0,6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});

  // Category data rows 12-27 (index 11-26) — alternating + currency formatting
  const CATS = ['Venue','Invitations','Decorations','Food','Drinks','Cake & Desserts',
    'Entertainment','Music','Audio / Visual','Gifts','Awards','Photography',
    'Supplies','Transportation','Setup & Cleanup','Miscellaneous'];
  CATS.forEach((_, i) => {
    const ri = 11 + i;
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,0,1),
      cell: { userEnteredFormat: {
        backgroundColor: ri%2===1 ? hex(C.altRow) : hex(C.panel),
        textFormat: { bold:true, fontSize:9, fontFamily:'Arial' },
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat)'
    }});
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,1,2),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.input),
        textFormat: { fontSize:9, fontFamily:'Arial' },
        numberFormat: { type:'CURRENCY', pattern:'$#,##0.00' }
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,numberFormat)'
    }});
    [2,3].forEach(ci => {
      reqs.push({ repeatCell: {
        range: gridRange(BE,ri,ri+1,ci,ci+1),
        cell: { userEnteredFormat: {
          backgroundColor: hex(C.formula),
          textFormat: { fontSize:9, fontFamily:'Arial' },
          numberFormat: { type:'CURRENCY', pattern:'$#,##0.00' }
        }},
        fields:'userEnteredFormat(backgroundColor,textFormat,numberFormat)'
      }});
    });
    // Variance % col E (index 4)
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,4,5),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.formula),
        textFormat: { fontSize:9, fontFamily:'Arial' },
        numberFormat: { type:'PERCENT', pattern:'0.0%' }
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,numberFormat)'
    }});
    // Status col F (index 5)
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,5,6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(C.formula),
        textFormat: { bold:true, fontSize:9, fontFamily:'Arial' },
        horizontalAlignment:'CENTER'
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
    }});
    // alt-row cols 6-13
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,6,14),
      cell: { userEnteredFormat: {
        backgroundColor: ri%2===1 ? hex(C.altRow) : hex(C.panel),
        textFormat: { fontSize:9, fontFamily:'Arial' }
      }},
      fields:'userEnteredFormat(backgroundColor,textFormat)'
    }});
  });

  // Totals row 28 (index 27) — bold
  reqs.push({ repeatCell: {
    range: gridRange(BE,27,28,0,6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      numberFormat: { type:'CURRENCY', pattern:'$#,##0.00' }
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,numberFormat)'
  }});

  // Status CF rows 12-27 col F
  const statusCF = [
    { val:'Under Budget', bg:C.success },
    { val:'Near Budget',  bg:C.warning },
    { val:'Over Budget',  bg:C.attention },
    { val:'No Budget Set',bg:C.border },
  ];
  statusCF.forEach((r, idx) => {
    reqs.push({ addConditionalFormatRule: { index:idx, rule: {
      ranges:[gridRange(BE,11,27,5,6)],
      booleanRule: {
        condition:{ type:'TEXT_EQ', values:[{userEnteredValue:r.val}] },
        format:{ backgroundColor:hex(r.bg) }
      }
    }}});
  });

  // Expense Log section header row 29 (index 28)
  reqs.push({ mergeCells: { range: gridRange(BE,28,29,0,14), mergeType:'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(BE,28,29,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.secondary),
      textFormat: { bold:true, fontSize:10, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'LEFT', verticalAlignment:'MIDDLE'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:28, endIndex:29 },
    properties: { pixelSize:28 }, fields:'pixelSize'
  }});

  // Expense log column headers row 30 (index 29)
  reqs.push({ repeatCell: {
    range: gridRange(BE,29,30,0,14),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold:true, fontSize:9, foregroundColor:hex(C.white), fontFamily:'Arial' },
      horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE', wrapStrategy:'WRAP'
    }},
    fields:'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)'
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId:BE, dimension:'ROWS', startIndex:29, endIndex:30 },
    properties: { pixelSize:40 }, fields:'pixelSize'
  }});

  // Alternating rows in expense log (index 30-328)
  for (let ri = 30; ri < 329; ri += 2) {
    reqs.push({ repeatCell: {
      range: gridRange(BE,ri,ri+1,0,14),
      cell: { userEnteredFormat: { backgroundColor: hex(C.altRow) }},
      fields:'userEnteredFormat(backgroundColor)'
    }});
  }

  // Formula cols A, H styling
  [0,7].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(BE,30,329,ci,ci+1),
      cell: { userEnteredFormat: { backgroundColor: hex(C.formula), textFormat:{fontSize:9,fontFamily:'Arial'} }},
      fields:'userEnteredFormat(backgroundColor,textFormat)'
    }});
  });

  // Currency for G, H, I cols in expense log
  [6,7,8].forEach(ci => {
    reqs.push({ repeatCell: {
      range: gridRange(BE,30,329,ci,ci+1),
      cell: { userEnteredFormat: {
        textFormat:{fontSize:9,fontFamily:'Arial'},
        numberFormat:{ type:'CURRENCY', pattern:'$#,##0.00' }
      }},
      fields:'userEnteredFormat(textFormat,numberFormat)'
    }});
  });

  // Payment Status CF col J (index 9)
  const paymentCF = [
    { val:'Paid',           bg:C.success },
    { val:'Deposit Paid',   bg:C.mutedBlue },
    { val:'Partially Paid', bg:C.warning },
    { val:'Planned',        bg:C.bg },
    { val:'Cancelled',      bg:C.border },
    { val:'Refunded',       bg:C.mutedMauve },
  ];
  paymentCF.forEach((r, idx) => {
    reqs.push({ addConditionalFormatRule: { index:4+idx, rule: {
      ranges:[gridRange(BE,30,329,9,10)],
      booleanRule: {
        condition:{ type:'TEXT_EQ', values:[{userEnteredValue:r.val}] },
        format:{ backgroundColor:hex(r.bg) }
      }
    }}});
  });

  // Data validations
  const dvReqs = [];
  // Receipt checkbox col M (index 12)
  dvReqs.push({ setDataValidation: {
    range: gridRange(BE,30,329,12,13),
    rule: { condition:{ type:'BOOLEAN' }, showCustomUi:true }
  }});
  // Payment Status dropdown col J (index 9)
  dvReqs.push({ setDataValidation: {
    range: gridRange(BE,30,329,9,10),
    rule: { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$I$2:$I$7` }] }, showCustomUi:true, strict:false }
  }});
  // Category dropdown col C (index 2)
  dvReqs.push({ setDataValidation: {
    range: gridRange(BE,30,329,2,3),
    rule: { condition:{ type:'ONE_OF_RANGE', values:[{ userEnteredValue:`=${R}!$G$2:$G$17` }] }, showCustomUi:true, strict:false }
  }});

  // Column widths
  const colW = [80,90,120,140,200,60,80,100,100,110,90,90,70,160];
  colW.forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId:BE, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
      properties: { pixelSize:w }, fields:'pixelSize'
    }});
  });

  await batchUpdate(id, reqs);
  await batchUpdate(id, dvReqs);
  console.log('Budget & Expenses formatting applied');
})();
