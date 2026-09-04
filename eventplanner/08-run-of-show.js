'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const ROS = sheetMap['⏱️ Run-of-Show'];

(async () => {
  const reqs = [];

  // Row heights
  reqs.push({ updateDimensionProperties: { range: { sheetId: ROS, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 52 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: ROS, dimension: 'ROWS', startIndex: 1, endIndex: 3 }, properties: { pixelSize: 28 }, fields: 'pixelSize' } });
  reqs.push({ updateDimensionProperties: { range: { sheetId: ROS, dimension: 'ROWS', startIndex: 3, endIndex: 45 }, properties: { pixelSize: 24 }, fields: 'pixelSize' } });

  // Column widths: A(#30) B(Segment 200) C(Type 120) D(Event 180) E(Start 90) F(End 90) G(Duration 80) H(Location 140) I(Lead 140) J(Notes 220) K(Status 100) L(Alert 70)
  const colW = [30, 200, 120, 180, 90, 90, 80, 140, 140, 220, 100, 70];
  colW.forEach((w,i) => {
    reqs.push({ updateDimensionProperties: { range: { sheetId: ROS, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 }, properties: { pixelSize: w }, fields: 'pixelSize' } });
  });

  // Title row
  reqs.push({ mergeCells: { range: gridRange(ROS, 0, 1, 0, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(ROS, 0, 1, 0, 12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.slateBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Event selector row 2
  reqs.push({ mergeCells: { range: gridRange(ROS, 1, 2, 0, 3), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(ROS, 1, 2, 3, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(ROS, 1, 2, 7, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(ROS, 1, 2, 0, 3),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(ROS, 1, 2, 3, 7),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.ivory),
      textFormat: { foregroundColor: hex(C.slateBlue), bold: true, fontSize: 11 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(ROS, 1, 2, 7, 12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Sub-header row 3 (date, totals)
  reqs.push({ mergeCells: { range: gridRange(ROS, 2, 3, 0, 3), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(ROS, 2, 3, 3, 7), mergeType: 'MERGE_ALL' } });
  reqs.push({ mergeCells: { range: gridRange(ROS, 2, 3, 7, 12), mergeType: 'MERGE_ALL' } });
  reqs.push({ repeatCell: {
    range: gridRange(ROS, 2, 3, 0, 12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.champagne),
      textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Header row 4
  reqs.push({ repeatCell: {
    range: gridRange(ROS, 3, 4, 0, 12),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.slateBlue),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Data rows alternating
  for (let r = 4; r < 45; r++) {
    const bg = r % 2 === 0 ? C.champagne : C.ivory;
    reqs.push({ repeatCell: {
      range: gridRange(ROS, r, r+1, 0, 12),
      cell: { userEnteredFormat: { backgroundColor: hex(bg), textFormat: { foregroundColor: hex(C.darkPlum), fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Time format cols E, F (indices 4, 5)
  [4, 5].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(ROS, 4, 45, c, c+1),
      cell: { userEnteredFormat: { numberFormat: { type: 'TIME', pattern: 'h:mm AM/PM' } } },
      fields: 'userEnteredFormat.numberFormat',
    }});
  });

  // Duration col G (index 6) — formula, paleRose
  reqs.push({ repeatCell: {
    range: gridRange(ROS, 4, 45, 6, 7),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose), horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
  }});

  // Alert col L (index 11) — formula, paleRose
  reqs.push({ repeatCell: {
    range: gridRange(ROS, 4, 45, 11, 12),
    cell: { userEnteredFormat: { backgroundColor: hex(C.paleRose), horizontalAlignment: 'CENTER' } },
    fields: 'userEnteredFormat(backgroundColor,horizontalAlignment)',
  }});

  // Center align cols A, C, E, F, K
  [0, 2, 4, 5, 10].forEach(c => {
    reqs.push({ repeatCell: { range: gridRange(ROS, 3, 45, c, c+1), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } }, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(ROS, 3, 45, 0, 12),
    innerHorizontal: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
    bottom: { style: 'SOLID', color: hex(C.dustyRose), width: 1 },
  }});

  await batchUpdate(id, reqs, 'ros-format');

  const data = [];
  data.push({ range: `'⏱️ Run-of-Show'!A1`, values: [['⏱️ Run-of-Show — Event Schedule & Timeline']] });
  data.push({ range: `'⏱️ Run-of-Show'!A2`, values: [['  📅 Event:']] });
  data.push({ range: `'⏱️ Run-of-Show'!D2`, values: [['Tech Summit 2025']] });
  data.push({ range: `'⏱️ Run-of-Show'!H2`, values: [['Date: 15 September 2025 | Venue: Grand Convention Centre']] });
  data.push({ range: `'⏱️ Run-of-Show'!A3`, values: [['  🕐 Segments:']] });
  data.push({ range: `'⏱️ Run-of-Show'!D3`, values: [['=COUNTA(B5:B45)']] });
  data.push({ range: `'⏱️ Run-of-Show'!H3`, values: [['Total Run Time: =TEXT(IFERROR(SUMIF(C5:C45,"<>Setup",G5:G45)*24,0),"0.0")&" hours"']] });
  data.push({ range: `'⏱️ Run-of-Show'!A4:L4`, values: [['#','Segment / Activity','Type','Event','Start Time','End Time','Duration','Location','Lead Person','Notes','Status','⚠️']] });

  // Times as decimal fractions of a day
  const t = (h, m) => `${h/24 + m/1440}`;

  const segments = [
    [1,'Venue Setup & AV Check','Setup','Tech Summit 2025','0.291666667','0.354166667','=TEXT((F5-E5)*24,"0.0")&"h"','Main Hall','Mark Davis','Full tech check + mic test','Done',''],
    [2,'Volunteer Briefing','Setup','Tech Summit 2025','0.354166667','0.375','=TEXT((F6-E6)*24,"0.0")&"h"','Foyer','Lisa Chen','All volunteers to gather at foyer entrance','Done',''],
    [3,'Registration & Welcome Desk Opens','Registration','Tech Summit 2025','0.375','0.458333333','=TEXT((F7-E7)*24,"0.0")&"h"','Main Entrance','Lisa Chen','Badge collection, swag bags','Done',''],
    [4,'Morning Tea & Networking','Break','Tech Summit 2025','0.375','0.416666667','=TEXT((F8-E8)*24,"0.0")&"h"','Atrium','','Catering set up 30 min prior','Done',''],
    [5,'Opening Address — CEO','Opening','Tech Summit 2025','0.458333333','0.479166667','=TEXT((F9-E9)*24,"0.0")&"h"','Main Stage','Sarah Johnson','Welcome + housekeeping','Done','=IFERROR(IF(A9<>A8,"",IF(C9<D8,"⚠️ OVERLAP",IF(C9>(D8+TIME(0,15,0)),"⚡ GAP",""))),"")'],
    [6,'Keynote: James Hartley — AI Future','Keynote','Tech Summit 2025','0.479166667','0.541666667','=TEXT((F10-E10)*24,"0.0")&"h"','Main Stage','James Hartley','Run slides 10 min before start','Done','=IFERROR(IF(A10<>A9,"",IF(C10<D9,"⚠️ OVERLAP",IF(C10>(D9+TIME(0,15,0)),"⚡ GAP",""))),"")'],
    [7,'Q&A — Keynote','Panel','Tech Summit 2025','0.541666667','0.5625','=TEXT((F11-E11)*24,"0.0")&"h"','Main Stage','Oliver Chen','Mic runners to be ready','Done','=IFERROR(IF(A11<>A10,"",IF(C11<D10,"⚠️ OVERLAP",IF(C11>(D10+TIME(0,15,0)),"⚡ GAP",""))),"")'],
    [8,'Lunch Break & Sponsor Expo','Meal','Tech Summit 2025','0.5625','0.625','=TEXT((F12-E12)*24,"0.0")&"h"','Atrium + Foyer','Mark Davis','Sponsor tables to be ready by 1pm','Done','=IFERROR(IF(A12<>A11,"",IF(C12<D11,"⚠️ OVERLAP",IF(C12>(D11+TIME(0,15,0)),"⚡ GAP",""))),"")'],
    [9,'Workshop A: Cloud Architecture','Workshop','Tech Summit 2025','0.625','0.708333333','=TEXT((F13-E13)*24,"0.0")&"h"','Room 1A','Priya Sharma','Max 40 attendees','In Progress','=IFERROR(IF(A13<>A12,"",IF(C13<D12,"⚠️ OVERLAP",IF(C13>(D12+TIME(0,15,0)),"⚡ GAP",""))),"")'],
    [10,'Workshop B: FinTech Innovation','Workshop','Tech Summit 2025','0.625','0.708333333','=TEXT((F14-E14)*24,"0.0")&"h"','Room 2B','Ethan Kowalski','Max 35 attendees','In Progress','=IFERROR(IF(A14<>A13,"",IF(C14<D13,"⚠️ OVERLAP",IF(C14>(D13+TIME(0,15,0)),"⚡ GAP",""))),"")'],
    [11,'Afternoon Tea','Break','Tech Summit 2025','0.708333333','0.729166667','=TEXT((F15-E15)*24,"0.0")&"h"','Atrium','','','Not Started','=IFERROR(IF(A15<>A14,"",IF(C15<D14,"⚠️ OVERLAP",IF(C15>(D14+TIME(0,15,0)),"⚡ GAP",""))),"")'],
    [12,'Panel: Future of Work','Panel','Tech Summit 2025','0.729166667','0.8125','=TEXT((F16-E16)*24,"0.0")&"h"','Main Stage','Oliver Chen','4 panellists confirmed','Not Started','=IFERROR(IF(A16<>A15,"",IF(C16<D15,"⚠️ OVERLAP",IF(C16>(D15+TIME(0,15,0)),"⚡ GAP",""))),"")'],
    [13,'Closing Remarks & Prize Draw','Closing','Tech Summit 2025','0.8125','0.833333333','=TEXT((F17-E17)*24,"0.0")&"h"','Main Stage','Sarah Johnson','Prize draw for all attendees','Not Started','=IFERROR(IF(A17<>A16,"",IF(C17<D16,"⚠️ OVERLAP",IF(C17>(D16+TIME(0,15,0)),"⚡ GAP",""))),"")'],
    [14,'Venue Teardown & Pack Out','Teardown','Tech Summit 2025','0.833333333','0.916666667','=TEXT((F18-E18)*24,"0.0")&"h"','All Areas','Mark Davis','AV crew to dismantle + venue clean','Not Started','=IFERROR(IF(A18<>A17,"",IF(C18<D17,"⚠️ OVERLAP",IF(C18>(D17+TIME(0,15,0)),"⚡ GAP",""))),"")'],
  ];

  data.push({ range: `'⏱️ Run-of-Show'!A5:L18`, values: segments });

  await valuesBatchUpdate(id, data, 'ros-values');
  console.log('Run-of-Show complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
