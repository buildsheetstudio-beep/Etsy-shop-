'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const ITN = sheetMap['🗓️ Event Itinerary & Activities'];

// Itinerary: [day, startTime, endTime, activity, type, ledBy, location, supplies, notes]
const ITINERARY = [
  ['Day 1','2:00 PM','5:00 PM','Guest Arrival & Check-In','Welcome','Organiser','Driveway + Cabins','Welcome bags, name tags, parking guide',''],
  ['Day 1','5:00 PM','5:30 PM','Welcome Drinks & Mingling','Welcome','All','Back Deck','Drinks table, ice, glasses',''],
  ['Day 1','5:30 PM','5:45 PM','Official Welcome Speech & Toast','Speech','Organiser','Back Deck','Microphone optional',''],
  ['Day 1','5:45 PM','6:00 PM','Icebreaker: Two Truths & A Lie','Games','Co-Organiser','Back Deck','Game cards',''],
  ['Day 1','6:00 PM','7:30 PM','BBQ Dinner — Meats & Salads','Meal','BBQ Team','Outdoor Area','BBQ, tongs, platters, salads',''],
  ['Day 1','7:30 PM','8:00 PM','Family Trivia Round 1','Games','Organiser','Dining Area','Quiz sheets, pens, answer key',''],
  ['Day 1','8:00 PM','8:30 PM','Family Photo Slideshow','Photo','Co-Organiser','Living Room','Laptop + TV/projector',''],
  ['Day 1','8:30 PM','10:00 PM','Evening — Bonfire & Free Time','Free Time','All','Fire Pit Area','Firewood, marshmallows, chairs',''],
  ['Day 2','7:30 AM','8:30 AM','Breakfast — Communal Buffet','Meal','Kitchen Team','Dining Area','Eggs, toast, fruit, coffee',''],
  ['Day 2','8:30 AM','9:30 AM','Morning Walk & Explore','Activity','Optional','Local Trails','Water bottles, hats, sunscreen',''],
  ['Day 2','9:30 AM','10:30 AM','Lawn Games — Bocce, Cricket, Frisbee','Games','Co-Organiser','Back Lawn','Bocce set, cricket gear, frisbees',''],
  ['Day 2','10:30 AM','11:00 AM','Family Trivia Round 2 — Final','Games','Organiser','Back Deck','Quiz sheets, prizes, scoreboard',''],
  ['Day 2','11:00 AM','11:30 AM','Prize Ceremony + Funny Awards','Games','All','Back Deck','Prizes, certificate printouts',''],
  ['Day 2','11:30 AM','12:00 PM','Group Family Photo — All Members','Photo','Photographer','Back Lawn','Camera, tripod, photo props',''],
  ['Day 2','12:00 PM','1:30 PM','Farewell Lunch — Leftovers + Fresh','Meal','Kitchen Team','Outdoor Tables','Leftovers, fresh rolls, salads',''],
  ['Day 2','1:30 PM','2:00 PM','Closing Words & Memory Sharing','Speech','Organiser','Back Deck','Keepsake cards for memories',''],
  ['Day 2','2:00 PM','2:30 PM','T-Shirt Distribution + Gift Packs','Activity','Organiser','Driveway','Reunion t-shirts, gift packs',''],
  ['Day 2','2:30 PM','4:00 PM','Guest Departures & Farewells','Farewell','All','Driveway','Transport coordination',''],
  ['Day 2','4:00 PM','5:00 PM','Pack Down & Venue Clean Up','Setup','Helpers','All Areas','Bins, boxes, cleaning supplies',''],
  ['Day 2','5:00 PM','5:30 PM','Final Organiser Check — Venue Handback','Setup','Organiser','All Areas','Checklist, keys',''],
];

// Activities: [name, type, description, ages, duration, supplies, status]
const ACTIVITIES = [
  ['Family Trivia Quiz','Trivia','20-question quiz covering family history, funny moments, and general knowledge. Divide into teams by family branch. Host reads questions; teams submit written answers.','All Ages',45,'Printed quiz sheets, pens, answer key, scoreboard, prizes','Planned'],
  ['Two Truths & A Lie — Family Edition','Games','Each person shares 3 statements about themselves — 2 true, 1 false. Others guess which is the lie. Great icebreaker for the first evening.','All Ages',20,'Game cards or paper slips, pens','Planned'],
  ['Bocce Ball Tournament','Outdoor','Teams of 2 compete in a round-robin bocce tournament. Bracket drawn randomly — winners progress. Losers bracket consolation games.','All Ages',60,'Bocce ball set, scoresheet, prizes','Supplies Bought'],
  ['Backyard Cricket','Outdoor','Classic Australian backyard cricket. Set up a pitch, rotate batting and fielding. Modified rules include younger kids (soft ball for children).','All Ages',60,'Cricket set (soft ball), witches hats for stumps, sunscreen','Supplies Bought'],
  ['Giant Jenga','Games','Oversized Jenga tower — pull blocks without toppling the tower. Write family challenge forfeits on blocks for extra fun.','All Ages',30,'Giant Jenga set, marker for writing forfeits','Planned'],
  ['Family Photo Challenge','Creative','Give each family group a list of 15 photo challenges to complete across the reunion (silliest face, oldest + youngest, everyone named David, etc.). Best album wins.','All Ages',0,'Phone camera, challenge list, judge for final album','Planned'],
  ['Kids Treasure Hunt','For Kids','Hidden clue treasure hunt across the property. Kids work in teams following riddles to find the treasure. Clues pre-hidden by adults.','Kids Only',40,'Clue cards, small treasure/prize bags, hiding spots set up in advance','Planned'],
  ['Family Recipe Swap','Creative','Each family unit brings a printed copy of their most treasured family recipe to exchange. Bonus: bring a small sample dish. Organiser compiles into a recipe booklet.','All Ages',30,'Printed recipe cards, folder or booklet','Planned'],
  ['Memory Lane Video','Indoor','Play a compilation video of old family photos and videos. Guests shout out names, dates, and stories. Bring tissues — this one is emotional!','All Ages',20,'Laptop + TV/projector, video file prepared in advance','Planned'],
  ['Volleyball / Badminton','Outdoor','Set up a net for rotating volleyball or badminton. Teams of 4–6 rotate each game. No scorekeeping — just fun.','Teens + Adults',45,'Volleyball/badminton net, ball/shuttlecocks','Planned'],
  ['Funny Family Awards Night','Games','Prepare 15+ funny award categories (Best Dad Joke, Most Likely to Be Late, Best Cook, etc.). Family votes in advance. MC announces winners with trophy or printed certificate.','All Ages',30,'Award categories, certificates or trophies, voting slips','Planned'],
  ['Bonfire + Storytelling','Outdoor','End-of-day bonfire. Each family elder shares one family story or memory. Younger guests contribute questions. Roast marshmallows.','All Ages',90,'Firewood, fire starter, marshmallows, skewers, seating','Supplies Bought'],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'🗓️ Event Itinerary & Activities'";
  const ITNCOLS = 11; // A-K
  const ACTCOLS = 8;  // A-H

  // Column widths for itinerary cols
  [80,90,90,90,220,120,140,140,200,80,180].forEach((w,i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: ITN, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row 0: title banner
  reqs.push({ mergeCells: { range: gridRange(ITN,0,1,0,ITNCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ITN,0,1,0,ITNCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A1`, values: [['  🗓️  Event Itinerary & Activities — Henderson Family Reunion 2025']] });

  // Row 1: ITINERARY section header (0-indexed)
  reqs.push({ mergeCells: { range: gridRange(ITN,1,2,0,ITNCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ITN,1,2,0,ITNCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['  🕐  ITINERARY — RUN OF SHOW']] });

  // Row 2: Itinerary column headers (0-indexed = GSheets row 3)
  const itnHdrs = ['DAY','START TIME','END TIME','DURATION (MINS)','ACTIVITY / SEGMENT','TYPE','LED BY','LOCATION','SUPPLIES','TIME CHECK','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(ITN,2,3,0,ITNCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'ROWS', startIndex: 2, endIndex: 3 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A3`, values: [itnHdrs] });

  // Itinerary data rows (0-indexed 3-22 = GSheets rows 4-23)
  reqs.push({ repeatCell: {
    range: gridRange(ITN,3,3+ITINERARY.length,0,ITNCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.white),
      textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
      verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
  }});
  for (let i = 0; i < ITINERARY.length; i++) {
    if (i % 2 === 1) {
      reqs.push({ repeatCell: {
        range: gridRange(ITN,3+i,4+i,0,ITNCOLS),
        cell: { userEnteredFormat: { backgroundColor: hex(C.warmCream) }},
        fields: 'userEnteredFormat.backgroundColor',
      }});
    }
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'ROWS', startIndex: 3, endIndex: 3+ITINERARY.length },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});

  // Time format cols B, C (indices 1, 2)
  reqs.push({ repeatCell: {
    range: gridRange(ITN,3,3+ITINERARY.length,1,3),
    cell: { userEnteredFormat: { numberFormat: { type: 'TIME', pattern: 'h:mm AM/PM' }, horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat(numberFormat,horizontalAlignment)',
  }});
  // Center col D (duration), A (day), J (time check)
  reqs.push({ repeatCell: {
    range: gridRange(ITN,3,3+ITINERARY.length,0,1),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(ITN,3,3+ITINERARY.length,3,4),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(ITN,3,3+ITINERARY.length,9,10),
    cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
    fields: 'userEnteredFormat.horizontalAlignment',
  }});

  // Itinerary data values
  for (let i = 0; i < ITINERARY.length; i++) {
    const [day, start, end, activity, type, ledBy, location, supplies, notes] = ITINERARY[i];
    const row = 4 + i;
    // Duration formula: (C-B)*1440
    const durFormula = `=IFERROR(ROUND((C${row}-B${row})*1440,0),"")`;
    // Time check: overlap within same day
    const checkFormula = i === 0 ? '' : `=IFERROR(IF(B${row}="","",IF(AND(A${row}=A${row-1},B${row}<C${row-1}),"⚠️ OVERLAP","")),"")`;
    vals.push({ range: `${TAB}!A${row}`, values: [[day, start, end, durFormula, activity, type, ledBy, location, supplies, checkFormula, notes]] });
  }

  // Borders for itinerary
  reqs.push({ updateBorders: {
    range: gridRange(ITN,2,3+ITINERARY.length,0,ITNCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  // ── ACTIVITIES SECTION ──
  // Section header at row 3+ITINERARY.length (0-indexed) = row 23 (0-indexed)
  const actHeaderRi = 3 + ITINERARY.length; // = 23
  const actHeaderGs = actHeaderRi + 1;       // = 24

  reqs.push({ mergeCells: { range: gridRange(ITN,actHeaderRi,actHeaderRi+1,0,ACTCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ITN,actHeaderRi,actHeaderRi+1,0,ACTCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.midGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'ROWS', startIndex: actHeaderRi, endIndex: actHeaderRi+1 },
    properties: { pixelSize: 30 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A${actHeaderGs}`, values: [['  🎯  ACTIVITIES & GAMES PLANNER']] });

  // Activities col headers
  const actHdrs = ['ACTIVITY NAME','TYPE','DESCRIPTION','AGES','DURATION (MINS)','SUPPLIES','STATUS','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(ITN,actHeaderRi+1,actHeaderRi+2,0,ACTCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.forestGreen),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'ROWS', startIndex: actHeaderRi+1, endIndex: actHeaderRi+2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!A${actHeaderGs+1}`, values: [actHdrs] });

  // Activities data rows
  const actDataStartRi = actHeaderRi + 2;
  for (let i = 0; i < ACTIVITIES.length; i++) {
    const ri  = actDataStartRi + i;
    const row = ri + 1;
    const bg  = i % 2 === 0 ? C.white : C.warmCream;
    reqs.push({ repeatCell: {
      range: gridRange(ITN, ri, ri+1, 0, ACTCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.bodyText), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
    const [actName, type, desc, ages, duration, supplies, status] = ACTIVITIES[i];
    vals.push({ range: `${TAB}!A${row}`, values: [[actName, type, desc, ages, duration, supplies, status, '']] });
  }
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'ROWS', startIndex: actDataStartRi, endIndex: actDataStartRi+ACTIVITIES.length },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});

  // Description col C wrap text
  reqs.push({ repeatCell: {
    range: gridRange(ITN,actDataStartRi,actDataStartRi+ACTIVITIES.length,2,3),
    cell: { userEnteredFormat: { wrapStrategy: 'WRAP' }},
    fields: 'userEnteredFormat.wrapStrategy',
  }});

  // Borders for activities
  reqs.push({ updateBorders: {
    range: gridRange(ITN,actHeaderRi+1,actDataStartRi+ACTIVITIES.length,0,ACTCOLS),
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'itinerary-format');
  await valuesBatchUpdate(id, vals, 'itinerary-values');
  console.log('Event Itinerary & Activities complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
