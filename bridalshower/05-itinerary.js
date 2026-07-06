'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const ITN = sheetMap['🗓️ Party Itinerary & Games'];

const SEGMENTS = [
  ['11:00','11:30','Welcome & Arrivals','Welcome','Sophie Clarke','Floral arch, welcome board, prosecco on arrival'],
  ['11:30','12:00','Icebreaker — How Well Do You Know the Bride?','Icebreaker','Emma Grant','Quiz sheets, pens'],
  ['12:00','12:15','Bridesmaid Introductions & Thank Yous','Welcome','Charlotte Harrison','Microphone'],
  ['12:15','13:00','Afternoon Tea — First Seating','Food & Drinks','Caterer','Tiered stands, champagne flutes'],
  ['13:00','13:30','Bridal Shower Bingo','Classic','Lily Thompson','Bingo cards, dabbers, prizes'],
  ['13:30','14:00','Wishes for the Bride Activity','Keepsake','Emma Grant','Wish cards, pen cups, display jar'],
  ['14:00','14:30','Afternoon Tea — Second Seating & Cake Cutting','Food & Drinks','Caterer','Cake stand, cake knife'],
  ['14:30','15:15','Group Photo & Garden Walk','Photo','Sophie Clarke','Photographer, bouquet props'],
  ['15:15','15:45','Wedding Trivia: How Well Do They Know Each Other?','Trivia','Lily Thompson','Trivia cards, scoreboard'],
  ['15:45','16:15','Gift Opening','Gift Opening','Charlotte Harrison','Gift table, card box'],
  ['16:15','16:45','Speeches & Toasts','Speeches','Sophie Clarke','Champagne, microphone'],
  ['16:45','17:00','Prize Draw & Favours','Prize Game','Emma Grant','Prize hamper, favour bags'],
  ['17:00','17:20','Farewell & Group Hug','Farewell','All','Polaroid camera, disposable cameras'],
  ['17:20','17:30','Venue Pack-Down','Setup','Bridesmaids','Boxes, bags'],
];

const GAMES = [
  ['Bridal Shower Bingo','Classic','Lily Thompson',30,'Bingo cards, daubers, small prizes','Classic crowd-pleaser; prizes for first full card'],
  ['How Well Do You Know the Bride?','Trivia','Emma Grant',20,'Quiz sheets, pens','20 questions about the bride; highest score wins'],
  ['He Said / She Said','Trivia','Sophie Clarke',15,'Cards with quotes','Guess who said what — hilarious results'],
  ['Bridal Word Scramble','Classic','Emma Grant',10,'Printed sheets, pens','Unscramble wedding words; fastest wins'],
  ['Ring Hunt','Classic','Lily Thompson',20,'Hidden rings (confetti rings)','Hidden around venue; most rings wins prize'],
  ['Wishes for the Bride','Keepsake','Sophie Clarke',20,'Wish cards, display jar','Guests write hopes & wishes — keepsake for bride'],
  ['Wedding Trivia: Couple Edition','Trivia','Lily Thompson',20,'Trivia cards','Fun facts about the couple'],
  ['Advice for the Bride','Keepsake','Emma Grant',15,'Advice cards','Written advice — funny & heartfelt mix'],
  ['Pin the Veil on the Bride','Creative','Sophie Clarke',15,'Illustrated poster, veils, blindfold','Blindfolded party game version of classic'],
  ['Prize Draw','Prize Game','Emma Grant',10,'Ticket roll, prize hamper','Ticket per attended game; grand prize draw'],
];

(async () => {
  const reqs = [];
  const vals = [];
  const TAB = "'🗓️ Party Itinerary & Games'";

  // Column widths
  const colW = [80,80,90,180,130,130,220,130];
  colW.forEach((w,i) => reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'COLUMNS', startIndex: i, endIndex: i+1 },
    properties: { pixelSize: w }, fields: 'pixelSize',
  }}));

  // Row 0: Title
  reqs.push({ mergeCells: { range: gridRange(ITN,0,1,1,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ITN,0,1,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 14 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 40 }, fields: 'pixelSize',
  }});
  vals.push({ range: `${TAB}!B1`, values: [['🗓️ Party Itinerary & Games']] });

  // Row 1: Party info strip
  reqs.push({ repeatCell: {
    range: gridRange(ITN,1,2,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.parchment),
      textFormat: { foregroundColor: hex(C.muted), fontSize: 9 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A2`, values: [['Date:','','Venue:','','Start:','11:00 AM','End:','5:30 PM']] });

  // Row 2: section header — Itinerary
  reqs.push({ mergeCells: { range: gridRange(ITN,2,3,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ITN,2,3,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A3`, values: [['PARTY ITINERARY']] });

  // Row 3: col headers
  const itnHeaders = ['START','END','DURATION','ACTIVITY','TYPE','LED BY','SUPPLIES / NOTES','TIME CHECK'];
  reqs.push({ repeatCell: {
    range: gridRange(ITN,3,4,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A4`, values: [itnHeaders] });

  // Segments rows 5-18 (0-indexed 4-17)
  for (let i = 0; i < SEGMENTS.length; i++) {
    const [start, end, activity, type, ledBy, notes] = SEGMENTS[i];
    const row = 5 + i;
    const ri  = 4 + i;
    const bg  = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(ITN, ri, ri+1, 0, 8),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    const overlapFormula = i === 0 ? '—'
      : `=IFERROR(IF(TIMEVALUE(A${row})<TIMEVALUE(B${row-1}),"⚠ OVERLAP","✓"),"")`;
    vals.push({ range: `${TAB}!A${row}`, values: [[
      start, end,
      `=IFERROR(TEXT(TIMEVALUE(B${row})-TIMEVALUE(A${row}),"h:mm"),"")`,
      activity, type, ledBy, notes,
      overlapFormula,
    ]]});
  }

  // Divider row (0-indexed 18)
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: ITN, dimension: 'ROWS', startIndex: 18, endIndex: 19 },
    properties: { pixelSize: 8 }, fields: 'pixelSize',
  }});
  reqs.push({ repeatCell: {
    range: gridRange(ITN,18,19,0,8),
    cell: { userEnteredFormat: { backgroundColor: hex(C.gold) }},
    fields: 'userEnteredFormat(backgroundColor)',
  }});

  // Row 20: Games section header (0-indexed 19)
  reqs.push({ mergeCells: { range: gridRange(ITN,19,20,0,8), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(ITN,19,20,0,8),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.medDustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat)',
  }});
  vals.push({ range: `${TAB}!A20`, values: [['GAMES PLANNER']] });

  // Row 21: game col headers (0-indexed 20)
  const gameHeaders = ['GAME NAME','TYPE','HOSTED BY','DURATION (MIN)','SUPPLIES','NOTES'];
  reqs.push({ repeatCell: {
    range: gridRange(ITN,20,21,0,6),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.dustyRose),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});
  vals.push({ range: `${TAB}!A21`, values: [gameHeaders] });

  // Games rows 22-31 (0-indexed 21-30)
  for (let i = 0; i < GAMES.length; i++) {
    const row = 22 + i;
    const ri  = 21 + i;
    const bg  = i % 2 === 0 ? C.ivory : C.parchment;
    reqs.push({ repeatCell: {
      range: gridRange(ITN, ri, ri+1, 0, 6),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    }});
    vals.push({ range: `${TAB}!A${row}`, values: [GAMES[i]] });
  }

  await batchUpdate(id, reqs, 'itinerary-format');
  await valuesBatchUpdate(id, vals, 'itinerary-values');
  console.log('Party Itinerary & Games complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
