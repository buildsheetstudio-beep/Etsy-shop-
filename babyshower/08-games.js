'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const GA = sheetMap['🎮 Games & Activities'];

// Layout:
// Row 1: Title (A1:I1)
// Row 2: Col headers
// Rows 3-14: 12 games
// Cols: A=Game Name, B=Type, C=Description, D=Supplies Needed, E=Duration, F=# Players, G=Prize, H=Host, I=Status

const NCOLS = 9;

// [name, type, description, supplies, duration, players, prize, host, status]
const GAMES = [
  [
    'Baby Bingo',
    'Classic',
    'Guests fill in bingo cards with predicted gifts as Sarah opens presents. First to get bingo wins!',
    'Pre-printed bingo cards x20, daubers or pens, call sheet',
    '20 min',
    'All',
    'Luxe body lotion gift set',
    'Emma Clarke',
    'Planned',
  ],
  [
    'Guess the Baby Food',
    'Interactive',
    'Unlabelled jars of baby food are passed around. Guests taste and identify the flavour. Highest score wins.',
    '10 unlabelled baby food jars, spoons x20, score sheets',
    '20 min',
    'All',
    'Artisan chocolate box',
    'Mia Patel',
    'Planned',
  ],
  [
    'Baby Trivia',
    'Trivia',
    'Multiple-choice trivia questions about pregnancy, babies, and Sarah. Teams or individuals compete.',
    'Printed trivia sheets x20, pens, answer key',
    '20 min',
    'All',
    'Pamper hamper',
    'Zoe Roberts',
    'Planned',
  ],
  [
    'Baby Price is Right',
    'Prize Game',
    'Guests guess the retail price of baby items. Closest guess without going over wins each round.',
    '8 baby items with tags removed, price cards, whiteboard for guesses',
    '20 min',
    'All',
    'Candle & bath set',
    'Amy Johnson',
    'Planned',
  ],
  [
    'Nappy Poetry',
    'Creative',
    'Each guest writes a funny poem or rhyme on the inside of a nappy for Sarah to read during night feeds.',
    'Nappies x20, permanent markers, reading basket',
    '15 min',
    'All',
    'No prize — keepsake activity',
    'Lily Thompson',
    'Planned',
  ],
  [
    'Baby Measurements',
    'Interactive',
    'Guests cut a piece of string to guess the circumference of Sarah\'s belly. Closest to actual wins.',
    'Ball of string, scissors, measuring tape',
    '10 min',
    'All',
    'Gourmet biscuit tin',
    'Emma Clarke',
    'Planned',
  ],
  [
    'Name that Nursery Rhyme',
    'Trivia',
    'Guests hear the first line of a nursery rhyme and name it. Can be played in teams.',
    'Printed rhyme cards, pens, score sheets',
    '15 min',
    'All',
    'Pretty notebook & pen set',
    'Lucy Chen',
    'Planned',
  ],
  [
    'Mum\'s Words of Wisdom',
    'Keepsake',
    'Each guest writes their best piece of parenting advice on a card for Sarah to keep.',
    'Decorative advice cards x20, pens, keepsake box',
    '10 min',
    'All',
    'No prize — keepsake activity',
    'Margaret Mitchell',
    'Planned',
  ],
  [
    'Baby Draw — No Looking!',
    'Creative',
    'Guests balance a paper plate on their head and draw a baby without looking. Funniest drawing wins.',
    'Paper plates x20, pens or markers',
    '10 min',
    'All',
    'Funny mug + tea set',
    'Claire Mitchell',
    'Planned',
  ],
  [
    'Two Truths & a Baby Lie',
    'Interactive',
    'Each guest shares two true facts and one lie about their baby or childhood experiences. Group guesses the lie.',
    'No supplies needed — conversational',
    '15 min',
    'All',
    'Wine + chocolates',
    'Emma Clarke',
    'Planned',
  ],
  [
    'Baby Photo Match',
    'Classic',
    'Guests match baby photos of attendees (collected in advance) to the correct adult. Highest score wins.',
    'Printed baby photos of guests, numbered sheets, answer sheets',
    '15 min',
    'All',
    'Spa day voucher',
    'Lily Thompson',
    'Planned',
  ],
  [
    'Baby Shower Scavenger Hunt',
    'Team',
    'Teams race to find or photograph items in the venue matching clues. Fun icebreaker if guests arrive at staggered times.',
    'Printed clue cards x4 sets, small prizes per team member',
    '15 min',
    'Teams of 4-5',
    'Team prizes — bath bombs x4',
    'Emma Clarke',
    'Planned',
  ],
];

(async () => {
  const reqs = [];

  // Column widths: A=200, B=100, C=320, D=280, E=80, F=80, G=180, H=120, I=90
  [200,100,320,280,80,80,180,120,90].forEach((w, i) => {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: GA, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
      properties: { pixelSize: w }, fields: 'pixelSize',
    }});
  });

  // Row heights
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GA, dimension: 'ROWS', startIndex: 0, endIndex: 1 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GA, dimension: 'ROWS', startIndex: 1, endIndex: 2 },
    properties: { pixelSize: 28 }, fields: 'pixelSize',
  }});
  reqs.push({ updateDimensionProperties: {
    range: { sheetId: GA, dimension: 'ROWS', startIndex: 2, endIndex: 14 },
    properties: { pixelSize: 52 }, fields: 'pixelSize',
  }});

  // Title row
  reqs.push({ mergeCells: { range: gridRange(GA, 0, 1, 0, NCOLS), mergeType: 'MERGE_ALL' }});
  reqs.push({ repeatCell: {
    range: gridRange(GA, 0, 1, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 16 },
      horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
  }});

  // Col headers row 2 (index 1)
  reqs.push({ repeatCell: {
    range: gridRange(GA, 1, 2, 0, NCOLS),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.lilac),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 9 },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
      wrapStrategy: 'WRAP',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)',
  }});

  // Data rows 3-14 (indices 2-13)
  for (let r = 2; r < 14; r++) {
    const bg = (r % 2 === 0) ? C.white : C.petalPink;
    reqs.push({ repeatCell: {
      range: gridRange(GA, r, r + 1, 0, NCOLS),
      cell: { userEnteredFormat: {
        backgroundColor: hex(bg),
        textFormat: { foregroundColor: hex(C.text), fontSize: 9 },
        verticalAlignment: 'MIDDLE',
      }},
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    }});
  }

  // Wrap for description (col C=2) and supplies (col D=3)
  [2, 3].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(GA, 2, 14, c, c + 1),
      cell: { userEnteredFormat: { wrapStrategy: 'WRAP' }},
      fields: 'userEnteredFormat.wrapStrategy',
    }});
  });

  // Center for duration, players, status
  [4, 5, 8].forEach(c => {
    reqs.push({ repeatCell: {
      range: gridRange(GA, 2, 14, c, c + 1),
      cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' }},
      fields: 'userEnteredFormat.horizontalAlignment',
    }});
  });

  // Borders
  reqs.push({ updateBorders: {
    range: gridRange(GA, 1, 14, 0, NCOLS),
    top:    { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    bottom: { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    left:   { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    right:  { style: 'SOLID_MEDIUM', color: hex(C.lilac), width: 2 },
    innerHorizontal: { style: 'SOLID', color: hex(C.border), width: 1 },
    innerVertical:   { style: 'SOLID', color: hex(C.border), width: 1 },
  }});

  await batchUpdate(id, reqs, 'games-format');

  // Values
  const vals = [];
  vals.push({ range: "'🎮 Games & Activities'!A1", values: [['  🎮  Games & Activities']] });
  vals.push({ range: "'🎮 Games & Activities'!A2", values: [['GAME NAME','TYPE','DESCRIPTION','SUPPLIES NEEDED','DURATION','PLAYERS','PRIZE','HOST','STATUS']] });
  vals.push({ range: "'🎮 Games & Activities'!A3", values: GAMES });

  await valuesBatchUpdate(id, vals, 'games-values');
  console.log('Games & Activities complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
