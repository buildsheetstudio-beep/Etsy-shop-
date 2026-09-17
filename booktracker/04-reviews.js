'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Book Review & Notes'];
const S = "'Book Review & Notes'";
const LIB = "'Master Book Library'";

// Cols: A=Book ID  B=Title (VLOOKUP)  C=Author (VLOOKUP)  D=Genre (VLOOKUP)
//       E=Date Finished (VLOOKUP)  F=Rating (VLOOKUP)
//       G=Overall Review  H=Plot (1-5)  I=Characters (1-5)  J=Writing Style (1-5)
//       K=Pacing (1-5)  L=Avg Detail Rating (formula)  M=Favorite Quote
//       N=Key Themes  O=Would Recommend (Yes/No)  P=Recommend To
//       Q=Contains Spoilers (Yes/No)  R=Review Date  S=Extended Notes

// [BookID, Review, PlotRating, CharRating, WritingRating, PacingRating, FavoriteQuote, Themes, Recommend, RecommendTo, Spoilers, ReviewDate, Notes]
const REVIEWS = [
  ['BOOK-00001','A tour de force of world-building. Vasquez creates a magic system rooted in grief that feels utterly original. The political intrigue rivals anything in classic fantasy.',5,5,5,4,'"The throne is not made of iron — it is made of every choice you refused to make."','Power, sacrifice, grief, destiny','Yes','Fantasy fans, lovers of epic world-building','No','2023-01-28','Cannot wait to see how the trilogy ends.'],
  ['BOOK-00003','The trilogy\'s middle chapter somehow outdoes the opener. Character arcs land with real weight, and the magic system\'s expansion is logical and thrilling.',5,4,5,5,'"She did not need a crown to rule — she needed them to fear her love of justice."','Legacy, loyalty, power','Yes','Epic fantasy readers','No','2023-03-24','Ember Crown is my favourite of the three so far.'],
  ['BOOK-00005','Sharma\'s debut is a revelation. The portal fantasy concept is deceptively simple, but she uses it to explore belonging and identity with quiet brilliance.',5,5,5,5,'"The compass did not point north. It pointed home — and she was only now learning what that meant."','Belonging, identity, wonder','Yes','All readers — a near-universal book','No','2024-01-25','Best book I have read in recent memory.'],
  ['BOOK-00011','Moore perfects the cozy mystery formula: a protagonist you want to befriend, a village that feels lived-in, and a mystery you\'ll genuinely not solve until the final chapter.',5,5,4,5,'"A letter unsent is a conversation forever postponed."','Community, secrets, justice','Yes','Cozy mystery lovers, fans of gentle crime fiction','No','2023-07-18','Starting the series from the beginning again.'],
  ['BOOK-00015','Verlaine writes crime fiction that reads like literary fiction. The murder mystery is almost beside the point — it\'s really a meditation on art, obsession, and loss.',5,4,5,3,'"Every masterpiece contains the artist\'s darkest hour, pressed between the brushstrokes."','Art, obsession, identity, loss','Yes','Literary-minded mystery readers','No','2024-02-27','A genuine masterclass in atmosphere.'],
  ['BOOK-00021','Ito\'s space opera combines the scope of classic SF with the intimacy of a character study. The quantum mechanics are hand-wavy in the best way — a prop for a deeply human story.',5,5,5,4,'"We do not drift in space — we drift in time, and every star we see is already the past."','Identity, time, sacrifice','Yes','All SF readers, even genre newcomers','No','2023-12-22','Starting book 2 immediately after finishing this.'],
  ['BOOK-00025','Tanaka\'s magnum opus. No other SF novel I have read comes close to the emotional devastation of its final act. Prepare to be changed.',5,5,5,4,'"The wormhole does not care what you lost to reach it."','Colonialism, identity, sacrifice, love','Yes','Serious SF readers — not a light read','No','2025-01-24','Cried for a day after finishing this.'],
  ['BOOK-00031','Mercer\'s prose is among the finest I have encountered. Every sentence earns its place. A love story, a historical novel, and a meditation on memory that transcends all three.',5,5,5,4,'"She had mapped the world for him. She never thought to map her own heart."','Love, memory, cartography, loss','Yes','Literary fiction readers, lovers of beautiful prose','No','2024-03-15','Immediately put on my permanent favorites shelf.'],
  ['BOOK-00033','Laurent\'s epistolary novel is a masterpiece of restraint. What is left unsaid haunts every letter. I read the final three in one sitting, unable to stop.',5,5,5,5,'"The sea does not answer letters — but I keep writing anyway."','Longing, distance, love, grief','Yes','Lovers of quiet, literary fiction','No','2024-10-16','Will be rereading this every few years.'],
  ['BOOK-00038','Romano\'s novel announces a major talent. The third daughter is a voice I will remember long after the final page — fierce, wounded, and ultimately triumphant.',5,5,5,4,'"To be the third daughter is to be the one who survives — if you are clever enough to choose it."','Family, ambition, gender, resilience','Yes','Fans of family sagas and literary fiction','No','2025-08-20','One of the best books of this year for me.'],
  ['BOOK-00041','Nassar resurrects the medieval Silk Road in breathtaking detail. A merchant\'s journey becomes a pilgrimage through cultures, faiths, and the nature of trust.',5,5,4,5,'"Trade is not the exchange of goods — it is the exchange of worlds."','Trade, faith, identity, trust','Yes','Historical fiction fans, readers who love travel narratives','No','2023-06-01','The kind of book that makes you want to study history.'],
  ['BOOK-00042','Petit\'s WWII resistance novel is the most emotionally devastating historical fiction I have read. The characters feel utterly real — their losses, personal.',5,5,5,4,'"Resistance is not loud. It is the quiet refusal to forget who you were before they came."','Resistance, war, identity, courage','Yes','All readers — an important book','No','2023-10-10','Recommended to five people already.'],
  ['BOOK-00046','Kowalczyk\'s Warsaw Letters is both heart-wrenching and uplifting. The epistolary structure creates intimacy that a third-person narrative could never achieve.',5,5,5,4,'"Every letter is a small act of survival — proof that we were here and that we loved."','Survival, love, war, memory','Yes','WWII fiction readers, history enthusiasts','No','2025-03-10','Paired beautifully with The Warsaw Uprising documentary.'],
  ['BOOK-00048','Yamamoto\'s portrait of the Meiji empress is a triumph of historical imagination. Imperial intrigue has never been this psychologically rich.',5,5,5,5,'"She wore the empire\'s expectations like a second skin — suffocating but beautiful."','Power, duty, feminism, history','Yes','Historical fiction fans, lovers of Japanese history','No','2025-09-07','My favourite historical fiction of 2025.'],
  ['BOOK-00051','Fontaine captures the soul of a Parisian bookshop with such warmth that I immediately went looking for the address (fictional, alas). A perfect cozy romance.',4,4,5,4,'"Every book is a love letter to its ideal reader. This shop, she thought, was written for me."','Books, love, Paris, belonging','Yes','Romance readers, bibliophiles, Francophiles','No','2023-08-14','Would read again on a rainy afternoon.'],
  ['BOOK-00057','Hughes writes second-chance romance with a lightness that disguises its emotional depth. The final chapters are genuinely moving.',5,4,4,5,'"Another sunrise means another chance. She had stopped believing that — until him."','Second chances, grief, hope, love','Yes','Romance fans who want emotional depth','No','2024-12-02','A book I will think about for a long time.'],
  ['BOOK-00061','Stone\'s debut horror novel announces a terrifying new voice. The forest feels like a living, malevolent character. Read this one with the lights on.',5,4,5,4,'"The dark between the trees is not empty — it is full of everything you tried to forget."','Fear, isolation, trauma, survival','Yes','Horror fans — strong content warnings apply','No','2023-10-30','My favourite horror read of the year.'],
  ['BOOK-00065','Marsh resurrects the rural-horror tradition with astonishing craft. The slow build of dread is perfectly calibrated — you feel the wrongness before you see it.',5,4,5,4,'"Something was waking up — and it had been asleep long enough to be hungry."','Supernatural evil, community, dread','Yes','Horror fans who appreciate atmospheric slow burns','No','2024-11-03','Perfect autumn reading.'],
  ['BOOK-00067','Thornton\'s gothic horror is the kind of book you can reread every October and find something new each time. A modern classic of the genre.',5,5,5,4,'"The house did not haunt her. She came to understand, eventually, that she haunted the house."','Grief, haunting, identity, place','Yes','Gothic horror fans — an essential text','No','2025-10-06','Annual reread. Gets better every time.'],
  ['BOOK-00071','Professor Liu makes a compelling academic case for deep, slow reading in an age of distraction. Changed how I approach every book I pick up.',5,4,4,5,'"To read deeply is to be in conversation with the dead — and to let them change your mind."','Reading, attention, cognition, culture','Yes','Readers who want to read better, book clubs','No','2023-12-05','Immediately changed my reading habits.'],
  ['BOOK-00074','Park synthesizes neuroscience and literary criticism with remarkable clarity. Explains why stories are so fundamental to human cognition.',5,4,5,4,'"The brain does not distinguish between reading about an experience and having it — which is why stories matter."','Reading, neuroscience, empathy','Yes','Anyone curious about the science of storytelling','No','2024-12-18','A book that every book lover should read.'],
  ['BOOK-00077','Callahan makes the most elegant case I have read for why fiction matters. Quietly revolutionary, even if the arguments have been made before.',5,4,5,5,'"Fiction is not an escape from reality. It is our most reliable map of it."','Literature, empathy, imagination','Yes','Literary readers, English teachers, book skeptics','No','2025-02-10','Short enough to finish in an afternoon. Dense enough to think about for months.'],
  ['BOOK-00081','Rivera\'s YA debut is inventive, emotionally resonant, and introduces a protagonist I fell in love with by page three. The map magic is inspired.',5,5,5,4,'"Not all maps show you where to go. Some show you who you are."','Identity, courage, friendship, maps','Yes','YA readers and adults who love coming-of-age stories','No','2023-05-12','My favourite YA debut in years.'],
  ['BOOK-00084','Chen\'s magic system is the cleverest I\'ve encountered in YA fantasy. The gift system creates genuine moral complexity rather than simple power fantasy.',5,5,4,5,'"A gift is only a gift if you choose it. Otherwise, it is a burden with a bow."','Gifts, choice, power, responsibility','Yes','YA fantasy readers, fans of complex magic systems','No','2025-08-05','Will be watching for Chen\'s next book eagerly.'],
  ['BOOK-00087','Donaghue\'s dystopian YA takes the genre\'s conventions and dismantles them intelligently. The final act rejects easy heroics in favour of something harder and truer.',5,5,5,4,'"The last choice is never the one they tell you about in the training. It\'s the one you make alone."','Courage, sacrifice, truth, dystopia','Yes','YA readers ready for something challenging','No','2025-12-22','A stunning achievement in YA fiction.'],
  ['BOOK-00091','Lively\'s memoir reads like a love letter to libraries and to the readers who need them. Warm, wise, and unexpectedly moving.',5,5,5,4,'"Every book on the shelf is a reader who once needed exactly this — and found it."','Libraries, community, reading, memory','Yes','Librarians, book lovers, readers who love behind-the-scenes stories','No','2023-11-20','Immediately recommended to a librarian friend.'],
  ['BOOK-00094','Oswald\'s literary memoir is unlike any I have read — he doesn\'t just cite books, he inhabits them, shows how they shaped his thinking, his grief, his love.',5,5,5,5,'"I did not choose the books that made me. They chose me — and I was lucky enough to open them."','Reading, memory, loss, identity','Yes','All readers — an extraordinary memoir','No','2025-10-02','One of my favourite books of any year.'],
  ['BOOK-00098','Watercolor World is the most beautiful graphic novel I have encountered — every panel is a work of art. Lin\'s wordless pages communicate more than most novels.',5,4,5,5,'"[No words — just a double-page spread of autumn light through an open window]"','Life, beauty, stillness, art','Yes','Graphic novel fans, artists, anyone who appreciates beauty','No','2024-02-07','One of the most beautiful objects I own.'],
  ['BOOK-00103','Kim\'s techniques transformed my relationship to reading. The evidence-based approach to retention actually works, and the book itself models its own advice.',4,4,5,4,'"Retention is not about reading faster. It is about asking better questions."','Learning, memory, reading','Yes','Students, avid readers, anyone who reads for learning','No','2025-04-08','Changed how I annotate and review books.'],
  ['BOOK-00109','Hassan\'s novel about diaspora identity is quiet but devastating. The prose is restrained to the point of austerity, but every sentence earns its place.',4,5,5,3,'"To live between two cultures is to be at home in neither — and in both — simultaneously."','Diaspora, identity, family, belonging','Yes','Literary fiction readers, diaspora readers','No','2026-06-11','Discovered through a recommendation from my reading group.'],
  ['BOOK-00112','Espinoza\'s magical realism is the most assured debut in the genre since Allende. The painted house becomes a character as vivid as any of the humans in it.',5,5,5,4,'"Every house remembers the hands that built it. This one refused to forget."','Memory, place, magic, family','Yes','Magical realism fans, Latin American fiction readers','No','2026-08-05','Found this at a local indie bookshop — a wonderful discovery.'],
];

(async () => {
  const fmt  = [];
  const vals = [];

  // Background wash
  fmt.push({ repeatCell: { range: gridRange(SID,0,600,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Georgia', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title
  vals.push({ range: `${S}!A1`, values: [['✍️ Book Review & Notes']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' }});

  // Subtitle
  vals.push({ range: `${S}!A2`, values: [['Capture your thoughts on every book — ratings, quotes, themes, and full reviews. Use Book ID to link back to your library.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.goldTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Stat bar row 3
  vals.push({ range: `${S}!A3`, values: [[
    '✍️ Total Reviews',
    `=COUNTA(${S}!A6:A506)`,
    '⭐ Reviews Rated 5',
    `=SUMPRODUCT((${S}!F6:F506=5)*ISNUMBER(${S}!F6:F506))`,
    '📖 Avg Plot Rating',
    `=IFERROR(ROUND(AVERAGE(${S}!H6:H506),1),"—")`,
    '🎭 Avg Characters',
    `=IFERROR(ROUND(AVERAGE(${S}!I6:I506),1),"—")`,
    '✍ Avg Writing',
    `=IFERROR(ROUND(AVERAGE(${S}!J6:J506),1),"—")`,
  ]] });
  for (let pair = 0; pair < 5; pair++) {
    const c1 = pair * 2;
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c1+1), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c1+1), cell: { userEnteredFormat: {
      backgroundColor: hex(pair % 2 === 0 ? C.primary : C.secondary),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1+1,c1+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1+1,c1+2), cell: { userEnteredFormat: {
      backgroundColor: hex(pair % 2 === 0 ? C.wineTint : C.goldTint),
      textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 38 }, fields: 'pixelSize' }});

  // Empty row 4
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Header row 5
  const HDR = ['Book ID','Title (Auto)','Author (Auto)','Genre (Auto)','Date Finished (Auto)','Rating (Auto)','Overall Review','Plot (1-5)','Characters (1-5)','Writing Style (1-5)','Pacing (1-5)','Avg Detail Rating','Favorite Quote','Key Themes','Would Recommend?','Recommend To','Contains Spoilers?','Review Date','Extended Notes'];
  vals.push({ range: `${S}!A5`, values: [HDR] });
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // VLOOKUP formula columns for B-F (auto-populated from Library)
  // Col B (1): Title — VLOOKUP(A6, MBL!A:B, 2, FALSE)
  // Col C (2): Author — VLOOKUP(A6, MBL!A:C, 3, FALSE)
  // Col D (3): Genre — VLOOKUP(A6, MBL!A:D, 4, FALSE)
  // Col E (4): Date Finished — VLOOKUP(A6, MBL!A:O, 15, FALSE)
  // Col F (5): Rating — VLOOKUP(A6, MBL!A:P, 16, FALSE)
  // Col L (11): Avg Detail Rating = IFERROR(ROUND(AVERAGE(H6:K6),1),"")
  [[1,'=IFERROR(VLOOKUP(A6,\'Master Book Library\'!$A:$B,2,FALSE),"")'],
   [2,'=IFERROR(VLOOKUP(A6,\'Master Book Library\'!$A:$C,3,FALSE),"")'],
   [3,'=IFERROR(VLOOKUP(A6,\'Master Book Library\'!$A:$D,4,FALSE),"")'],
   [4,'=IFERROR(VLOOKUP(A6,\'Master Book Library\'!$A:$O,15,FALSE),"")'],
   [5,'=IFERROR(VLOOKUP(A6,\'Master Book Library\'!$A:$P,16,FALSE),"")'],
  ].forEach(([ci, formula]) => {
    fmt.push({ repeatCell: {
      range: gridRange(SID,5,506,ci,ci+1),
      cell: { userEnteredValue: { formulaValue: formula }, userEnteredFormat: {
        backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial' },
      }},
      fields: 'userEnteredValue,userEnteredFormat',
    }});
  });
  fmt.push({ repeatCell: {
    range: gridRange(SID,5,506,11,12),
    cell: { userEnteredValue: { formulaValue: '=IFERROR(ROUND(AVERAGE(H6:K6),1),"")' }, userEnteredFormat: {
      backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER',
      numberFormat: { type: 'NUMBER', pattern: '0.0' },
    }},
    fields: 'userEnteredValue,userEnteredFormat',
  }});

  // Alternating rows + input highlights for data area
  for (let i = 0; i < 250; i++) {
    const bg = i % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,5+i,6+i,0,20), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  // Input cols: A(0), G(6)-K(10), M(12)-R(17), S(18)
  [0,6,7,8,9,10,12,13,14,15,16,17,18].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID,5,506,ci,ci+1), cell: { userEnteredFormat: {
      backgroundColor: hex(C.input),
    }}, fields: 'userEnteredFormat.backgroundColor' }});
  });
  // Rating detail cols H-K center
  [7,8,9,10,11].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID,5,506,ci,ci+1), cell: { userEnteredFormat: {
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // Column widths
  const widths = [90,180,130,100,90,65,280,60,65,65,55,60,220,150,80,120,60,90,250];
  widths.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 506 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Freeze
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' }});

  await batchUpdate(id, fmt, '04-reviews format');

  // Sample reviews
  REVIEWS.forEach((rev, i) => {
    const row = 6 + i;
    // [BookID, Review, Plot, Char, Writing, Pacing, Quote, Themes, Recommend, RecommendTo, Spoilers, ReviewDate, Notes]
    vals.push({ range: `${S}!A${row}`, values: [[rev[0]]] });
    vals.push({ range: `${S}!G${row}:K${row}`, values: [[rev[1], rev[2], rev[3], rev[4], rev[5]]] });
    vals.push({ range: `${S}!M${row}:S${row}`, values: [[rev[6], rev[7], rev[8], rev[9], rev[10], rev[11], rev[12]]] });
  });

  await valuesBatchUpdate(id, vals, '04-reviews values');
  console.log(`✅  Book Review & Notes done — ${REVIEWS.length} reviews loaded.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
