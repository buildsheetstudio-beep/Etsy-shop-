'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Wishlist'];
const S = "'Wishlist'";

// Cols: A=Wishlist ID  B=Title  C=Author  D=Genre  E=Sub-genre  F=Format
//       G=Series Name  H=Series #  I=Year Published  J=Pages (approx)
//       K=Priority  L=Source  M=Reason to Read  N=Recommended By
//       O=Acquired?  P=Date Added  Q=Date Acquired  R=Est. Cost  S=Notes
// [Title, Author, Genre, SubGenre, Format, Series, SeriesNum, Year, Pages, Priority, Source, Reason, RecommendedBy, DateAdded, EstCost, Notes]
const WISHLIST = [
  ['The Dragon\'s Promise','Mei Qian','Fantasy','Epic Fantasy','Hardcover','',null,2024,567,'High','Online Review','Stunning reviews from trusted bloggers','TBD','2026-05-12',24.99,'Available at my local bookshop.'],
  ['The Crimson Laboratory','Arthur Pine','Mystery','Historical Mystery','Paperback','The Pine Mysteries',2,2023,389,'High','Friend Recommendation','Sequel to The Clockmaker\'s Secret — must read','Colleague at work','2026-04-10',14.99,'Part of Wells Mysteries companion series.'],
  ['Saturn\'s Children','Felix Vance','Science Fiction','Hard SF','Paperback','',null,2022,534,'Medium','Best-of List','Hugo Award winner 2022','Best of Year list','2026-03-22',13.99,'Award-winning hard SF — exactly my wheelhouse.'],
  ['The Painter\'s Model','Cecile Blanc','Literary Fiction','Art Historical','Hardcover','',null,2023,412,'High','Bookstore Browse','Set in 19th-century Paris art world','Browsing at indie shop','2026-02-18',26.99,'Beautiful cover art — discovered at an indie bookshop.'],
  ['Blood and Bounty','Rowan Steele','Historical Fiction','Pirate Adventure','Paperback','',null,2024,498,'Medium','Social Media','Pirate historical fiction — unique setting','BookTok recommendation','2026-06-05',15.99,'Going on my reading list for next winter.'],
  ['Love Uncharted','Isabela Costa','Romance','Adventure Romance','eBook','',null,2024,312,'Low','Online Review','Light summer read — looks fun','Goodreads review','2026-07-01',7.99,'Summer reading candidate.'],
  ['The Descent','Cole Hamlin','Horror','Psychological Horror','Paperback','',null,2023,445,'High','Friend Recommendation','Described as the scariest book published this year','My horror-reading friend','2026-10-01',14.99,'Saving for October reading.'],
  ['Atlas of Reading','Dr. Nina Patel','Non-Fiction','Reading Science','Hardcover','',null,2024,289,'High','Best-of List','Best book about reading since Chen Liu\'s work','Year\'s Best Non-Fiction list','2026-04-28',28.99,'Companion to The Art of Deep Reading.'],
  ['Kingdom of Ash and Starlight','Tarla Nieves','Fantasy','YA Fantasy','Hardcover','The Ashlight Trilogy',1,2025,456,'High','Social Media','Phenomenal debut getting massive buzz','BookTok','2026-08-20',22.99,'Pre-ordered — arrives in October.'],
  ['The Tokyo Conspiracy','Hiroki Bando','Mystery','Political Thriller','Paperback','',null,2023,423,'Medium','Online Review','Tokyo setting + geopolitical thriller plot','Crime Fiction review site','2026-05-30',14.99,'Set in my dream destination — cannot wait.'],
  ['Void Walkers','Zara Osei','Science Fiction','Space Fantasy','eBook','The Void Series',1,2025,512,'High','Author Follow','New book from my favourite SF debut author','Following the author online','2026-09-01',9.99,'Instant pre-order when announced.'],
  ['The Second Marriage','Clara Voss','Literary Fiction','Domestic Fiction','Paperback','',null,2023,334,'Low','Book Club','Chosen for upcoming book club meeting','Book club coordinator','2026-09-10',13.99,'Book club pick for October.'],
  ['The Ottoman Jewel','Aysel Demir','Historical Fiction','Ottoman Empire','Hardcover','',null,2024,589,'High','Personal Discovery','Ottoman history — a period I want to explore more','Self-discovered browsing online','2026-06-18',27.99,'The Ottoman setting is completely unique in my library.'],
  ['Midnight Waltz','Sophie Marceau','Romance','Regency Romance','Paperback','',null,2024,356,'Medium','Friend Recommendation','Regency ballroom romance — classic setup','Friend who loves Heyer','2026-07-15',14.99,'Apparently the dancing scenes are extraordinary.'],
  ['The Undertaker\'s Smile','Desmond Wraith','Horror','Gothic Horror','Hardcover','',null,2023,512,'High','Online Review','Exceptional gothic horror debut — multiple award nods','Horror blog I follow','2026-09-05',25.99,'Going straight to October reading pile.'],
  ['Think Like a Reader','Vera Ashford','Non-Fiction','Critical Reading','Paperback','',null,2024,267,'Medium','Best-of List','Practical guide to analytical reading','Reading improvement list','2026-08-01',13.99,'Short but highly practical.'],
  ['The Phoenix Accord','Dion Larkin','Fantasy','Political Fantasy','Paperback','The Accord Trilogy',1,2023,634,'High','Online Review','Intricate political worldbuilding — complex and rewarding','Fantasy review blog','2026-04-15',15.99,'Starting a new series in this genre.'],
  ['Harbor Street Blues','Rita Falk','Mystery','Bavarian Noir','Paperback','Eberhofer Mystery',1,2019,278,'Medium','Friend Recommendation','Funny + dark Bavarian village mysteries','Friend who lived in Germany','2026-05-20',13.99,'Apparently very funny — different from my usual mystery fare.'],
  ['The Martian Gardens','Cassie Elroy','Science Fiction','Biopunk','eBook','',null,2024,445,'High','Best-of List','Mars terraforming + botanical worldbuilding','SF Year\'s Best list','2026-07-22',8.99,'The premise of botanists on Mars is delightful.'],
  ['Paper and Stone','Jamie Reyes','Literary Fiction','Bicultural Fiction','Paperback','',null,2023,389,'Medium','Online Review','Mexican-American family saga across three generations','Literary fiction blog','2026-03-15',14.99,'Themes of heritage and belonging resonate with me.'],
  ['Daughters of the Nile','Amira Ghali','Historical Fiction','Ancient Egypt','Hardcover','',null,2024,623,'High','Personal Discovery','Female pharaohs of ancient Egypt — my period of history','Found at museum bookshop','2026-06-08',26.99,'Found at the Egypt exhibition — must read.'],
  ['The Proposal','Lily Porter','Romance','Workplace Romance','eBook','',null,2024,312,'Low','Social Media','Light workplace romance — good for a quick read','Instagram recommendation','2026-08-05',7.99,'Saving for when I need something easy.'],
  ['Red Hollow','Marcus Dread','Horror','Appalachian Horror','Paperback','',null,2023,467,'High','Friend Recommendation','Regional horror with strong sense of place','Horror fan friend','2026-09-15',14.99,'October reading list.'],
  ['Books as Medicine','Dr. Alma Chen','Non-Fiction','Bibliotherapy','Paperback','',null,2024,234,'High','Personal Discovery','Bibliotherapy — the idea that books can heal','Self-discovered browsing academic section','2026-07-10',14.99,'The concept of reading as therapy deeply interests me.'],
  ['The Fallen Star','Orin Blake','Fantasy','Sky Fantasy','Hardcover','',null,2025,578,'Medium','Online Review','Aerial fantasy world with original magic system','Fantasy magazine review','2026-08-28',24.99,'Unusual sky-based fantasy setting.'],
  ['The Silent Evidence','Harriet Quinn','Mystery','Forensic Mystery','Paperback','Quinn Investigations',1,2024,398,'High','Best-of List','Forensic scientist protagonist — love this sub-genre','Mystery year\'s best','2026-05-08',14.99,'First in a new forensic mystery series.'],
  ['Terra Nova','Riku Mäkinen','Science Fiction','Climate Fiction','eBook','',null,2023,512,'Medium','Online Review','Climate fiction from a Finnish perspective — interesting viewpoint','Climate fiction blog','2026-06-25',8.99,'Cli-fi is an area I want to read more.'],
  ['The Garden Party','Vivienne Leclerc','Literary Fiction','French Literary','Paperback','',null,2024,334,'Low','Friend Recommendation','Novella-length French literary fiction — elegant','French literature fan friend','2026-09-20',12.99,'Short and supposedly perfect.'],
  ['The Bronze Archer','Dominic Fabio','Historical Fiction','Ancient Greece','Hardcover','',null,2023,578,'High','Personal Discovery','Bronze Age Aegean — my favourite historical period','History podcast episode on Bronze Age','2026-04-05',26.99,'Prompted by my obsession with Bronze Age Greece.'],
  ['Always You','Natasha Bloom','Romance','Friends-to-Lovers','Paperback','',null,2024,345,'Medium','Social Media','Friends-to-lovers romance — a favourite trope','Bookstagram post','2026-07-28',13.99,'The classic trope, well-reviewed execution.'],
  ['Last Light','Warren Greer','Horror','Post-Apocalyptic Horror','Paperback','',null,2023,498,'High','Personal Discovery','End-of-world survival horror with literary ambitions','Found in a used bookshop','2026-08-18',12.99,'Found a signed copy at a second-hand bookshop.'],
  ['The Year of Reading Dangerously','Hugo Mills','Non-Fiction','Reading Memoir','eBook','',null,2023,267,'Medium','Best-of List','Year reading books outside comfort zone — relatable project','Reading challenge list','2026-06-30',7.99,'Inspired by my own genre challenge this year.'],
  ['The Shattered Court','Alara Storm','Fantasy','Court Intrigue Fantasy','Paperback','',null,2025,612,'High','Author Follow','New author I discovered through a newsletter — stunning debut','Fantasy newsletter','2026-09-03',15.99,'Just announced — on pre-order.'],
  ['The Ice Garden','Fiona Blackwood','Mystery','Botanical Mystery','Hardcover','',null,2024,367,'Medium','Bookstore Browse','Plant poisonings in a Victorian botanical garden — irresistible','Browsing at bookshop','2026-05-18',24.99,'Cozy mystery meets Victorian botany — perfect.'],
  ['The Stargazer\'s Apprentice','Leo Vega','Young Adult','YA Historical','Paperback','',null,2024,412,'Medium','Friend Recommendation','YA set in Renaissance observatory — sounds wonderful','Friend with great YA taste','2026-08-14',14.99,'For when I want a lighter historical read.'],
  ['The Vanished Hours','Petra Strand','Thriller','Scandinavian Noir','eBook','',null,2024,423,'High','Online Review','Swedish thriller with an unreliable narrator — gripping premise','Nordic Noir blog','2026-07-08',8.99,'Nordic noir is reliably excellent.'],
  ['The Lighthouse Letters','Mira Osei','Literary Fiction','Epistolary','Hardcover','',null,2025,312,'High','Personal Discovery','Follow-up to an author whose work I adore','Author newsletter','2026-09-08',24.99,'Pre-ordered immediately when announced.'],
  ['Sovereign','Kieran Voss','Fantasy','Military Fantasy','Paperback','The Iron Wars',1,2024,689,'Medium','Friend Recommendation','Big military fantasy series — friend is obsessed','Friend who reads only fantasy','2026-06-20',15.99,'For when I want something epic and long.'],
  ['The Glass Ocean','Simone Rowe','Historical Fiction','Maritime History','Hardcover','',null,2023,534,'High','Best-of List','1920s ocean liner historical fiction — sounds cinematic','Historical fiction year\'s best','2026-04-30',26.99,'The maritime setting is underused in historical fiction.'],
  ['Digital Minds','Dr. Priya Nair','Non-Fiction','AI & Society','Paperback','',null,2025,298,'Medium','Online Review','Accessible AI book from a humanities perspective','Technology book review','2026-08-25',14.99,'Want to understand AI through a literary lens.'],
  ['The Coral Throne','Aisha Osei','Fantasy','Ocean Fantasy','Hardcover','',null,2025,523,'High','Social Media','Underwater fantasy with incredible worldbuilding','BookTok','2026-09-01',24.99,'Underwater fantasy — completely untested territory for me.'],
  ['Before the Storm','Harriet Crane','Literary Fiction','Pre-War Fiction','Paperback','',null,2024,378,'Medium','Bookstore Browse','1930s interwar literary fiction — elegiac and precise','Browsing recommended shelf','2026-07-20',14.99,'Perfect autumn reading material.'],
];

(async () => {
  const fmt  = [];
  const vals = [];

  // Background
  fmt.push({ repeatCell: { range: gridRange(SID,0,600,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Georgia', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // Title
  vals.push({ range: `${S}!A1`, values: [['🌟 Wishlist']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' }});

  // Subtitle
  vals.push({ range: `${S}!A2`, values: [['All the books you want to read next — track priority, source, and acquisition status. Mark Acquired? when you get the book.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,20), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.goldTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // Stats row 3
  vals.push({ range: `${S}!A3`, values: [[
    '🌟 Wishlist Size',
    `=COUNTA(${S}!B6:B506)`,
    '🔴 High Priority',
    `=SUMPRODUCT((${S}!K6:K506="High")*ISBLANK(${S}!O6:O506))`,
    '🟡 Medium Priority',
    `=SUMPRODUCT((${S}!K6:K506="Medium")*ISBLANK(${S}!O6:O506))`,
    '✅ Acquired',
    `=SUMPRODUCT((${S}!O6:O506=TRUE)*1)`,
    '💰 Total Wishlist Est. Value',
    `=IFERROR(SUMIF(${S}!O6:O506,FALSE,${S}!R6:R506),"—")`,
  ]] });
  for (let pair = 0; pair < 5; pair++) {
    const c1 = pair * 2;
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c1+1), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c1+1), cell: { userEnteredFormat: {
      backgroundColor: hex(pair % 2 === 0 ? C.secondary : C.primary),
      textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1+1,c1+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1+1,c1+2), cell: { userEnteredFormat: {
      backgroundColor: hex(pair % 2 === 0 ? C.goldTint : C.wineTint),
      textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 38 }, fields: 'pixelSize' }});

  // Empty row 4
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // Wishlist ID formula (col A) from row 6
  fmt.push({ repeatCell: {
    range: gridRange(SID,5,506,0,1),
    cell: { userEnteredValue: { formulaValue: '=IF(B6="","","WL-"&TEXT(ROW()-5,"00000"))' }, userEnteredFormat: {
      backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial', bold: true },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredValue,userEnteredFormat',
  }});

  // Header row 5
  const HDR = ['Wishlist ID','Title','Author','Genre','Sub-genre','Format','Series','Series #','Year Published','Pages (approx)','Priority','Source','Reason to Read','Recommended By','Acquired?','Date Added','Date Acquired','Est. Cost','Notes'];
  vals.push({ range: `${S}!A5`, values: [HDR] });
  fmt.push({ repeatCell: { range: gridRange(SID,4,5,0,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.secondary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 5 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // Alternating rows
  for (let i = 0; i < 300; i++) {
    const bg = i % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,5+i,6+i,0,20), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  // Input cols
  fmt.push({ repeatCell: { range: gridRange(SID,5,506,1,20), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
  }}, fields: 'userEnteredFormat.backgroundColor' }});
  // Cost col R (17): currency format
  fmt.push({ repeatCell: { range: gridRange(SID,5,506,17,18), cell: { userEnteredFormat: {
    numberFormat: { type: 'CURRENCY', pattern: '"$"#,##0.00' },
  }}, fields: 'userEnteredFormat.numberFormat' }});

  // Column widths
  const widths = [85,200,150,110,110,80,130,55,70,70,70,120,200,130,65,85,85,75,200];
  widths.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 5, endIndex: 506 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // Freeze
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 5 } }, fields: 'gridProperties.frozenRowCount' }});

  await batchUpdate(id, fmt, '05-wishlist format');

  // Data: [Title, Author, Genre, SubGenre, Format, Series, SeriesNum, Year, Pages, Priority, Source, Reason, RecommendedBy, DateAdded, EstCost, Notes]
  WISHLIST.forEach((book, i) => {
    const row = 6 + i;
    // cols B(1)-N(13): Title through RecommendedBy = indices 0-12 (13 values)
    // col O(14): Acquired? = FALSE (checkbox — set by validation)
    // col P(15): DateAdded = index 13
    // col Q(16): DateAcquired = '' (empty)
    // col R(17): EstCost = index 14
    // col S(18): Notes = index 15
    const bToN = book.slice(0, 13);
    const pOnward = [book[13], '', book[14], book[15]];
    vals.push({ range: `${S}!B${row}:N${row}`, values: [bToN] });
    vals.push({ range: `${S}!P${row}:S${row}`, values: [pOnward] });
  });

  await valuesBatchUpdate(id, vals, '05-wishlist values');
  console.log(`✅  Wishlist done — ${WISHLIST.length} books loaded.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
