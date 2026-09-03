'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const SID = sheetMap['Master Book Library'];
const S = "'Master Book Library'";

// Cols: B=Title C=Author D=Genre E=SubGenre F=Format G=Series H=SeriesNum
//       I=YearPub J=Language K=Publisher L=Shelf M=Status N=DateStart O=DateEnd
//       P=Rating Q=TotalPages R=PagesRead [S=formula] T=DateAdded U=Fav V=Reread W=TimesReread X=Notes
// 23 values per row (B-X, with '' placeholder at index 17 for col S which gets a formula)
const BOOKS = [
  // Fantasy (10)
  ['The Midnight Throne','Elena Vasquez','Fantasy','Epic Fantasy','Paperback','The Throne Saga',1,2021,'English','Ember Press','Favorites','Finished','2023-01-05','2023-01-28',5,487,487,'','2022-12-20',true,true,2,'Breathtaking world-building. A new all-time favorite.'],
  ['Shadow of the Seventh Moon','Marcus Chen','Fantasy','Dark Fantasy','Hardcover','',null,2019,'English','Orbit Books','Fiction','Finished','2023-02-01','2023-02-19',4,612,612,'','2023-01-15',false,false,0,'Gripping, though the ending felt rushed.'],
  ['The Ember Crown','Lydia Hartwell','Fantasy','Epic Fantasy','Paperback','The Throne Saga',2,2022,'English','Ember Press','Favorites','Finished','2023-03-02','2023-03-24',5,534,534,'','2023-01-28',true,false,0,'Even better than book 1. Cannot wait for the conclusion.'],
  ['The Glass Sorceress','Nina Huang','Fantasy','Magical Realism','eBook','',null,2021,'English','Tor Books','Fiction','Finished','2023-04-10','2023-04-28',4,398,398,'','2023-03-25',false,false,0,'Beautiful prose and a clever magic system.'],
  ['The Silver Compass','Priya Sharma','Fantasy','Portal Fantasy','Paperback','',null,2024,'English','Del Rey','Favorites','Finished','2024-01-08','2024-01-25',5,445,445,'','2023-12-10',true,false,0,'Instant favorite — pure wonder and heart.'],
  ['Echoes of the Ancient Sea','Tobias Rein','Fantasy','Mythic Fantasy','Hardcover','',null,2023,'English','Penguin Fantasy','Fiction','Finished','2023-08-14','2023-09-02',4,602,602,'','2023-07-20',false,false,0,'Gorgeous mythology-inspired storytelling.'],
  ['Realm of the Bone Keeper','Dylan McCabe','Fantasy','Grimdark','Paperback','',null,2022,'English','Solaris','Fiction','DNF',  '2023-05-05','',0,721,180,'','2023-04-30',false,false,0,'Too dark and slow for me. Set aside at 25%.'],
  ['Whisper of the Forgotten Gods','Amara Okafor','Fantasy','Mythic Fantasy','eBook','',null,2023,'English','Angry Robot','Fiction','Finished','2023-06-01','2023-06-22',4,568,568,'','2023-05-28',false,false,0,'Rich West African–inspired mythology.'],
  ['Daughters of the Iron Gate','Sofia Petrov','Fantasy','Epic Fantasy','Hardcover','The Iron Gate Chronicles',1,2020,'English','Tor Books','Fiction','Reading','2026-08-01','',0,789,312,'','2026-07-15',false,false,0,'Recommended by my reading group.'],
  ['Wolf and Winter','Astrid Nord','Fantasy','Norse Fantasy','Paperback','',null,2022,'English','Head of Zeus','Fiction','Paused','2025-11-10','',0,534,200,'','2025-10-25',false,false,0,'Put aside during busy season. Want to finish.'],

  // Mystery / Thriller (10)
  ['The Last Letter','Catherine Moore','Mystery','Cozy Mystery','Paperback','',null,2020,'English','Minotaur','Fiction','Finished','2023-07-03','2023-07-18',5,345,345,'','2023-06-25',true,true,1,'Perfect cozy mystery — could not put it down.'],
  ['Poisoned Waters','Henri Dumont','Mystery','Crime Fiction','Hardcover','',null,2022,'English','Quercus','Fiction','Finished','2023-09-05','2023-09-19',4,387,387,'','2023-08-30',false,false,0,'Atmospheric French setting, clever plot.'],
  ['The Clockmaker\'s Secret','Agatha Wells','Mystery','Historical Mystery','Paperback','The Wells Mysteries',1,2019,'English','HarperCollins','Fiction','Finished','2023-10-01','2023-10-15',5,412,412,'','2023-09-20',true,false,0,'Classic whodunnit with a brilliant twist.'],
  ['The Vanishing Hour','Stella Kowalski','Thriller','Psychological Thriller','eBook','',null,2021,'English','Simon & Schuster','Fiction','Finished','2023-11-03','2023-11-17',4,467,467,'','2023-10-20',false,false,0,'Kept me guessing until the very end.'],
  ['The Curator\'s Body','Margot Verlaine','Mystery','Crime Fiction','Hardcover','',null,2021,'English','Faber & Faber','Favorites','Finished','2024-02-12','2024-02-27',5,478,478,'','2024-01-28',true,false,0,'Exquisite prose and a plot that stuns.'],
  ['No One Listens','Rico Martinez','Thriller','Domestic Thriller','Paperback','',null,2023,'English','St. Martin\'s Press','Fiction','Finished','2024-03-04','2024-03-16',3,298,298,'','2024-02-28',false,false,0,'Decent but predictable ending.'],
  ['The Red Envelope','Li Wei','Mystery','Crime Fiction','eBook','',null,2022,'English','Soho Crime','Fiction','Finished','2024-04-08','2024-04-22',4,423,423,'','2024-03-30',false,false,0,'Excellent multicultural setting and pacing.'],
  ['Cold Case: Alderton','Frank Bellamy','Thriller','Legal Thriller','Paperback','',null,2023,'English','Little Brown','Fiction','Finished','2024-05-01','2024-05-14',3,389,389,'','2024-04-25',false,false,0,'Good courtroom drama, slightly formulaic.'],
  ['Midnight at the Mortuary','Solomon Pierce','Mystery','Cozy Mystery','Paperback','Mortuary Mysteries',1,2024,'English','Berkley','Fiction','Finished','2024-06-03','2024-06-17',4,334,334,'','2024-05-25',false,false,0,'A fun, quirky series opener.'],
  ['Buried in Secrets','Nadine Frost','Thriller','Psychological Thriller','Hardcover','',null,2024,'English','Penguin Books','Fiction','Finished','2024-07-01','2024-07-15',4,351,351,'','2024-06-20',false,false,0,'Atmospheric and genuinely unsettling.'],

  // Science Fiction (10)
  ['Quantum Drift','Zara Ito','Science Fiction','Space Opera','Hardcover','Drift Series',1,2021,'English','Tor Books','Favorites','Finished','2023-12-01','2023-12-22',5,512,512,'','2023-11-20',true,false,0,'Best sci-fi I have read in years.'],
  ['The Last Colony','Erik Svensson','Science Fiction','Space Opera','Paperback','',null,2019,'English','Orbit Books','Fiction','Finished','2024-01-28','2024-02-10',4,678,678,'','2024-01-15',false,false,0,'Epic scale and thoughtful colonialism themes.'],
  ['Neural Storm','Anika Patel','Science Fiction','Cyberpunk','eBook','',null,2023,'English','Del Rey','Fiction','Finished','2024-08-05','2024-08-21',4,445,445,'','2024-07-28',false,false,0,'Fast-paced and visually imaginative.'],
  ['The Terraformers','Claudia Reyes','Science Fiction','Hard SF','Paperback','',null,2022,'English','Tordotcom','Fiction','Finished','2024-09-03','2024-09-20',3,589,589,'','2024-08-25',false,false,0,'Dense but rewarding. Slow start.'],
  ['Children of the Wormhole','Yuki Tanaka','Science Fiction','Space Opera','Hardcover','',null,2020,'English','Tor Books','Favorites','Finished','2025-01-06','2025-01-24',5,634,634,'','2024-12-28',true,false,0,'Emotionally devastating and beautiful.'],
  ['The Memory Engine','Dmitri Volkov','Science Fiction','Cyberpunk','eBook','',null,2023,'English','Angry Robot','Fiction','Finished','2025-02-03','2025-02-18',4,487,487,'','2025-01-25',false,false,0,'Explores memory and identity brilliantly.'],
  ['Singularity\'s Edge','Clara Bishop','Science Fiction','Hard SF','Paperback','',null,2022,'English','Baen Books','Fiction','Finished','2025-03-03','2025-03-18',4,398,398,'','2025-02-20',false,false,0,'Gripping AI ethics narrative.'],
  ['Far Horizons','Paulo Soto','Science Fiction','First Contact','eBook','',null,2021,'English','Tor Books','Fiction','Finished','2025-04-07','2025-04-23',3,567,567,'','2025-03-25',false,false,0,'Interesting but the finale underwhelmed.'],
  ['The Signal','Nadia Khalil','Science Fiction','Biopunk','Paperback','',null,2024,'English','Head of Zeus','Fiction','Finished','2025-05-05','2025-05-21',4,412,412,'','2025-04-25',false,false,0,'Original premise, strong character work.'],
  ['Starfall Protocol','James Osei','Science Fiction','Military SF','Hardcover','Drift Series',2,2024,'English','Tor Books','Fiction','Reading','2026-09-01','',0,523,220,'','2026-08-25',false,false,0,'Sequel to Quantum Drift — action packed so far.'],

  // Literary Fiction (10)
  ['The Cartographer\'s Wife','Helena Mercer','Literary Fiction','Contemporary','Hardcover','',null,2020,'English','Knopf','Favorites','Finished','2024-02-28','2024-03-15',5,378,378,'','2024-02-15',true,true,1,'One of the most beautifully written books I have ever read.'],
  ['Songs We Never Sang','Dominic Abara','Literary Fiction','Family Saga','Paperback','',null,2022,'English','Riverhead','Fiction','Finished','2024-03-18','2024-03-31',4,312,312,'','2024-03-10',false,false,0,'Tender and honest family portrait.'],
  ['Letters From the Shore','Isabelle Laurent','Literary Fiction','Epistolary','eBook','',null,2021,'English','Bloomsbury','Favorites','Finished','2024-10-01','2024-10-16',5,289,289,'','2024-09-22',true,false,0,'Moved me to tears. Stunningly written.'],
  ['The Amber Years','Soo-Young Park','Literary Fiction','Coming-of-Age','Paperback','',null,2023,'English','Holt','Fiction','Finished','2024-11-04','2024-11-18',4,445,445,'','2024-10-20',false,false,0,'Vivid and emotionally resonant.'],
  ['Where the Light Falls','Gabriel Nwosu','Literary Fiction','Contemporary','Hardcover','',null,2024,'English','Farrar Straus','Fiction','Finished','2025-06-02','2025-06-18',4,367,367,'','2025-05-25',false,false,0,'Quiet but deeply affecting.'],
  ['The Last Summer','Vera Kossuth','Literary Fiction','Historical','eBook','',null,2019,'English','Penguin','Fiction','Finished','2025-07-07','2025-07-21',3,423,423,'','2025-06-30',false,false,0,'Atmospheric but overlong in the middle.'],
  ['The Third Daughter','Chiara Romano','Literary Fiction','Family Saga','Paperback','',null,2023,'English','Riverhead','Favorites','Finished','2025-08-04','2025-08-20',5,512,512,'','2025-07-25',true,false,0,'Lyrical and devastating. Highly recommended.'],
  ['Under the Olive Tree','Rashida Osman','Literary Fiction','Contemporary','Hardcover','',null,2021,'English','Bloomsbury','Fiction','Finished','2025-09-01','2025-09-16',4,398,398,'','2025-08-25',false,false,0,'Evocative and warm — beautiful read.'],
  ['The Quiet Undoing','Peter Strand','Literary Fiction','Psychological','eBook','',null,2024,'English','Picador','Fiction','Reading','2026-09-01','',0,334,180,'','2026-08-28',false,false,0,'Eerie and slow-building. Loving it.'],
  ['The Long Way Home','Sienna MacPherson','Literary Fiction','Coming-of-Age','Paperback','',null,2022,'English','Little Brown','Favorites','Finished','2025-10-06','2025-10-22',5,367,367,'','2025-09-30',true,false,0,'Perfect. Will gift copies to everyone I know.'],

  // Historical Fiction (10)
  ['The Silk Road Merchant','Leila Nassar','Historical Fiction','Ancient World','Hardcover','',null,2020,'English','Knopf','Favorites','Finished','2023-05-10','2023-06-01',5,623,623,'','2023-04-28',true,false,0,'Immersive journey through medieval Central Asia.'],
  ['Daughters of the Resistance','Yvonne Petit','Historical Fiction','WWII','Paperback','',null,2021,'English','HarperCollins','Favorites','Finished','2023-09-22','2023-10-10',5,578,578,'','2023-09-15',true,false,0,'Deeply moving and expertly researched.'],
  ['The Empress\'s Garden','Ming-Li Zhao','Historical Fiction','Imperial China','Hardcover','',null,2022,'English','Knopf','Fiction','Finished','2024-04-24','2024-05-10',4,689,689,'','2024-04-15',false,false,0,'Richly detailed palace intrigue.'],
  ['Iron and Ashes','Cormac O\'Brien','Historical Fiction','Medieval','Paperback','',null,2019,'English','Macmillan','Fiction','Finished','2024-06-20','2024-07-08',4,745,745,'','2024-06-10',false,false,0,'Brutal but gripping medieval epic.'],
  ['The Warsaw Letters','Anna Kowalczyk','Historical Fiction','WWII','eBook','',null,2022,'English','Penguin','Favorites','Finished','2025-02-20','2025-03-10',5,612,612,'','2025-02-12',true,false,0,'Heartbreaking and beautifully rendered.'],
  ['The Tudor Spy','Harriet Vance','Historical Fiction','Tudor England','Paperback','The Tudor Files',1,2023,'English','Hodder','Fiction','Finished','2025-04-25','2025-05-09',4,534,534,'','2025-04-18',false,false,0,'Clever plotting in a vivid historical setting.'],
  ['A Viking\'s Promise','Lars Eriksen','Historical Fiction','Viking Age','Hardcover','',null,2021,'English','Bantam','Fiction','Finished','2025-06-23','2025-07-08',3,487,487,'','2025-06-15',false,false,0,'Solid adventure, somewhat predictable.'],
  ['The Last Empress','Sachi Yamamoto','Historical Fiction','Meiji Japan','eBook','',null,2024,'English','Tor Books','Favorites','Finished','2025-08-22','2025-09-07',5,656,656,'','2025-08-15',true,false,0,'Stunning. The best historical fiction I have read this year.'],
  ['Rome\'s Forgotten','Julius Mancini','Historical Fiction','Ancient Rome','Paperback','',null,2020,'English','Knopf','Fiction','Finished','2025-10-23','2025-11-09',4,589,589,'','2025-10-15',false,false,0,'Atmospheric and well-researched.'],
  ['The Alchemist\'s Daughter','Beatrice Holden','Historical Fiction','Renaissance','Hardcover','',null,2022,'English','Doubleday','Fiction','Finished','2026-01-08','2026-01-26',4,567,567,'','2025-12-28',false,false,0,'Fascinating historical detail woven into a gripping plot.'],

  // Romance (10)
  ['The Bookshop in Paris','Amélie Fontaine','Romance','Contemporary','Paperback','',null,2022,'French','Gallimard','Favorites','Finished','2023-08-01','2023-08-14',5,312,312,'','2023-07-25',true,false,0,'Charming and warm — the perfect summer read.'],
  ['When Stars Align','Rosalind Chase','Romance','Contemporary','eBook','',null,2023,'English','Avon','Fiction','Finished','2024-07-16','2024-07-28',4,345,345,'','2024-07-08',false,false,0,'Sweet and satisfying romance arc.'],
  ['The Unexpected Mr. Collins','Penelope Hart','Romance','Regency Romance','Paperback','',null,2021,'English','Berkley','Fiction','Finished','2024-08-22','2024-09-03',4,412,412,'','2024-08-15',false,false,0,'Witty Austen-adjacent romance — great fun.'],
  ['Love in Translation','Yumi Sasaki','Romance','Contemporary','eBook','',null,2023,'Japanese','Kodansha','Fiction','Finished','2024-09-25','2024-10-07',4,334,334,'','2024-09-18',false,false,0,'Cross-cultural love story — touching and funny.'],
  ['One More Sunrise','Clara Hughes','Romance','Second Chance','Paperback','',null,2024,'English','Harlequin','Favorites','Finished','2024-11-20','2024-12-02',5,298,298,'','2024-11-12',true,false,0,'Emotional and beautifully paced second-chance romance.'],
  ['Dancing in the Rain','Santiago Vargas','Romance','Sports Romance','eBook','',null,2023,'Spanish','Planeta','Fiction','Finished','2025-03-20','2025-04-02',3,345,345,'','2025-03-12',false,false,0,'Fun premise, a bit formulaic in the resolution.'],
  ['Always and Forever','Daniella Fox','Romance','Contemporary','Paperback','',null,2024,'English','Avon','Fiction','Finished','2025-05-26','2025-06-09',4,378,378,'','2025-05-20',false,false,0,'Enjoyable modern romance with great chemistry.'],
  ['Finding Forever','Maya Johnson','Romance','Small Town Romance','eBook','',null,2021,'English','Kensington','Fiction','Finished','2025-07-25','2025-08-07',3,289,289,'','2025-07-18',false,false,0,'Charming but predictable.'],
  ['The Second Chance','Oliver Wright','Romance','Second Chance','Paperback','',null,2024,'English','St. Martin\'s','Fiction','Finished','2025-11-17','2025-12-01',4,378,378,'','2025-11-10',false,false,0,'Warm and genuinely touching reunion story.'],
  ['The Summer We Found Each Other','Bella Montoya','Romance','Beach Read','Hardcover','',null,2022,'English','Doubleday','Fiction','Finished','2026-06-10','2026-06-24',4,367,367,'','2026-06-03',false,false,0,'The perfect summer read.'],

  // Horror (8)
  ['The Dark Between Trees','Caleb Stone','Horror','Atmospheric Horror','Paperback','',null,2022,'English','Tor Nightfire','Fiction','Finished','2023-10-16','2023-10-30',5,456,456,'','2023-10-10',true,false,0,'Read in one October weekend. Terrifying.'],
  ['The Waking Dark','Isaiah Marsh','Horror','Supernatural Horror','Hardcover','',null,2022,'English','Del Rey','Fiction','Finished','2024-10-19','2024-11-03',5,534,534,'','2024-10-12',true,false,0,'Genuinely scary in the best way.'],
  ['The Haunting of Hill Manor','Priscilla Thornton','Horror','Gothic Horror','Paperback','',null,2021,'English','Cemetery Dance','Favorites','Finished','2025-09-22','2025-10-06',5,467,467,'','2025-09-15',true,true,1,'Reread every October. A masterpiece.'],
  ['Hollowed','Miriam Westley','Horror','Folk Horror','eBook','',null,2023,'English','Tor Nightfire','Fiction','Finished','2024-10-07','2024-10-20',4,398,398,'','2024-09-30',false,false,0,'Eerie and slow-burn atmospheric horror.'],
  ['Teeth and Shadows','Rosa Delgado','Horror','Vampire Horror','Paperback','',null,2024,'English','Tor Nightfire','Fiction','Finished','2025-10-24','2025-11-07',4,387,387,'','2025-10-17',false,false,0,'Fresh take on vampire mythology.'],
  ['Skin Deep','Lena Voss','Horror','Body Horror','eBook','',null,2023,'English','Titan Books','Fiction','Finished','2025-11-11','2025-11-24',4,445,445,'','2025-11-04',false,false,0,'Deeply unsettling — not for the faint of heart.'],
  ['Pale Fire Rising','Elliot Crane','Horror','Supernatural Horror','Paperback','',null,2022,'English','Solaris','Fiction','Finished','2026-02-03','2026-02-18',4,423,423,'','2026-01-28',false,false,0,'Excellent debut horror novel.'],
  ['Midnight\'s Children','Declan Murray','Horror','Atmospheric Horror','Hardcover','',null,2024,'English','St. Martin\'s','Fiction','Reading','2026-09-02','',0,389,150,'','2026-08-30',false,false,0,'Just started — creepy atmosphere so far.'],

  // Non-Fiction (8)
  ['The Art of Deep Reading','Prof. Chen Liu','Non-Fiction','Reading & Books','Hardcover','',null,2021,'English','Oxford UP','Non-Fiction','Finished','2023-11-20','2023-12-05',5,256,256,'','2023-11-10',true,false,0,'Transformative book about how we engage with literature.'],
  ['How Books Changed the World','Samantha Blake','Non-Fiction','History of Books','Paperback','',null,2020,'English','Harvard UP','Non-Fiction','Finished','2024-01-03','2024-01-15',4,312,312,'','2023-12-28',false,false,0,'Fascinating history of print culture.'],
  ['The Power of Stories','Dr. Marcus Gray','Non-Fiction','Cognitive Science','eBook','',null,2022,'English','MIT Press','Non-Fiction','Finished','2024-03-22','2024-04-05',4,289,289,'','2024-03-15',false,false,0,'Neuroscience of narrative — eye-opening.'],
  ['Reading Minds: The Science of Stories','Dr. Avery Park','Non-Fiction','Psychology','Hardcover','',null,2023,'English','Princeton UP','Non-Fiction','Finished','2024-12-04','2024-12-18',5,334,334,'','2024-11-25',true,false,0,'A must-read for any book lover.'],
  ['Why We Read Fiction','Nora Callahan','Non-Fiction','Literary Theory','Paperback','',null,2024,'English','Columbia UP','Non-Fiction','Finished','2025-01-27','2025-02-10',5,267,267,'','2025-01-20',true,false,0,'Elegant argument for the value of literary fiction.'],
  ['The Novel\'s Secret History','Victor Randall','Non-Fiction','Literary History','eBook','',null,2022,'English','Yale UP','Non-Fiction','Finished','2025-08-26','2025-09-09',4,423,423,'','2025-08-18',false,false,0,'Scholarly but accessible history of the novel form.'],
  ['Reading in the Digital Age','Felicia Strom','Non-Fiction','Media Studies','Paperback','',null,2023,'English','MIT Press','Non-Fiction','Finished','2026-03-03','2026-03-18',3,298,298,'','2026-02-25',false,false,0,'Interesting but slightly overstated conclusions.'],
  ['Book Culture: A Global History','Amara Diallo','Non-Fiction','Cultural History','Hardcover','',null,2024,'English','Routledge','Non-Fiction','Finished','2026-04-01','2026-04-17',4,512,512,'','2026-03-25',false,false,0,'Comprehensive global survey — excellent reference.'],

  // Young Adult (8)
  ['The Map to Tomorrow','Zoe Rivera','Young Adult','YA Fantasy','Paperback','',null,2022,'English','Hyperion','Fiction','Finished','2023-04-29','2023-05-12',5,367,367,'','2023-04-22',true,false,0,'Inventive YA fantasy with a memorable protagonist.'],
  ['Storm and Starlight','Asha Mehta','Young Adult','YA Sci-Fi','eBook','',null,2023,'English','Tor Teen','Fiction','Finished','2024-06-24','2024-07-08',4,398,398,'','2024-06-18',false,false,0,'Exciting and emotionally gripping YA adventure.'],
  ['Rise of the Gifted','Tara Chen','Young Adult','YA Fantasy','Hardcover','',null,2024,'English','Scholastic','Fiction','Finished','2025-07-22','2025-08-05',5,456,456,'','2025-07-15',true,false,0,'Best YA I have read this year — compelling magic system.'],
  ['The Forgotten Kingdom','Liam Prescott','Young Adult','YA Fantasy','Paperback','',null,2021,'English','HarperTeen','Fiction','Finished','2025-05-26','2025-06-09',4,423,423,'','2025-05-19',false,false,0,'Solid YA fantasy with good pacing.'],
  ['The Last Choice','Finn Donaghue','Young Adult','YA Dystopia','Hardcover','',null,2024,'English','Little Brown','Favorites','Finished','2025-12-08','2025-12-22',5,412,412,'','2025-12-01',true,false,0,'Stunning and thought-provoking YA dystopia.'],
  ['Spark and Shadow','Celeste Marino','Young Adult','YA Fantasy','Paperback','',null,2022,'English','Scholastic','Fiction','Finished','2026-02-20','2026-03-06',4,345,345,'','2026-02-14',false,false,0,'Fun and fast-paced YA fantasy.'],
  ['A Song for the Broken','Jasmine Reed','Young Adult','YA Contemporary','eBook','',null,2023,'English','HarperCollins','Fiction','Finished','2026-04-21','2026-05-06',3,389,389,'','2026-04-14',false,false,0,'Emotionally honest but uneven pacing.'],
  ['Light in the Dark','Amani Osei','Young Adult','YA Contemporary','Paperback','',null,2024,'English','Scholastic','Fiction','To Read','','',0,378,0,'','2026-07-05',false,false,0,'Recommended by a friend. On my to-read pile.'],

  // Biography & Memoir (6)
  ['My Years in the Archive','Constance Lively','Memoir','Literary Memoir','Hardcover','',null,2021,'English','Knopf','Non-Fiction','Finished','2023-11-06','2023-11-20',5,312,312,'','2023-10-30',true,false,0,'A librarian\'s life in books — utterly charming.'],
  ['The Bookseller\'s Memoir','Roland Hunt','Memoir','Business Memoir','Paperback','',null,2022,'English','Bloomsbury','Non-Fiction','Finished','2024-12-20','2025-01-04',4,278,278,'','2024-12-12',false,false,0,'Warm portrait of an independent bookshop.'],
  ['Shelf Life: A Literary Memoir','Patrick Oswald','Memoir','Literary Memoir','eBook','',null,2024,'English','Farrar Straus','Non-Fiction','Finished','2025-09-18','2025-10-02',5,289,289,'','2025-09-10',true,false,0,'A life told through the books that shaped it. Extraordinary.'],
  ['Reading My Mother','Elena Sato','Memoir','Personal Memoir','Paperback','',null,2022,'English','Penguin','Non-Fiction','Finished','2025-11-24','2025-12-08',4,334,334,'','2025-11-17',false,false,0,'Tender and moving family memoir.'],
  ['The Library Keeper','George Addison','Biography','Literary Biography','Hardcover','',null,2023,'English','Oxford UP','Non-Fiction','Finished','2026-05-05','2026-05-21',3,267,267,'','2026-04-28',false,false,0,'Interesting subject, slightly dry writing style.'],
  ['Pages of My Life','Adriana Cruz','Memoir','Personal Memoir','eBook','',null,2023,'English','Vintage','Non-Fiction','Finished','2026-07-02','2026-07-17',4,356,356,'','2026-06-25',false,false,0,'Honest and beautifully written memoir.'],

  // Graphic Novel (5)
  ['The Ink Dimension','Alex Torres','Graphic Novel','Science Fiction GN','Paperback','',null,2022,'English','Image Comics','Fiction','Finished','2023-07-24','2023-07-28',4,234,234,'','2023-07-20',false,false,0,'Stunning artwork and a clever multiverse premise.'],
  ['Watercolor World','Mei Lin','Graphic Novel','Slice of Life GN','Hardcover','',null,2023,'English','Drawn & Quarterly','Favorites','Finished','2024-02-04','2024-02-07',5,198,198,'','2024-01-30',true,false,0,'Pure visual poetry. Gorgeous in every way.'],
  ['The Glass City','Damien Kowalski','Graphic Novel','Fantasy GN','Paperback','',null,2021,'English','Dark Horse','Fiction','Finished','2024-11-25','2024-11-28',4,256,256,'','2024-11-20',false,false,0,'Intricate worldbuilding through stunning artwork.'],
  ['Night Pages','Camille Deschamps','Graphic Novel','Horror GN','eBook','',null,2024,'English','Image Comics','Fiction','Finished','2025-10-13','2025-10-16',4,212,212,'','2025-10-08',false,false,0,'Effectively creepy October read.'],
  ['Neon Garden','Kenji Watanabe','Graphic Novel','Cyberpunk GN','Hardcover','',null,2023,'Japanese','Kodansha','Fiction','Finished','2026-06-03','2026-06-06',4,198,198,'','2026-05-28',false,false,0,'Visually spectacular cyberpunk graphic novel.'],

  // Self-Help (4)
  ['The Reading Life','Dr. Jean Ellis','Self-Help','Reading Habits','Paperback','',null,2022,'English','Hay House','Non-Fiction','Finished','2024-09-11','2024-09-22',4,234,234,'','2024-09-04',false,false,0,'Practical tips for building a sustainable reading habit.'],
  ['How to Remember Everything You Read','Dr. Sarah Kim','Self-Help','Learning','eBook','',null,2024,'English','Penguin Life','Non-Fiction','Finished','2025-03-26','2025-04-08',5,245,245,'','2025-03-18',true,false,0,'Game-changing reading techniques. Highly recommend.'],
  ['The Bookworm\'s Blueprint','Chris Hanson','Self-Help','Reading Habits','Paperback','',null,2022,'English','Portfolio','Non-Fiction','Finished','2026-01-28','2026-02-12',3,212,212,'','2026-01-20',false,false,0,'Some useful tips but repetitive in places.'],
  ['365 Books, 365 Days','Marcus O\'Connell','Self-Help','Reading Goals','Hardcover','',null,2023,'English','Chronicle','Non-Fiction','Finished','2026-08-05','2026-08-20',3,198,198,'','2026-07-28',false,false,0,'Fun challenge concept but unrealistic for most readers.'],

  // Additional To Read / upcoming (21 more for 125+ total)
  ['Echoes of the Empire','Renata Sousa','Fantasy','Epic Fantasy','Paperback','',null,2024,'English','Ember Press','To Read Soon','To Read','','',0,589,0,'','2026-05-10',false,false,0,'Sequel to the Ember Crown series — cannot wait.'],
  ['The Lighthouse Keeper\'s Son','Owen Bray','Literary Fiction','Contemporary','Hardcover','',null,2023,'English','Riverhead','To Read Soon','To Read','','',0,423,0,'','2026-04-22',false,false,0,'Longlisted for a major literary prize.'],
  ['The Crystal Court','Elena Vasquez','Fantasy','Epic Fantasy','Hardcover','The Throne Saga',3,2025,'English','Ember Press','To Read Soon','To Read','','',0,612,0,'','2026-09-01',false,false,0,'Just released — the Throne Saga finale!'],
  ['The City of Glass Towers','Fatima El-Amin','Science Fiction','Dystopian SF','Hardcover','',null,2024,'English','Tor Books','To Read Soon','To Read','','',0,612,0,'','2026-03-18',false,false,0,'Hugely buzzed about dystopian debut.'],
  ['Crimson Tide','Rhett Barlow','Thriller','Crime Thriller','Paperback','',null,2024,'English','Putnam','To Read Soon','To Read','','',0,378,0,'','2026-05-25',false,false,0,'Gift from a colleague. Looks gripping.'],
  ['The Pearl Diver','Nadia Haddad','Historical Fiction','Gulf History','Paperback','',null,2023,'English','Bloomsbury','Fiction','Paused','2026-03-01','',0,489,120,'','2026-02-20',false,false,0,'Beautiful writing but set aside mid-read.'],
  ['Fractured Kingdom','Seun Adeyemi','Fantasy','African Fantasy','Hardcover','',null,2023,'English','Tor Books','To Read Soon','To Read','','',0,678,0,'','2026-06-15',false,false,0,'West African mythological fantasy — very excited.'],
  ['The Mage\'s Burden','Finn Kelley','Fantasy','Epic Fantasy','eBook','The Mage Chronicles',1,2023,'English','Angry Robot','To Read Soon','To Read','','',0,456,0,'','2026-07-18',false,false,0,'First in a new epic fantasy series.'],
  ['Beneath the Surface','Laila Hassan','Literary Fiction','Contemporary','Paperback','',null,2022,'English','Holt','Fiction','Finished','2026-05-27','2026-06-11',4,345,345,'','2026-05-20',false,false,0,'Quietly powerful look at diaspora identity.'],
  ['Forgotten Tides','Blake Morrison','Historical Fiction','Age of Sail','Hardcover','',null,2024,'English','Macmillan','To Read Soon','To Read','','',0,523,0,'','2026-08-10',false,false,0,'Maritime historical fiction — love this setting.'],
  ['Where Ravens Fly','Tilda Svengren','Mystery','Nordic Noir','Paperback','',null,2024,'English','Penguin Scandinavia','To Read Soon','To Read','','',0,412,0,'','2026-08-22',false,false,0,'Nordic noir — perfect autumn read.'],
  ['The Painted House','Rosa Espinoza','Literary Fiction','Magical Realism','Paperback','',null,2023,'English','Vintage','Fiction','Finished','2026-07-21','2026-08-05',4,378,378,'','2026-07-14',false,false,0,'Lush magical realism from a compelling new voice.'],
  ['Beyond the Edge','Yusuf Al-Rashid','Science Fiction','Space Exploration','eBook','',null,2024,'English','Gollancz','To Read Soon','To Read','','',0,534,0,'','2026-09-02',false,false,0,'Acquired at a bookshop — intriguing premise.'],
  ['The Parchment Letters','Olivia Strand','Historical Fiction','Tudor England','Hardcover','The Tudor Files',2,2024,'English','Hodder','Fiction','Reading','2026-08-10','',0,512,280,'','2026-08-04',false,false,0,'Tudor Files sequel. Even better than book 1!'],
  ['Glass and Stone','Adriano Ferreira','Fantasy','Dark Fantasy','Paperback','',null,2023,'English','Orbit Books','Fiction','Finished','2026-03-24','2026-04-08',3,489,489,'','2026-03-17',false,false,0,'Interesting magic system but uneven character work.'],
  ['The Last Midnight','Cassandra Vale','Mystery','Gothic Mystery','Hardcover','',null,2024,'English','Minotaur','To Read Soon','To Read','','',0,356,0,'','2026-06-30',false,false,0,'Gothic mystery with incredible reviews.'],
  ['Aurora Falls','Jenny Kwan','Romance','Paranormal Romance','Paperback','',null,2024,'English','Berkley','To Read Soon','To Read','','',0,312,0,'','2026-07-22',false,false,0,'Paranormal romance recommended by a friend.'],
  ['The Mind\'s Eye','Dr. Reginald Stone','Non-Fiction','Neuroscience','Hardcover','',null,2023,'English','Basic Books','To Read Soon','To Read','','',0,289,0,'','2026-08-12',false,false,0,'Popular neuroscience — very well reviewed.'],
  ['The Hollow Season','Nigel Crowe','Horror','Folk Horror','Paperback','',null,2021,'English','Cemetery Dance','Fiction','Finished','2024-10-25','2024-11-08',3,512,512,'','2024-10-18',false,false,0,'Interesting folk horror premise but pacing issues.'],
  ['The Feeding Ground','Vanessa King','Horror','Creature Horror','eBook','',null,2023,'English','Tor Nightfire','Fiction','Finished','2025-10-09','2025-10-22',3,512,512,'','2025-10-02',false,false,0,'Tense creature feature but lacked depth.'],
  ['The Byzantine Code','Eleni Papadakis','Historical Fiction','Byzantine Empire','Paperback','',null,2023,'English','Penguin','Fiction','Finished','2026-07-08','2026-07-23',4,478,478,'','2026-07-01',false,false,0,'Fascinating window into Byzantine history and politics.'],
];

(async () => {
  const fmt  = [];
  const vals = [];

  // ── Background wash ───────────────────────────────────────────────────────
  fmt.push({ repeatCell: { range: gridRange(SID,0,1100,0,24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg), textFormat: { fontSize: 10, fontFamily: 'Georgia', foregroundColor: hex(C.text) },
  }}, fields: 'userEnteredFormat(backgroundColor,textFormat)' }});

  // ── Title row ─────────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A1`, values: [['📚 Master Book Library']] });
  fmt.push({ mergeCells: { range: gridRange(SID,0,1,0,24), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,0,1,0,24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 16, foregroundColor: hex(C.white), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 42 }, fields: 'pixelSize' }});

  // ── Subtitle row ──────────────────────────────────────────────────────────
  vals.push({ range: `${S}!A2`, values: [['Your personal reading universe — track every book you\'ve read, are reading, or want to read.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,1,2,0,24), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,1,2,0,24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.goldTint), textFormat: { italic: true, fontSize: 10, foregroundColor: hex(C.secText), fontFamily: 'Georgia' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 26 }, fields: 'pixelSize' }});

  // ── Stat cards row 3 (quick stats via formulas) ───────────────────────────
  const LIB = S;
  vals.push({ range: `${S}!A3`, values: [[
    '📖 Total Books',
    `=COUNTA(${LIB}!B8:B1008)`,
    '✅ Finished',
    `=SUMPRODUCT((${LIB}!M8:M1008="Finished")*1)`,
    '⭐ Avg Rating',
    `=IFERROR(ROUND(AVERAGEIF(${LIB}!M8:M1008,"Finished",${LIB}!P8:P1008),1),"—")`,
    '📄 Pages Read',
    `=IFERROR(SUMIF(${LIB}!M8:M1008,"Finished",${LIB}!Q8:Q1008),"—")`,
    '📅 Read This Year',
    `=SUMPRODUCT((YEAR(IFERROR(DATEVALUE(${LIB}!O8:O1008),0))=YEAR(TODAY()))*(${LIB}!M8:M1008="Finished"))`,
    '💝 Favorites',
    `=SUMPRODUCT((${LIB}!U8:U1008=TRUE)*1)`,
    '🔖 Reading Now',
    `=SUMPRODUCT((${LIB}!M8:M1008="Reading")*1)`,
    '📋 To Read',
    `=SUMPRODUCT((${LIB}!M8:M1008="To Read")*1)`,
  ]] });
  // Format stat cards — pairs of (label, value) in groups of 2 cols
  for (let pair = 0; pair < 8; pair++) {
    const c1 = pair * 3;
    const bgLabel = pair % 2 === 0 ? C.primary : C.secondary;
    const bgValue = pair % 2 === 0 ? C.wineTint : C.goldTint;
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1,c1+1), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1,c1+1), cell: { userEnteredFormat: {
      backgroundColor: hex(bgLabel), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
    fmt.push({ mergeCells: { range: gridRange(SID,2,3,c1+1,c1+2), mergeType: 'MERGE_ALL' }});
    fmt.push({ repeatCell: { range: gridRange(SID,2,3,c1+1,c1+2), cell: { userEnteredFormat: {
      backgroundColor: hex(bgValue), textFormat: { bold: true, fontSize: 12, foregroundColor: hex(C.text), fontFamily: 'Arial' },
      horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 38 }, fields: 'pixelSize' }});

  // Rows 4-6: tip text + empty buffer
  vals.push({ range: `${S}!A4`, values: [['💡 HOW TO USE: Enter each book once here — all other tabs link to this library via Book ID. Yellow cells = enter data. Blue cells = auto-calculated. Checkboxes auto-format.']] });
  fmt.push({ mergeCells: { range: gridRange(SID,3,4,0,24), mergeType: 'MERGE_ALL' }});
  fmt.push({ repeatCell: { range: gridRange(SID,3,4,0,24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.goldTint), textFormat: { italic: true, fontSize: 9, foregroundColor: hex(C.text), fontFamily: 'Arial' },
    horizontalAlignment: 'LEFT', verticalAlignment: 'MIDDLE',
    padding: { left: 8 },
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 24 }, fields: 'pixelSize' }});

  // Rows 5-6: empty
  fmt.push({ repeatCell: { range: gridRange(SID,4,6,0,24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.bg),
  }}, fields: 'userEnteredFormat.backgroundColor' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 4, endIndex: 6 }, properties: { pixelSize: 8 }, fields: 'pixelSize' }});

  // ── Header row (row 7, 0-indexed = 6) ─────────────────────────────────────
  const HEADERS = ['Book ID','Title','Author','Genre','Sub-genre','Format','Series Name','Series #','Year Published','Language','Publisher','Shelf','Status','Date Started','Date Finished','Rating (1-5)','Total Pages','Pages Read','Progress %','Date Added','Favorite?','Reread?','Times Reread','Notes'];
  vals.push({ range: `${S}!A7`, values: [HEADERS] });
  fmt.push({ repeatCell: { range: gridRange(SID,6,7,0,24), cell: { userEnteredFormat: {
    backgroundColor: hex(C.primary), textFormat: { bold: true, fontSize: 9, foregroundColor: hex(C.white), fontFamily: 'Arial' },
    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
  }}, fields: 'userEnteredFormat' }});
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 6, endIndex: 7 }, properties: { pixelSize: 36 }, fields: 'pixelSize' }});

  // ── Formula columns A (Book ID) and S (Progress %) ─────────────────────────
  fmt.push({ repeatCell: {
    range: gridRange(SID,7,1008,0,1),
    cell: { userEnteredValue: { formulaValue: '=IF(B8="","","BOOK-"&TEXT(ROW()-7,"00000"))' }, userEnteredFormat: {
      backgroundColor: hex(C.formula), textFormat: { fontSize: 9, fontFamily: 'Arial', bold: true },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredValue,userEnteredFormat',
  }});
  fmt.push({ repeatCell: {
    range: gridRange(SID,7,1008,18,19),
    cell: { userEnteredValue: { formulaValue: '=IFERROR(IF(M8="Finished",1,IF(M8="To Read",0,MIN(1,R8/Q8))),0)' }, userEnteredFormat: {
      backgroundColor: hex(C.formula), numberFormat: { type: 'PERCENT', pattern: '0%' },
      textFormat: { fontSize: 9, fontFamily: 'Arial' }, horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredValue,userEnteredFormat',
  }});

  // ── Data row alternating background ──────────────────────────────────────
  for (let i = 0; i < 500; i++) {
    const bg = i % 2 === 0 ? C.white : C.altRow;
    fmt.push({ repeatCell: { range: gridRange(SID,7+i,8+i,0,24), cell: { userEnteredFormat: {
      backgroundColor: hex(bg), textFormat: { fontSize: 9, fontFamily: 'Arial' }, verticalAlignment: 'MIDDLE',
    }}, fields: 'userEnteredFormat' }});
  }

  // Input cell highlight (yellow) for B-R, T-W columns (not formula cols A and S)
  // Input: cols B(1)-N(13), P(15)-R(17) + T(19)-W(22)
  // Formula/auto: A(0), S(18), X(23) is notes (white)
  fmt.push({ repeatCell: { range: gridRange(SID,7,1008,1,14), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
  }}, fields: 'userEnteredFormat.backgroundColor' }});
  fmt.push({ repeatCell: { range: gridRange(SID,7,1008,15,18), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
  }}, fields: 'userEnteredFormat.backgroundColor' }});
  fmt.push({ repeatCell: { range: gridRange(SID,7,1008,19,23), cell: { userEnteredFormat: {
    backgroundColor: hex(C.input),
  }}, fields: 'userEnteredFormat.backgroundColor' }});

  // Rating column P (15): center
  fmt.push({ repeatCell: { range: gridRange(SID,7,1008,15,16), cell: { userEnteredFormat: {
    horizontalAlignment: 'CENTER',
  }}, fields: 'userEnteredFormat.horizontalAlignment' }});
  // Numeric cols Q(16), R(17), W(22): center
  [16,17,22].forEach(ci => {
    fmt.push({ repeatCell: { range: gridRange(SID,7,1008,ci,ci+1), cell: { userEnteredFormat: {
      horizontalAlignment: 'CENTER',
    }}, fields: 'userEnteredFormat.horizontalAlignment' }});
  });

  // ── Column widths ──────────────────────────────────────────────────────────
  const widths = [90,220,160,120,110,90,140,60,70,80,110,90,80,90,90,65,70,70,70,90,60,60,70,200];
  widths.forEach((px, ci) => {
    fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'COLUMNS', startIndex: ci, endIndex: ci+1 }, properties: { pixelSize: px }, fields: 'pixelSize' }});
  });

  // Row heights for data rows
  fmt.push({ updateDimensionProperties: { range: { sheetId: SID, dimension: 'ROWS', startIndex: 7, endIndex: 1008 }, properties: { pixelSize: 22 }, fields: 'pixelSize' }});

  // ── Freeze rows and columns ───────────────────────────────────────────────
  fmt.push({ updateSheetProperties: { properties: { sheetId: SID, gridProperties: { frozenRowCount: 7 } }, fields: 'gridProperties.frozenRowCount' }});

  await batchUpdate(id, fmt, '03-library format');

  // ── Sample data ───────────────────────────────────────────────────────────
  // Each book: 23 values for cols B-X (B=idx0, S at idx17 is '' placeholder - formula takes over)
  BOOKS.forEach((book, i) => {
    const row = 8 + i;
    // Split at S (index 17): B:R = indices 0-16, T:X = indices 17-21
    // Actual book array: [B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,'',T,U,V,W,X]
    // idx:               [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22]
    const bToR = book.slice(0, 17);  // B through R (skip S placeholder)
    const tToX = book.slice(18);     // T through X
    vals.push({ range: `${S}!B${row}:R${row}`, values: [bToR] });
    vals.push({ range: `${S}!T${row}:X${row}`, values: [tToX] });
  });

  await valuesBatchUpdate(id, vals, '03-library values');
  console.log(`✅  Master Book Library done — ${BOOKS.length} books loaded.`);
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });
