'use strict';
const { batchUpdate, valuesBatchUpdate, gridRange, hex, C } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const RF = sheetMap['Reference Data'];
const S  = "'Reference Data'";

// Named ranges will be created so other tabs can use DATA VALIDATION dropdown lists.
// All lists are in columns A-T, each list occupies one column, header in row 1.

const lists = {
  // Col A: RSVP Status
  RSVPStatus: ['Attending', 'Not Attending', 'Pending', 'No Response'],
  // Col B: Meal Preference
  MealPref: ['Chicken', 'Fish', 'Beef', 'Vegetarian', 'Vegan', 'Kids Meal', 'TBD'],
  // Col C: Dietary Restriction tags
  DietaryTags: ['Gluten-Free', 'Nut Allergy', 'Dairy-Free', 'Shellfish Allergy', 'Soy-Free', 'Egg-Free', 'Halal', 'Kosher', 'Low-Sodium', 'Diabetic', 'None'],
  // Col D: Guest Relationship
  Relationship: ["Bride's Family", "Groom's Family", 'Mutual Friend', "Bride's Friend", "Groom's Friend", 'Colleague', 'Neighbor', 'Other'],
  // Col E: Guest Group
  GuestGroup: ['Immediate Family', 'Extended Family', 'Wedding Party', 'Friends', 'Colleagues', 'Neighbors', 'Children', 'Other'],
  // Col F: Side
  Side: ['Bride', 'Groom', 'Both'],
  // Col G: Table Shape
  TableShape: ['Round', 'Rectangle', 'Square', 'Oval', 'Banquet'],
  // Col H: Table Zone
  TableZone: ['Head Table', 'Family Front', 'Center', 'Window', 'Garden', 'Balcony', 'Corner', 'Near Dance Floor', 'Near Bar', 'Other'],
  // Col I: Vendor Category
  VendorCategory: ['Caterer', 'Photographer', 'Videographer', 'Florist', 'DJ / Band', 'Hair & Makeup', 'Officiant', 'Transportation', 'Cake / Bakery', 'Photo Booth', 'Lighting', 'Stationery', 'Rentals', 'Coordinator / Planner', 'Ceremony Venue', 'Reception Venue', 'Hotel Room Block', 'Honeymoon', 'Other'],
  // Col J: Booking Status
  BookingStatus: ['Booked', 'Deposit Paid', 'Inquired', 'Quoted', 'Declined', 'Backup Option', 'Cancelled'],
  // Col K: Payment Method
  PaymentMethod: ['Credit Card', 'Debit Card', 'Bank Transfer', 'Check', 'Cash', 'Venmo', 'Zelle', 'PayPal', 'Other'],
  // Col L: Task Priority
  TaskPriority: ['Critical', 'High', 'Medium', 'Low'],
  // Col M: Task Status
  TaskStatus: ['Not Started', 'In Progress', 'Completed', 'Blocked', 'Delegated', 'N/A'],
  // Col N: Planning Phase
  PlanningPhase: ['1 - Early Planning (12+ months)', '2 - Venue & Date (10-12 months)', '3 - Key Vendors (8-10 months)', '4 - Guest & Invitations (6-8 months)', '5 - Details & Decor (4-6 months)', '6 - Attire & Beauty (3-4 months)', '7 - Final Confirmations (1-3 months)', '8 - Last Month (4 weeks out)', '9 - Final Week', '10 - Wedding Day'],
  // Col O: Budget Category
  BudgetCategory: ['Venue', 'Catering', 'Photography', 'Videography', 'Florals', 'Music / Entertainment', 'Hair & Makeup', 'Attire & Accessories', 'Invitations & Stationery', 'Transportation', 'Cake & Desserts', 'Officiant', 'Decorations & Rentals', 'Favors & Gifts', 'Honeymoon', 'Accommodations', 'Miscellaneous'],
  // Col P: Reception Segment Type
  ReceptionSegment: ['Cocktail Hour', 'Guest Arrival', 'Bridal Party Entrance', 'Couple Entrance', 'Welcome & Toasts', 'First Dance', 'Parent Dance', 'Wedding Party Dance', 'Dinner Service', 'Cake Cutting', 'Bouquet Toss', 'Garter Toss', 'Open Dancing', 'Special Performance', 'Departure / Send-Off', 'Photo Session', 'Ceremony', 'Rehearsal Dinner', 'Brunch', 'Other'],
  // Col Q: Drink Type
  DrinkType: ['Signature Cocktail', 'Beer', 'Wine (Red)', 'Wine (White)', 'Wine (Rosé)', 'Sparkling / Champagne', 'Mocktail', 'Soft Drink', 'Juice', 'Coffee / Tea', 'Water', 'Other'],
  // Col R: Cake Tier/Flavor
  CakeFlavor: ['Vanilla', 'Chocolate', 'Red Velvet', 'Lemon', 'Carrot', 'Almond', 'Champagne', 'Funfetti', 'Marble', 'Custom'],
  // Col S: Delivery Status
  DeliveryStatus: ['Mailed', 'Hand Delivered', 'Returned Undelivered', 'Lost', 'Not Yet Sent'],
  // Col T: Seating Confirmation
  SeatingStatus: ['Confirmed', 'Draft', 'Needs Review', 'TBD'],
};

(async () => {
  const data = [];
  const reqs = [];

  const entries = Object.entries(lists);
  entries.forEach(([name, items], colIdx) => {
    // Header
    data.push({ range: `${S}!${colLetter(colIdx)}1`, values: [[name]] });
    // Values
    items.forEach((item, rowIdx) => {
      data.push({ range: `${S}!${colLetter(colIdx)}${rowIdx + 2}`, values: [[item]] });
    });
  });

  // Style: header row bold + primary background
  reqs.push({ repeatCell: {
    range: gridRange(RF, 0, 1, 0, entries.length),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.primary),
      textFormat: { bold: true, foregroundColor: hex(C.white), fontSize: 10 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});

  // Freeze header row
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: RF, gridProperties: { frozenRowCount: 1 } },
    fields: 'gridProperties.frozenRowCount',
  }});

  // Hide Reference Data tab
  reqs.push({ updateSheetProperties: {
    properties: { sheetId: RF, hidden: true },
    fields: 'hidden',
  }});

  await valuesBatchUpdate(id, data, 'ref-data');
  await batchUpdate(id, reqs, 'ref-format');

  // Named ranges for use in data validation
  const namedRangeReqs = entries.map(([name, items], colIdx) => ({
    addNamedRange: {
      namedRange: {
        name: `REF_${name.replace(/[^A-Za-z0-9]/g, '_')}`,
        range: gridRange(RF, 1, items.length + 1, colIdx, colIdx + 1),
      },
    },
  }));

  await batchUpdate(id, namedRangeReqs, 'named-ranges');

  console.log('✓ Reference Data tab built with', entries.length, 'lookup lists');
})().catch(e => { console.error(e.errors || e.message || e); process.exit(1); });

function colLetter(n) {
  let s = '';
  n += 1;
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}
