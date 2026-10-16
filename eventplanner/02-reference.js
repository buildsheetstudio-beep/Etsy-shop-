'use strict';
const { batchUpdate, valuesBatchUpdate, hex, C, gridRange } = require('./lib');
const fs = require('fs');
const { id, sheetMap } = JSON.parse(fs.readFileSync(__dirname + '/spreadsheet.json'));
const REF = sheetMap['📋 Reference Data'];

// Column layout:
// A: Event Category    B: Event Type       C: Event Status     D: Event Priority
// E: Task Category     F: Task Status      G: Task Priority    H: Vendor Category
// I: Payment Status    J: RSVP Status      K: Dietary Req      L: Ticket Type
// M: Segment Type      N: Budget Category  O: Note Type        P: Contact Role
// Q: Satisfaction      R: Attendance Status

const cols = {
  eventCategory:   ['Corporate', 'Social', 'Wedding', 'Conference', 'Workshop', 'Fundraiser', 'Birthday', 'Anniversary', 'Holiday Party', 'Product Launch', 'Team Building', 'Other'],
  eventType:       ['In-Person', 'Virtual', 'Hybrid'],
  eventStatus:     ['Planning', 'Confirmed', 'In Progress', 'Completed', 'Cancelled', 'Postponed'],
  eventPriority:   ['High', 'Medium', 'Low'],
  taskCategory:    ['Venue', 'Catering', 'Audio/Visual', 'Decor', 'Marketing', 'Logistics', 'Guest Management', 'Vendor', 'Budget', 'Other'],
  taskStatus:      ['Not Started', 'In Progress', 'Done', 'Blocked', 'Deferred'],
  taskPriority:    ['High', 'Medium', 'Low'],
  vendorCategory:  ['Venue', 'Catering', 'Photography', 'Videography', 'Florist', 'DJ/Band', 'Décor', 'Lighting', 'AV/Tech', 'Staffing', 'Transport', 'Printing', 'Security', 'Other'],
  paymentStatus:   ['Not Paid', 'Deposit Paid', 'Partially Paid', 'Paid in Full', 'Overdue', 'Refunded'],
  rsvpStatus:      ['Pending', 'Confirmed', 'Declined', 'Waitlisted', 'No Show'],
  dietaryReq:      ['None', 'Vegetarian', 'Vegan', 'Gluten-Free', 'Halal', 'Kosher', 'Dairy-Free', 'Nut Allergy', 'Other'],
  ticketType:      ['General Admission', 'VIP', 'Early Bird', 'Speaker', 'Sponsor', 'Complimentary', 'Staff'],
  segmentType:     ['Setup', 'Registration', 'Opening', 'Keynote', 'Panel', 'Workshop', 'Break', 'Meal', 'Networking', 'Award', 'Entertainment', 'Closing', 'Teardown', 'Other'],
  budgetCategory:  ['Venue', 'Catering', 'AV/Tech', 'Décor', 'Marketing', 'Staffing', 'Transport', 'Accommodation', 'Entertainment', 'Printing', 'Contingency', 'Other'],
  noteType:        ['General', 'Action Item', 'Decision', 'Risk', 'Lesson Learned', 'Idea', 'Follow-Up'],
  contactRole:     ['Organiser', 'Co-Organiser', 'Venue Contact', 'Caterer', 'AV Contact', 'Sponsor', 'Speaker', 'Volunteer Lead', 'Emergency Contact', 'Other'],
  satisfaction:    ['1', '2', '3', '4', '5'],
  attendanceStatus:['Attended', 'No Show', 'Partial'],
};

(async () => {
  const reqs = [];

  // Header row bg + bold
  reqs.push({ repeatCell: {
    range: gridRange(REF, 0, 1, 0, 22),
    cell: { userEnteredFormat: {
      backgroundColor: hex(C.deepBlush),
      textFormat: { foregroundColor: hex(C.white), bold: true, fontSize: 10 },
      horizontalAlignment: 'CENTER',
    }},
    fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
  }});

  // Column widths — 130px each
  for (let c = 0; c < 18; c++) {
    reqs.push({ updateDimensionProperties: {
      range: { sheetId: REF, dimension: 'COLUMNS', startIndex: c, endIndex: c+1 },
      properties: { pixelSize: 130 },
      fields: 'pixelSize',
    }});
  }

  await batchUpdate(id, reqs, 'ref-format');

  // Values
  const headers = [
    'Event Category','Event Type','Event Status','Event Priority',
    'Task Category','Task Status','Task Priority','Vendor Category',
    'Payment Status','RSVP Status','Dietary Req','Ticket Type',
    'Segment Type','Budget Category','Note Type','Contact Role',
    'Satisfaction','Attendance Status',
  ];

  const data = [];
  data.push({ range: `'📋 Reference Data'!A1:R1`, values: [headers] });

  const colKeys = Object.keys(cols);
  colKeys.forEach((key, ci) => {
    const colL = String.fromCharCode(65 + ci);
    const vals = cols[key];
    data.push({ range: `'📋 Reference Data'!${colL}2:${colL}${vals.length+1}`, values: vals.map(v => [v]) });
  });

  await valuesBatchUpdate(id, data, 'ref-values');
  console.log('Reference Data complete');
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
