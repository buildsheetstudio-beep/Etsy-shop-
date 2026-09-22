'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const CHUNK  = 400;
const VCHUNK = 200;

function getAuth() {
  const s = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const { client_id, client_secret, redirect_uris } = s.installed || s.web;
  const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
  return auth;
}

async function batchUpdate(spreadsheetId, requests) {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });
  for (let i = 0; i < requests.length; i += CHUNK) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: requests.slice(i, i + CHUNK) } });
    if (i + CHUNK < requests.length) await new Promise(r => setTimeout(r, 300));
  }
}

async function valuesBatchUpdate(spreadsheetId, data) {
  const sheets = google.sheets({ version: 'v4', auth: getAuth() });
  for (let i = 0; i < data.length; i += VCHUNK) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId, requestBody: { valueInputOption: 'USER_ENTERED', data: data.slice(i, i + VCHUNK) }
    });
    if (i + VCHUNK < data.length) await new Promise(r => setTimeout(r, 300));
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

function hex(h) {
  return { red: parseInt(h.slice(1,3),16)/255, green: parseInt(h.slice(3,5),16)/255, blue: parseInt(h.slice(5,7),16)/255 };
}

const C = {
  primary:   '#264F55',
  secondary: '#D78758',
  bg:        '#F5F2EC',
  panel:     '#FFFFFF',
  success:   '#8EAD91',
  warning:   '#D4B168',
  attention: '#B86F68',
  mutedBlue: '#AFC4D0',
  mutedGold: '#D4BE85',
  mainText:  '#2C3334',
  secText:   '#697273',
  border:    '#C9C5BC',
  white:     '#FFFFFF',
  input:     '#FEFCE8',
  formula:   '#EBF3F5',
  altRow:    '#EFECE5',
};

const SHOP_TYPES = ['Digital Products','Handmade','Print on Demand','Vintage','Craft Supplies','Mixed Shop','Other'];

const INCOME_SOURCES = ['Etsy Order','Etsy Shipping Collected','Etsy Refund Adjustment','Off-Etsy Sale','Wholesale','Custom Order','Licensing','Affiliate Income','Other Business Income'];

const EXPENSE_CATS = [
  'Etsy Transaction Fees','Etsy Listing Fees','Etsy Payment Processing Fees','Etsy Ads','Offsite Ads',
  'Refunds','Materials','Packaging','Shipping Labels','Postage','Production Partner Costs',
  'Print-on-Demand Costs','Software & Subscriptions','Design Assets','Contractors','Photography',
  'Equipment','Office Supplies','Education & Training','Professional Fees','Banking Fees',
  'Internet & Phone','Mileage & Travel','Insurance','Licenses & Permits','Rent / Workspace',
  'Taxes Paid','Other Business Expense',
];

const FEE_TYPES = [
  'Transaction Fee','Listing Fee','Payment Processing Fee','Etsy Ads','Offsite Ads',
  'Shipping Fee','Regulatory Operating Fee','Currency Conversion Fee','Refund','Credit','Other Etsy Fee',
];

const PRODUCT_TYPES = ['Digital Download','Physical Handmade','Print on Demand','Vintage','Craft Supply','Custom Order','Bundle','Other'];

const TAX_GROUP_MAP = {
  'Etsy Transaction Fees':          'Banking & Payment Fees',
  'Etsy Listing Fees':              'Banking & Payment Fees',
  'Etsy Payment Processing Fees':   'Banking & Payment Fees',
  'Etsy Ads':                       'Advertising & Marketing',
  'Offsite Ads':                    'Advertising & Marketing',
  'Refunds':                        'Other / Review Required',
  'Materials':                      'Cost of Goods Sold / Production',
  'Packaging':                      'Shipping & Fulfillment',
  'Shipping Labels':                'Shipping & Fulfillment',
  'Postage':                        'Shipping & Fulfillment',
  'Production Partner Costs':       'Cost of Goods Sold / Production',
  'Print-on-Demand Costs':          'Cost of Goods Sold / Production',
  'Software & Subscriptions':       'Software & Subscriptions',
  'Design Assets':                  'Cost of Goods Sold / Production',
  'Contractors':                    'Professional Services',
  'Photography':                    'Advertising & Marketing',
  'Equipment':                      'Office & Equipment',
  'Office Supplies':                'Office & Equipment',
  'Education & Training':           'Education',
  'Professional Fees':              'Professional Services',
  'Banking Fees':                   'Banking & Payment Fees',
  'Internet & Phone':               'Office & Equipment',
  'Mileage & Travel':               'Travel & Mileage',
  'Insurance':                      'Insurance & Licenses',
  'Licenses & Permits':             'Insurance & Licenses',
  'Rent / Workspace':               'Workspace',
  'Taxes Paid':                     'Taxes & Government Fees',
  'Other Business Expense':         'Other / Review Required',
};

const TAX_GROUPS = [
  'Advertising & Marketing','Cost of Goods Sold / Production','Shipping & Fulfillment',
  'Software & Subscriptions','Office & Equipment','Professional Services','Education',
  'Banking & Payment Fees','Travel & Mileage','Insurance & Licenses','Workspace',
  'Taxes & Government Fees','Other / Review Required',
];

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

const INCOME_STATUSES  = ['Received','Pending','Refunded','Partially Refunded','Cancelled'];
const EXPENSE_STATUSES = ['Paid','Pending','Partially Paid','Refunded','Cancelled'];
const IMPORT_STATUSES  = ['Ready','Missing Required Field','Duplicate Suspected','Review Needed','Imported','Ignored'];

// Fictional listings for Willow Paper Studio
const LISTINGS = [
  { id:'1001', title:'2026 Printable Annual Planner',          type:'Digital Download', cat:'Planners',    cost:0,    pod:0   },
  { id:'1002', title:'Botanical Wedding Invitation Suite',      type:'Digital Download', cat:'Stationery',  cost:0,    pod:0   },
  { id:'1003', title:'Minimalist Daily Planner Pad',            type:'Print on Demand',  cat:'Planners',    cost:0,    pod:5.50},
  { id:'1004', title:'Watercolor Floral Wall Art Bundle',       type:'Digital Download', cat:'Art Prints',  cost:0,    pod:0   },
  { id:'1005', title:'Custom Business Card Template',           type:'Digital Download', cat:'Stationery',  cost:0,    pod:0   },
  { id:'1006', title:'Self-Care Journal Printable',             type:'Digital Download', cat:'Journals',    cost:0,    pod:0   },
  { id:'1007', title:'Kids Birthday Party Printable Pack',      type:'Digital Download', cat:'Party Decor', cost:0,    pod:0   },
  { id:'1008', title:'Gradient Abstract Art Print',             type:'Print on Demand',  cat:'Art Prints',  cost:0,    pod:8.00},
  { id:'1009', title:'Recipe Card Template Set',                type:'Digital Download', cat:'Kitchen',     cost:0,    pod:0   },
  { id:'1010', title:'Yoga & Wellness Planner',                 type:'Digital Download', cat:'Planners',    cost:0,    pod:0   },
  { id:'1011', title:'Budget Tracker Spreadsheet',              type:'Digital Download', cat:'Finance',     cost:0,    pod:0   },
  { id:'1012', title:'Mushroom Forest Watercolor Print',        type:'Digital Download', cat:'Art Prints',  cost:0,    pod:0   },
  { id:'1013', title:'Baby Shower Games Bundle',                type:'Digital Download', cat:'Party Decor', cost:0,    pod:0   },
  { id:'1014', title:'Custom Wedding Day Itinerary',            type:'Custom Order',     cat:'Stationery',  cost:2.50, pod:0   },
  { id:'1015', title:'90-Day Goal Setting Workbook',            type:'Digital Download', cat:'Journals',    cost:0,    pod:0   },
  { id:'1016', title:'Tropical Botanical Print Trio',           type:'Print on Demand',  cat:'Art Prints',  cost:0,    pod:11.00},
  { id:'1017', title:'Weekly Meal Planning Bundle',             type:'Digital Download', cat:'Kitchen',     cost:0,    pod:0   },
  { id:'1018', title:'Boho Nursery Wall Art Set',               type:'Digital Download', cat:'Art Prints',  cost:0,    pod:0   },
  { id:'1019', title:'Teacher Planner Bundle 2026',             type:'Digital Download', cat:'Education',   cost:0,    pod:0   },
  { id:'1020', title:'Gratitude Journal 365 Prompts',           type:'Digital Download', cat:'Journals',    cost:0,    pod:0   },
];

module.exports = {
  batchUpdate, valuesBatchUpdate, gridRange, hex, C,
  SHOP_TYPES, INCOME_SOURCES, EXPENSE_CATS, FEE_TYPES, PRODUCT_TYPES,
  TAX_GROUP_MAP, TAX_GROUPS, MONTHS, INCOME_STATUSES, EXPENSE_STATUSES, IMPORT_STATUSES,
  LISTINGS,
};
