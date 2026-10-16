'use strict';
const { google } = require('googleapis');
const fs = require('fs');

const CHUNK  = 400;
const VCHUNK = 200;

function getAuth() {
  const s = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
  const { client_id, client_secret, redirect_uris } = s.installed;
  const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
  return auth;
}

async function batchUpdate(spreadsheetId, requests, label) {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < requests.length; i += CHUNK) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId, requestBody: { requests: requests.slice(i, i + CHUNK) }
    });
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < data.length; i += VCHUNK) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: data.slice(i, i + VCHUNK) }
    });
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

function hex(h) {
  return { red: parseInt(h.slice(1,3),16)/255, green: parseInt(h.slice(3,5),16)/255, blue: parseInt(h.slice(5,7),16)/255 };
}

// Farm palette
const C = {
  primary:   '#315A6B',  // Deep Field Blue
  secondary: '#B86F3C',  // Harvest Copper
  bg:        '#F7F2E8',  // Warm Oat
  panel:     '#FFFDFC',  // Light Panel
  success:   '#8FAF88',  // sage green
  warning:   '#D8B267',  // amber
  attention: '#B87362',  // muted rust
  mutedBlue: '#AFC7D2',
  mainText:  '#2D3436',
  secText:   '#66706F',
  border:    '#C9C5BA',
  white:     '#FFFFFF',
  input:     '#FEFBE8',  // pale warm yellow for input cells
  formula:   '#EBF2F5',  // pale blue-gray for formula cells
  // Row alternates
  altRow:    '#F0EDE6',
};

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

const ENTERPRISES = ['Crops','Livestock','Dairy','Poultry','Produce','Orchard','Greenhouse','Farm Stand','Custom Farm Work','Agritourism','Hay & Forage','Apiary','Nursery','Other Farm Income','General Farm'];

const INCOME_CATS = ['Crop Sales','Livestock Sales','Dairy Sales','Egg & Poultry Sales','Produce Sales','Farm Stand Sales','Hay & Forage Sales','Nursery & Plant Sales','Custom Farm Services','Agritourism Income','Government Program Payments','Insurance Proceeds','Equipment Rental Income','Breeding & Stud Fees','Value-Added Product Sales','Other Farm Income','Refunds & Rebates'];

const EXPENSE_CATS = ['Seed & Plants','Feed','Fertilizer','Soil Amendments','Chemicals & Pest Control','Veterinary','Livestock Purchases','Breeding','Fuel','Equipment Repairs','Small Tools','Equipment Rental','Irrigation','Utilities','Labor','Contract Services','Insurance','Property Taxes','Rent & Lease','Loan Interest','Supplies','Packaging','Market & Selling Fees','Transportation','Professional Fees','Office & Software','Education & Training','Licenses & Permits','Repairs & Maintenance','Other Farm Expense'];

const TAX_GROUPS = {
  'Seed & Plants':           'Cost of Goods / Production Inputs',
  'Feed':                    'Livestock & Animal Care',
  'Fertilizer':              'Cost of Goods / Production Inputs',
  'Soil Amendments':         'Cost of Goods / Production Inputs',
  'Chemicals & Pest Control':'Cost of Goods / Production Inputs',
  'Veterinary':              'Livestock & Animal Care',
  'Livestock Purchases':     'Livestock & Animal Care',
  'Breeding':                'Livestock & Animal Care',
  'Fuel':                    'Vehicle & Transportation',
  'Equipment Repairs':       'Equipment & Repairs',
  'Small Tools':             'Equipment & Repairs',
  'Equipment Rental':        'Equipment & Repairs',
  'Irrigation':              'Land & Facilities',
  'Utilities':               'Utilities',
  'Labor':                   'Labor & Contract Services',
  'Contract Services':       'Labor & Contract Services',
  'Insurance':               'Insurance & Professional',
  'Property Taxes':          'Land & Facilities',
  'Rent & Lease':            'Land & Facilities',
  'Loan Interest':           'Land & Facilities',
  'Supplies':                'Selling & Administrative',
  'Packaging':               'Selling & Administrative',
  'Market & Selling Fees':   'Selling & Administrative',
  'Transportation':          'Vehicle & Transportation',
  'Professional Fees':       'Insurance & Professional',
  'Office & Software':       'Selling & Administrative',
  'Education & Training':    'Selling & Administrative',
  'Licenses & Permits':      'Insurance & Professional',
  'Repairs & Maintenance':   'Equipment & Repairs',
  'Other Farm Expense':      'Other / Review Required',
};

const PAYMENT_METHODS = ['Cash','Checking Account','Savings Account','Credit Card','Debit Card','Check','Digital Payment','Farm Account'];

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C, MONTHS, ENTERPRISES, INCOME_CATS, EXPENSE_CATS, TAX_GROUPS, PAYMENT_METHODS };
