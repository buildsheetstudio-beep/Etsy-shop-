'use strict';
const { google } = require('googleapis');
const fs = require('fs');

const CHUNK  = 400;
const VCHUNK = 200;

function getAuth() {
  const secret = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
  const { client_id, client_secret, redirect_uris } = secret.installed;
  const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  auth.setCredentials(JSON.parse(fs.readFileSync(__dirname + '/tokens.json')));
  return auth;
}

async function batchUpdate(spreadsheetId, requests, label) {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });
  for (let i = 0; i < requests.length; i += CHUNK) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: requests.slice(i, i + CHUNK) }
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
  return {
    red:   parseInt(h.slice(1,3),16)/255,
    green: parseInt(h.slice(3,5),16)/255,
    blue:  parseInt(h.slice(5,7),16)/255,
  };
}

const C = {
  bg:        '#FAFAF8',
  mainText:  '#2B2B2B',
  secondary: '#666666',
  border:    '#A9A9A9',
  headerGray:'#EDEDEA',
  completed: '#DCEBD8',
  completedText: '#6F786D',
  today:     '#E5F1F8',
  weekend:   '#F5EFE6',
  notes:     '#FFFDF7',
  footer:    '#F3E5EA',
  white:     '#FFFFFF',
  black:     '#1F1F1F',
  // 6 pastel member colors
  member: ['#F4C7AB','#E8B7C5','#BFD7EA','#D7C6E8','#C7E5D1','#F3E4A6'],
};

module.exports = { batchUpdate, valuesBatchUpdate, gridRange, hex, C };
