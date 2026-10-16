'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const C = {
  primary:    '#6E2B3F',  // Deep Wine
  secondary:  '#B8965A',  // Aged Gold
  accent:     '#4A7C59',  // Forest Green
  bg:         '#FAF8F3',  // Cream Paper
  white:      '#FFFFFF',
  wineTint:   '#F4E8EC',  // Light Wine Tint
  goldTint:   '#F7EDD8',  // Light Gold Tint
  greenTint:  '#E3EDE7',  // Light Green Tint
  text:       '#2C2822',  // Dark Ink
  secText:    '#7A6E65',  // Faded Ink
  border:     '#D5CAC0',  // Parchment Edge
  success:    '#4A7C59',  // Forest Green
  warning:    '#C17B2E',  // Amber
  attention:  '#8B3A3A',  // Deep Red
  info:       '#4A6E8C',  // Steel Blue
  input:      '#FFFBF0',  // Warm White — input cells
  formula:    '#EDF3F0',  // Pale green — formula cells
  altRow:     '#F5F2EE',  // Slightly darker cream
  hdrDark:    '#6E2B3F',  // = primary
  hdrLight:   '#F4E8EC',  // = wineTint
  amber:      '#C17B2E',
  gray:       '#D5CAC0',
};

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red: r, green: g, blue: b };
}

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const { client_id, client_secret, redirect_uris } = secret.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  const token = JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json')));
  oAuth2Client.setCredentials(token);
  return oAuth2Client;
}

async function getSheets() {
  const auth = await getAuth();
  return google.sheets({ version: 'v4', auth });
}

const CHUNK = 400;
async function batchUpdate(spreadsheetId, requests, label) {
  const sheets = await getSheets();
  for (let i = 0; i < requests.length; i += CHUNK) {
    const slice = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: slice } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}`);
  }
}

const VCHUNK = 200;
async function valuesBatchUpdate(spreadsheetId, data, label) {
  const sheets = await getSheets();
  for (let i = 0; i < data.length; i += VCHUNK) {
    const slice = data.slice(i, i + VCHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: slice },
    });
    if (data.length > VCHUNK) console.log(`  ${label} value chunk ${Math.floor(i/VCHUNK)+1}`);
  }
}

function gridRange(sheetId, r1, r2, c1, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2, startColumnIndex: c1, endColumnIndex: c2 };
}

module.exports = { C, hex, getAuth, getSheets, batchUpdate, valuesBatchUpdate, gridRange };
