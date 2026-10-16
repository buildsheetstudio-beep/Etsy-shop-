'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// Palette: Deep Ocean Blue-Teal #1B5E7A + Vibrant Orange #E07030 + Clean Off-White #F6F8F7
// Distinctness: Blue-teal + orange combo unused in catalog; prior workout planner used cobalt/lime
const C = {
  primary:      '#1B5E7A',  // Deep Ocean Blue-Teal
  secondary:    '#E07030',  // Vibrant Orange
  accent:       '#3A7D5C',  // Forest Green
  bg:           '#F6F8F7',  // Clean Off-White
  white:        '#FFFFFF',
  primaryTint:  '#E0EDF3',  // Light blue-teal tint
  secondaryTint:'#FDF0E7',  // Light orange tint
  greenTint:    '#E2EFE8',  // Light green tint
  text:         '#1A2D35',  // Dark Navy
  secText:      '#5D7580',  // Gray-blue
  border:       '#C5D3D9',  // Light border
  success:      '#3A7D5C',  // Forest Green
  warning:      '#D4852A',  // Amber
  attention:    '#8B3A30',  // Muted Rust
  info:         '#3A9BAB',  // Light Teal
  input:        '#EFF6F9',  // Very light teal input
  formula:      '#EDF4F0',  // Light formula cells
  altRow:       '#F0F4F5',  // Alternating row
  restDay:      '#E8EDF0',  // Rest day gray-blue
  planned:      '#E0EDF3',  // Planned info tint
  partial:      '#FDF4E3',  // Partial amber tint
  skipped:      '#EBEBEB',  // Skipped neutral gray
  completed:    '#E2EFE8',  // Completed success tint
  amber:        '#D4852A',
  gray:         '#C5D3D9',
};

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return { red: r, green: g, blue: b };
}

async function getAuth() {
  const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
  const creds = secret.installed || secret.web;
  const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);
  oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
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
