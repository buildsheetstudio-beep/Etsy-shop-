'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// Palette: Deep Cinnamon #7D4B2A + Cool Slate Teal #4A7A8C + Warm Off-White #F7F4F0
// Distinctness: warm-earth primary (first warm brown/sienna primary in catalog)
const C = {
  primary:      '#7D4B2A',  // Deep Cinnamon
  secondary:    '#4A7A8C',  // Cool Slate Teal
  accent:       '#8C6A4A',  // Warm Tan
  bg:           '#F7F4F0',  // Warm Off-White
  white:        '#FFFFFF',
  primaryTint:  '#F2E8E2',  // Light cinnamon tint
  secondaryTint:'#E4EEF1',  // Light teal tint
  warmTint:     '#F5EFE8',  // Warm cream
  text:         '#2D1E14',  // Deep warm brown
  secText:      '#6B5A4E',  // Warm medium brown
  border:       '#D6C8BF',  // Light warm border
  success:      '#5B7A4A',  // Muted olive (achievement, not directional)
  warning:      '#C4892A',  // Amber-gold
  attention:    '#8B3A30',  // Muted rust
  info:         '#4A7A8C',  // Slate Teal
  input:        '#FAF7F4',  // Very warm light input
  formula:      '#EEE9E4',  // Warm formula cells
  altRow:       '#F4F0EC',  // Alternating row
  neutral:      '#E8E2DC',  // Neutral warm gray
  muted:        '#D4C8BC',  // Muted warm
  achieved:     '#E6EEE0',  // Soft muted sage (achieved goals)
  active:       '#E4EEF1',  // Active teal tint
  planned:      '#EDF3F5',  // Upcoming
  partial:      '#FDF4E3',  // Partial amber
  paused:       '#EFEFEF',  // Paused neutral
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
