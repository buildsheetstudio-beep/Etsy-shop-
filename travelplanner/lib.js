const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const CREDS = JSON.parse(fs.readFileSync(__dirname + '/client_secret.json'));
const TOKENS = JSON.parse(fs.readFileSync(__dirname + '/tokens.json'));

const auth = new google.auth.OAuth2(
  CREDS.installed.client_id,
  CREDS.installed.client_secret,
  CREDS.installed.redirect_uris[0],
);
auth.setCredentials(TOKENS);

const sheets = google.sheets({ version: 'v4', auth });

function hex(h) {
  h = h.replace('#', '');
  return { red: parseInt(h.slice(0,2),16)/255, green: parseInt(h.slice(2,4),16)/255, blue: parseInt(h.slice(4,6),16)/255 };
}

const COLOR = {
  deepNavy: '#0D1B3E',
  midnightBlue: '#1A2F5A',
  antiqueGold: '#C9A84C',
  champagne: '#F5EDD5',
  lightGold: '#F0E0A0',
  forestGreen: '#2D6A4F',
  lightGreen: '#D1F0E3',
  amber: '#B5830A',
  lightAmber: '#FFF3CD',
  deepCrimson: '#8B1A1A',
  lightRed: '#FFE4E4',
  nearBlack: '#1C1C1E',
  white: '#FFFFFF',
  softGrey: '#F0F0F0',
  midGrey: '#94A3B8',
  lightBlue: '#E8F0F8',
  inputBg: '#FFFDF5',
  formulaBg: '#F0F4FA',
  border: '#C9A84C',
  subheaderBg: '#1A2F5A',
  altRow: '#F8F6F0',
  paleNavy: '#E8EDF5',
  gold20: '#F5EDD5',
};

function colLetter(n) {
  let s = '';
  n++;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

function gridRange(sheetId, r0, r1, c0, c1) {
  return { sheetId, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 };
}

const CHUNK = 400;
const VAL_CHUNK = 200;

async function batchUpdate(spreadsheetId, requests, label) {
  for (let i = 0; i < requests.length; i += CHUNK) {
    const chunk = requests.slice(i, i + CHUNK);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: chunk } });
    if (requests.length > CHUNK) console.log(`  ${label} chunk ${Math.floor(i/CHUNK)+1}/${Math.ceil(requests.length/CHUNK)}`);
  }
}

async function valuesBatchUpdate(spreadsheetId, data, label) {
  for (let i = 0; i < data.length; i += VAL_CHUNK) {
    const chunk = data.slice(i, i + VAL_CHUNK);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: chunk },
    });
    if (data.length > VAL_CHUNK) console.log(`  ${label} chunk ${Math.floor(i/VAL_CHUNK)+1}/${Math.ceil(data.length/VAL_CHUNK)}`);
  }
}

function colFormulaRange(sheetId, startRow, endRow, col, formula) {
  const reqs = [];
  for (let r = startRow; r < endRow; r++) {
    reqs.push({
      updateCells: {
        range: gridRange(sheetId, r, r+1, col, col+1),
        rows: [{ values: [{ userEnteredValue: { formulaValue: formula.replace(/ROW_OFFSET/g, String(r)) } }] }],
        fields: 'userEnteredValue',
      },
    });
  }
  return reqs;
}

function vr(range, values) { return { range, values }; }

async function clearConditionalFormats(spreadsheetId, sheetId) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets(properties(sheetId),conditionalFormats)' });
  const sheet = meta.data.sheets.find(s => s.properties.sheetId === sheetId);
  const rules = (sheet && sheet.conditionalFormats) || [];
  if (!rules.length) return;
  await batchUpdate(spreadsheetId, [{ deleteConditionalFormatRule: { sheetId, index: 0 } }].concat(
    rules.slice(1).map(() => ({ deleteConditionalFormatRule: { sheetId, index: 0 } }))
  ), `clear-cf-${sheetId}`);
}

module.exports = { sheets, auth, hex, COLOR, colLetter, gridRange, batchUpdate, valuesBatchUpdate, colFormulaRange, vr, clearConditionalFormats };
