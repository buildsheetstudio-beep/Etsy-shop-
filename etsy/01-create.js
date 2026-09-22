'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

const s = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
const { client_id, client_secret, redirect_uris } = s.installed || s.web;
const auth = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
auth.setCredentials(JSON.parse(fs.readFileSync(path.join(__dirname, 'tokens.json'))));
const sheets = google.sheets({ version: 'v4', auth });

const TABS = [
  { title: 'Reference Data',         color: '#888888' },
  { title: 'Business Setup',         color: '#264F55' },
  { title: 'Etsy CSV Staging',       color: '#AFC4D0' },
  { title: 'Income & Orders',        color: '#8EAD91' },
  { title: 'Expense Log',            color: '#B86F68' },
  { title: 'Monthly Dashboard',      color: '#264F55' },
  { title: 'Quarterly Summary',      color: '#264F55' },
  { title: 'Annual Summary',         color: '#264F55' },
  { title: 'Multi-Year Comparison',  color: '#D4B168' },
  { title: 'Product Profitability',  color: '#8EAD91' },
  { title: 'Etsy Fee Analysis',      color: '#D78758' },
  { title: 'Tax Deduction Summary',  color: '#AFC4D0' },
  { title: 'Dashboard',              color: '#264F55' },
];

(async () => {
  const res = await sheets.spreadsheets.create({
    requestBody: {
      properties: { title: 'Ultimate Etsy Seller Bookkeeping System — Willow Paper Studio' },
      sheets: TABS.map((t, i) => {
        const c = t.color.replace('#', '');
        return {
          properties: {
            sheetId: i,
            index:   i,
            title:   t.title,
            tabColorStyle: { rgbColor: {
              red:   parseInt(c.slice(0,2),16)/255,
              green: parseInt(c.slice(2,4),16)/255,
              blue:  parseInt(c.slice(4,6),16)/255,
            }},
            gridProperties: { rowCount: 500, columnCount: 30 },
          },
        };
      }),
    },
  });

  const id  = res.data.spreadsheetId;
  const url = res.data.spreadsheetUrl;
  const sheetMap = {};
  res.data.sheets.forEach(s => { sheetMap[s.properties.title] = s.properties.sheetId; });

  fs.writeFileSync(path.join(__dirname, 'spreadsheet.json'), JSON.stringify({ id, url, sheetMap }, null, 2));
  console.log('Created:', url);
  console.log('ID:', id);
  console.log('sheetMap:', JSON.stringify(sheetMap));
})().catch(e => { console.error(e.errors || e.message); process.exit(1); });
