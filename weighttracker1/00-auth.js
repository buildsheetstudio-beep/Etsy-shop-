'use strict';
// Run: node 00-auth.js
// Opens a browser URL — paste the code to get fresh tokens.json
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const secret = JSON.parse(fs.readFileSync(path.join(__dirname, 'client_secret.json')));
const creds = secret.installed || secret.web;
const oAuth2Client = new google.auth.OAuth2(creds.client_id, creds.client_secret, creds.redirect_uris[0]);

const url = oAuth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: 'consent' });
console.log('\nOpen this URL in your browser:\n');
console.log(url);
console.log('\nAfter authorizing, paste the code here:');

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
rl.question('Code: ', async (code) => {
  rl.close();
  try {
    const { tokens } = await oAuth2Client.getToken(code.trim());
    fs.writeFileSync(path.join(__dirname, 'tokens.json'), JSON.stringify(tokens, null, 2));
    console.log('\n✅  tokens.json saved. Run 01-create.js next.');
  } catch (e) {
    console.error('Error getting tokens:', e.message);
    process.exit(1);
  }
});
