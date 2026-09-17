'use strict';
const { google } = require('googleapis');
const fs = require('fs');
const readline = require('readline');
const path = require('path');

const SECRET = path.join(__dirname, 'client_secret.json');
const TOKEN  = path.join(__dirname, 'tokens.json');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

const { client_id, client_secret, redirect_uris } = JSON.parse(fs.readFileSync(SECRET)).installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

const authUrl = oAuth2Client.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: 'consent' });
console.log('\n=== AUTHORIZATION REQUIRED ===');
console.log('Visit this URL in your browser:\n');
console.log(authUrl);
console.log('\nThen paste the authorization code below:\n');

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
rl.question('Authorization code: ', async (code) => {
  rl.close();
  try {
    const { tokens } = await oAuth2Client.getToken(code.trim());
    fs.writeFileSync(TOKEN, JSON.stringify(tokens, null, 2));
    console.log('\n✅ tokens.json saved successfully! Run 01-create.js now.');
  } catch (e) {
    console.error('Error getting tokens:', e.message);
  }
});
