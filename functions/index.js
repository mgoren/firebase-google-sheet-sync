/**
 * Copyright 2017 Google Inc. All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
'use strict';

// trigger function that copies new Firebase data to a Google Sheet

const functions = require('firebase-functions');
const admin = require('firebase-admin');
admin.initializeApp();
const {OAuth2Client} = require('google-auth-library');
const {google} = require('googleapis');

// TODO: Use firebase functions:config:set to configure your googleapi object:
// googleapi.client_id = Google API client ID,
// googleapi.client_secret = client secret, and
// googleapi.sheet_id = Google Sheet id (long string in middle of sheet URL)
const CONFIG_CLIENT_ID = functions.config().googleapi.client_id;
const CONFIG_CLIENT_SECRET = functions.config().googleapi.client_secret;
const CONFIG_SHEET_ID = functions.config().googleapi.sheet_id;

// TODO: Change this if necessary to match your Firebase database path
// watchedpaths.data_path = Firebase path for data to be synced to Google Sheet
const CONFIG_DATA_PATH = '/orders';

// The OAuth Callback Redirect.
const FUNCTIONS_REDIRECT = `https://${process.env.GCLOUD_PROJECT}.firebaseapp.com/oauthcallback`;

// setup for authGoogleAPI
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const functionsOauthClient = new OAuth2Client(CONFIG_CLIENT_ID, CONFIG_CLIENT_SECRET,
  FUNCTIONS_REDIRECT);

// OAuth token cached locally.
let oauthTokens = null;

// visit the URL for this Function to request tokens
exports.authgoogleapi = functions.https.onRequest((req, res) => {
  res.set('Cache-Control', 'private, max-age=0, s-maxage=0');
  res.redirect(functionsOauthClient.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent',
  }));
});

// setup for OauthCallback
const DB_TOKEN_PATH = '/api_tokens';

// after you grant access, you will be redirected to the URL for this Function
// this Function stores the tokens to your Firebase database
exports.oauthcallback = functions.https.onRequest(async (req, res) => {
  res.set('Cache-Control', 'private, max-age=0, s-maxage=0');
  const code = `${req.query.code}`;
  try {
    const { tokens } = await functionsOauthClient.getToken(code);
    // Now tokens contains an access_token and an optional refresh_token. Save them.
    await admin.database().ref(DB_TOKEN_PATH).set(tokens);
    res.status(200).send('App successfully configured with new Credentials. '
        + 'You can now close this page.');
  } catch (error) {
    res.status(400).send(error);
  }
});

// trigger function to write to Sheet when new data comes in on CONFIG_DATA_PATH
exports.appendrecordtospreadsheet = functions.database.ref(`${CONFIG_DATA_PATH}/{ITEM}`).onCreate(
    (snap) => {
      const newRecord = snap.val();
      const createdAt = new Date(newRecord.timestamp).toLocaleDateString();
      const orders = splitOrder(newRecord);
      const promises = orders.map((order) => {
        const { first, last, nametag, email, phone, address, city, state, zip, country, volunteer, share, comments, admissionQuantity, admissionCost, donation, total, deposit, owed, purchaser, paypalEmail } = order;
        // fields must be in the same order as the columns in the spreadsheet
        const fields = {
          first,
          last,
          nametag,
          email,
          phone,
          address,
          city,
          state,
          zip,
          country,
          volunteer: volunteer.join(', '),
          share: share.join(', '),
          comments,
          admissionQuantity,
          admissionCost,
          donation,
          total,
          deposit,
          owed,
          purchaser,
          createdAt,
          paypalEmail
        };
        return appendPromise({
          spreadsheetId: CONFIG_SHEET_ID,
          range: 'A:M',
          valueInputOption: 'USER_ENTERED',
          insertDataOption: 'INSERT_ROWS',
          resource: {
            values: [Object.values(fields)]
          }
        });
      });
      return Promise.all(promises);
    });

// accepts an append request, returns a Promise to append it, enriching it with auth
function appendPromise(requestWithoutAuth) {
  return new Promise((resolve, reject) => {
    return getAuthorizedClient().then((client) => {
      const sheets = google.sheets('v4');
      const request = requestWithoutAuth;
      request.auth = client;
      return sheets.spreadsheets.values.append(request, (err, response) => {
        if (err) {
          functions.logger.log(`The API returned an error: ${err}`);
          return reject(err);
        }
        return resolve(response.data);
      });
    });
  });
}

// checks if oauthTokens have been loaded into memory, and if not, retrieves them
async function getAuthorizedClient() {
  if (oauthTokens) {
    return functionsOauthClient;
  }
  const snapshot = await admin.database().ref(DB_TOKEN_PATH).once('value');
  oauthTokens = snapshot.val();
  functionsOauthClient.setCredentials(oauthTokens);
  return functionsOauthClient;
}
const PERSON_FIELDS = ['first', 'last', 'nametag', 'email', 'phone', 'address', 'city', 'state', 'zip', 'country'];

function splitOrder(order) {
  let orders = [];
  const { volunteer, share, comments, admissionQuantity, admissionCost, donation, total, deposit, paypalEmail } = order;
  const owed = total - deposit;
  const purchaser = order.people[0].first + ' ' + order.people[0].last;
  for (const person of order.people) {
    const {first, last, nametag, email, phone, address, city, state, zip, country} = person;
    if (person.index === 0) {
      orders.push({
        first,
        last,
        nametag,
        email,
        phone,
        address,
        city,
        state,
        zip,
        country,
        volunteer,
        share,
        comments,
        admissionQuantity,
        admissionCost,
        donation,
        total,
        deposit,
        owed,
        paypalEmail
      });
    } else {
      orders.push({
        first,
        last,
        nametag,
        email,
        phone,
        address,
        city,
        state,
        zip,
        country,
        purchaser,
      });
    }
  }
  return orders;
}
