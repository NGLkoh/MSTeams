// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const config = {
  appId: 'd1cb90c1-529b-42d6-b19f-5b403cc47bce',
  redirectUri: 'http://localhost:3000/api/callback',
  scopes: [
    'User.Read',
    'MailboxSettings.Read', 
    'Calendars.ReadWrite',
    'OnlineMeetings.ReadWrite'
  ]
};

export default config;