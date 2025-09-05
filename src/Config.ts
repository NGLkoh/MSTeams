
const config = {
  appId: 'e01235ac-3f91-4685-8653-1ca3dae4b93f',
  redirectUri: 'http://localhost:3000/api/callback',
  scopes: [
    'User.Read',
    'MailboxSettings.Read', 
    'Calendars.ReadWrite',
    'OnlineMeetings.ReadWrite'
  ]
};

export default config;