import * as msal from '@azure/msal-node';
import dotenv from 'dotenv';

dotenv.config();

// MSAL configuration for Microsoft OAuth
const msalConfig = {
  auth: {
    clientId: process.env.MICROSOFT_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID || 'common'}`,
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel, message, containsPii) {
        if (!containsPii) {
          console.log(message);
        }
      },
      piiLoggingEnabled: false,
      logLevel: msal.LogLevel.Warning,
    }
  }
};

// Initialize the MSAL confidential client application
const confidentialClientApplication = new msal.ConfidentialClientApplication(msalConfig);

// Microsoft Graph API scopes for Teams access
// Based on available permissions in Azure AD
const MICROSOFT_SCOPES = [
  // ===== OpenID / Identity =====
  'openid',                       // Sign users in
  'profile',                      // Basic profile
  'email',                        // Email address
  'offline_access',               // Refresh token support

  // ===== User =====
  'https://graph.microsoft.com/User.Read', // Sign in and read user profile

  // ===== Calendars =====
  'https://graph.microsoft.com/Calendars.Read',              // Read user calendars
  'https://graph.microsoft.com/Calendars.Read.Shared',       // Read shared calendars
  'https://graph.microsoft.com/Calendars.ReadBasic',         // Read basic calendar details
  'https://graph.microsoft.com/Calendars.ReadWrite',         // Full access to user calendars
  'https://graph.microsoft.com/Calendars.ReadWrite.Shared',  // Read/write shared calendars

  // ===== Teams & Channels =====
  'https://graph.microsoft.com/Team.ReadBasic.All',           // Read team names & descriptions
  'https://graph.microsoft.com/Channel.ReadBasic.All',        // Read channel names & descriptions
  'https://graph.microsoft.com/ChannelMessage.Read.All',      // Read channel messages (Admin consent)
  'https://graph.microsoft.com/ChannelMessage.Send',          // Send channel messages

  // ===== Chats =====
  'https://graph.microsoft.com/Chat.Create',                  // Create chats
  'https://graph.microsoft.com/Chat.Read',                    // Read chat messages
  'https://graph.microsoft.com/Chat.ReadBasic',               // Read chat names & members
  'https://graph.microsoft.com/Chat.ReadWrite',               // Read/write chat messages
];


// OAuth configuration
const microsoftOAuthClient = {
  clientId: process.env.MICROSOFT_CLIENT_ID,
  clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
  redirectUri: process.env.MICROSOFT_REDIRECT_URI,
  authority: `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID || 'common'}`,
  scopes: MICROSOFT_SCOPES
};

export { confidentialClientApplication, MICROSOFT_SCOPES, microsoftOAuthClient };
