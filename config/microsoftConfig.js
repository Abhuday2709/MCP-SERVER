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
// Using ONLY permissions that don't require admin consent
const MICROSOFT_SCOPES = [
  'https://graph.microsoft.com/User.Read',           // ✓ No admin consent needed
  'https://graph.microsoft.com/Chat.Read',           // ✓ No admin consent needed
  'https://graph.microsoft.com/Chat.ReadWrite',      // ✓ No admin consent needed
  'https://graph.microsoft.com/ChatMessage.Read',    // ✓ No admin consent needed
  'https://graph.microsoft.com/ChatMessage.Send',    // ✓ No admin consent needed
  'offline_access',                                   // ✓ No admin consent needed
  'openid',                                           // ✓ No admin consent needed
  'profile',                                          // ✓ No admin consent needed
  'email'                                             // ✓ No admin consent needed
  
  // Removed permissions that require admin consent:
  // - Team.ReadBasic.All (requires admin)
  // - Channel.ReadBasic.All (requires admin)
  // - ChannelMessage.Read.All (requires admin)
  // - ChannelMessage.Send (requires admin)
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
