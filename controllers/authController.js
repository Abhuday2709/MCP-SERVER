import { oauth2Client, SCOPES as GOOGLE_SCOPES } from '../config/googleConfig.js';
import { google } from 'googleapis';
import { generateToken, setTokenCookie, clearTokenCookie, verifyToken } from '../utils/jwt.js';
import redis from '../config/redisClient.js';
import { confidentialClientApplication, MICROSOFT_SCOPES } from '../config/microsoftConfig.js';
import { Client } from '@microsoft/microsoft-graph-client';

// Redis key prefix for user tokens
const USER_TOKEN_PREFIX = 'user:token:';

// Google Authentication
export const getGoogleAuthUrl = (req, res) => {
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: GOOGLE_SCOPES,
    prompt: 'consent'
  });
  res.json({ authUrl });
};

export const googleCallback = async (req, res) => {
  const { code } = req.query;
  
  try {
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);
    
    // Get user info
    const oauth2 = google.oauth2({ version: 'v2', auth: oauth2Client });
    const { data } = await oauth2.userinfo.get();
    
    // Create JWT payload
    const payload = {
      userId: data.id,
      email: data.email,
      name: data.name,
      picture: data.picture,
      provider: 'google'
    };
    
    // Generate JWT
    const jwtToken = generateToken(payload);
    
    // Store Google tokens in Redis
    const userData = {
      googleTokens: tokens,
      googleAccessToken: tokens.access_token,
      googleRefreshToken: tokens.refresh_token,
      user: payload
    };
    
    await redis.set(
      `${USER_TOKEN_PREFIX}${data.id}`,
      JSON.stringify(userData),
      'EX',
      60 * 60 * 24 * 30 // 30 days expiration
    );
    
    console.log(`Stored tokens in Redis for user ${data.id}`);
    
    // Set JWT in cookie
    setTokenCookie(res, jwtToken);
    
    res.redirect(`${process.env.FRONTEND_URL}/auth/success?provider=google`);
  } catch (error) {
    console.error('Error in Google callback:', error);
    res.redirect(`${process.env.FRONTEND_URL}/auth/error?provider=google`);
  }
};

// Microsoft Authentication
export const getMicrosoftAuthUrl = (req, res) => {
  const authCodeUrlParameters = {
    scopes: MICROSOFT_SCOPES,
    redirectUri: process.env.MICROSOFT_REDIRECT_URI,
  };

  confidentialClientApplication.getAuthCodeUrl(authCodeUrlParameters)
    .then((authUrl) => {
      res.json({ authUrl });
    })
    .catch((error) => {
      console.error('Error generating Microsoft auth URL:', error);
      res.status(500).json({ error: 'Failed to generate authorization URL' });
    });
};

export const microsoftCallback = async (req, res) => {
  const { code } = req.query;
  
  try {
    // Exchange authorization code for tokens
    const tokenRequest = {
      code: code,
      scopes: MICROSOFT_SCOPES,
      redirectUri: process.env.MICROSOFT_REDIRECT_URI,
    };

    const response = await confidentialClientApplication.acquireTokenByCode(tokenRequest);
    
    // Get user profile from Microsoft Graph API
    const client = Client.init({
      authProvider: (done) => {
        done(null, response.accessToken);
      }
    });

    const userProfile = await client.api('/me').get();
    
    // Create JWT payload
    const payload = {
      userId: userProfile.id,
      email: userProfile.mail || userProfile.userPrincipalName,
      name: userProfile.displayName,
      provider: 'microsoft'
    };
    
    // Generate JWT
    const jwtToken = generateToken(payload);
    
    // Store Microsoft tokens in Redis
    const userData = {
      microsoftTokens: {
        accessToken: response.accessToken,
        refreshToken: response.refreshToken || null,
        expiresOn: response.expiresOn ? response.expiresOn.getTime() : null
      },
      microsoftAccessToken: response.accessToken,
      user: payload
    };
    
    await redis.set(
      `${USER_TOKEN_PREFIX}${userProfile.id}`,
      JSON.stringify(userData),
      'EX',
      60 * 60 * 24 * 30 // 30 days expiration
    );
    
    console.log(`Stored Microsoft tokens in Redis for user ${userProfile.id}`);
    
    // Set JWT in cookie
    setTokenCookie(res, jwtToken);
    
    res.redirect(`${process.env.FRONTEND_URL}/auth/success?provider=microsoft`);
  } catch (error) {
    console.error('Error in Microsoft callback:', error);
    res.redirect(`${process.env.FRONTEND_URL}/auth/error?provider=microsoft`);
  }
};

// Get current auth status
export const getAuthStatus = async (req, res) => {
  const token = req.cookies.auth_token;
  
  if (!token) {
    return res.json({
      google: { authenticated: false, user: null },
      microsoft: { authenticated: false, user: null }
    });
  }

  const decoded = verifyToken(token);
  
  if (!decoded) {
    clearTokenCookie(res);
    return res.json({
      google: { authenticated: false, user: null },
      microsoft: { authenticated: false, user: null }
    });
  }

  const response = {
    google: { authenticated: false, user: null },
    microsoft: { authenticated: false, user: null }
  };

  // Check if user data exists in Redis
  try {
    const userDataStr = await redis.get(`${USER_TOKEN_PREFIX}${decoded.userId}`);
    
    if (userDataStr) {
      const userData = JSON.parse(userDataStr);
      
      if (decoded.provider === 'google' && userData.googleTokens) {
        response.google = {
          authenticated: true,
          user: {
            email: decoded.email,
            name: decoded.name,
            picture: decoded.picture
          }
        };
      } else if (decoded.provider === 'microsoft' && userData.microsoftTokens) {
        response.microsoft = {
          authenticated: true,
          user: {
            email: decoded.email,
            name: decoded.name
          }
        };
      }
    }
  } catch (error) {
    console.error('Error fetching user data from Redis:', error);
  }

  res.json(response);
};

// Logout
export const logout = async (req, res) => {
  const { provider } = req.params;
  const token = req.cookies.auth_token;
  
  if (token) {
    const decoded = verifyToken(token);
    if (decoded && decoded.userId) {
      // If provider is specified, only delete that provider's tokens
      if (provider === 'microsoft' || provider === 'google') {
        const userDataStr = await redis.get(`${USER_TOKEN_PREFIX}${decoded.userId}`);
        
        if (userDataStr) {
          const userData = JSON.parse(userDataStr);
          
          if (provider === 'microsoft') {
            delete userData.microsoftTokens;
            delete userData.microsoftAccessToken;
            console.log(`Deleted Microsoft tokens from Redis for user ${decoded.userId}`);
          } else if (provider === 'google') {
            delete userData.googleTokens;
            delete userData.googleAccessToken;
            delete userData.googleRefreshToken;
            console.log(`Deleted Google tokens from Redis for user ${decoded.userId}`);
          }
          
          // If no tokens remain, delete the entire key
          if (!userData.googleTokens && !userData.microsoftTokens) {
            await redis.del(`${USER_TOKEN_PREFIX}${decoded.userId}`);
          } else {
            // Otherwise, update with remaining tokens
            await redis.set(
              `${USER_TOKEN_PREFIX}${decoded.userId}`,
              JSON.stringify(userData),
              'EX',
              60 * 60 * 24 * 30
            );
          }
        }
      } else {
        // No provider specified, delete all tokens
        await redis.del(`${USER_TOKEN_PREFIX}${decoded.userId}`);
        console.log(`Deleted all tokens from Redis for user ${decoded.userId}`);
      }
    }
  }
  
  clearTokenCookie(res);
  res.json({ success: true });
};

// Get Google tokens for API calls (protected route)
export const getGoogleTokens = async (req, res) => {
  const userDataStr = await redis.get(`${USER_TOKEN_PREFIX}${req.user.userId}`);
  
  if (!userDataStr) {
    return res.status(404).json({ error: 'Google tokens not found' });
  }
  
  const userData = JSON.parse(userDataStr);
  
  if (!userData.googleTokens) {
    return res.status(404).json({ error: 'Google tokens not found' });
  }
  
  res.json({ tokens: userData.googleTokens });
};

// Get Microsoft tokens for API calls (protected route)
export const getMicrosoftTokens = async (req, res) => {
  const userDataStr = await redis.get(`${USER_TOKEN_PREFIX}${req.user.userId}`);
  
  if (!userDataStr) {
    return res.status(404).json({ error: 'Microsoft tokens not found' });
  }
  
  const userData = JSON.parse(userDataStr);
  
  if (!userData.microsoftTokens) {
    return res.status(404).json({ error: 'Microsoft tokens not found' });
  }
  
  res.json({ tokens: userData.microsoftTokens });
};