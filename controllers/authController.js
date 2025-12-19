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
    
    console.log('\n=== GOOGLE LOGIN ===');
    console.log('Email from Google:', data.email);
    
    // Check if user already has a session (JWT cookie from previous login)
    const existingToken = req.cookies.auth_token;
    let sessionId;
    let existingData = {};
    
    if (existingToken) {
      const existingDecoded = verifyToken(existingToken);
      if (existingDecoded && existingDecoded.sessionId) {
        // User already has a session - use that session ID
        sessionId = existingDecoded.sessionId;
        console.log('Found existing session:', sessionId);
        
        const existingDataStr = await redis.get(`${USER_TOKEN_PREFIX}${sessionId}`);
        existingData = existingDataStr ? JSON.parse(existingDataStr) : {};
        //console.log('Existing data from session:', existingData);
      }
    }
    
    // If no existing session, create new session ID
    if (!sessionId) {
      sessionId = `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
      console.log('Creating new session:', sessionId);
    }
    
    console.log('Redis key will be:', `${USER_TOKEN_PREFIX}${sessionId}`);
    console.log('Has Microsoft tokens?', !!existingData.microsoftTokens);
    console.log('Has Google tokens?', !!existingData.googleTokens);
    
    // Create JWT payload with session ID
    const payload = {
      sessionId: sessionId,
      userId: sessionId, // Keep for backward compatibility
      email: data.email,
      name: data.name,
      picture: data.picture,
      provider: 'google',
      googleId: data.id,
      googleEmail: data.email
    };      
    // Merge Google tokens with existing data
    const userData = {
      ...existingData,
      googleTokens: tokens,
      googleAccessToken: tokens.access_token,
      googleRefreshToken: tokens.refresh_token,
      googleId: data.id,
      googleEmail: data.email,
      sessionId: sessionId,
      // Keep the most recent name/picture
      name: data.name,
      picture: data.picture
    };
    
    await redis.set(
      `${USER_TOKEN_PREFIX}${sessionId}`,
      JSON.stringify(userData),
      'EX',
      60 * 60 * 24 * 30 // 30 days expiration
    );
    
    console.log(`Stored/updated Google tokens in Redis for session ${sessionId}`);
    console.log('Final userData has Microsoft tokens?', !!userData.microsoftTokens);
    console.log('Final userData has Google tokens?', !!userData.googleTokens);
    
    // Generate and set new JWT with session ID
    const jwtToken = generateToken(payload);
    setTokenCookie(res, jwtToken);
    console.log(`Set JWT cookie with session ${sessionId}`);
    
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
    const userEmail = (userProfile.mail || userProfile.userPrincipalName).toLowerCase();
    
    console.log('\n=== MICROSOFT LOGIN ===');
    console.log('Email from Microsoft:', userEmail);
    
    // Check if user already has a session (JWT cookie from previous login)
    const existingToken = req.cookies.auth_token;
    let sessionId;
    let existingData = {};
    
    if (existingToken) {
      const existingDecoded = verifyToken(existingToken);
      if (existingDecoded && existingDecoded.sessionId) {
        // User already has a session - use that session ID
        sessionId = existingDecoded.sessionId;
        console.log('Found existing session:', sessionId);
        
        const existingDataStr = await redis.get(`${USER_TOKEN_PREFIX}${sessionId}`);
        existingData = existingDataStr ? JSON.parse(existingDataStr) : {};
        //console.log('Existing data from session:', existingData);
      }
    }
    
    // If no existing session, create new session ID
    if (!sessionId) {
      sessionId = `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
      console.log('Creating new session:', sessionId);
    }
    
    console.log('Redis key will be:', `${USER_TOKEN_PREFIX}${sessionId}`);
    console.log('Has Google tokens?', !!existingData.googleTokens);
    console.log('Has Microsoft tokens?', !!existingData.microsoftTokens);
    
    // Create JWT payload with session ID
    const payload = {
      sessionId: sessionId,
      userId: sessionId, // Keep for backward compatibility
      email: userEmail,
      name: userProfile.displayName,
      provider: 'microsoft',
      microsoftId: userProfile.id,
      microsoftEmail: userEmail
    };                
    // Merge Microsoft tokens with existing data
    const userData = {
      ...existingData,
      microsoftTokens: {
        accessToken: response.accessToken,
        refreshToken: response.refreshToken || null,
        expiresOn: response.expiresOn ? response.expiresOn.getTime() : null
      },
      microsoftAccessToken: response.accessToken,
      microsoftId: userProfile.id,
      microsoftEmail: userEmail,
      sessionId: sessionId,
      // Keep the most recent name
      name: userProfile.displayName
    };
                                                            
    await redis.set(
      `${USER_TOKEN_PREFIX}${sessionId}`,
      JSON.stringify(userData),
      'EX',
      60 * 60 * 24 * 30 // 30 days expiration
    );
    
    console.log(`Stored/updated Microsoft tokens in Redis for session ${sessionId}`);
    console.log('Final userData has Google tokens?', !!userData.googleTokens);
    console.log('Final userData has Microsoft tokens?', !!userData.microsoftTokens);
    
    // Generate and set new JWT with session ID
    const jwtToken = generateToken(payload);
    setTokenCookie(res, jwtToken);
    console.log(`Set JWT cookie with session ${sessionId}`);
    
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
    // Use sessionId if available, otherwise fall back to userId
    const redisKey = decoded.sessionId || decoded.userId;
    const userDataStr = await redis.get(`${USER_TOKEN_PREFIX}${redisKey}`);
    
    if (userDataStr) {
      const userData = JSON.parse(userDataStr);
      
      // Check for Google authentication
      if (userData.googleTokens) {
        response.google = {
          authenticated: true,
          user: {
            email: userData.googleEmail || userData.email,
            name: userData.name,
            picture: userData.picture
          }
        };
      }
      
      // Check for Microsoft authentication
      if (userData.microsoftTokens) {
        response.microsoft = {
          authenticated: true,
          user: {
            email: userData.microsoftEmail || userData.email,
            name: userData.name
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