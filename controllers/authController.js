import { oauth2Client, SCOPES as GOOGLE_SCOPES } from '../config/googleConfig.js';
import { google } from 'googleapis';
import { generateToken, setTokenCookie, clearTokenCookie, verifyToken } from '../utils/jwt.js';

// In-memory token storage (use Redis in production)
const userTokens = new Map();

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
    
    // Store Google tokens (use Redis/DB in production)
    userTokens.set(data.id, {
      googleTokens: tokens,
      user: payload
    });
    
    // Set JWT in cookie
    setTokenCookie(res, jwtToken);
    
    res.redirect(`${process.env.FRONTEND_URL}/auth/success?provider=google`);
  } catch (error) {
    console.error('Error in Google callback:', error);
    res.redirect(`${process.env.FRONTEND_URL}/auth/error?provider=google`);
  }
};

// Get current auth status
export const getAuthStatus = (req, res) => {
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

  if (decoded.provider === 'google') {
    response.google = {
      authenticated: true,
      user: {
        email: decoded.email,
        name: decoded.name,
        picture: decoded.picture
      }
    };
  }

  res.json(response);
};

// Logout
export const logout = (req, res) => {
  const { provider } = req.params;
  const token = req.cookies.auth_token;
  
  if (token) {
    const decoded = verifyToken(token);
    if (decoded && decoded.userId) {
      userTokens.delete(decoded.userId);
    }
  }
  
  clearTokenCookie(res);
  res.json({ success: true });
};

// Get Google tokens for API calls (protected route)
export const getGoogleTokens = (req, res) => {
  const userData = userTokens.get(req.user.userId);
  
  if (!userData || !userData.googleTokens) {
    return res.status(404).json({ error: 'Google tokens not found' });
  }
  
  res.json({ tokens: userData.googleTokens });
};