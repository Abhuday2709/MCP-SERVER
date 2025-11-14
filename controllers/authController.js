import { oauth2Client, SCOPES as GOOGLE_SCOPES } from '../config/googleConfig.js';
// import { pca, SCOPES as MS_SCOPES } from '../config/microsoft.config.js';
import { google } from 'googleapis';

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
    
    // Store in session
    req.session.googleTokens = tokens;
    req.session.googleUser = {
      email: data.email,
      name: data.name,
      picture: data.picture
    };
    
    // Redirect to frontend
    res.redirect(`${process.env.FRONTEND_URL}/auth/success?provider=google`);
  } catch (error) {
    console.error('Error in Google callback:', error);
    res.redirect(`${process.env.FRONTEND_URL}/auth/error?provider=google`);
  }
};


// Get current session status
export const getAuthStatus = (req, res) => {
  res.json({
    google: {
      authenticated: !!req.session.googleTokens,
      user: req.session.googleUser || null
    },
    microsoft: {
      authenticated: !!req.session.microsoftTokens,
      user: req.session.microsoftUser || null
    }
  });
};

// Logout
export const logout = (req, res) => {
  const { provider } = req.params;
  
  if (provider === 'google') {
    delete req.session.googleTokens;
    delete req.session.googleUser;
  } else if (provider === 'microsoft') {
    delete req.session.microsoftTokens;
    delete req.session.microsoftUser;
  } else {
    req.session.destroy();
  }
  
  res.json({ success: true });
};