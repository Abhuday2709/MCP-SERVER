import express from 'express';
import {
  getGoogleAuthUrl,
  googleCallback,
  getAuthStatus,
  logout,
  getGoogleTokens,
  getMicrosoftAuthUrl,
  microsoftCallback,
  getMicrosoftTokens
} from '../controllers/authController.js';
import { authenticateToken } from '../middleware/auth.js';

const router = express.Router();

// Google routes
router.get('/google', getGoogleAuthUrl);
router.get('/google/callback', googleCallback);

// Microsoft routes
router.get('/microsoft', getMicrosoftAuthUrl);
router.get('/microsoft/callback', microsoftCallback);

// Status and logout
router.get('/status', getAuthStatus);
router.post('/logout/:provider', logout);

// Protected route to get Google tokens
router.get('/google/tokens', authenticateToken, getGoogleTokens);

// Protected route to get Microsoft tokens
router.get('/microsoft/tokens', authenticateToken, getMicrosoftTokens);

export default router;