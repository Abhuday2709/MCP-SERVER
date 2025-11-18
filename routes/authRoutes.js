import express from 'express';
import {
  getGoogleAuthUrl,
  googleCallback,
  getAuthStatus,
  logout,
  getGoogleTokens
} from '../controllers/authController.js';
import { authenticateToken } from '../middleware/auth.js';

const router = express.Router();

// Google routes
router.get('/google', getGoogleAuthUrl);
router.get('/google/callback', googleCallback);

// Status and logout
router.get('/status', getAuthStatus);
router.post('/logout/:provider', logout);

// Protected route to get Google tokens
router.get('/google/tokens', authenticateToken, getGoogleTokens);

export default router;