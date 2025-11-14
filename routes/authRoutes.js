import express from 'express';
import {
  getGoogleAuthUrl,
  googleCallback,
  getAuthStatus,
  logout
} from '../controllers/authController.js';

const router = express.Router();

// Google routes
router.get('/google', getGoogleAuthUrl);
router.get('/google/callback', googleCallback);

// Status and logout
router.get('/status', getAuthStatus);
router.post('/logout/:provider', logout);

export default router;