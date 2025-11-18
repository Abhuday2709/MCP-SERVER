import { verifyToken } from '../utils/jwt.js';

export const authenticateToken = (req, res, next) => {
  const token = req.cookies.auth_token;

  if (!token) {
    return res.status(401).json({ error: 'Access denied. No token provided.' });
  }

  const decoded = verifyToken(token);
  
  if (!decoded) {
    return res.status(403).json({ error: 'Invalid or expired token.' });
  }

  req.user = decoded;
  next();
};