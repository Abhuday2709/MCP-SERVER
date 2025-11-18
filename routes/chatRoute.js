import express from "express"
import { aiResponse, getChatStatus } from "../controllers/chatController.js"

const router = express.Router()

// Unified smart chat endpoint (auto-detects MCP usage)
router.post('/chat', aiResponse);

// Get chat status (what tools are available)
router.get('/chat/status', getChatStatus);

export default router;
