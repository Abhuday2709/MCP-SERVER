import express from "express"
import { aiResponse } from "../controllers/chatController.js"

const router = express.Router()

router.post('/chat',aiResponse);

export default router;