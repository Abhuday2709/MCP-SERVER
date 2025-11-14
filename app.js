import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import { GoogleGenerativeAI } from '@google/generative-ai';
import chatRouter from "./routes/chatRoute.js"

dotenv.config();

const app = express();

app.use(cors());
app.use(express.json());

app.use('/api',chatRouter);

app.get('/', (req, res) => {
    res.send('MCP Chatbot Server is running');
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});
