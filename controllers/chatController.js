import { GoogleGenerativeAI } from '@google/generative-ai';
import dotenv from 'dotenv';
dotenv.config();
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

export async function aiResponse(req,res){
    try {
        const { message, conversationHistory } = req.body;

        if (!message) {
            return res.status(400).json({ error: 'Message is required' });
        }

        // Get the generative model
        const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

        // Build conversation context
        let prompt = message;
        if (conversationHistory && conversationHistory.length > 0) {
            const context = conversationHistory
                .slice(-5) // Last 5 messages for context
                .map(msg => `${msg.type === 'user' ? 'User' : 'Assistant'}: ${msg.content}`)
                .join('\n');
            prompt = `Previous conversation:\n${context}\n\nUser: ${message}`;
        }

        // Generate response from Gemini
        const result = await model.generateContent(prompt);
        const response = await result.response;
        const text = response.text();

        res.json({ 
            success: true, 
            response: text 
        });

    } catch (error) {
        console.error('Error with Gemini API:', error);
        res.status(500).json({ 
            success: false, 
            error: 'Failed to generate response',
            message: error.message 
        });
    }
}