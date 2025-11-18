import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import cookieParser from 'cookie-parser';
import chatRouter from "./routes/chatRoute.js"
import authRouter from "./routes/authRoutes.js"

dotenv.config();

const app = express();

app.use(cors({
    origin: process.env.FRONTEND_URL || 'http://localhost:5173',
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));    
app.use(express.json());
app.use(cookieParser());

app.use('/api', chatRouter);
app.use('/api/auth', authRouter);

app.get('/', (req, res) => {
    res.send('MCP Chatbot Server is running');
});

const PORT = process.env.PORT || 3000;

const server = app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});

// Graceful shutdown - disconnect MCP clients
process.on('SIGINT', async () => {
  console.log('\nShutting down gracefully...');
  try {
    const mcpClient = (await import('./mcp/client/mcp-client.js')).default;
    await mcpClient.disconnect();
    console.log('MCP clients disconnected');
  } catch (error) {
    console.error('Error disconnecting MCP clients:', error);
  }
  server.close(() => {
    console.log('Server closed');
    process.exit(0);
  });
});

process.on('SIGTERM', async () => {
  console.log('\nSIGTERM received, shutting down gracefully...');
  try {
    const mcpClient = (await import('./mcp/client/mcp-client.js')).default;
    await mcpClient.disconnect();
  } catch (error) {
    console.error('Error disconnecting MCP clients:', error);
  }
  server.close(() => {
    console.log('Server closed');
    process.exit(0);
  });
});