import { GoogleGenerativeAI } from '@google/generative-ai';
import dotenv from 'dotenv';
import { verifyToken } from '../utils/jwt.js';
import { userTokens } from './authController.js';
import mcpClient from '../mcp/client/mcp-client.js';

dotenv.config();
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// Helper to get user's Google access token
function getUserAccessToken(req) {
  const token = req.cookies.auth_token;
  console.log("cookies",req.cookies);
  
  if (!token) return null;
  
  const decoded = verifyToken(token);
  console.log("decoded",decoded);
  
  if (!decoded) return null;
  
  const userData = userTokens.get(decoded.userId);
  console.log("userData",userData);
  
  return userData?.googleAccessToken || null;
}

// Helper to check if query is email-related
function isEmailRelatedQuery(message) {
  const emailKeywords = [
    'email', 'mail', 'inbox', 'message', 'send', 'search',
    'from:', 'to:', 'subject:', 'unread', 'read', 'sent',
    'gmail', 'compose', 'draft', 'attachment', 'reply',
    'forward', 'delete', 'archive', 'spam', 'folder',
    'recent emails', 'latest emails', 'my emails', 'check mail'
  ];
  
  const lowerMessage = message.toLowerCase();
  return emailKeywords.some(keyword => lowerMessage.includes(keyword));
}

// UNIFIED: Smart chat that auto-detects whether to use MCP
export async function aiResponse(req, res) {
  try {
    const { message, conversationHistory = [] } = req.body;

    if (!message) {
      return res.status(400).json({ error: 'Message is required' });
    }

    // Get user's Google access token (if available)
    const googleAccessToken = getUserAccessToken(req);
    const isAuthenticated = !!googleAccessToken;
    
    // Check if this is an email-related query
    const isEmailQuery = isEmailRelatedQuery(message);

    console.log(`Query: "${message.substring(0, 50)}..."`);
    console.log(`Is authenticated: ${isAuthenticated}`);
    console.log(`Is email query: ${isEmailQuery}`);

    // If authenticated AND email query, try to use MCP
    if (isAuthenticated && isEmailQuery) {
      try {
        console.log('Attempting to use MCP tools...');
        
        // Get available MCP tools
        const mcpTools = await mcpClient.getAvailableTools(['gmail']);
        
        if (mcpTools.length > 0) {
          console.log(`Loaded ${mcpTools.length} MCP tools, using MCP-enhanced response`);
          return await handleMCPResponse(req, res, message, conversationHistory, googleAccessToken, mcpTools);
        } else {
          console.log('No MCP tools available, falling back to standard response');
        }
      } catch (mcpError) {
        console.error('MCP error, falling back to standard response:', mcpError.message);
        // Fall through to standard response
      }
    }

    // Standard Gemini response (no MCP)
    console.log('Using standard Gemini response');
    return await handleStandardResponse(req, res, message, conversationHistory);

  } catch (error) {
    console.error('Error in aiResponse:', error);
    res.status(500).json({ 
      success: false, 
      error: 'Failed to generate response',
      message: error.message 
    });
  }
}

// Handle standard Gemini response (no MCP)
async function handleStandardResponse(req, res, message, conversationHistory) {
  const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

  // Build conversation context
  let prompt = message;
  if (conversationHistory && conversationHistory.length > 0) {
    const context = conversationHistory
      .slice(-5)
      .map(msg => `${msg.type === 'user' ? 'User' : 'Assistant'}: ${msg.content}`)
      .join('\n');
    prompt = `Previous conversation:\n${context}\n\nUser: ${message}`;
  }

  // Generate response from Gemini
  const result = await model.generateContent(prompt);
  const response = await result.response;
  const text = response.text();

  return res.json({ 
    success: true, 
    response: text,
    usedMCP: false,
    mode: 'standard'
  });
}

// Handle MCP-enhanced response
async function handleMCPResponse(req, res, message, conversationHistory, googleAccessToken, mcpTools) {
  const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

  // Build conversation context
  const context = conversationHistory
    .slice(-5)
    .map(msg => `${msg.type === 'user' ? 'User' : 'Assistant'}: ${msg.content}`)
    .join('\n');

  // Create a prompt that tells Gemini about available tools
  const toolDescriptions = mcpTools.map(tool => 
    `- ${tool.name}: ${tool.description}\n  Parameters: ${JSON.stringify(tool.inputSchema.properties, null, 2)}`
  ).join('\n\n');

  const enhancedPrompt = `You are an email assistant with access to the user's Gmail account through these tools:

${toolDescriptions}

Previous conversation:
${context}

User query: ${message}

Analyze the user's query and determine:
1. Does this require accessing Gmail data? 
2. If yes, which tool should be used?
3. What parameters should be passed?

Respond in JSON format:
{
  "needsTool": true/false,
  "toolName": "exact_tool_name" (only if needsTool is true),
  "toolArgs": { /* parameters as specified in tool schema */ } (only if needsTool is true),
  "reasoning": "brief explanation"
}

If no tool is needed (e.g., general question about emails), respond:
{
  "needsTool": false,
  "response": "your helpful response to the user"
}

IMPORTANT: 
- Use exact tool names from the list above
- Match parameter types exactly (string, number, boolean)
- For date filters, use format YYYY/MM/DD
- Be specific with parameters`;

  const result = await model.generateContent(enhancedPrompt);
  const geminiResponse = await result.response;
  let responseText = geminiResponse.text();

  // Clean up markdown code blocks if present
  responseText = responseText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();

  console.log('Gemini decision:', responseText);

  let decision;
  try {
    decision = JSON.parse(responseText);
  } catch (parseError) {
    console.error('Failed to parse Gemini response as JSON:', parseError);
    // If parsing fails, treat as standard response
    return res.json({
      success: true,
      response: responseText,
      usedMCP: false,
      mode: 'standard',
      parseError: true
    });
  }

  // If Gemini decided to use a tool
  if (decision.needsTool && decision.toolName) {
    console.log(`Executing MCP tool: ${decision.toolName}`);
    console.log('Tool arguments:', JSON.stringify(decision.toolArgs, null, 2));

    try {
      // Execute the MCP tool
      const toolResult = await mcpClient.callTool(
        'gmail',
        decision.toolName,
        decision.toolArgs,
        googleAccessToken
      );

      // Extract text from tool result
      const toolResultText = toolResult.content
        .filter(c => c.type === 'text')
        .map(c => c.text)
        .join('\n');

      console.log('Tool result received (first 200 chars):', toolResultText.substring(0, 200));

      // Parse the tool result
      let parsedResult;
      try {
        parsedResult = JSON.parse(toolResultText);
      } catch {
        parsedResult = { data: toolResultText };
      }

      // Ask Gemini to format the result for the user
      const formattingPrompt = `The user asked: "${message}"

I executed the Gmail tool "${decision.toolName}" and got this result:
${JSON.stringify(parsedResult, null, 2)}

Please provide a clear, friendly, and well-formatted response to the user based on this data. 
- If there are emails, list them clearly with sender, subject, and date
- If there's an error, explain it helpfully
- Be conversational and helpful
- Format the response in a readable way`;

      const formattingResult = await model.generateContent(formattingPrompt);
      const finalResponse = await formattingResult.response;
      const formattedText = finalResponse.text();

      return res.json({
        success: true,
        response: formattedText,
        usedMCP: true,
        mode: 'mcp',
        toolName: decision.toolName,
        rawData: parsedResult // Optional: include raw data for frontend
      });

    } catch (toolError) {
      console.error('Tool execution error:', toolError);
      
      // Provide helpful error message
      let errorMessage = `I tried to access your Gmail but encountered an error: ${toolError.message}.`;
      
      if (toolError.message.includes('Access token')) {
        errorMessage += ' Please try logging in again.';
      } else if (toolError.message.includes('quota')) {
        errorMessage += ' Gmail API quota may be exceeded. Please try again later.';
      }

      return res.json({
        success: true,
        response: errorMessage,
        usedMCP: false,
        mode: 'error',
        error: toolError.message
      });
    }
  } else {
    // Direct response from Gemini (no tool needed)
    return res.json({
      success: true,
      response: decision.response || 'I can help you with that, but I need more specific information.',
      usedMCP: false,
      mode: 'standard',
      reasoning: decision.reasoning
    });
  }
}

// Get chat status with MCP tools
export async function getChatStatus(req, res) {
  try {
    const googleAccessToken = getUserAccessToken(req);
    
    if (!googleAccessToken) {
      return res.json({
        authenticated: false,
        availableTools: [],
        capabilities: ['standard_chat']
      });
    }

    const mcpTools = await mcpClient.getAvailableTools(['gmail']);

    res.json({
      authenticated: true,
      connectedProviders: ['gmail'],
      availableTools: mcpTools.map(t => ({
        name: t.name,
        description: t.description
      })),
      capabilities: ['standard_chat', 'email_access', 'mcp_tools']
    });

  } catch (error) {
    console.error('Error getting chat status:', error);
    res.status(500).json({ 
      error: 'Failed to get chat status' 
    });
  }
}