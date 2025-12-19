import { GoogleGenerativeAI } from '@google/generative-ai';
import dotenv from 'dotenv';
import { verifyToken } from '../utils/jwt.js';
import redis from '../config/redisClient.js';
import mcpClient from '../mcp/client/mcp-client.js';

dotenv.config();
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// Redis key prefix for user tokens
const USER_TOKEN_PREFIX = 'user:token:';

// Helper to get user's access tokens (both Google and Microsoft)
async function getUserAccessToken(req) {
  const token = req.cookies.auth_token;
  console.log("cookies", req.cookies);
  
  if (!token) {
    return {
      googleAccessToken: null,
      microsoftAccessToken: null
    };
  }
  
  const decoded = verifyToken(token);
  console.log("decoded", decoded);
  
  if (!decoded) {
    return {
      googleAccessToken: null,
      microsoftAccessToken: null
    };
  }
  
  const userDataStr = await redis.get(`${USER_TOKEN_PREFIX}${decoded.userId}`);
  
  if (!userDataStr) {
    return {
      googleAccessToken: null,
      microsoftAccessToken: null
    };
  }
  
  const userData = JSON.parse(userDataStr);
  console.log("userData", userData);
  
  return {
    googleAccessToken: userData?.googleAccessToken || null,
    microsoftAccessToken: userData?.microsoftAccessToken || null
  };
}

// AI-based query classification
async function classifyQueryWithAI(message, conversationHistory) {
  try {
    const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

    // Build conversation context
    const context = conversationHistory && conversationHistory.length > 0
      ? conversationHistory
          .slice(-5)
          .map(msg => `${msg.type === 'user' ? 'User' : 'Assistant'}: ${msg.content}`)
          .join('\n')
      : 'No previous conversation';

    const classificationPrompt = `You are an intelligent query classifier. Analyze the user's query and conversation history to determine if it's related to:
1. **Email/Gmail** - queries about emails, sending messages, checking inbox, composing, searching emails, etc.
2. **Microsoft Teams** - queries about Teams chats, channels, meetings, Teams messages, calendar events, etc.
3. **Both** - queries that could involve both email and Teams
4. **Neither** - general queries not related to email or Teams

Previous conversation:
${context}

Current user query: "${message}"

IMPORTANT RULES:
- If the query mentions "email", "mail", "gmail", "inbox", "send to email" → classify as EMAIL
- If the query mentions "teams", "chat", "channel", "meeting", "calendar" → classify as TEAMS
- If the query mentions both or could apply to both → classify as BOTH
- Consider conversation context - if they were talking about Teams before, a follow-up might be Teams-related
- Be smart about ambiguous queries like "send a message" - check context for clues
- General questions like "hello", "how are you", "what can you do" → classify as NEITHER

Respond ONLY with a valid JSON object in this exact format:
{
  "isEmail": true/false,
  "isTeams": true/false,
  "confidence": "high/medium/low",
  "reasoning": "brief explanation"
}

Examples:

Query: "Show me my recent emails"
Response: {"isEmail": true, "isTeams": false, "confidence": "high", "reasoning": "Explicitly asks for emails"}

Query: "List my Teams chats"
Response: {"isEmail": false, "isTeams": true, "confidence": "high", "reasoning": "Explicitly asks for Teams chats"}

Query: "Send a message to John"
Response: {"isEmail": true, "isTeams": true, "confidence": "medium", "reasoning": "Ambiguous - could be email or Teams message"}

Query: "What's the weather today?"
Response: {"isEmail": false, "isTeams": false, "confidence": "high", "reasoning": "General question, not related to email or Teams"}

Query: "Create a meeting tomorrow at 3pm"
Response: {"isEmail": false, "isTeams": true, "confidence": "high", "reasoning": "Meeting creation is a Teams/Calendar feature"}

Now classify this query and respond with JSON only:`;

    const result = await model.generateContent(classificationPrompt);
    const response = await result.response;
    const text = response.text().trim();

    // Extract JSON from the response (handle markdown code blocks)
    let jsonText = text;
    if (text.includes('```json')) {
      jsonText = text.match(/```json\s*([\s\S]*?)\s*```/)?.[1] || text;
    } else if (text.includes('```')) {
      jsonText = text.match(/```\s*([\s\S]*?)\s*```/)?.[1] || text;
    }

    const classification = JSON.parse(jsonText);
    
    console.log('AI Classification Result:', classification);
    
    return {
      isEmail: classification.isEmail || false,
      isTeams: classification.isTeams || false,
      confidence: classification.confidence || 'low',
      reasoning: classification.reasoning || 'No reasoning provided'
    };
  } catch (error) {
    console.error('Error in AI classification, using fallback:', error.message);
    
    // Fallback to simple keyword matching if AI classification fails
    const lowerMessage = message.toLowerCase();
    const hasEmailKeywords = ['email', 'mail', 'gmail', 'inbox', 'send to'].some(k => lowerMessage.includes(k));
    const hasTeamsKeywords = ['teams', 'team', 'chat', 'channel', 'meeting', 'calendar'].some(k => lowerMessage.includes(k));
    
    return {
      isEmail: hasEmailKeywords,
      isTeams: hasTeamsKeywords,
      confidence: 'low',
      reasoning: 'Fallback keyword matching due to AI error'
    };
  }
}

// UNIFIED: Smart chat that auto-detects whether to use MCP
export async function aiResponse(req, res) {
  try {
    const { message, conversationHistory = [] } = req.body;

    if (!message) {
      return res.status(400).json({ error: 'Message is required' });
    }

    // Get user's access tokens (both Google and Microsoft)
    const { googleAccessToken, microsoftAccessToken } = await getUserAccessToken(req);
    const isAuthenticated = !!(googleAccessToken || microsoftAccessToken);
    
    // Use AI to intelligently classify the query
    const classification = await classifyQueryWithAI(message, conversationHistory);
    const isEmailQuery = classification.isEmail;
    const isTeamsQuery = classification.isTeams;

    console.log(`Query: "${message.substring(0, 50)}..."`);
    console.log(`Is authenticated (Google): ${!!googleAccessToken}`);
    console.log(`Is authenticated (Microsoft): ${!!microsoftAccessToken}`);
    console.log(`AI Classification:`, {
      isEmail: isEmailQuery,
      isTeams: isTeamsQuery,
      confidence: classification.confidence,
      reasoning: classification.reasoning
    });

    // Determine which providers to use based on authentication and query type
    const providers = [];
    if (googleAccessToken && isEmailQuery) providers.push('gmail');
    if (microsoftAccessToken && isTeamsQuery) providers.push('teams');

    // If we have providers to use, try MCP
    if (providers.length > 0) {
      try {
        console.log(`Attempting to use MCP tools for providers: ${providers.join(', ')}`);
        
        // Get available MCP tools for the relevant providers
        const mcpTools = await mcpClient.getAvailableTools(providers);
        
        if (mcpTools.length > 0) {
          console.log(`Loaded ${mcpTools.length} MCP tools, using MCP-enhanced response`);
          
          // Determine which access token to pass based on the primary provider
          const primaryProvider = providers[0];
          const accessToken = primaryProvider === 'gmail' ? googleAccessToken : microsoftAccessToken;
          
          return await handleMCPResponse(
            req, 
            res, 
            message, 
            conversationHistory, 
            accessToken, 
            mcpTools,
            primaryProvider,
            { googleAccessToken, microsoftAccessToken }
          );
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
async function handleMCPResponse(req, res, message, conversationHistory, accessToken, mcpTools, primaryProvider, allTokens) {
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

  // Determine assistant type based on available tools
  const hasGmailTools = mcpTools.some(t => t.name.startsWith('gmail_'));
  const hasTeamsTools = mcpTools.some(t => t.name.startsWith('teams_'));
  
  let assistantType = 'assistant';
  if (hasGmailTools && hasTeamsTools) {
    assistantType = 'email and Teams assistant';
  } else if (hasGmailTools) {
    assistantType = 'email assistant';
  } else if (hasTeamsTools) {
    assistantType = 'Teams assistant';
  }

  const enhancedPrompt = `You are an ${assistantType} with access to the user's account through these tools:

${toolDescriptions}

Previous conversation:
${context}

User query: ${message}

Analyze the user's query and determine:
1. Does this require accessing Gmail data? 
2. If yes, which tool should be used?
3. What parameters should be passed?

EXAMPLES:

Gmail Examples:

Example 1 - Sending email:
User: "Send an email to john@example.com with subject Meeting saying Let's meet tomorrow at 3pm"
Response:
{
  "needsTool": true,
  "toolName": "gmail_send_message",
  "toolArgs": {
    "to": "john@example.com",
    "subject": "Meeting",
    "body": "Let's meet tomorrow at 3pm"
  },
  "reasoning": "User wants to send an email"
}

Example 2 - Sending email with full message:
User: "Email sarah@test.com about the project update. Tell her the project is on track and we'll deliver by Friday"
Response:
{
  "needsTool": true,
  "toolName": "gmail_send_message",
  "toolArgs": {
    "to": "sarah@test.com",
    "subject": "Project Update",
    "body": "Hi Sarah,\\n\\nThe project is on track and we'll deliver by Friday.\\n\\nBest regards"
  },
  "reasoning": "User wants to send an email about project update"
}

Example 3 - Listing emails:
User: "Show me my recent emails"
Response:
{
  "needsTool": true,
  "toolName": "gmail_list_messages",
  "toolArgs": {
    "maxResults": 10
  },
  "reasoning": "User wants to see recent emails"
}

Teams Examples:

Example 4 - Listing Teams chats:
User: "Show me my recent Teams chats"
Response:
{
  "needsTool": true,
  "toolName": "teams_list_chats",
  "toolArgs": {
    "maxResults": 10
  },
  "reasoning": "User wants to see recent Teams chats"
}

Example 5 - Sending Teams message:
User: "Send a Teams message to chat abc123 saying Great work on the presentation"
Response:
{
  "needsTool": true,
  "toolName": "teams_send_message",
  "toolArgs": {
    "chatId": "abc123",
    "message": "Great work on the presentation"
  },
  "reasoning": "User wants to send a Teams message"
}

Example 6 - Listing Teams channels:
User: "Show me the channels in my Marketing team"
Response:
{
  "needsTool": true,
  "toolName": "teams_list_teams",
  "toolArgs": {},
  "reasoning": "User wants to see their teams first to find the Marketing team"
}

Example 7 - Posting to Teams channel:
User: "Post to the General channel in team xyz789 saying Meeting at 2pm today"
Response:
{
  "needsTool": true,
  "toolName": "teams_post_channel_message",
  "toolArgs": {
    "teamId": "xyz789",
    "channelId": "general",
    "message": "Meeting at 2pm today"
  },
  "reasoning": "User wants to post a message to a Teams channel"
}

Now respond in JSON format for the user's query above:
{
  "needsTool": true/false,
  "toolName": "exact_tool_name" (only if needsTool is true),
  "toolArgs": { /* parameters as specified in tool schema */ } (only if needsTool is true),
  "reasoning": "brief explanation"
}

CRITICAL RULES:
- For gmail_send_message: ALWAYS include a non-empty "body" field with the actual message content
- For teams_send_message: ALWAYS include a non-empty "message" field with the actual message content
- For teams_post_channel_message: ALWAYS include a non-empty "message" field with the actual message content
- Extract the email/message content from what the user wants to say
- Never leave body/message as null, empty string, or undefined
- If user doesn't specify content, infer a reasonable message based on the context`;

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
    
    // Validate required fields for gmail_send_message
    if (decision.toolName === 'gmail_send_message') {
      if (!decision.toolArgs.body || decision.toolArgs.body.trim() === '') {
        console.error('ERROR: Body is empty or null for gmail_send_message');
        console.error('Original user message:', message);
        return res.json({
          success: false,
          response: 'I couldn\'t extract the email body from your message. Please specify what you want to say in the email.',
          usedMCP: false,
          mode: 'error'
        });
      }
    }
    
    // Validate required fields for Teams message tools
    if (decision.toolName === 'teams_send_message' || decision.toolName === 'teams_post_channel_message') {
      if (!decision.toolArgs.message || decision.toolArgs.message.trim() === '') {
        console.error(`ERROR: Message is empty or null for ${decision.toolName}`);
        console.error('Original user message:', message);
        return res.json({
          success: false,
          response: 'I couldn\'t extract the message content from your request. Please specify what you want to say.',
          usedMCP: false,
          mode: 'error'
        });
      }
    }

    try {
      // Determine which provider and access token to use based on tool name
      let provider = primaryProvider;
      let tokenToUse = accessToken;
      
      if (decision.toolName.startsWith('gmail_')) {
        provider = 'gmail';
        tokenToUse = allTokens.googleAccessToken;
        
        // Check if Gmail token is available
        if (!tokenToUse) {
          console.error('ERROR: Gmail tool requested but user not authenticated with Google');
          return res.json({
            success: false,
            response: 'To use Gmail features, please sign in with your Google account first. Click the "Sign in with Google" button.',
            usedMCP: false,
            mode: 'auth_required',
            authRequired: 'google'
          });
        }
      } else if (decision.toolName.startsWith('teams_')) {
        provider = 'teams';
        tokenToUse = allTokens.microsoftAccessToken;
        
        // Check if Microsoft token is available
        if (!tokenToUse) {
          console.error('ERROR: Teams tool requested but user not authenticated with Microsoft');
          return res.json({
            success: false,
            response: 'To use Microsoft Teams features, please sign in with your Microsoft account first. Click the "Sign in with Microsoft" button.',
            usedMCP: false,
            mode: 'auth_required',
            authRequired: 'microsoft'
          });
        }
      }
      
      // Execute the MCP tool
      const toolResult = await mcpClient.callTool(
        provider,
        decision.toolName,
        decision.toolArgs,
        tokenToUse
      );

      // Extract text from tool result
      const toolResultText = toolResult.content
        .filter(c => c.type === 'text')
        .map(c => c.text)
        .join('\n');
      console.log('Tool result received:', toolResultText);

      // Parse the tool result
      let parsedResult;
      try {
        parsedResult = JSON.parse(toolResultText);
      } catch {
        parsedResult = { data: toolResultText };
      }

      // Ask Gemini to format the result for the user
      const toolType = decision.toolName.startsWith('gmail_') ? 'Gmail' : 
                       decision.toolName.startsWith('teams_') ? 'Teams' : 'tool';
      
      const formattingPrompt = `The user asked: "${message}"

I executed the ${toolType} tool "${decision.toolName}" and got this result:
${JSON.stringify(parsedResult, null, 2)}

Please provide a clear, friendly, and well-formatted response to the user based on this data.

FORMATTING GUIDELINES:
- Use clear section headers followed by a colon (e.g., "Your Recent Emails:")
- For emails: Format as a list with each email on a new line
  * Include: From, Subject, and Date
  * Use bullet points (- or •) for lists
- For Teams chats/messages: Format similarly with relevant details
- For Teams channels: List them clearly with team names
- For success messages (like sent email): Be concise and confirmatory
- Use blank lines between sections for readability
- If there's an error, explain it helpfully with suggestions
- Be conversational and friendly
- Use numbered lists (1., 2., 3.) for step-by-step information

EXAMPLE FORMAT:

Your Recent Emails:

• From: John Smith <john@example.com>
  Subject: Meeting Tomorrow
  Date: Dec 17, 2025

• From: Sarah Lee <sarah@company.com>
  Subject: Project Update
  Date: Dec 16, 2025

Please format the response according to these guidelines.`;
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
    const { googleAccessToken, microsoftAccessToken } = await getUserAccessToken(req);
    
    // Determine which providers are authenticated
    const providers = [];
    if (googleAccessToken) providers.push('gmail');
    if (microsoftAccessToken) providers.push('teams');
    
    if (providers.length === 0) {
      return res.json({
        authenticated: false,
        connectedProviders: [],
        availableTools: [],
        capabilities: ['standard_chat']
      });
    }

    // Get available MCP tools for all authenticated providers
    const mcpTools = await mcpClient.getAvailableTools(providers);

    // Build capabilities list
    const capabilities = ['standard_chat', 'mcp_tools'];
    if (googleAccessToken) capabilities.push('email_access');
    if (microsoftAccessToken) capabilities.push('teams_access');

    res.json({
      authenticated: true,
      connectedProviders: providers,
      availableTools: mcpTools.map(t => ({
        name: t.name,
        description: t.description
      })),
      capabilities
    });

  } catch (error) {
    console.error('Error getting chat status:', error);
    res.status(500).json({ 
      error: 'Failed to get chat status' 
    });
  }
}