# MCP Chatbot Server

A Node.js/Express backend server that integrates Google Gemini AI with the Model Context Protocol (MCP) to provide intelligent chatbot capabilities with Gmail integration through OAuth2 authentication.

## 🏗️ Architecture

### Technology Stack

- **Runtime**: Node.js (ES Modules)
- **Framework**: Express.js 5.x
- **AI Model**: Google Gemini 2.5 Flash
- **Authentication**: JWT with HTTP-only cookies
- **OAuth2**: Google OAuth2 for Gmail access
- **Database/Cache**: Redis (Upstash compatible)
- **Protocol**: Model Context Protocol (MCP) SDK
- **API Client**: Google APIs (googleapis)

### Project Structure

```
server/
├── config/
│   ├── googleConfig.js      # Google OAuth2 configuration
│   └── redisClient.js        # Redis connection setup
├── controllers/
│   ├── authController.js     # Authentication logic
│   └── chatController.js     # AI chat & MCP integration
├── mcp/
│   ├── client/
│   │   └── mcp-client.js     # MCP client manager
│   └── servers/
│       └── gmail-server.js   # Gmail MCP server implementation
├── middleware/
│   └── auth.js               # JWT authentication middleware
├── routes/
│   ├── authRoutes.js         # Auth endpoints
│   └── chatRoute.js          # Chat endpoints
├── utils/
│   ├── jwt.js                # JWT helper functions
│   └── testMCP.js            # MCP testing utilities
├── .env                      # Environment variables
├── app.js                    # Main application entry
└── package.json              # Dependencies
```

## 🎯 Core Features

### 1. **Intelligent AI Chat with MCP Integration**
- **Automatic Tool Detection**: Analyzes user queries to determine if Gmail tools are needed
- **Context-Aware Responses**: Maintains conversation history (last 5 messages)
- **Dual Mode Operation**:
  - Standard mode: Direct Gemini AI responses
  - MCP mode: Gmail tool execution with AI interpretation
- **Smart Query Classification**: Detects email-related keywords automatically

### 2. **Google OAuth2 Authentication**
- Secure authorization flow with Google
- JWT token generation and validation
- Redis-based token storage (30-day expiration)
- HTTP-only cookie management
- Support for refresh tokens

### 3. **Gmail Integration via MCP**
Four powerful Gmail tools:
- **`gmail_list_messages`**: List emails with filters
- **`gmail_get_message`**: Get full email details
- **`gmail_search_messages`**: Advanced search with multiple criteria
- **`gmail_send_message`**: Send emails programmatically

### 4. **Model Context Protocol (MCP) Architecture**
- Client-server communication via stdio transport
- Dynamic tool discovery and execution
- Tool result formatting with AI
- Graceful error handling and fallback

### 5. **Security Features**
- JWT authentication with 7-day expiration
- HTTP-only cookies (XSS protection)
- CORS with credentials support
- Secure token storage in Redis
- Environment-based security configurations

## 📋 Prerequisites

- **Node.js**: v18.0.0 or higher
- **Redis**: Local instance or Upstash cloud Redis
- **Google Cloud Project**: With OAuth2 credentials
- **Gmail API**: Enabled in Google Cloud Console

## 🛠️ Installation

### 1. Clone and Navigate

```bash
cd server
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Configure Google Cloud

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a new project or select existing
3. Enable **Gmail API**
4. Go to **Credentials** → **Create Credentials** → **OAuth 2.0 Client ID**
5. Configure OAuth consent screen:
   - Add scopes: `gmail.readonly`, `gmail.send`, `userinfo.email`, `userinfo.profile`
   - Add test users (during development)
6. Create OAuth 2.0 Client ID:
   - Application type: **Web application**
   - Authorized redirect URIs: `http://localhost:3000/api/auth/google/callback`
7. Download credentials (Client ID and Secret)

### 4. Set Up Redis

**Option A: Local Redis**
```bash
# Install Redis
# Windows: Download from https://redis.io/download
# macOS: brew install redis
# Linux: sudo apt-get install redis-server

# Start Redis
redis-server
```

**Option B: Upstash (Recommended for Production)**
1. Go to [Upstash Console](https://console.upstash.com)
2. Create a new Redis database
3. Copy the Redis URL

### 5. Configure Environment Variables

Create a `.env` file in the `server` directory:

```env
# Server Configuration
PORT=3000
NODE_ENV=development
FRONTEND_URL=http://localhost:5173

# Google OAuth2
GOOGLE_CLIENT_ID=your-google-client-id.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=your-google-client-secret
GOOGLE_REDIRECT_URI=http://localhost:3000/api/auth/google/callback

# JWT Secret
JWT_SECRET=your-super-secret-jwt-key-min-32-characters-long

# Google Gemini AI
GEMINI_API_KEY=your-gemini-api-key

# Redis
REDIS_URL=redis://localhost:6379
# For Upstash: redis://default:your-password@your-endpoint.upstash.io:6379
```

### 6. Get Gemini API Key

1. Go to [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Create a new API key
3. Copy and paste into `.env`

### 7. Start the Server

**Development mode (with auto-reload):**
```bash
npm run dev
```

**Production mode:**
```bash
npm start
```

The server will start on `http://localhost:3000`

## 🔐 Authentication Flow

```
┌─────────────┐
│   Frontend  │
└──────┬──────┘
       │ 1. GET /api/auth/google
       ▼
┌─────────────┐
│   Backend   │ 2. Returns authUrl
└──────┬──────┘
       │ 3. Redirect to Google
       ▼
┌─────────────┐
│   Google    │ 4. User logs in & grants permissions
└──────┬──────┘
       │ 5. Redirects to callback with code
       ▼
┌─────────────┐
│   Backend   │ 6. Exchange code for tokens
└──────┬──────┘ 7. Get user info
       │         8. Generate JWT
       │         9. Store tokens in Redis
       │         10. Set JWT cookie
       ▼
┌─────────────┐
│   Frontend  │ 11. Redirects to success page
└──────┬──────┘
       │ 12. GET /api/auth/status
       ▼
┌─────────────┐
│   Backend   │ 13. Returns user info from JWT
└─────────────┘
```

---

## 🤖 AI Chat Flow

### Standard Chat (No Gmail)

```
User Message → Gemini AI → Response
```

### MCP-Enhanced Chat (Gmail Access)

```
User: "Show my recent emails"
       │
       ▼
┌──────────────────────┐
│  Keyword Detection   │ → Email-related? Yes
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Gemini Analyzes     │ → Which tool? gmail_list_messages
│  & Decides Tool      │ → Parameters? { maxResults: 10 }
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  MCP Client calls    │ → Passes access token
│  gmail-server.js     │
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Gmail API Request   │ → Returns email data
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  Gemini Formats      │ → Human-readable response
│  Result for User     │
└──────────┬───────────┘
           │
           ▼
       Response to User
```

---

## 🔒 Security Best Practices

### 1. **Environment Variables**
- Never commit `.env` to version control
- Use strong, random JWT secrets (32+ characters)
- Rotate secrets periodically

### 2. **JWT Configuration**
```javascript
// Production settings
{
  httpOnly: true,        // Prevents XSS attacks
  secure: true,          // HTTPS only
  sameSite: 'none',      // Cross-site requests
  maxAge: 7 days         // Auto expiration
}
```

### 3. **CORS Configuration**
```javascript
{
  origin: process.env.FRONTEND_URL,  // Specific origin only
  credentials: true,                  // Allow cookies
  methods: ['GET', 'POST', ...],     // Explicit methods
  allowedHeaders: [...]               // Explicit headers
}
```

### 4. **Token Storage**
- JWT in HTTP-only cookies (client can't access)
- OAuth tokens in Redis with TTL
- Automatic cleanup on logout

### 5. **API Rate Limiting** (Recommended)
```bash
npm install express-rate-limit
```

```javascript
import rateLimit from 'express-rate-limit';

const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100 // limit each IP to 100 requests per windowMs
});

app.use('/api/', limiter);
```

---

## 📄 License

This project is licensed under the ISC License.

---

## 🔗 Related Documentation

- [Frontend Client Documentation](../client/README.md)
- [Model Context Protocol Specification](https://modelcontextprotocol.io)
- [Google OAuth2 Documentation](https://developers.google.com/identity/protocols/oauth2)
- [Gmail API Reference](https://developers.google.com/gmail/api)
- [Google Gemini AI Documentation](https://ai.google.dev/docs)
- [Redis Documentation](https://redis.io/docs)
- [Express.js Guide](https://expressjs.com)

---

## 📞 Support

For issues and questions:
- Open an issue on GitHub
- Check existing issues for solutions
- Review troubleshooting section above
- Check server logs for detailed errors

---

**Built with ❤️ using Node.js, Express, and Google Gemini AI**
