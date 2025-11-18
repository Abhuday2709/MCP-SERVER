import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';

class GmailMCPServer {
  constructor() {
    this.server = new Server(
      {
        name: 'gmail-mcp-server',
        version: '1.0.0',
      },
      {
        capabilities: {
          tools: {},
        },
      }
    );

    this.setupTools();
    this.setupErrorHandling();
  }

  setupTools() {
    // List available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: 'gmail_list_messages',
            description: 'List emails from Gmail inbox with optional filters',
            inputSchema: {
              type: 'object',
              properties: {
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of messages to return (default: 10)',
                  default: 10,
                },
                query: {
                  type: 'string',
                  description: 'Gmail search query (e.g., "from:user@example.com", "is:unread", "after:2024/01/01")',
                },
              },
            },
          },
          {
            name: 'gmail_get_message',
            description: 'Get full details of a specific email by ID',
            inputSchema: {
              type: 'object',
              properties: {
                messageId: {
                  type: 'string',
                  description: 'The ID of the message to retrieve',
                },
              },
              required: ['messageId'],
            },
          },
          {
            name: 'gmail_search_messages',
            description: 'Search emails with specific criteria',
            inputSchema: {
              type: 'object',
              properties: {
                from: {
                  type: 'string',
                  description: 'Filter by sender email',
                },
                to: {
                  type: 'string',
                  description: 'Filter by recipient email',
                },
                subject: {
                  type: 'string',
                  description: 'Filter by subject keywords',
                },
                after: {
                  type: 'string',
                  description: 'Date after (YYYY/MM/DD)',
                },
                before: {
                  type: 'string',
                  description: 'Date before (YYYY/MM/DD)',
                },
                hasAttachment: {
                  type: 'boolean',
                  description: 'Filter messages with attachments',
                },
                isUnread: {
                  type: 'boolean',
                  description: 'Filter unread messages',
                },
                maxResults: {
                  type: 'number',
                  description: 'Maximum results (default: 20)',
                  default: 20,
                },
              },
            },
          },
          {
            name: 'gmail_send_message',
            description: 'Send an email via Gmail',
            inputSchema: {
              type: 'object',
              properties: {
                to: {
                  type: 'string',
                  description: 'Recipient email address',
                },
                subject: {
                  type: 'string',
                  description: 'Email subject',
                },
                body: {
                  type: 'string',
                  description: 'Email body (plain text or HTML)',
                },
                cc: {
                  type: 'string',
                  description: 'CC recipients (comma-separated)',
                },
              },
              required: ['to', 'subject', 'body'],
            },
          },
        ],
      };
    });

    // Handle tool calls
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case 'gmail_list_messages':
            return await this.listMessages(args);
          case 'gmail_get_message':
            return await this.getMessage(args);
          case 'gmail_search_messages':
            return await this.searchMessages(args);
          case 'gmail_send_message':
            return await this.sendMessage(args);
          default:
            throw new Error(`Unknown tool: ${name}`);
        }
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: `Error: ${error.message}`,
            },
          ],
          isError: true,
        };
      }
    });
  }

  async listMessages(args) {
    const { maxResults = 10, query = '' } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: accessToken });

    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

    const response = await gmail.users.messages.list({
      userId: 'me',
      maxResults,
      q: query,
    });

    const messages = response.data.messages || [];

    if (messages.length === 0) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              totalMessages: 0,
              messages: [],
              message: 'No messages found',
            }, null, 2),
          },
        ],
      };
    }

    // Fetch details for each message (limit to 5 for performance)
    const detailedMessages = await Promise.all(
      messages.slice(0, Math.min(5, messages.length)).map(async (msg) => {
        try {
          const detail = await gmail.users.messages.get({
            userId: 'me',
            id: msg.id,
            format: 'metadata',
            metadataHeaders: ['From', 'To', 'Subject', 'Date'],
          });

          const headers = detail.data.payload.headers;
          const getHeader = (name) => headers.find(h => h.name === name)?.value || '';

          return {
            id: msg.id,
            threadId: msg.threadId,
            from: getHeader('From'),
            to: getHeader('To'),
            subject: getHeader('Subject'),
            date: getHeader('Date'),
            snippet: detail.data.snippet,
          };
        } catch (err) {
          console.error(`Error fetching message ${msg.id}:`, err.message);
          return null;
        }
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalMessages: messages.length,
            messages: detailedMessages.filter(m => m !== null),
          }, null, 2),
        },
      ],
    };
  }

  async getMessage(args) {
    const { messageId } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: accessToken });

    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

    const response = await gmail.users.messages.get({
      userId: 'me',
      id: messageId,
      format: 'full',
    });

    const message = response.data;
    const headers = message.payload.headers;
    const getHeader = (name) => headers.find(h => h.name === name)?.value || '';

    // Extract body
    let body = '';
    if (message.payload.parts) {
      const part = message.payload.parts.find(p => p.mimeType === 'text/plain') ||
                    message.payload.parts.find(p => p.mimeType === 'text/html');
      if (part && part.body.data) {
        body = Buffer.from(part.body.data, 'base64').toString('utf-8');
      }
    } else if (message.payload.body.data) {
      body = Buffer.from(message.payload.body.data, 'base64').toString('utf-8');
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            id: message.id,
            threadId: message.threadId,
            from: getHeader('From'),
            to: getHeader('To'),
            subject: getHeader('Subject'),
            date: getHeader('Date'),
            body: body.substring(0, 1000), // Limit body length
            snippet: message.snippet,
            labelIds: message.labelIds,
          }, null, 2),
        },
      ],
    };
  }

  async searchMessages(args) {
    const {
      from,
      to,
      subject,
      after,
      before,
      hasAttachment,
      isUnread,
      maxResults = 20,
    } = args;

    // Build Gmail search query
    let queryParts = [];
    if (from) queryParts.push(`from:${from}`);
    if (to) queryParts.push(`to:${to}`);
    if (subject) queryParts.push(`subject:${subject}`);
    if (after) queryParts.push(`after:${after}`);
    if (before) queryParts.push(`before:${before}`);
    if (hasAttachment) queryParts.push('has:attachment');
    if (isUnread) queryParts.push('is:unread');

    const query = queryParts.join(' ');

    return await this.listMessages({ maxResults, query, _auth: args._auth });
  }

  async sendMessage(args) {
    const { to, subject, body, cc } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: accessToken });

    const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

    // Create email
    const email = [
      `To: ${to}`,
      cc ? `Cc: ${cc}` : '',
      `Subject: ${subject}`,
      '',
      body,
    ].filter(Boolean).join('\n');

    const encodedEmail = Buffer.from(email)
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');

    const response = await gmail.users.messages.send({
      userId: 'me',
      requestBody: {
        raw: encodedEmail,
      },
    });

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: response.data.id,
            message: 'Email sent successfully',
          }, null, 2),
        },
      ],
    };
  }

  setupErrorHandling() {
    this.server.onerror = (error) => {
      console.error('[Gmail MCP Server Error]', error);
    };

    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Gmail MCP Server running on stdio');
  }
}

// Start server
const server = new GmailMCPServer();
server.run().catch(console.error);