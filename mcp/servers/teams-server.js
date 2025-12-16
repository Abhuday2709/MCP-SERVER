import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';

class TeamsMCPServer {
  constructor() {
    this.server = new Server(
      {
        name: 'teams-mcp-server',
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

  /**
   * Centralized Graph API handler
   * Handles authentication, error handling, and rate limiting
   */
  async executeGraphAPICall(accessToken, endpoint, method = 'GET', body = null) {
    if (!accessToken) {
      throw new Error('Access token is required');
    }

    try {
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken);
        },
      });

      let request = client.api(endpoint);

      // Execute based on method
      let response;
      switch (method.toUpperCase()) {
        case 'GET':
          response = await request.get();
          break;
        case 'POST':
          response = await request.post(body);
          break;
        case 'PATCH':
          response = await request.patch(body);
          break;
        case 'DELETE':
          response = await request.delete();
          break;
        default:
          throw new Error(`Unsupported HTTP method: ${method}`);
      }

      return response;
    } catch (error) {
      // Handle specific Graph API errors
      if (error.statusCode === 401) {
        throw new Error('Your Microsoft Teams session has expired. Please log in again.');
      } else if (error.statusCode === 403) {
        throw new Error("You don't have permission to access this Teams resource.");
      } else if (error.statusCode === 429) {
        throw new Error('Microsoft Teams API rate limit reached. Please try again in a moment.');
      } else if (error.statusCode === 404) {
        throw new Error('The requested Teams resource was not found.');
      } else {
        throw new Error(`Microsoft Graph API error: ${error.message}`);
      }
    }
  }

  async listChats(args) {
    const { maxResults = 10 } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    const response = await this.executeGraphAPICall(
      accessToken,
      `/me/chats?$top=${maxResults}&$expand=members`
    );

    const chats = response.value || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalChats: chats.length,
            chats: chats.map(chat => ({
              id: chat.id,
              topic: chat.topic || 'No topic',
              chatType: chat.chatType,
              members: chat.members?.map(m => ({
                userId: m.userId,
                displayName: m.displayName,
                email: m.email,
              })) || [],
            })),
          }, null, 2),
        },
      ],
    };
  }

  async listMessages(args) {
    const { chatId, maxResults = 20 } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!chatId) {
      throw new Error('chatId is required');
    }

    const response = await this.executeGraphAPICall(
      accessToken,
      `/me/chats/${chatId}/messages?$top=${maxResults}&$orderby=createdDateTime desc`
    );

    const messages = response.value || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalMessages: messages.length,
            chatId,
            messages: messages.map(msg => ({
              id: msg.id,
              from: {
                userId: msg.from?.user?.id,
                displayName: msg.from?.user?.displayName,
              },
              body: {
                content: msg.body?.content,
                contentType: msg.body?.contentType,
              },
              createdDateTime: msg.createdDateTime,
              importance: msg.importance,
            })),
          }, null, 2),
        },
      ],
    };
  }

  async sendMessage(args) {
    const { chatId, message } = args;
    const { accessToken } = args._auth || {};

    console.log('[Teams Server] sendMessage called with args:', {
      chatId,
      message: message ? `"${message.substring(0, 50)}..."` : 'NULL/EMPTY',
      hasAccessToken: !!accessToken
    });

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!chatId) {
      throw new Error('chatId is required');
    }

    if (!message || message.trim() === '') {
      throw new Error('Message body is required and cannot be empty');
    }

    const body = {
      body: {
        content: message,
        contentType: 'text',
      },
    };

    const response = await this.executeGraphAPICall(
      accessToken,
      `/me/chats/${chatId}/messages`,
      'POST',
      body
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: response.id,
            message: 'Message sent successfully',
          }, null, 2),
        },
      ],
    };
  }

  async listTeams(args) {
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    const response = await this.executeGraphAPICall(
      accessToken,
      '/me/joinedTeams'
    );

    const teams = response.value || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalTeams: teams.length,
            teams: teams.map(team => ({
              id: team.id,
              displayName: team.displayName,
              description: team.description,
            })),
          }, null, 2),
        },
      ],
    };
  }

  async listChannels(args) {
    const { teamId } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!teamId) {
      throw new Error('teamId is required');
    }

    const response = await this.executeGraphAPICall(
      accessToken,
      `/teams/${teamId}/channels`
    );

    const channels = response.value || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalChannels: channels.length,
            teamId,
            channels: channels.map(channel => ({
              id: channel.id,
              displayName: channel.displayName,
              description: channel.description,
              membershipType: channel.membershipType,
            })),
          }, null, 2),
        },
      ],
    };
  }

  async getChannelMessages(args) {
    const { teamId, channelId, maxResults = 20 } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!teamId) {
      throw new Error('teamId is required');
    }

    if (!channelId) {
      throw new Error('channelId is required');
    }

    const response = await this.executeGraphAPICall(
      accessToken,
      `/teams/${teamId}/channels/${channelId}/messages?$top=${maxResults}&$orderby=createdDateTime desc`
    );

    const messages = response.value || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalMessages: messages.length,
            teamId,
            channelId,
            messages: messages.map(msg => ({
              id: msg.id,
              from: {
                userId: msg.from?.user?.id,
                displayName: msg.from?.user?.displayName,
              },
              body: {
                content: msg.body?.content,
                contentType: msg.body?.contentType,
              },
              createdDateTime: msg.createdDateTime,
              importance: msg.importance,
            })),
          }, null, 2),
        },
      ],
    };
  }

  async postChannelMessage(args) {
    const { teamId, channelId, message } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!teamId) {
      throw new Error('teamId is required');
    }

    if (!channelId) {
      throw new Error('channelId is required');
    }

    if (!message || message.trim() === '') {
      throw new Error('Message body is required and cannot be empty');
    }

    const body = {
      body: {
        content: message,
        contentType: 'text',
      },
    };

    const response = await this.executeGraphAPICall(
      accessToken,
      `/teams/${teamId}/channels/${channelId}/messages`,
      'POST',
      body
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: response.id,
            message: 'Channel message posted successfully',
          }, null, 2),
        },
      ],
    };
  }

  async searchMessages(args) {
    const { query, maxResults = 20, from, after, before } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!query) {
      throw new Error('query is required');
    }

    // Get all chats first
    const chatsResponse = await this.executeGraphAPICall(
      accessToken,
      '/me/chats?$top=50'
    );

    const chats = chatsResponse.value || [];
    const allMessages = [];

    // Search through each chat's messages
    for (const chat of chats) {
      try {
        const messagesResponse = await this.executeGraphAPICall(
          accessToken,
          `/me/chats/${chat.id}/messages?$top=50&$orderby=createdDateTime desc`
        );

        const messages = messagesResponse.value || [];

        // Filter messages based on criteria
        const filteredMessages = messages.filter(msg => {
          const content = msg.body?.content?.toLowerCase() || '';
          const matchesQuery = content.includes(query.toLowerCase());

          // Filter by sender if specified
          if (from && msg.from?.user?.displayName) {
            const matchesSender = msg.from.user.displayName.toLowerCase().includes(from.toLowerCase());
            if (!matchesSender) return false;
          }

          // Filter by date range if specified
          if (after) {
            const messageDate = new Date(msg.createdDateTime);
            const afterDate = new Date(after);
            if (messageDate < afterDate) return false;
          }

          if (before) {
            const messageDate = new Date(msg.createdDateTime);
            const beforeDate = new Date(before);
            if (messageDate > beforeDate) return false;
          }

          return matchesQuery;
        });

        // Add chat context to messages
        filteredMessages.forEach(msg => {
          allMessages.push({
            ...msg,
            chatId: chat.id,
            chatTopic: chat.topic || 'No topic',
          });
        });

        // Stop if we have enough results
        if (allMessages.length >= maxResults) {
          break;
        }
      } catch (error) {
        // Continue with other chats if one fails
        console.error(`Error searching chat ${chat.id}:`, error.message);
      }
    }

    // Limit results
    const limitedMessages = allMessages.slice(0, maxResults);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalResults: limitedMessages.length,
            query,
            messages: limitedMessages.map(msg => ({
              id: msg.id,
              chatId: msg.chatId,
              chatTopic: msg.chatTopic,
              from: {
                userId: msg.from?.user?.id,
                displayName: msg.from?.user?.displayName,
              },
              body: {
                content: msg.body?.content,
                contentType: msg.body?.contentType,
              },
              createdDateTime: msg.createdDateTime,
              importance: msg.importance,
            })),
          }, null, 2),
        },
      ],
    };
  }

  setupTools() {
    // Register all tools with their schemas
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: 'teams_list_chats',
            description: "List user's recent Teams chats",
            inputSchema: {
              type: 'object',
              properties: {
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of chats to return (default: 10)',
                  default: 10,
                },
              },
            },
          },
          {
            name: 'teams_list_messages',
            description: 'List messages from a specific Teams chat',
            inputSchema: {
              type: 'object',
              properties: {
                chatId: {
                  type: 'string',
                  description: 'The ID of the chat to retrieve messages from',
                },
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of messages to return (default: 20)',
                  default: 20,
                },
              },
              required: ['chatId'],
            },
          },
          {
            name: 'teams_send_message',
            description: 'Send a message to a Teams chat',
            inputSchema: {
              type: 'object',
              properties: {
                chatId: {
                  type: 'string',
                  description: 'The ID of the chat to send the message to',
                },
                message: {
                  type: 'string',
                  description: 'The message content to send',
                },
              },
              required: ['chatId', 'message'],
            },
          },
          {
            name: 'teams_list_teams',
            description: "List teams the user is a member of",
            inputSchema: {
              type: 'object',
              properties: {},
            },
          },
          {
            name: 'teams_list_channels',
            description: 'List channels in a specific team',
            inputSchema: {
              type: 'object',
              properties: {
                teamId: {
                  type: 'string',
                  description: 'The ID of the team to list channels from',
                },
              },
              required: ['teamId'],
            },
          },
          {
            name: 'teams_get_channel_messages',
            description: 'Get messages from a specific Teams channel',
            inputSchema: {
              type: 'object',
              properties: {
                teamId: {
                  type: 'string',
                  description: 'The ID of the team',
                },
                channelId: {
                  type: 'string',
                  description: 'The ID of the channel',
                },
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of messages to return (default: 20)',
                  default: 20,
                },
              },
              required: ['teamId', 'channelId'],
            },
          },
          {
            name: 'teams_post_channel_message',
            description: 'Post a message to a Teams channel',
            inputSchema: {
              type: 'object',
              properties: {
                teamId: {
                  type: 'string',
                  description: 'The ID of the team',
                },
                channelId: {
                  type: 'string',
                  description: 'The ID of the channel',
                },
                message: {
                  type: 'string',
                  description: 'The message content to post',
                },
              },
              required: ['teamId', 'channelId', 'message'],
            },
          },
          {
            name: 'teams_search_messages',
            description: 'Search messages across Teams chats with filters',
            inputSchema: {
              type: 'object',
              properties: {
                query: {
                  type: 'string',
                  description: 'Search query to match message content',
                },
                from: {
                  type: 'string',
                  description: 'Filter by sender display name',
                },
                after: {
                  type: 'string',
                  description: 'Filter messages after this date (ISO 8601 format)',
                },
                before: {
                  type: 'string',
                  description: 'Filter messages before this date (ISO 8601 format)',
                },
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of results to return (default: 20)',
                  default: 20,
                },
              },
              required: ['query'],
            },
          },
        ],
      };
    });

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      try {
        switch (name) {
          case 'teams_list_chats':
            return await this.listChats(args);
          case 'teams_list_messages':
            return await this.listMessages(args);
          case 'teams_send_message':
            return await this.sendMessage(args);
          case 'teams_list_teams':
            return await this.listTeams(args);
          case 'teams_list_channels':
            return await this.listChannels(args);
          case 'teams_get_channel_messages':
            return await this.getChannelMessages(args);
          case 'teams_post_channel_message':
            return await this.postChannelMessage(args);
          case 'teams_search_messages':
            return await this.searchMessages(args);
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

  setupErrorHandling() {
    this.server.onerror = (error) => {
      console.error('[Teams MCP Server Error]', error);
    };

    process.on('SIGINT', async () => {
      await this.server.close();
      process.exit(0);
    });

    process.on('SIGTERM', async () => {
      await this.server.close();
      process.exit(0);
    });
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Teams MCP Server running on stdio');
  }
}

// Start server
const server = new TeamsMCPServer();
server.run().catch(console.error);
