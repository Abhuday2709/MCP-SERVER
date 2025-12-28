import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { Client } from '@microsoft/microsoft-graph-client';

// Constants for retry logic and rate limiting
const MAX_RETRIES = 10;
const RETRY_DELAY_MS = 1000;
const RATE_LIMIT_DELAY_MS = 5000;

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

    // Cache for reducing API calls
    this.cache = {
      chats: null,
      chatsExpiry: 0,
      teams: null,
      teamsExpiry: 0,
      user: null,
      userExpiry: 0,
    };
    this.CACHE_TTL = 60000; // 1 minute cache

    this.setupTools();
    this.setupErrorHandling();
  }

  /**
   * Helper to delay execution
   */
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Clear cache when needed
   */
  clearCache() {
    this.cache = {
      chats: null,
      chatsExpiry: 0,
      teams: null,
      teamsExpiry: 0,
      user: null,
      userExpiry: 0,
    };
  }

  /**
   * Get cached data or fetch new
   */
  async getCachedOrFetch(cacheKey, fetchFn, accessToken) {
    const now = Date.now();
    if (this.cache[cacheKey] && this.cache[`${cacheKey}Expiry`] > now) {
      return this.cache[cacheKey];
    }
    
    const data = await fetchFn(accessToken);
    this.cache[cacheKey] = data;
    this.cache[`${cacheKey}Expiry`] = now + this.CACHE_TTL;
    return data;
  }

  /**
   * Centralized Graph API handler with retry logic
   * Handles authentication, error handling, rate limiting, and retries
   */
  async executeGraphAPICall(accessToken, endpoint, method = 'GET', body = null, retryCount = 0) {
    if (!accessToken) {
      throw new Error('Access token is required. Please sign in with Microsoft.');
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
      const statusCode = error.statusCode || error.code;
      
      // Handle rate limiting with exponential backoff
      if (statusCode === 429) {
        if (retryCount < MAX_RETRIES) {
          const retryAfter = error.headers?.['retry-after'] 
            ? parseInt(error.headers['retry-after']) * 1000 
            : RATE_LIMIT_DELAY_MS * Math.pow(2, retryCount);
          
          console.warn(`Rate limited. Retrying after ${retryAfter}ms (attempt ${retryCount + 1}/${MAX_RETRIES})`);
          await this.delay(retryAfter);
          return this.executeGraphAPICall(accessToken, endpoint, method, body, retryCount + 1);
        }
        throw new Error('Microsoft Teams API rate limit exceeded. Please try again later.');
      }

      // Handle transient errors with retry
      if ((statusCode === 503 || statusCode === 504) && retryCount < MAX_RETRIES) {
        const retryDelay = RETRY_DELAY_MS * Math.pow(2, retryCount);
        console.warn(`Service unavailable. Retrying after ${retryDelay}ms (attempt ${retryCount + 1}/${MAX_RETRIES})`);
        await this.delay(retryDelay);
        return this.executeGraphAPICall(accessToken, endpoint, method, body, retryCount + 1);
      }

      // Handle specific Graph API errors with user-friendly messages
      if (statusCode === 401) {
        this.clearCache();
        throw new Error('Your Microsoft session has expired. Please sign in again.');
      } else if (statusCode === 403) {
        throw new Error("You don't have permission to access this resource. Please check your permissions.");
      } else if (statusCode === 404) {
        throw new Error('The requested resource was not found.');
      } else if (statusCode === 400) {
        const errorMessage = error.body?.error?.message || error.message;
        throw new Error(`Invalid request: ${errorMessage}`);
      } else {
        // Log the full error for debugging but return a clean message
        console.error('Graph API Error:', {
          statusCode,
          message: error.message,
          body: error.body,
          endpoint,
          method
        });
        throw new Error(`Microsoft Graph API error: ${error.message || 'Unknown error occurred'}`);
      }
    }
  }

  /**
   * Safely extract text content from message body
   */
  extractMessageContent(body) {
    if (!body) return '';
    
    let content = body.content || '';
    
    // Strip HTML tags if content type is HTML
    if (body.contentType === 'html') {
      content = content
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .trim();
    }
    
    return content;
  }

  /**
   * Format date for display
   */
  formatDate(dateString) {
    if (!dateString) return null;
    try {
      return new Date(dateString).toISOString();
    } catch {
      return dateString;
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
      `/me/chats?$top=${Math.min(maxResults, 50)}&$expand=members&$orderby=lastMessagePreview/createdDateTime desc`
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
              topic: chat.topic || this.getChatDisplayName(chat),
              chatType: chat.chatType,
              lastMessagePreview: chat.lastMessagePreview?.body?.content 
                ? this.extractMessageContent(chat.lastMessagePreview.body).substring(0, 100)
                : null,
              lastMessageDate: this.formatDate(chat.lastMessagePreview?.createdDateTime),
              members: chat.members?.map(m => ({
                userId: m.userId,
                displayName: m.displayName,
                email: m.email,
              })).filter(m => m.displayName) || [],
            })),
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Get a display name for chats without topics
   */
  getChatDisplayName(chat) {
    if (chat.topic) return chat.topic;
    
    if (chat.chatType === 'oneOnOne' && chat.members) {
      // For 1:1 chats, return the other person's name
      const otherMembers = chat.members.filter(m => m.displayName);
      if (otherMembers.length > 0) {
        return otherMembers.map(m => m.displayName).join(', ');
      }
    }
    
    if (chat.members && chat.members.length > 0) {
      const names = chat.members
        .filter(m => m.displayName)
        .map(m => m.displayName)
        .slice(0, 3);
      return names.length > 0 ? names.join(', ') + (chat.members.length > 3 ? '...' : '') : 'Unnamed chat';
    }
    
    return 'Unnamed chat';
  }

  async listMessages(args) {
    const { chatId, maxResults = 25 } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!chatId) {
      throw new Error('chatId is required');
    }

    const response = await this.executeGraphAPICall(
      accessToken,
      `/me/chats/${chatId}/messages?$top=${Math.min(maxResults, 50)}&$orderby=createdDateTime desc`
    );

    const messages = response.value || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalMessages: messages.length,
            chatId,
            messages: messages
              .filter(msg => msg.messageType === 'message') // Filter out system messages
              .map(msg => ({
                id: msg.id,
                from: {
                  userId: msg.from?.user?.id,
                  displayName: msg.from?.user?.displayName || 'Unknown',
                },
                body: this.extractMessageContent(msg.body),
                createdDateTime: this.formatDate(msg.createdDateTime),
                importance: msg.importance,
              })),
          }, null, 2),
        },
      ],
    };
  }

  async findChatByName(args) {
    const { chatName, participantName } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!chatName && !participantName) {
      throw new Error('Either chatName or participantName is required');
    }

    // Get all chats with members expanded
    const response = await this.executeGraphAPICall(
      accessToken,
      '/me/chats?$expand=members&$top=50&$orderby=lastMessagePreview/createdDateTime desc'
    );

    const chats = response.value || [];
    const searchTerm = (chatName || participantName).toLowerCase().trim();

    // Find matching chats with improved matching logic
    const matches = chats.filter(chat => {
      // Exact match by topic/name
      if (chat.topic && chat.topic.toLowerCase().trim() === searchTerm) {
        return true;
      }
      
      // Partial match by topic/name
      if (chat.topic && chat.topic.toLowerCase().includes(searchTerm)) {
        return true;
      }

      // Match by participant name or email
      if (chat.members) {
        return chat.members.some(member => 
          (member.displayName && member.displayName.toLowerCase().includes(searchTerm)) ||
          (member.email && member.email.toLowerCase().includes(searchTerm))
        );
      }

      return false;
    });

    // Sort matches: exact matches first, then by last message date
    matches.sort((a, b) => {
      const aExact = a.topic?.toLowerCase().trim() === searchTerm;
      const bExact = b.topic?.toLowerCase().trim() === searchTerm;
      if (aExact && !bExact) return -1;
      if (!aExact && bExact) return 1;
      return 0;
    });

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            searchTerm: chatName || participantName,
            totalMatches: matches.length,
            matches: matches.slice(0, 10).map(chat => ({
              id: chat.id,
              topic: chat.topic || this.getChatDisplayName(chat),
              chatType: chat.chatType,
              members: chat.members?.map(m => ({
                displayName: m.displayName,
                email: m.email,
              })).filter(m => m.displayName || m.email) || [],
            })),
          }, null, 2),
        },
      ],
    };
  }

  async createOneOnOneChat(args) {
    const { userId, userEmail } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!userId && !userEmail) {
      throw new Error('Either userId or userEmail is required');
    }

    // If we have email but not userId, look up the user
    let targetUserId = userId;
    if (!targetUserId && userEmail) {
      try {
        const userResponse = await this.executeGraphAPICall(
          accessToken,
          `/users/${encodeURIComponent(userEmail)}`
        );
        targetUserId = userResponse.id;
      } catch (error) {
        throw new Error(`Could not find user with email: ${userEmail}. ${error.message}`);
      }
    }

    // Get current user ID for the chat member binding
    const meResponse = await this.executeGraphAPICall(accessToken, '/me');

    // Create a one-on-one chat
    const chatData = {
      chatType: 'oneOnOne',
      members: [
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: ['owner'],
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${meResponse.id}')`
        },
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: ['owner'],
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${targetUserId}')`
        }
      ]
    };

    const chat = await this.executeGraphAPICall(
      accessToken,
      '/chats',
      'POST',
      chatData
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            chatId: chat.id,
            message: 'One-on-one chat created successfully',
          }, null, 2),
        },
      ],
    };
  }

  async sendMessage(args) {
    const {
      teamName,
      chatId,
      chatName,
      participantName,
      participantEmail,
      participants = [],
      message
    } = args;

    const { accessToken } = args._auth || {};

    if (!accessToken) throw new Error('Access token is required');
    if (!message || !message.trim()) throw new Error('Message content is required');

    const cleanMessage = message.trim();

    /* --------------------------------------------------
       1️⃣ TEAM/CHANNEL MESSAGE FLOW (HIGHEST PRIORITY)
    -------------------------------------------------- */
    if (teamName) {
      console.log('[Teams] Sending to team:', teamName);

      const teamsRes = await this.executeGraphAPICall(
        accessToken,
        '/me/joinedTeams'
      );

      const team = teamsRes.value?.find(
        t => t.displayName.toLowerCase() === teamName.toLowerCase()
      );

      if (!team) {
        const availableTeams = teamsRes.value?.map(t => t.displayName).join(', ') || 'none';
        throw new Error(`Team "${teamName}" not found. Available teams: ${availableTeams}`);
      }

      const channelsRes = await this.executeGraphAPICall(
        accessToken,
        `/teams/${team.id}/channels`
      );

      const channel = channelsRes.value.find(c => c.displayName === 'General') || channelsRes.value[0];

      if (!channel) {
        throw new Error(`No channels found in team "${teamName}"`);
      }

      const response = await this.executeGraphAPICall(
        accessToken,
        `/teams/${team.id}/channels/${channel.id}/messages`,
        'POST',
        {
          body: {
            contentType: 'text',
            content: cleanMessage
          }
        }
      );

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              type: 'channel',
              team: team.displayName,
              channel: channel.displayName,
              messageId: response.id,
              message: `Message sent to ${channel.displayName} in ${team.displayName}`
            }, null, 2)
          }
        ]
      };
    }

    /* --------------------------------------------------
       2️⃣ DIRECT CHAT MESSAGE FLOW
    -------------------------------------------------- */
    let targetChatId = chatId;

    // Find existing chat if needed
    if (!targetChatId && (chatName || participantName || participantEmail)) {
      const chatsRes = await this.executeGraphAPICall(
        accessToken,
        '/me/chats?$expand=members&$top=50'
      );

      // Search by chat name/topic
      if (chatName && !participantName && !participantEmail) {
        const searchTerm = chatName.toLowerCase().trim();
        
        const chat = chatsRes.value.find(c =>
          c.topic && c.topic.toLowerCase().trim() === searchTerm
        ) || chatsRes.value.find(c =>
          c.topic && c.topic.toLowerCase().includes(searchTerm)
        );

        if (chat) {
          targetChatId = chat.id;
          console.log('[Teams] Found chat by name:', chat.topic);
        } else {
          const availableChats = chatsRes.value
            .filter(c => c.topic)
            .map(c => c.topic)
            .join(', ');
          throw new Error(`Chat "${chatName}" not found. Available named chats: ${availableChats || 'none'}`);
        }
      }
      // Search by participant
      else if (participantName || participantEmail) {
        const searchTerm = (participantName || participantEmail).toLowerCase().trim();
        
        // Prefer one-on-one chats for individual messages
        const oneOnOneChat = chatsRes.value.find(c =>
          c.chatType === 'oneOnOne' &&
          c.members?.some(m =>
            (m.displayName && m.displayName.toLowerCase().includes(searchTerm)) ||
            (m.email && m.email.toLowerCase().includes(searchTerm))
          )
        );

        if (oneOnOneChat) {
          targetChatId = oneOnOneChat.id;
          console.log('[Teams] Found one-on-one chat with:', searchTerm);
        } else {
          // Try any chat with this participant
          const anyChat = chatsRes.value.find(c =>
            c.members?.some(m =>
              (m.displayName && m.displayName.toLowerCase().includes(searchTerm)) ||
              (m.email && m.email.toLowerCase().includes(searchTerm))
            )
          );

          if (anyChat) {
            targetChatId = anyChat.id;
            console.log('[Teams] Found group chat with participant:', searchTerm);
          }
        }
      }
    }

    // Create new chat if not found and we have enough info
    if (!targetChatId) {
      const targetEmail = participantEmail || participants[0]?.email;
      
      if (!targetEmail) {
        throw new Error('Could not find an existing chat. Please provide a valid email address to create a new chat.');
      }

      console.log('[Teams] Creating new chat with:', targetEmail);

      // Look up the user
      let targetUser;
      try {
        targetUser = await this.executeGraphAPICall(
          accessToken,
          `/users/${encodeURIComponent(targetEmail)}`
        );
      } catch (error) {
        throw new Error(`Could not find user with email "${targetEmail}". Please verify the email address.`);
      }

      // Get current user
      const meResponse = await this.executeGraphAPICall(accessToken, '/me');

      // Build member list
      const memberBindings = [
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: ['owner'],
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${meResponse.id}')`
        },
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: ['owner'],
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${targetUser.id}')`
        }
      ];

      // Add additional participants if provided
      for (const p of participants.slice(1)) {
        if (p.email) {
          try {
            const user = await this.executeGraphAPICall(
              accessToken,
              `/users/${encodeURIComponent(p.email)}`
            );
            memberBindings.push({
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: ['owner'],
              'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user.id}')`
            });
          } catch (error) {
            console.warn(`Could not add participant ${p.email}: ${error.message}`);
          }
        }
      }

      const chat = await this.executeGraphAPICall(
        accessToken,
        '/chats',
        'POST',
        {
          chatType: memberBindings.length > 2 ? 'group' : 'oneOnOne',
          members: memberBindings
        }
      );

      targetChatId = chat.id;
    }

    /* --------------------------------------------------
       3️⃣ SEND MESSAGE TO CHAT
    -------------------------------------------------- */
    const response = await this.executeGraphAPICall(
      accessToken,
      `/me/chats/${targetChatId}/messages`,
      'POST',
      {
        body: {
          contentType: 'text',
          content: cleanMessage
        }
      }
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            type: 'chat',
            chatId: targetChatId,
            messageId: response.id,
            message: 'Message sent successfully'
          }, null, 2)
        }
      ]
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
              description: team.description || 'No description',
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
      throw new Error('teamId is required. Use teams_list_teams first to get the team ID.');
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
              description: channel.description || 'No description',
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
      `/teams/${teamId}/channels/${channelId}/messages?$top=${Math.min(maxResults, 50)}&$orderby=createdDateTime desc`
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
            messages: messages
              .filter(msg => msg.messageType === 'message')
              .map(msg => ({
                id: msg.id,
                from: {
                  userId: msg.from?.user?.id,
                  displayName: msg.from?.user?.displayName || 'Unknown',
                },
                body: this.extractMessageContent(msg.body),
                createdDateTime: this.formatDate(msg.createdDateTime),
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
      throw new Error('Message content is required');
    }

    const response = await this.executeGraphAPICall(
      accessToken,
      `/teams/${teamId}/channels/${channelId}/messages`,
      'POST',
      {
        body: {
          content: message.trim(),
          contentType: 'text',
        },
      }
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
      throw new Error('Search query is required');
    }

    const searchQuery = query.toLowerCase().trim();

    // Get chats
    const chatsResponse = await this.executeGraphAPICall(
      accessToken,
      '/me/chats?$top=30'
    );

    const chats = chatsResponse.value || [];
    const allMessages = [];

    // Search through each chat's messages
    for (const chat of chats) {
      if (allMessages.length >= maxResults) break;

      try {
        const messagesResponse = await this.executeGraphAPICall(
          accessToken,
          `/me/chats/${chat.id}/messages?$top=30&$orderby=createdDateTime desc`
        );

        const messages = messagesResponse.value || [];

        // Filter messages
        const filteredMessages = messages.filter(msg => {
          if (msg.messageType !== 'message') return false;
          
          const content = this.extractMessageContent(msg.body).toLowerCase();
          if (!content.includes(searchQuery)) return false;

          // Filter by sender
          if (from && msg.from?.user?.displayName) {
            if (!msg.from.user.displayName.toLowerCase().includes(from.toLowerCase())) {
              return false;
            }
          }

          // Filter by date
          if (after) {
            const messageDate = new Date(msg.createdDateTime);
            if (messageDate < new Date(after)) return false;
          }

          if (before) {
            const messageDate = new Date(msg.createdDateTime);
            if (messageDate > new Date(before)) return false;
          }

          return true;
        });

        // Add chat context
        filteredMessages.forEach(msg => {
          if (allMessages.length < maxResults) {
            allMessages.push({
              ...msg,
              chatId: chat.id,
              chatTopic: chat.topic || this.getChatDisplayName(chat),
            });
          }
        });

      } catch (error) {
        console.warn(`Error searching chat ${chat.id}:`, error.message);
      }
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalResults: allMessages.length,
            query,
            messages: allMessages.map(msg => ({
              id: msg.id,
              chatId: msg.chatId,
              chatTopic: msg.chatTopic,
              from: {
                userId: msg.from?.user?.id,
                displayName: msg.from?.user?.displayName || 'Unknown',
              },
              body: this.extractMessageContent(msg.body),
              createdDateTime: this.formatDate(msg.createdDateTime),
            })),
          }, null, 2),
        },
      ],
    };
  }

  async listCalendars(args) {
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    const response = await this.executeGraphAPICall(
      accessToken,
      '/me/calendars'
    );

    const calendars = response.value || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalCalendars: calendars.length,
            calendars: calendars.map(cal => ({
              id: cal.id,
              name: cal.name,
              color: cal.color,
              isDefaultCalendar: cal.isDefaultCalendar,
              canEdit: cal.canEdit,
              canShare: cal.canShare,
              owner: cal.owner?.name,
            })),
          }, null, 2),
        },
      ],
    };
  }

  async listCalendarEvents(args) {
    const { calendarId, startDateTime, endDateTime, specificDate, maxResults = 50 } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    // Calculate date ranges
    let start, end;
    
    if (specificDate) {
      const date = new Date(specificDate);
      if (isNaN(date.getTime())) {
        throw new Error('Invalid specificDate format. Use ISO 8601 format (YYYY-MM-DD).');
      }
      start = new Date(date.setHours(0, 0, 0, 0)).toISOString();
      end = new Date(date.setHours(23, 59, 59, 999)).toISOString();
    } else if (startDateTime && endDateTime) {
      start = startDateTime;
      end = endDateTime;
    } else {
      // Default: today and next 7 days
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const nextWeek = new Date(today);
      nextWeek.setDate(today.getDate() + 7);
      nextWeek.setHours(23, 59, 59, 999);
      
      start = today.toISOString();
      end = nextWeek.toISOString();
    }

    let endpoint = calendarId 
      ? `/me/calendars/${calendarId}/calendarView`
      : '/me/calendarView';

    const queryParams = [
      `startDateTime=${encodeURIComponent(start)}`,
      `endDateTime=${encodeURIComponent(end)}`,
      `$top=${Math.min(maxResults, 100)}`,
      '$orderby=start/dateTime',
      '$select=id,subject,start,end,location,isOnlineMeeting,onlineMeetingUrl,organizer,attendees,bodyPreview'
    ];

    endpoint += `?${queryParams.join('&')}`;

    const response = await this.executeGraphAPICall(accessToken, endpoint);
    const events = response.value || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalEvents: events.length,
            dateRange: { start, end },
            events: events.map(event => ({
              id: event.id,
              subject: event.subject || 'No subject',
              start: event.start,
              end: event.end,
              location: event.location?.displayName,
              isOnlineMeeting: event.isOnlineMeeting,
              onlineMeetingUrl: event.onlineMeetingUrl,
              organizer: event.organizer?.emailAddress,
              attendees: event.attendees?.map(a => ({
                name: a.emailAddress?.name,
                email: a.emailAddress?.address,
                status: a.status?.response,
              })),
              bodyPreview: event.bodyPreview?.substring(0, 200),
            })),
          }, null, 2),
        },
      ],
    };
  }

  async getCalendarEvent(args) {
    const { eventId, calendarId } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!eventId) {
      throw new Error('eventId is required');
    }

    const endpoint = calendarId
      ? `/me/calendars/${calendarId}/events/${eventId}`
      : `/me/calendar/events/${eventId}`;

    const event = await this.executeGraphAPICall(accessToken, endpoint);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            id: event.id,
            subject: event.subject,
            body: this.extractMessageContent(event.body),
            bodyType: event.body?.contentType,
            start: event.start,
            end: event.end,
            location: event.location,
            isOnlineMeeting: event.isOnlineMeeting,
            onlineMeeting: event.onlineMeeting,
            organizer: event.organizer,
            attendees: event.attendees,
            recurrence: event.recurrence,
            categories: event.categories,
            importance: event.importance,
            sensitivity: event.sensitivity,
          }, null, 2),
        },
      ],
    };
  }

  async createCalendarEvent(args) {
    const { 
      subject, 
      body, 
      startDateTime, 
      endDateTime, 
      attendees = [], 
      location,
      isOnlineMeeting = false,
      calendarId,
      timeZone = 'UTC'
    } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!subject) {
      throw new Error('Event subject is required');
    }

    if (!startDateTime || !endDateTime) {
      throw new Error('Start and end date/time are required (ISO 8601 format)');
    }

    // Validate dates
    const startDate = new Date(startDateTime);
    const endDate = new Date(endDateTime);
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      throw new Error('Invalid date format. Use ISO 8601 format.');
    }

    if (endDate <= startDate) {
      throw new Error('End time must be after start time');
    }

    const eventData = {
      subject,
      body: {
        contentType: 'HTML',
        content: body || '',
      },
      start: {
        dateTime: startDateTime,
        timeZone: timeZone,
      },
      end: {
        dateTime: endDateTime,
        timeZone: timeZone,
      },
      attendees: attendees.map(email => ({
        emailAddress: {
          address: email,
        },
        type: 'required',
      })),
      isOnlineMeeting,
    };

    if (location) {
      eventData.location = {
        displayName: location,
      };
    }

    if (isOnlineMeeting) {
      eventData.onlineMeetingProvider = 'teamsForBusiness';
    }

    const endpoint = calendarId 
      ? `/me/calendars/${calendarId}/events`
      : '/me/calendar/events';

    const event = await this.executeGraphAPICall(
      accessToken,
      endpoint,
      'POST',
      eventData
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            eventId: event.id,
            subject: event.subject,
            start: event.start,
            end: event.end,
            joinUrl: event.onlineMeeting?.joinUrl,
            message: 'Calendar event created successfully',
          }, null, 2),
        },
      ],
    };
  }

  async updateCalendarEvent(args) {
    const { 
      eventId,
      calendarId,
      subject, 
      body, 
      startDateTime, 
      endDateTime, 
      attendees, 
      location,
      timeZone = 'UTC'
    } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!eventId) {
      throw new Error('eventId is required');
    }

    const updateData = {};

    if (subject) updateData.subject = subject;
    if (body) {
      updateData.body = {
        contentType: 'HTML',
        content: body,
      };
    }
    if (startDateTime) {
      updateData.start = {
        dateTime: startDateTime,
        timeZone: timeZone,
      };
    }
    if (endDateTime) {
      updateData.end = {
        dateTime: endDateTime,
        timeZone: timeZone,
      };
    }
    if (attendees) {
      updateData.attendees = attendees.map(email => ({
        emailAddress: {
          address: email,
        },
        type: 'required',
      }));
    }
    if (location) {
      updateData.location = {
        displayName: location,
      };
    }

    if (Object.keys(updateData).length === 0) {
      throw new Error('At least one field to update is required');
    }

    const endpoint = calendarId
      ? `/me/calendars/${calendarId}/events/${eventId}`
      : `/me/calendar/events/${eventId}`;

    const event = await this.executeGraphAPICall(
      accessToken,
      endpoint,
      'PATCH',
      updateData
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            eventId: event.id,
            message: 'Calendar event updated successfully',
          }, null, 2),
        },
      ],
    };
  }

  async deleteCalendarEvent(args) {
    const { eventId, calendarId } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!eventId) {
      throw new Error('eventId is required');
    }

    const endpoint = calendarId
      ? `/me/calendars/${calendarId}/events/${eventId}`
      : `/me/calendar/events/${eventId}`;

    await this.executeGraphAPICall(
      accessToken,
      endpoint,
      'DELETE'
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            message: 'Calendar event deleted successfully',
          }, null, 2),
        },
      ],
    };
  }

  async getUserProfile(args) {
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    const user = await this.executeGraphAPICall(accessToken, '/me?$select=id,displayName,mail,userPrincipalName,jobTitle,officeLocation,mobilePhone,businessPhones');

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            id: user.id,
            displayName: user.displayName,
            email: user.mail || user.userPrincipalName,
            jobTitle: user.jobTitle,
            officeLocation: user.officeLocation,
            mobilePhone: user.mobilePhone,
            businessPhones: user.businessPhones,
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
            description: "List user's recent Teams chats, ordered by most recent activity",
            inputSchema: {
              type: 'object',
              properties: {
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of chats to return (default: 10, max: 50)',
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
                  description: 'Maximum number of messages to return (default: 20, max: 50)',
                  default: 20,
                },
              },
              required: ['chatId'],
            },
          },
          {
            name: 'teams_find_chat_by_name',
            description: 'Find a Teams chat by its name/topic or participant name. Returns chat details including all members with their display names and email addresses.',
            inputSchema: {
              type: 'object',
              properties: {
                chatName: {
                  type: 'string',
                  description: 'The name/topic of the chat to find (e.g., "COE Team", "Project Alpha")',
                },
                participantName: {
                  type: 'string',
                  description: 'Name or email of a participant in the chat',
                },
              },
            },
          },
          {
            name: 'teams_create_one_on_one_chat',
            description: 'Create a new one-on-one chat with a specific user',
            inputSchema: {
              type: 'object',
              properties: {
                userId: {
                  type: 'string',
                  description: 'The Microsoft user ID of the person to chat with',
                },
                userEmail: {
                  type: 'string',
                  description: 'The email address of the person to chat with (alternative to userId)',
                },
              },
            },
          },
          {
            name: 'teams_send_message',
            description: 'Send a message to a Teams chat. Can send to group chats by name, or to one-on-one chats by participant. New chats are created automatically if needed.',
            inputSchema: {
              type: 'object',
              properties: {
                chatId: {
                  type: 'string',
                  description: 'The ID of the chat (if known)',
                },
                chatName: {
                  type: 'string',
                  description: 'The name/topic of a group chat',
                },
                participantName: {
                  type: 'string',
                  description: 'Name of a person for one-on-one chat',
                },
                participantEmail: {
                  type: 'string',
                  description: 'Email of a person for one-on-one chat (preferred)',
                },
                message: {
                  type: 'string',
                  description: 'The message content to send',
                },
              },
              required: ['message'],
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
                  description: 'The ID of the team',
                },
              },
              required: ['teamId'],
            },
          },
          {
            name: 'teams_get_channel_messages',
            description: 'Get messages from a Teams channel',
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
                  description: 'Maximum number of messages (default: 20)',
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
            description: 'Search messages across Teams chats',
            inputSchema: {
              type: 'object',
              properties: {
                query: {
                  type: 'string',
                  description: 'Search query',
                },
                from: {
                  type: 'string',
                  description: 'Filter by sender name',
                },
                after: {
                  type: 'string',
                  description: 'Filter after date (ISO 8601)',
                },
                before: {
                  type: 'string',
                  description: 'Filter before date (ISO 8601)',
                },
                maxResults: {
                  type: 'number',
                  description: 'Maximum results (default: 20)',
                  default: 20,
                },
              },
              required: ['query'],
            },
          },
          {
            name: 'teams_list_calendars',
            description: 'List all calendars for the user',
            inputSchema: {
              type: 'object',
              properties: {},
            },
          },
          {
            name: 'teams_list_calendar_events',
            description: 'List calendar events. Defaults to next 7 days.',
            inputSchema: {
              type: 'object',
              properties: {
                calendarId: {
                  type: 'string',
                  description: 'Optional calendar ID',
                },
                specificDate: {
                  type: 'string',
                  description: 'Get events for a specific date (YYYY-MM-DD)',
                },
                startDateTime: {
                  type: 'string',
                  description: 'Start of date range (ISO 8601)',
                },
                endDateTime: {
                  type: 'string',
                  description: 'End of date range (ISO 8601)',
                },
                maxResults: {
                  type: 'number',
                  description: 'Maximum events (default: 50)',
                  default: 50,
                },
              },
            },
          },
          {
            name: 'teams_get_calendar_event',
            description: 'Get details of a specific calendar event',
            inputSchema: {
              type: 'object',
              properties: {
                eventId: {
                  type: 'string',
                  description: 'The event ID',
                },
                calendarId: {
                  type: 'string',
                  description: 'Optional calendar ID',
                },
              },
              required: ['eventId'],
            },
          },
          {
            name: 'teams_create_calendar_event',
            description: 'Create a calendar event or Teams meeting',
            inputSchema: {
              type: 'object',
              properties: {
                subject: {
                  type: 'string',
                  description: 'Event title',
                },
                body: {
                  type: 'string',
                  description: 'Event description',
                },
                startDateTime: {
                  type: 'string',
                  description: 'Start time (ISO 8601)',
                },
                endDateTime: {
                  type: 'string',
                  description: 'End time (ISO 8601)',
                },
                attendees: {
                  type: 'array',
                  description: 'Attendee emails',
                  items: { type: 'string' },
                },
                location: {
                  type: 'string',
                  description: 'Location',
                },
                isOnlineMeeting: {
                  type: 'boolean',
                  description: 'Create Teams meeting',
                  default: false,
                },
                calendarId: {
                  type: 'string',
                  description: 'Optional calendar ID',
                },
                timeZone: {
                  type: 'string',
                  description: 'Time zone (default: UTC)',
                  default: 'UTC',
                },
              },
              required: ['subject', 'startDateTime', 'endDateTime'],
            },
          },
          {
            name: 'teams_update_calendar_event',
            description: 'Update a calendar event',
            inputSchema: {
              type: 'object',
              properties: {
                eventId: {
                  type: 'string',
                  description: 'Event ID to update',
                },
                calendarId: {
                  type: 'string',
                  description: 'Optional calendar ID',
                },
                subject: { type: 'string' },
                body: { type: 'string' },
                startDateTime: { type: 'string' },
                endDateTime: { type: 'string' },
                attendees: {
                  type: 'array',
                  items: { type: 'string' },
                },
                location: { type: 'string' },
                timeZone: {
                  type: 'string',
                  default: 'UTC',
                },
              },
              required: ['eventId'],
            },
          },
          {
            name: 'teams_delete_calendar_event',
            description: 'Delete a calendar event',
            inputSchema: {
              type: 'object',
              properties: {
                eventId: {
                  type: 'string',
                  description: 'Event ID to delete',
                },
                calendarId: {
                  type: 'string',
                  description: 'Optional calendar ID',
                },
              },
              required: ['eventId'],
            },
          },
          {
            name: 'teams_get_user_profile',
            description: 'Get the authenticated user profile',
            inputSchema: {
              type: 'object',
              properties: {},
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
          case 'teams_find_chat_by_name':
            return await this.findChatByName(args);
          case 'teams_list_messages':
            return await this.listMessages(args);
          case 'teams_create_one_on_one_chat':
            return await this.createOneOnOneChat(args);
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
          case 'teams_list_calendars':
            return await this.listCalendars(args);
          case 'teams_list_calendar_events':
            return await this.listCalendarEvents(args);
          case 'teams_get_calendar_event':
            return await this.getCalendarEvent(args);
          case 'teams_create_calendar_event':
            return await this.createCalendarEvent(args);
          case 'teams_update_calendar_event':
            return await this.updateCalendarEvent(args);
          case 'teams_delete_calendar_event':
            return await this.deleteCalendarEvent(args);
          case 'teams_get_user_profile':
            return await this.getUserProfile(args);
          default:
            throw new Error(`Unknown tool: ${name}`);
        }
      } catch (error) {
        console.error(`[Teams MCP] Error in ${name}:`, error.message);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: true,
                message: error.message,
                tool: name,
              }, null, 2),
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
      console.log('[Teams MCP] Shutting down...');
      await this.server.close();
      process.exit(0);
    });

    process.on('SIGTERM', async () => {
      console.log('[Teams MCP] Shutting down...');
      await this.server.close();
      process.exit(0);
    });

    process.on('uncaughtException', (error) => {
      console.error('[Teams MCP] Uncaught exception:', error);
    });

    process.on('unhandledRejection', (reason, promise) => {
      console.error('[Teams MCP] Unhandled rejection at:', promise, 'reason:', reason);
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