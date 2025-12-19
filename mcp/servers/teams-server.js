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

  async findChatByName(args) {
    const { chatName, participantName } = args;
    const { accessToken } = args._auth || {};

    if (!accessToken) {
      throw new Error('Access token is required');
    }

    if (!chatName && !participantName) {
      throw new Error('Either chatName or participantName is required');
    }

    // Get all chats
    const response = await this.executeGraphAPICall(
      accessToken,
      '/me/chats?$expand=members&$top=50'
    );

    const chats = response.value || [];
    const searchTerm = (chatName || participantName).toLowerCase();

    // Find matching chats
    const matches = chats.filter(chat => {
      // Match by topic/name
      if (chat.topic && chat.topic.toLowerCase().includes(searchTerm)) {
        return true;
      }

      // Match by participant name
      if (chat.members) {
        return chat.members.some(member => 
          member.displayName?.toLowerCase().includes(searchTerm) ||
          member.email?.toLowerCase().includes(searchTerm)
        );
      }

      return false;
    });

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            searchTerm: chatName || participantName,
            totalMatches: matches.length,
            matches: matches.map(chat => ({
              id: chat.id,
              topic: chat.topic || 'No topic',
              chatType: chat.chatType,
              members: chat.members?.map(m => ({
                displayName: m.displayName,
                email: m.email,
              })) || [],
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

    // If we have email but not userId, we need to look up the user
    let targetUserId = userId;
    if (!targetUserId && userEmail) {
      const userResponse = await this.executeGraphAPICall(
        accessToken,
        `/users/${userEmail}`
      );
      targetUserId = userResponse.id;
    }

    // Create a one-on-one chat
    const chatData = {
      chatType: 'oneOnOne',
      members: [
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
  if (!message || !message.trim()) throw new Error('Message body is required');

  /* --------------------------------------------------
     1️⃣ TEAM MESSAGE FLOW (HIGHEST PRIORITY)
  -------------------------------------------------- */
  if (teamName) {
    console.log('[Teams] Team name provided:', teamName);

    // Find team
    const teamsRes = await this.executeGraphAPICall(
      accessToken,
      `/me/joinedTeams`
    );

    const team = teamsRes.value?.find(
      t => t.displayName.toLowerCase() === teamName.toLowerCase()
    );

    if (!team) {
      throw new Error(`Team "${teamName}" not found or user not a member`);
    }

    // Get General channel
    const channelsRes = await this.executeGraphAPICall(
      accessToken,
      `/teams/${team.id}/channels`
    );

    const channel =
      channelsRes.value.find(c => c.displayName === 'General') ||
      channelsRes.value[0];

    if (!channel) {
      throw new Error(`No channels found in team "${teamName}"`);
    }

    // Send message to channel
    const response = await this.executeGraphAPICall(
      accessToken,
      `/teams/${team.id}/channels/${channel.id}/messages`,
      'POST',
      {
        body: {
          contentType: 'text',
          content: message
        }
      }
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            team: team.displayName,
            channel: channel.displayName,
            messageId: response.id
          }, null, 2)
        }
      ]
    };
  }

  /* --------------------------------------------------
     2️⃣ CHAT FLOW (CREATE OR FIND)
  -------------------------------------------------- */

  let targetChatId = chatId;

  // Find chat by name or participant
  if (!targetChatId && (chatName || participantName || participantEmail)) {
    const chatsRes = await this.executeGraphAPICall(
      accessToken,
      `/me/chats?$expand=members&$top=50`
    );

    const search = (chatName || participantName || participantEmail).toLowerCase();

    // If searching by participant (not chat name), prioritize one-on-one chats
    if (participantName || participantEmail) {
      // First, try to find a one-on-one chat with this participant
      const oneOnOneChat = chatsRes.value.find(c =>
        c.chatType === 'oneOnOne' &&
        c.members?.some(m =>
          m.displayName?.toLowerCase().includes(search) ||
          m.email?.toLowerCase().includes(search)
        )
      );

      if (oneOnOneChat) {
        targetChatId = oneOnOneChat.id;
        console.log('[Teams] Found one-on-one chat:', oneOnOneChat.id);
      }
    }

    // If no one-on-one chat found, or searching by chat name, find any matching chat
    if (!targetChatId) {
      const chat = chatsRes.value.find(c =>
        c.topic?.toLowerCase().includes(search) ||
        c.members?.some(m =>
          m.displayName?.toLowerCase().includes(search) ||
          m.email?.toLowerCase().includes(search)
        )
      );

      if (chat) {
        targetChatId = chat.id;
        console.log('[Teams] Found chat:', { id: chat.id, type: chat.chatType, topic: chat.topic });
      }
    }
  }

  // Create new chat if not found
  if (!targetChatId) {
    const memberBindings = [];

    for (const p of participants.length ? participants : [{ email: participantEmail }]) {
      const user = await this.executeGraphAPICall(
        accessToken,
        `/users/${encodeURIComponent(p.email)}`
      );

      memberBindings.push({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user.id}')`
      });
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
        content: message
      }
    }
  );

  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          success: true,
          chatId: targetChatId,
          messageId: response.id
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
      // If specific date provided, get events for that day only
      const date = new Date(specificDate);
      start = new Date(date.setHours(0, 0, 0, 0)).toISOString();
      end = new Date(date.setHours(23, 59, 59, 999)).toISOString();
    } else if (startDateTime && endDateTime) {
      // Use provided date range
      start = startDateTime;
      end = endDateTime;
    } else {
      // Default: Last 5 days to today
      const today = new Date();
      const fiveDaysAgo = new Date();
      fiveDaysAgo.setDate(today.getDate() - 5);
      fiveDaysAgo.setHours(0, 0, 0, 0);
      today.setHours(23, 59, 59, 999);
      
      start = fiveDaysAgo.toISOString();
      end = today.toISOString();
    }

    let endpoint = calendarId 
      ? `/me/calendars/${calendarId}/calendarView`
      : '/me/calendarView';

    // Use calendarView for date-based filtering (more efficient)
    const queryParams = [];
    queryParams.push(`startDateTime=${encodeURIComponent(start)}`);
    queryParams.push(`endDateTime=${encodeURIComponent(end)}`);
    queryParams.push(`$top=${maxResults}`);
    queryParams.push('$orderby=start/dateTime desc');

    endpoint += `?${queryParams.join('&')}`;

    const response = await this.executeGraphAPICall(accessToken, endpoint);

    const events = response.value || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalEvents: events.length,
            events: events.map(event => ({
              id: event.id,
              subject: event.subject,
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
              bodyPreview: event.bodyPreview,
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
            body: event.body?.content,
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

    if (!subject || !startDateTime || !endDateTime) {
      throw new Error('subject, startDateTime, and endDateTime are required');
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

    const user = await this.executeGraphAPICall(accessToken, '/me');

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
            name: 'teams_find_chat_by_name',
            description: 'Find a Teams chat by its name/topic or participant name',
            inputSchema: {
              type: 'object',
              properties: {
                chatName: {
                  type: 'string',
                  description: 'The name/topic of the chat to find',
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
            description: 'Send a message to a Teams chat. If the one-on-one chat does not exist, it will be created automatically when using participantEmail or participantName.',
            inputSchema: {
              type: 'object',
              properties: {
                chatId: {
                  type: 'string',
                  description: 'The ID of the chat (if known)',
                },
                chatName: {
                  type: 'string',
                  description: 'The name/topic of the chat to send message to (alternative to chatId)',
                },
                participantName: {
                  type: 'string',
                  description: 'Name of a participant to identify or create a one-on-one chat (alternative to chatId)',
                },
                participantEmail: {
                  type: 'string',
                  description: 'Email address of a participant to identify or create a one-on-one chat (alternative to chatId, preferred for one-on-one)',
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
            description: 'List events from user calendar. By default, fetches events from the last 5 days. Can filter by specific date or custom date range.',
            inputSchema: {
              type: 'object',
              properties: {
                calendarId: {
                  type: 'string',
                  description: 'Optional calendar ID. If omitted, uses default calendar',
                },
                specificDate: {
                  type: 'string',
                  description: 'Get events for a specific date only (ISO 8601 format: YYYY-MM-DD). If provided, startDateTime and endDateTime are ignored.',
                },
                startDateTime: {
                  type: 'string',
                  description: 'Custom start date for filtering (ISO 8601 format). Only used if specificDate is not provided.',
                },
                endDateTime: {
                  type: 'string',
                  description: 'Custom end date for filtering (ISO 8601 format). Only used if specificDate is not provided.',
                },
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of events to return (default: 50)',
                  default: 50,
                },
              },
            },
          },
          {
            name: 'teams_get_calendar_event',
            description: 'Get detailed information about a specific calendar event',
            inputSchema: {
              type: 'object',
              properties: {
                eventId: {
                  type: 'string',
                  description: 'The ID of the calendar event',
                },
                calendarId: {
                  type: 'string',
                  description: 'Optional calendar ID. If omitted, uses default calendar',
                },
              },
              required: ['eventId'],
            },
          },
          {
            name: 'teams_create_calendar_event',
            description: 'Create a new calendar event or Teams meeting',
            inputSchema: {
              type: 'object',
              properties: {
                subject: {
                  type: 'string',
                  description: 'Event title/subject',
                },
                body: {
                  type: 'string',
                  description: 'Event description/body (HTML supported)',
                },
                startDateTime: {
                  type: 'string',
                  description: 'Start date and time (ISO 8601 format)',
                },
                endDateTime: {
                  type: 'string',
                  description: 'End date and time (ISO 8601 format)',
                },
                attendees: {
                  type: 'array',
                  description: 'Array of attendee email addresses',
                  items: {
                    type: 'string',
                  },
                },
                location: {
                  type: 'string',
                  description: 'Physical location of the event',
                },
                isOnlineMeeting: {
                  type: 'boolean',
                  description: 'Create as Teams online meeting',
                  default: false,
                },
                calendarId: {
                  type: 'string',
                  description: 'Optional calendar ID. If omitted, uses default calendar',
                },
                timeZone: {
                  type: 'string',
                  description: 'Time zone (e.g., "UTC", "America/New_York")',
                  default: 'UTC',
                },
              },
              required: ['subject', 'startDateTime', 'endDateTime'],
            },
          },
          {
            name: 'teams_update_calendar_event',
            description: 'Update an existing calendar event',
            inputSchema: {
              type: 'object',
              properties: {
                eventId: {
                  type: 'string',
                  description: 'The ID of the event to update',
                },
                calendarId: {
                  type: 'string',
                  description: 'Optional calendar ID',
                },
                subject: {
                  type: 'string',
                  description: 'New event title/subject',
                },
                body: {
                  type: 'string',
                  description: 'New event description',
                },
                startDateTime: {
                  type: 'string',
                  description: 'New start date and time',
                },
                endDateTime: {
                  type: 'string',
                  description: 'New end date and time',
                },
                attendees: {
                  type: 'array',
                  description: 'Updated attendee list',
                  items: {
                    type: 'string',
                  },
                },
                location: {
                  type: 'string',
                  description: 'New location',
                },
                timeZone: {
                  type: 'string',
                  description: 'Time zone',
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
                  description: 'The ID of the event to delete',
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
            description: 'Get the authenticated user profile information',
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
