import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';

// Constants for retry logic and rate limiting
const MAX_RETRIES = 10;
const RETRY_DELAY_MS = 1000;
const RATE_LIMIT_DELAY_MS = 5000;

// Gmail category label mappings
const GMAIL_CATEGORIES = {
  primary: 'CATEGORY_PERSONAL',
  social: 'CATEGORY_SOCIAL',
  promotions: 'CATEGORY_PROMOTIONS',
  updates: 'CATEGORY_UPDATES',
  forums: 'CATEGORY_FORUMS',
};

// Common Gmail system labels
const GMAIL_SYSTEM_LABELS = {
  inbox: 'INBOX',
  sent: 'SENT',
  drafts: 'DRAFT',
  spam: 'SPAM',
  trash: 'TRASH',
  unread: 'UNREAD',
  starred: 'STARRED',
  important: 'IMPORTANT',
};

// Filter templates for common scenarios
const FILTER_TEMPLATES = {
  fromSender: (senderEmail, labelIds = [], archive = false) => ({
    criteria: { from: senderEmail },
    action: {
      addLabelIds: labelIds,
      removeLabelIds: archive ? ['INBOX'] : undefined,
    },
  }),
  withSubject: (subjectText, labelIds = [], markAsRead = false) => ({
    criteria: { subject: subjectText },
    action: {
      addLabelIds: labelIds,
      removeLabelIds: markAsRead ? ['UNREAD'] : undefined,
    },
  }),
  largeEmails: (sizeInBytes = 10485760, labelIds = []) => ({
    criteria: { size: sizeInBytes, sizeComparison: 'larger' },
    action: { addLabelIds: labelIds },
  }),
  containingText: (searchText, labelIds = [], markImportant = false) => ({
    criteria: { query: searchText },
    action: {
      addLabelIds: markImportant ? [...labelIds, 'IMPORTANT'] : labelIds,
    },
  }),
  mailingList: (listIdentifier, labelIds = [], archive = true) => ({
    criteria: { query: `list:${listIdentifier} OR subject:[${listIdentifier}]` },
    action: {
      addLabelIds: labelIds,
      removeLabelIds: archive ? ['INBOX'] : undefined,
    },
  }),
};

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

    // Cache for reducing API calls
    this.cache = {
      labels: null,
      labelsExpiry: 0,
      profile: null,
      profileExpiry: 0,
    };
    this.CACHE_TTL = 300000; // 5 minutes cache

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
      labels: null,
      labelsExpiry: 0,
      profile: null,
      profileExpiry: 0,
    };
  }

  /**
   * Create authenticated Gmail client
   */
  createGmailClient(accessToken) {
    if (!accessToken) {
      throw new Error('Access token is required. Please sign in with Google.');
    }

    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: accessToken });

    return google.gmail({ version: 'v1', auth: oauth2Client });
  }

  /**
   * Convert category name to Gmail label ID
   */
  getCategoryLabelId(category) {
    if (!category) return null;
    const normalized = category.toLowerCase().trim();
    return GMAIL_CATEGORIES[normalized] || null;
  }

  /**
   * Convert system label name to Gmail label ID
   */
  getSystemLabelId(label) {
    if (!label) return null;
    const normalized = label.toLowerCase().trim();
    return GMAIL_SYSTEM_LABELS[normalized] || label.toUpperCase();
  }

  /**
   * Build label IDs array from category and other labels
   */
  buildLabelIds(category, labelIds = [], includeInbox = true) {
    const labels = new Set();
    
    // Add inbox by default for category queries
    if (includeInbox && category) {
      labels.add('INBOX');
    }
    
    // Add category label
    const categoryLabel = this.getCategoryLabelId(category);
    if (categoryLabel) {
      labels.add(categoryLabel);
    }
    
    // Add any additional labels
    if (labelIds && Array.isArray(labelIds)) {
      labelIds.forEach(label => {
        const systemLabel = this.getSystemLabelId(label);
        if (systemLabel) {
          labels.add(systemLabel);
        }
      });
    }
    
    return Array.from(labels);
  }

  /**
   * Execute Gmail API call with retry logic
   */
  async executeWithRetry(apiCall, retryCount = 0) {
    try {
      return await apiCall();
    } catch (error) {
      const statusCode = error.code || error.response?.status;

      // Handle rate limiting with exponential backoff
      if (statusCode === 429) {
        if (retryCount < MAX_RETRIES) {
          const retryAfter = RATE_LIMIT_DELAY_MS * Math.pow(2, retryCount);
          console.warn(`Rate limited. Retrying after ${retryAfter}ms (attempt ${retryCount + 1}/${MAX_RETRIES})`);
          await this.delay(retryAfter);
          return this.executeWithRetry(apiCall, retryCount + 1);
        }
        throw new Error('Gmail API rate limit exceeded. Please try again later.');
      }

      // Handle transient errors with retry
      if ((statusCode === 503 || statusCode === 500) && retryCount < MAX_RETRIES) {
        const retryDelay = RETRY_DELAY_MS * Math.pow(2, retryCount);
        console.warn(`Service error. Retrying after ${retryDelay}ms (attempt ${retryCount + 1}/${MAX_RETRIES})`);
        await this.delay(retryDelay);
        return this.executeWithRetry(apiCall, retryCount + 1);
      }

      // Handle specific errors with user-friendly messages
      if (statusCode === 401) {
        this.clearCache();
        throw new Error('Your Google session has expired. Please sign in again.');
      } else if (statusCode === 403) {
        throw new Error("You don't have permission to access this Gmail resource. Please check your permissions.");
      } else if (statusCode === 404) {
        throw new Error('The requested email or resource was not found.');
      } else if (statusCode === 400) {
        throw new Error(`Invalid request: ${error.message}`);
      }

      // Log and rethrow
      console.error('Gmail API Error:', {
        statusCode,
        message: error.message,
      });
      throw new Error(`Gmail API error: ${error.message || 'Unknown error occurred'}`);
    }
  }

  /**
   * Safely extract email header value
   */
  getHeader(headers, name) {
    if (!headers || !Array.isArray(headers)) return '';
    const header = headers.find(h => h.name?.toLowerCase() === name.toLowerCase());
    return header?.value || '';
  }

  /**
   * Extract plain text body from email payload
   */
  extractBody(payload) {
    if (!payload) return '';

    // Direct body data
    if (payload.body?.data) {
      try {
        return Buffer.from(payload.body.data, 'base64').toString('utf-8');
      } catch {
        return '';
      }
    }

    // Multi-part message
    if (payload.parts && Array.isArray(payload.parts)) {
      // Prefer plain text over HTML
      const textPart = payload.parts.find(p => p.mimeType === 'text/plain');
      if (textPart?.body?.data) {
        try {
          return Buffer.from(textPart.body.data, 'base64').toString('utf-8');
        } catch {
          return '';
        }
      }

      // Fall back to HTML
      const htmlPart = payload.parts.find(p => p.mimeType === 'text/html');
      if (htmlPart?.body?.data) {
        try {
          const html = Buffer.from(htmlPart.body.data, 'base64').toString('utf-8');
          // Strip HTML tags for plain text representation
          return html
            .replace(/<[^>]*>/g, '')
            .replace(/&nbsp;/g, ' ')
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .trim();
        } catch {
          return '';
        }
      }

      // Recursively check nested parts
      for (const part of payload.parts) {
        if (part.parts) {
          const nestedBody = this.extractBody(part);
          if (nestedBody) return nestedBody;
        }
      }
    }

    return '';
  }

  /**
   * Detect category from label IDs
   */
  detectCategory(labelIds) {
    if (!labelIds || !Array.isArray(labelIds)) return null;
    
    for (const [category, labelId] of Object.entries(GMAIL_CATEGORIES)) {
      if (labelIds.includes(labelId)) {
        return category;
      }
    }
    return null;
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

  /**
   * Parse email address from "Name <email@domain.com>" format
   */
  parseEmailAddress(addressString) {
    if (!addressString) return { name: '', email: '' };
    
    const match = addressString.match(/^(.+?)\s*<(.+?)>$/);
    if (match) {
      return {
        name: match[1].trim().replace(/^["']|["']$/g, ''),
        email: match[2].trim(),
      };
    }
    
    return {
      name: '',
      email: addressString.trim(),
    };
  }

  setupTools() {
    // List available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: 'gmail_list_messages',
            description: 'List emails from Gmail inbox with optional filters. Supports filtering by category (Primary, Social, Promotions, Updates, Forums).',
            inputSchema: {
              type: 'object',
              properties: {
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of messages to return (default: 10, max: 50)',
                  default: 10,
                },
                category: {
                  type: 'string',
                  enum: ['primary', 'social', 'promotions', 'updates', 'forums'],
                  description: 'Filter by Gmail category/tab (e.g., "primary", "social", "promotions", "updates", "forums")',
                },
                query: {
                  type: 'string',
                  description: 'Gmail search query (e.g., "from:user@example.com", "is:unread", "after:2024/01/01", "subject:meeting")',
                },
                labelIds: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Filter by label IDs (e.g., ["INBOX", "UNREAD", "STARRED"])',
                },
                includeSpamTrash: {
                  type: 'boolean',
                  description: 'Include messages from SPAM and TRASH (default: false)',
                  default: false,
                },
              },
            },
          },
          {
            name: 'gmail_list_by_category',
            description: 'List emails from a specific Gmail category tab (Primary, Social, Promotions, Updates, Forums). This is a convenience method for fetching emails by tab.',
            inputSchema: {
              type: 'object',
              properties: {
                category: {
                  type: 'string',
                  enum: ['primary', 'social', 'promotions', 'updates', 'forums'],
                  description: 'The Gmail category to fetch (required)',
                },
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of messages to return (default: 10, max: 50)',
                  default: 10,
                },
                unreadOnly: {
                  type: 'boolean',
                  description: 'Only fetch unread messages (default: false)',
                  default: false,
                },
              },
              required: ['category'],
            },
          },
          {
            name: 'gmail_get_message',
            description: 'Get full details of a specific email by ID, including the complete body content.',
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
            description: 'Search emails with specific criteria like sender, subject, date range, category, etc.',
            inputSchema: {
              type: 'object',
              properties: {
                from: {
                  type: 'string',
                  description: 'Filter by sender email or name',
                },
                to: {
                  type: 'string',
                  description: 'Filter by recipient email',
                },
                subject: {
                  type: 'string',
                  description: 'Filter by subject keywords',
                },
                body: {
                  type: 'string',
                  description: 'Search within email body',
                },
                category: {
                  type: 'string',
                  enum: ['primary', 'social', 'promotions', 'updates', 'forums'],
                  description: 'Filter by Gmail category',
                },
                after: {
                  type: 'string',
                  description: 'Date after (YYYY/MM/DD format)',
                },
                before: {
                  type: 'string',
                  description: 'Date before (YYYY/MM/DD format)',
                },
                hasAttachment: {
                  type: 'boolean',
                  description: 'Filter messages with attachments',
                },
                isUnread: {
                  type: 'boolean',
                  description: 'Filter unread messages only',
                },
                isStarred: {
                  type: 'boolean',
                  description: 'Filter starred messages only',
                },
                isImportant: {
                  type: 'boolean',
                  description: 'Filter important messages only',
                },
                label: {
                  type: 'string',
                  description: 'Filter by label name',
                },
                maxResults: {
                  type: 'number',
                  description: 'Maximum results (default: 20, max: 50)',
                  default: 20,
                },
              },
            },
          },
          {
            name: 'gmail_send_message',
            description: 'Send an email via Gmail. Supports plain text emails with optional CC recipients.',
            inputSchema: {
              type: 'object',
              properties: {
                to: {
                  type: 'string',
                  description: 'Recipient email address (required)',
                },
                subject: {
                  type: 'string',
                  description: 'Email subject line (required)',
                },
                body: {
                  type: 'string',
                  description: 'Email body content - plain text (required)',
                },
                cc: {
                  type: 'string',
                  description: 'CC recipients (comma-separated email addresses)',
                },
                bcc: {
                  type: 'string',
                  description: 'BCC recipients (comma-separated email addresses)',
                },
                replyTo: {
                  type: 'string',
                  description: 'Reply-To email address',
                },
              },
              required: ['to', 'subject', 'body'],
            },
          },
          {
            name: 'gmail_reply_to_message',
            description: 'Reply to an existing email thread',
            inputSchema: {
              type: 'object',
              properties: {
                messageId: {
                  type: 'string',
                  description: 'The ID of the message to reply to',
                },
                body: {
                  type: 'string',
                  description: 'Reply body content',
                },
                replyAll: {
                  type: 'boolean',
                  description: 'Reply to all recipients (default: false)',
                  default: false,
                },
              },
              required: ['messageId', 'body'],
            },
          },
          {
            name: 'gmail_get_thread',
            description: 'Get all messages in an email thread/conversation',
            inputSchema: {
              type: 'object',
              properties: {
                threadId: {
                  type: 'string',
                  description: 'The ID of the thread to retrieve',
                },
              },
              required: ['threadId'],
            },
          },
          {
            name: 'gmail_modify_labels',
            description: 'Add or remove labels from a message (mark as read/unread, star, archive, etc.)',
            inputSchema: {
              type: 'object',
              properties: {
                messageId: {
                  type: 'string',
                  description: 'The ID of the message to modify',
                },
                addLabelIds: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Label IDs to add (e.g., ["STARRED", "IMPORTANT"])',
                },
                removeLabelIds: {
                  type: 'array',
                  items: { type: 'string' },
                  description: 'Label IDs to remove (e.g., ["UNREAD", "INBOX"])',
                },
              },
              required: ['messageId'],
            },
          },
          {
            name: 'gmail_trash_message',
            description: 'Move a message to trash',
            inputSchema: {
              type: 'object',
              properties: {
                messageId: {
                  type: 'string',
                  description: 'The ID of the message to trash',
                },
              },
              required: ['messageId'],
            },
          },
          {
            name: 'gmail_list_labels',
            description: 'List all available Gmail labels/folders including categories',
            inputSchema: {
              type: 'object',
              properties: {},
            },
          },
          {
            name: 'gmail_get_profile',
            description: 'Get the Gmail profile information (email address, total messages, etc.)',
            inputSchema: {
              type: 'object',
              properties: {},
            },
          },
          {
            name: 'gmail_create_draft',
            description: 'Create a draft email',
            inputSchema: {
              type: 'object',
              properties: {
                to: {
                  type: 'string',
                  description: 'Recipient email address',
                },
                subject: {
                  type: 'string',
                  description: 'Email subject line',
                },
                body: {
                  type: 'string',
                  description: 'Email body content',
                },
                cc: {
                  type: 'string',
                  description: 'CC recipients (comma-separated)',
                },
              },
              required: ['to', 'subject', 'body'],
            },
          },
          {
            name: 'gmail_list_drafts',
            description: 'List all draft emails',
            inputSchema: {
              type: 'object',
              properties: {
                maxResults: {
                  type: 'number',
                  description: 'Maximum number of drafts to return (default: 10)',
                  default: 10,
                },
              },
            },
          },
          // Label Management Tools
          {
            name: 'gmail_create_label',
            description: 'Creates a new Gmail label',
            inputSchema: {
              type: 'object',
              properties: {
                name: {
                  type: 'string',
                  description: 'Name for the new label',
                },
                messageListVisibility: {
                  type: 'string',
                  enum: ['show', 'hide'],
                  description: 'Whether to show or hide the label in the message list',
                },
                labelListVisibility: {
                  type: 'string',
                  enum: ['labelShow', 'labelShowIfUnread', 'labelHide'],
                  description: 'Visibility of the label in the label list',
                },
              },
              required: ['name'],
            },
          },
          {
            name: 'gmail_update_label',
            description: 'Updates an existing Gmail label',
            inputSchema: {
              type: 'object',
              properties: {
                id: {
                  type: 'string',
                  description: 'ID of the label to update',
                },
                name: {
                  type: 'string',
                  description: 'New name for the label',
                },
                messageListVisibility: {
                  type: 'string',
                  enum: ['show', 'hide'],
                  description: 'Whether to show or hide the label in the message list',
                },
                labelListVisibility: {
                  type: 'string',
                  enum: ['labelShow', 'labelShowIfUnread', 'labelHide'],
                  description: 'Visibility of the label in the label list',
                },
              },
              required: ['id'],
            },
          },
          {
            name: 'gmail_delete_label',
            description: 'Deletes a Gmail label',
            inputSchema: {
              type: 'object',
              properties: {
                id: {
                  type: 'string',
                  description: 'ID of the label to delete',
                },
              },
              required: ['id'],
            },
          },
          {
            name: 'gmail_get_or_create_label',
            description: 'Gets an existing label by name or creates it if it does not exist',
            inputSchema: {
              type: 'object',
              properties: {
                name: {
                  type: 'string',
                  description: 'Name of the label to get or create',
                },
                messageListVisibility: {
                  type: 'string',
                  enum: ['show', 'hide'],
                  description: 'Whether to show or hide the label in the message list',
                },
                labelListVisibility: {
                  type: 'string',
                  enum: ['labelShow', 'labelShowIfUnread', 'labelHide'],
                  description: 'Visibility of the label in the label list',
                },
              },
              required: ['name'],
            },
          },
          {
            name: 'gmail_delete_message',
            description: 'Permanently deletes an email (cannot be undone)',
            inputSchema: {
              type: 'object',
              properties: {
                messageId: {
                  type: 'string',
                  description: 'The ID of the message to permanently delete',
                },
              },
              required: ['messageId'],
            },
          },
          // Filter Management Tools
          {
            name: 'gmail_create_filter',
            description: 'Creates a new Gmail filter with custom criteria and actions',
            inputSchema: {
              type: 'object',
              properties: {
                criteria: {
                  type: 'object',
                  description: 'Criteria for matching emails',
                  properties: {
                    from: { type: 'string', description: 'Sender email address to match' },
                    to: { type: 'string', description: 'Recipient email address to match' },
                    subject: { type: 'string', description: 'Subject text to match' },
                    query: { type: 'string', description: 'Gmail search query' },
                    negatedQuery: { type: 'string', description: 'Text that must NOT be present' },
                    hasAttachment: { type: 'boolean', description: 'Whether to match emails with attachments' },
                    excludeChats: { type: 'boolean', description: 'Whether to exclude chat messages' },
                    size: { type: 'number', description: 'Email size in bytes' },
                    sizeComparison: { type: 'string', enum: ['smaller', 'larger'], description: 'Size comparison operator' },
                  },
                },
                action: {
                  type: 'object',
                  description: 'Actions to perform on matching emails',
                  properties: {
                    addLabelIds: { type: 'array', items: { type: 'string' }, description: 'Label IDs to add to matching emails' },
                    removeLabelIds: { type: 'array', items: { type: 'string' }, description: 'Label IDs to remove from matching emails' },
                    forward: { type: 'string', description: 'Email address to forward matching emails to' },
                  },
                },
              },
              required: ['criteria', 'action'],
            },
          },
          {
            name: 'gmail_list_filters',
            description: 'Retrieves all Gmail filters',
            inputSchema: {
              type: 'object',
              properties: {},
            },
          },
          {
            name: 'gmail_get_filter',
            description: 'Gets details of a specific Gmail filter',
            inputSchema: {
              type: 'object',
              properties: {
                filterId: {
                  type: 'string',
                  description: 'ID of the filter to retrieve',
                },
              },
              required: ['filterId'],
            },
          },
          {
            name: 'gmail_delete_filter',
            description: 'Deletes a Gmail filter',
            inputSchema: {
              type: 'object',
              properties: {
                filterId: {
                  type: 'string',
                  description: 'ID of the filter to delete',
                },
              },
              required: ['filterId'],
            },
          },
          {
            name: 'gmail_create_filter_from_template',
            description: 'Creates a filter using a pre-defined template for common scenarios',
            inputSchema: {
              type: 'object',
              properties: {
                template: {
                  type: 'string',
                  enum: ['fromSender', 'withSubject', 'withAttachments', 'largeEmails', 'containingText', 'mailingList'],
                  description: 'Pre-defined filter template to use',
                },
                parameters: {
                  type: 'object',
                  description: 'Template-specific parameters',
                  properties: {
                    senderEmail: { type: 'string', description: 'Sender email (for fromSender template)' },
                    subjectText: { type: 'string', description: 'Subject text (for withSubject template)' },
                    searchText: { type: 'string', description: 'Text to search for (for containingText template)' },
                    listIdentifier: { type: 'string', description: 'Mailing list identifier (for mailingList template)' },
                    sizeInBytes: { type: 'number', description: 'Size threshold in bytes (for largeEmails template)' },
                    labelIds: { type: 'array', items: { type: 'string' }, description: 'Label IDs to apply' },
                    archive: { type: 'boolean', description: 'Whether to archive (skip inbox)' },
                    markAsRead: { type: 'boolean', description: 'Whether to mark as read' },
                    markImportant: { type: 'boolean', description: 'Whether to mark as important' },
                  },
                },
              },
              required: ['template', 'parameters'],
            },
          },
          {
            name: 'gmail_send_draft',
            description: 'Sends an existing draft email',
            inputSchema: {
              type: 'object',
              properties: {
                draftId: {
                  type: 'string',
                  description: 'The ID of the draft to send',
                },
              },
              required: ['draftId'],
            },
          },
          {
            name: 'gmail_delete_draft',
            description: 'Deletes a draft email',
            inputSchema: {
              type: 'object',
              properties: {
                draftId: {
                  type: 'string',
                  description: 'The ID of the draft to delete',
                },
              },
              required: ['draftId'],
            },
          },
          {
            name: 'gmail_update_draft',
            description: 'Updates an existing draft email',
            inputSchema: {
              type: 'object',
              properties: {
                draftId: {
                  type: 'string',
                  description: 'The ID of the draft to update',
                },
                to: {
                  type: 'string',
                  description: 'New recipient email address',
                },
                subject: {
                  type: 'string',
                  description: 'New email subject line',
                },
                body: {
                  type: 'string',
                  description: 'New email body content',
                },
                cc: {
                  type: 'string',
                  description: 'New CC recipients (comma-separated)',
                },
              },
              required: ['draftId'],
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
          case 'gmail_list_by_category':
            return await this.listByCategory(args);
          case 'gmail_get_message':
            return await this.getMessage(args);
          case 'gmail_search_messages':
            return await this.searchMessages(args);
          case 'gmail_send_message':
            return await this.sendMessage(args);
          case 'gmail_reply_to_message':
            return await this.replyToMessage(args);
          case 'gmail_get_thread':
            return await this.getThread(args);
          case 'gmail_modify_labels':
            return await this.modifyLabels(args);
          case 'gmail_trash_message':
            return await this.trashMessage(args);
          case 'gmail_delete_message':
            return await this.deleteMessage(args);
          case 'gmail_list_labels':
            return await this.listLabels(args);
          case 'gmail_get_profile':
            return await this.getProfile(args);
          case 'gmail_create_draft':
            return await this.createDraft(args);
          case 'gmail_list_drafts':
            return await this.listDrafts(args);
          case 'gmail_send_draft':
            return await this.sendDraft(args);
          case 'gmail_delete_draft':
            return await this.deleteDraft(args);
          case 'gmail_update_draft':
            return await this.updateDraft(args);
          // Label management
          case 'gmail_create_label':
            return await this.createLabel(args);
          case 'gmail_update_label':
            return await this.updateLabel(args);
          case 'gmail_delete_label':
            return await this.deleteLabel(args);
          case 'gmail_get_or_create_label':
            return await this.getOrCreateLabel(args);
          // Filter management
          case 'gmail_create_filter':
            return await this.createFilter(args);
          case 'gmail_list_filters':
            return await this.listFilters(args);
          case 'gmail_get_filter':
            return await this.getFilter(args);
          case 'gmail_delete_filter':
            return await this.deleteFilter(args);
          case 'gmail_create_filter_from_template':
            return await this.createFilterFromTemplate(args);
          default:
            throw new Error(`Unknown tool: ${name}`);
        }
      } catch (error) {
        console.error(`[Gmail MCP] Error in ${name}:`, error.message);
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

  async listMessages(args) {
    const { maxResults = 10, query = '', category, labelIds, includeSpamTrash = false } = args;
    const { accessToken } = args._auth || {};

    const gmail = this.createGmailClient(accessToken);

    const listParams = {
      userId: 'me',
      maxResults: Math.min(maxResults, 50),
      includeSpamTrash,
    };

    // Build query string - use category:X for accurate category filtering
    let finalQuery = query || '';
    
    // Add category filter to query (this is more accurate than using labelIds)
    // Use "in:inbox category:X" for accurate Gmail tab filtering
    if (category) {
      const normalizedCategory = category.toLowerCase().trim();
      const categoryQuery = `in:inbox category:${normalizedCategory}`;
      finalQuery = finalQuery ? `${finalQuery} ${categoryQuery}` : categoryQuery;
    }
    
    if (finalQuery) {
      listParams.q = finalQuery;
    }

    // Only use labelIds for non-category filters (like UNREAD, STARRED, etc.)
    // Don't use labelIds when category is specified to avoid conflicts
    if (!category && labelIds && Array.isArray(labelIds) && labelIds.length > 0) {
      const normalizedLabels = labelIds.map(l => this.getSystemLabelId(l));
      listParams.labelIds = normalizedLabels;
    }

    console.log('[Gmail] Listing messages with params:', JSON.stringify({
      ...listParams,
      category,
    }, null, 2));

    const response = await this.executeWithRetry(() => 
      gmail.users.messages.list(listParams)
    );

    // Log raw API response
    console.log('[Gmail] Raw list response:', JSON.stringify(response.data, null, 2));

    const messages = response.data.messages || [];

    if (messages.length === 0) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              totalMessages: 0,
              messages: [],
              category: category || 'all',
              message: `No messages found${category ? ` in ${category}` : ''} matching your criteria`,
            }, null, 2),
          },
        ],
      };
    }

    // Fetch details for each message (limit for performance)
    const fetchLimit = Math.min(messages.length, maxResults);
    const detailedMessages = await Promise.all(
      messages.slice(0, fetchLimit).map(async (msg) => {
        try {
          const detail = await this.executeWithRetry(() =>
            gmail.users.messages.get({
              userId: 'me',
              id: msg.id,
              format: 'metadata',
              metadataHeaders: ['From', 'To', 'Subject', 'Date', 'Cc'],
            })
          );

          // Log raw message detail from API
          console.log(`[Gmail] Raw message detail for ${msg.id}:`, JSON.stringify(detail.data, null, 2));

          const headers = detail.data.payload?.headers || [];
          const fromParsed = this.parseEmailAddress(this.getHeader(headers, 'From'));
          const detectedCategory = this.detectCategory(detail.data.labelIds);

          return {
            id: msg.id,
            threadId: msg.threadId,
            from: this.getHeader(headers, 'From'),
            fromName: fromParsed.name,
            fromEmail: fromParsed.email,
            to: this.getHeader(headers, 'To'),
            cc: this.getHeader(headers, 'Cc'),
            subject: this.getHeader(headers, 'Subject') || '(No Subject)',
            date: this.getHeader(headers, 'Date'),
            snippet: detail.data.snippet || '',
            labelIds: detail.data.labelIds || [],
            category: detectedCategory,
            isUnread: detail.data.labelIds?.includes('UNREAD') || false,
            isStarred: detail.data.labelIds?.includes('STARRED') || false,
            isImportant: detail.data.labelIds?.includes('IMPORTANT') || false,
          };
        } catch (err) {
          console.error(`Error fetching message ${msg.id}:`, err.message);
          return null;
        }
      })
    );

    const validMessages = detailedMessages.filter(m => m !== null);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalMessages: messages.length,
            returnedMessages: validMessages.length,
            category: category || 'all',
            messages: validMessages,
          }, null, 2),
        },
      ],
    };
  }

  async listByCategory(args) {
    const { category, maxResults = 10, unreadOnly = false } = args;
    const { accessToken } = args._auth || {};

    if (!category) {
      throw new Error('Category is required. Choose from: primary, social, promotions, updates, forums');
    }

    // Validate category
    const validCategories = ['primary', 'social', 'promotions', 'updates', 'forums'];
    if (!validCategories.includes(category.toLowerCase())) {
      throw new Error(`Invalid category: ${category}. Choose from: ${validCategories.join(', ')}`);
    }

    // Build query for category - this is the most reliable way
    let query = '';
    if (unreadOnly) {
      query = 'is:unread';
    }

    return await this.listMessages({
      maxResults,
      category: category.toLowerCase(),
      query,
      _auth: { accessToken },
    });
  }

  async getMessage(args) {
    const { messageId } = args;
    const { accessToken } = args._auth || {};

    if (!messageId) {
      throw new Error('messageId is required');
    }

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.messages.get({
        userId: 'me',
        id: messageId,
        format: 'full',
      })
    );

    const message = response.data;
    const headers = message.payload?.headers || [];
    const fromParsed = this.parseEmailAddress(this.getHeader(headers, 'From'));

    // Extract body content
    const body = this.extractBody(message.payload);

    // Detect category
    const category = this.detectCategory(message.labelIds);

    // Extract attachments info
    const attachments = [];
    if (message.payload?.parts) {
      for (const part of message.payload.parts) {
        if (part.filename && part.body?.attachmentId) {
          attachments.push({
            filename: part.filename,
            mimeType: part.mimeType,
            size: part.body.size,
            attachmentId: part.body.attachmentId,
          });
        }
      }
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            id: message.id,
            threadId: message.threadId,
            from: this.getHeader(headers, 'From'),
            fromName: fromParsed.name,
            fromEmail: fromParsed.email,
            to: this.getHeader(headers, 'To'),
            cc: this.getHeader(headers, 'Cc'),
            bcc: this.getHeader(headers, 'Bcc'),
            subject: this.getHeader(headers, 'Subject') || '(No Subject)',
            date: this.getHeader(headers, 'Date'),
            body: body.substring(0, 5000), // Limit body length
            snippet: message.snippet,
            labelIds: message.labelIds || [],
            category: category,
            isUnread: message.labelIds?.includes('UNREAD') || false,
            isStarred: message.labelIds?.includes('STARRED') || false,
            isImportant: message.labelIds?.includes('IMPORTANT') || false,
            hasAttachments: attachments.length > 0,
            attachments: attachments,
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
      body,
      category,
      after,
      before,
      hasAttachment,
      isUnread,
      isStarred,
      isImportant,
      label,
      maxResults = 20,
    } = args;

    // Build Gmail search query
    const queryParts = [];
    if (from) queryParts.push(`from:${from}`);
    if (to) queryParts.push(`to:${to}`);
    if (subject) queryParts.push(`subject:${subject}`);
    if (body) queryParts.push(`"${body}"`);
    // Don't add category here - it will be handled by listMessages
    if (after) queryParts.push(`after:${after}`);
    if (before) queryParts.push(`before:${before}`);
    if (hasAttachment) queryParts.push('has:attachment');
    if (isUnread) queryParts.push('is:unread');
    if (isStarred) queryParts.push('is:starred');
    if (isImportant) queryParts.push('is:important');
    if (label) queryParts.push(`label:${label}`);

    const query = queryParts.join(' ');

    return await this.listMessages({ 
      maxResults: Math.min(maxResults, 50), 
      query,
      category: category ? category.toLowerCase() : undefined,
      _auth: args._auth 
    });
  }

  async sendMessage(args) {
    const { to, subject, body, cc, bcc, replyTo } = args;
    const { accessToken } = args._auth || {};

    console.log('[Gmail Server] sendMessage called with args:', {
      to,
      subject,
      body: body ? `"${body.substring(0, 50)}..."` : 'NULL/EMPTY',
      cc,
      bcc,
      hasAccessToken: !!accessToken,
    });

    // Validate required fields
    if (!to || to.trim() === '') {
      throw new Error('Recipient email address (to) is required');
    }

    if (!subject) {
      throw new Error('Email subject is required');
    }

    if (!body || body.trim() === '') {
      throw new Error('Email body is required and cannot be empty');
    }

    // Basic email validation
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const toEmails = to.split(',').map(e => e.trim());
    for (const email of toEmails) {
      if (!emailRegex.test(email)) {
        throw new Error(`Invalid email address: ${email}`);
      }
    }

    const gmail = this.createGmailClient(accessToken);

    // Create email with proper MIME format
    const emailLines = [
      `To: ${to}`,
      cc ? `Cc: ${cc}` : null,
      bcc ? `Bcc: ${bcc}` : null,
      replyTo ? `Reply-To: ${replyTo}` : null,
      `Subject: ${subject}`,
      'Content-Type: text/plain; charset=utf-8',
      'MIME-Version: 1.0',
      '',
      body,
    ].filter(line => line !== null);

    const email = emailLines.join('\r\n');

    const encodedEmail = Buffer.from(email)
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');

    const response = await this.executeWithRetry(() =>
      gmail.users.messages.send({
        userId: 'me',
        requestBody: {
          raw: encodedEmail,
        },
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: response.data.id,
            threadId: response.data.threadId,
            to: to,
            subject: subject,
            message: 'Email sent successfully',
          }, null, 2),
        },
      ],
    };
  }

  async replyToMessage(args) {
    const { messageId, body, replyAll = false } = args;
    const { accessToken } = args._auth || {};

    if (!messageId) {
      throw new Error('messageId is required');
    }

    if (!body || body.trim() === '') {
      throw new Error('Reply body is required');
    }

    const gmail = this.createGmailClient(accessToken);

    // Get the original message
    const original = await this.executeWithRetry(() =>
      gmail.users.messages.get({
        userId: 'me',
        id: messageId,
        format: 'metadata',
        metadataHeaders: ['From', 'To', 'Cc', 'Subject', 'Message-ID', 'References'],
      })
    );

    const headers = original.data.payload?.headers || [];
    const originalFrom = this.getHeader(headers, 'From');
    const originalTo = this.getHeader(headers, 'To');
    const originalCc = this.getHeader(headers, 'Cc');
    const originalSubject = this.getHeader(headers, 'Subject');
    const messageIdHeader = this.getHeader(headers, 'Message-ID');
    const references = this.getHeader(headers, 'References');

    // Determine recipients
    const fromParsed = this.parseEmailAddress(originalFrom);
    let to = fromParsed.email;
    let cc = '';

    if (replyAll) {
      // Include all original recipients except self
      const allRecipients = new Set();
      if (originalTo) {
        originalTo.split(',').forEach(e => allRecipients.add(e.trim()));
      }
      if (originalCc) {
        originalCc.split(',').forEach(e => allRecipients.add(e.trim()));
      }
      // Remove the original sender from CC (they're already in To)
      allRecipients.delete(originalFrom);
      cc = Array.from(allRecipients).join(', ');
    }

    // Build subject with Re: prefix if not already present
    const subject = originalSubject.startsWith('Re:') 
      ? originalSubject 
      : `Re: ${originalSubject}`;

    // Build references header for threading
    const newReferences = references 
      ? `${references} ${messageIdHeader}` 
      : messageIdHeader;

    // Create reply email
    const emailLines = [
      `To: ${to}`,
      cc ? `Cc: ${cc}` : null,
      `Subject: ${subject}`,
      `In-Reply-To: ${messageIdHeader}`,
      `References: ${newReferences}`,
      'Content-Type: text/plain; charset=utf-8',
      'MIME-Version: 1.0',
      '',
      body,
    ].filter(line => line !== null);

    const email = emailLines.join('\r\n');

    const encodedEmail = Buffer.from(email)
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');

    const response = await this.executeWithRetry(() =>
      gmail.users.messages.send({
        userId: 'me',
        requestBody: {
          raw: encodedEmail,
          threadId: original.data.threadId,
        },
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: response.data.id,
            threadId: response.data.threadId,
            to: to,
            subject: subject,
            message: `Reply sent successfully${replyAll ? ' (replied to all)' : ''}`,
          }, null, 2),
        },
      ],
    };
  }

  async getThread(args) {
    const { threadId } = args;
    const { accessToken } = args._auth || {};

    if (!threadId) {
      throw new Error('threadId is required');
    }

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.threads.get({
        userId: 'me',
        id: threadId,
        format: 'full',
      })
    );

    const thread = response.data;
    const messages = thread.messages || [];

    const formattedMessages = messages.map(msg => {
      const headers = msg.payload?.headers || [];
      const fromParsed = this.parseEmailAddress(this.getHeader(headers, 'From'));
      const category = this.detectCategory(msg.labelIds);
      
      return {
        id: msg.id,
        from: this.getHeader(headers, 'From'),
        fromName: fromParsed.name,
        fromEmail: fromParsed.email,
        to: this.getHeader(headers, 'To'),
        subject: this.getHeader(headers, 'Subject'),
        date: this.getHeader(headers, 'Date'),
        body: this.extractBody(msg.payload).substring(0, 2000),
        snippet: msg.snippet,
        category: category,
      };
    });

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            threadId: thread.id,
            messageCount: messages.length,
            messages: formattedMessages,
          }, null, 2),
        },
      ],
    };
  }

  async modifyLabels(args) {
    const { messageId, addLabelIds = [], removeLabelIds = [] } = args;
    const { accessToken } = args._auth || {};

    if (!messageId) {
      throw new Error('messageId is required');
    }

    if (addLabelIds.length === 0 && removeLabelIds.length === 0) {
      throw new Error('At least one of addLabelIds or removeLabelIds is required');
    }

    const gmail = this.createGmailClient(accessToken);

    // Convert friendly names to label IDs
    const normalizedAddLabels = addLabelIds.map(l => this.getSystemLabelId(l));
    const normalizedRemoveLabels = removeLabelIds.map(l => this.getSystemLabelId(l));

    const response = await this.executeWithRetry(() =>
      gmail.users.messages.modify({
        userId: 'me',
        id: messageId,
        requestBody: {
          addLabelIds: normalizedAddLabels,
          removeLabelIds: normalizedRemoveLabels,
        },
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: response.data.id,
            labelIds: response.data.labelIds,
            message: 'Labels modified successfully',
          }, null, 2),
        },
      ],
    };
  }

  async trashMessage(args) {
    const { messageId } = args;
    const { accessToken } = args._auth || {};

    if (!messageId) {
      throw new Error('messageId is required');
    }

    const gmail = this.createGmailClient(accessToken);

    await this.executeWithRetry(() =>
      gmail.users.messages.trash({
        userId: 'me',
        id: messageId,
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: messageId,
            message: 'Message moved to trash',
          }, null, 2),
        },
      ],
    };
  }

  async listLabels(args) {
    const { accessToken } = args._auth || {};

    // Check cache
    const now = Date.now();
    if (this.cache.labels && this.cache.labelsExpiry > now) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(this.cache.labels, null, 2),
          },
        ],
      };
    }

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.labels.list({
        userId: 'me',
      })
    );

    const labels = response.data.labels || [];

    // Separate system labels, categories, and user labels
    const systemLabels = [];
    const categoryLabels = [];
    const userLabels = [];

    labels.forEach(label => {
      const labelInfo = {
        id: label.id,
        name: label.name,
        type: label.type,
        messageListVisibility: label.messageListVisibility,
        labelListVisibility: label.labelListVisibility,
      };

      if (label.id.startsWith('CATEGORY_')) {
        categoryLabels.push(labelInfo);
      } else if (label.type === 'system') {
        systemLabels.push(labelInfo);
      } else {
        userLabels.push(labelInfo);
      }
    });

    const result = {
      totalLabels: labels.length,
      systemLabels,
      categoryLabels,
      userLabels,
      availableCategories: Object.keys(GMAIL_CATEGORIES),
    };

    // Cache the result
    this.cache.labels = result;
    this.cache.labelsExpiry = now + this.CACHE_TTL;

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  }

  async getProfile(args) {
    const { accessToken } = args._auth || {};

    // Check cache
    const now = Date.now();
    if (this.cache.profile && this.cache.profileExpiry > now) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(this.cache.profile, null, 2),
          },
        ],
      };
    }

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.getProfile({
        userId: 'me',
      })
    );

    const profile = {
      emailAddress: response.data.emailAddress,
      messagesTotal: response.data.messagesTotal,
      threadsTotal: response.data.threadsTotal,
      historyId: response.data.historyId,
    };

    // Cache the result
    this.cache.profile = profile;
    this.cache.profileExpiry = now + this.CACHE_TTL;

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(profile, null, 2),
        },
      ],
    };
  }

  async createDraft(args) {
    const { to, subject, body, cc } = args;
    const { accessToken } = args._auth || {};

    if (!to || to.trim() === '') {
      throw new Error('Recipient email address (to) is required');
    }

    if (!subject) {
      throw new Error('Email subject is required');
    }

    if (!body || body.trim() === '') {
      throw new Error('Email body is required');
    }

    const gmail = this.createGmailClient(accessToken);

    // Create email with proper MIME format
    const emailLines = [
      `To: ${to}`,
      cc ? `Cc: ${cc}` : null,
      `Subject: ${subject}`,
      'Content-Type: text/plain; charset=utf-8',
      'MIME-Version: 1.0',
      '',
      body,
    ].filter(line => line !== null);

    const email = emailLines.join('\r\n');

    const encodedEmail = Buffer.from(email)
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');

    const response = await this.executeWithRetry(() =>
      gmail.users.drafts.create({
        userId: 'me',
        requestBody: {
          message: {
            raw: encodedEmail,
          },
        },
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            draftId: response.data.id,
            messageId: response.data.message?.id,
            to: to,
            subject: subject,
            message: 'Draft created successfully',
          }, null, 2),
        },
      ],
    };
  }

  async listDrafts(args) {
    const { maxResults = 10 } = args;
    const { accessToken } = args._auth || {};

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.drafts.list({
        userId: 'me',
        maxResults: Math.min(maxResults, 50),
      })
    );

    const drafts = response.data.drafts || [];

    if (drafts.length === 0) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              totalDrafts: 0,
              drafts: [],
              message: 'No drafts found',
            }, null, 2),
          },
        ],
      };
    }

    // Fetch details for each draft
    const detailedDrafts = await Promise.all(
      drafts.slice(0, Math.min(drafts.length, 10)).map(async (draft) => {
        try {
          const detail = await this.executeWithRetry(() =>
            gmail.users.drafts.get({
              userId: 'me',
              id: draft.id,
              format: 'metadata',
            })
          );

          const headers = detail.data.message?.payload?.headers || [];
          
          return {
            id: draft.id,
            messageId: detail.data.message?.id,
            to: this.getHeader(headers, 'To'),
            subject: this.getHeader(headers, 'Subject') || '(No Subject)',
            snippet: detail.data.message?.snippet || '',
          };
        } catch (err) {
          console.error(`Error fetching draft ${draft.id}:`, err.message);
          return null;
        }
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalDrafts: drafts.length,
            drafts: detailedDrafts.filter(d => d !== null),
          }, null, 2),
        },
      ],
    };
  }

  // ==================== NEW METHODS ====================

  /**
   * Permanently delete a message (cannot be undone)
   */
  async deleteMessage(args) {
    const { messageId } = args;
    const { accessToken } = args._auth || {};

    if (!messageId) {
      throw new Error('messageId is required');
    }

    const gmail = this.createGmailClient(accessToken);

    await this.executeWithRetry(() =>
      gmail.users.messages.delete({
        userId: 'me',
        id: messageId,
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: messageId,
            message: 'Message permanently deleted',
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Send an existing draft
   */
  async sendDraft(args) {
    const { draftId } = args;
    const { accessToken } = args._auth || {};

    if (!draftId) {
      throw new Error('draftId is required');
    }

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.drafts.send({
        userId: 'me',
        requestBody: {
          id: draftId,
        },
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            messageId: response.data.id,
            threadId: response.data.threadId,
            message: 'Draft sent successfully',
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Delete a draft
   */
  async deleteDraft(args) {
    const { draftId } = args;
    const { accessToken } = args._auth || {};

    if (!draftId) {
      throw new Error('draftId is required');
    }

    const gmail = this.createGmailClient(accessToken);

    await this.executeWithRetry(() =>
      gmail.users.drafts.delete({
        userId: 'me',
        id: draftId,
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            draftId: draftId,
            message: 'Draft deleted successfully',
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Update an existing draft
   */
  async updateDraft(args) {
    const { draftId, to, subject, body, cc } = args;
    const { accessToken } = args._auth || {};

    if (!draftId) {
      throw new Error('draftId is required');
    }

    const gmail = this.createGmailClient(accessToken);

    // Get the existing draft first
    const existingDraft = await this.executeWithRetry(() =>
      gmail.users.drafts.get({
        userId: 'me',
        id: draftId,
        format: 'metadata',
      })
    );

    const existingHeaders = existingDraft.data.message?.payload?.headers || [];
    
    // Use existing values if not provided
    const finalTo = to || this.getHeader(existingHeaders, 'To');
    const finalSubject = subject || this.getHeader(existingHeaders, 'Subject');
    const finalBody = body || '';
    const finalCc = cc || this.getHeader(existingHeaders, 'Cc');

    // Create updated email
    const emailLines = [
      `To: ${finalTo}`,
      finalCc ? `Cc: ${finalCc}` : null,
      `Subject: ${finalSubject}`,
      'Content-Type: text/plain; charset=utf-8',
      'MIME-Version: 1.0',
      '',
      finalBody,
    ].filter(line => line !== null);

    const email = emailLines.join('\r\n');

    const encodedEmail = Buffer.from(email)
      .toString('base64')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');

    const response = await this.executeWithRetry(() =>
      gmail.users.drafts.update({
        userId: 'me',
        id: draftId,
        requestBody: {
          message: {
            raw: encodedEmail,
          },
        },
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            draftId: response.data.id,
            message: 'Draft updated successfully',
          }, null, 2),
        },
      ],
    };
  }

  // ==================== LABEL MANAGEMENT ====================

  /**
   * Create a new Gmail label
   */
  async createLabel(args) {
    const { name, messageListVisibility = 'show', labelListVisibility = 'labelShow' } = args;
    const { accessToken } = args._auth || {};

    if (!name) {
      throw new Error('Label name is required');
    }

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.labels.create({
        userId: 'me',
        requestBody: {
          name,
          messageListVisibility,
          labelListVisibility,
        },
      })
    );

    // Clear cache
    this.cache.labels = null;

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            id: response.data.id,
            name: response.data.name,
            type: response.data.type,
            message: `Label "${name}" created successfully`,
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Update an existing Gmail label
   */
  async updateLabel(args) {
    const { id, name, messageListVisibility, labelListVisibility } = args;
    const { accessToken } = args._auth || {};

    if (!id) {
      throw new Error('Label ID is required');
    }

    const gmail = this.createGmailClient(accessToken);

    // Build update object with only provided fields
    const updates = {};
    if (name) updates.name = name;
    if (messageListVisibility) updates.messageListVisibility = messageListVisibility;
    if (labelListVisibility) updates.labelListVisibility = labelListVisibility;

    if (Object.keys(updates).length === 0) {
      throw new Error('At least one field to update is required');
    }

    const response = await this.executeWithRetry(() =>
      gmail.users.labels.update({
        userId: 'me',
        id: id,
        requestBody: updates,
      })
    );

    // Clear cache
    this.cache.labels = null;

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            id: response.data.id,
            name: response.data.name,
            type: response.data.type,
            message: 'Label updated successfully',
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Delete a Gmail label
   */
  async deleteLabel(args) {
    const { id } = args;
    const { accessToken } = args._auth || {};

    if (!id) {
      throw new Error('Label ID is required');
    }

    // Check if it's a system label
    if (id.startsWith('CATEGORY_') || GMAIL_SYSTEM_LABELS[id.toLowerCase()]) {
      throw new Error('Cannot delete system labels');
    }

    const gmail = this.createGmailClient(accessToken);

    await this.executeWithRetry(() =>
      gmail.users.labels.delete({
        userId: 'me',
        id: id,
      })
    );

    // Clear cache
    this.cache.labels = null;

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            id: id,
            message: 'Label deleted successfully',
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Get or create a label by name
   */
  async getOrCreateLabel(args) {
    const { name, messageListVisibility = 'show', labelListVisibility = 'labelShow' } = args;
    const { accessToken } = args._auth || {};

    if (!name) {
      throw new Error('Label name is required');
    }

    const gmail = this.createGmailClient(accessToken);

    // First, try to find the existing label
    const labelsResponse = await this.executeWithRetry(() =>
      gmail.users.labels.list({
        userId: 'me',
      })
    );

    const existingLabel = labelsResponse.data.labels?.find(
      label => label.name.toLowerCase() === name.toLowerCase()
    );

    if (existingLabel) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              success: true,
              action: 'found',
              id: existingLabel.id,
              name: existingLabel.name,
              type: existingLabel.type,
              message: `Found existing label "${name}"`,
            }, null, 2),
          },
        ],
      };
    }

    // Create new label if not found
    const response = await this.executeWithRetry(() =>
      gmail.users.labels.create({
        userId: 'me',
        requestBody: {
          name,
          messageListVisibility,
          labelListVisibility,
        },
      })
    );

    // Clear cache
    this.cache.labels = null;

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            action: 'created',
            id: response.data.id,
            name: response.data.name,
            type: response.data.type,
            message: `Created new label "${name}"`,
          }, null, 2),
        },
      ],
    };
  }

  // ==================== FILTER MANAGEMENT ====================

  /**
   * Create a Gmail filter
   */
  async createFilter(args) {
    const { criteria, action } = args;
    const { accessToken } = args._auth || {};

    if (!criteria || Object.keys(criteria).length === 0) {
      throw new Error('Filter criteria is required');
    }

    if (!action || Object.keys(action).length === 0) {
      throw new Error('Filter action is required');
    }

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.settings.filters.create({
        userId: 'me',
        requestBody: {
          criteria,
          action,
        },
      })
    );

    // Format criteria for display
    const criteriaText = Object.entries(criteria)
      .filter(([_, value]) => value !== undefined)
      .map(([key, value]) => `${key}: ${value}`)
      .join(', ');

    // Format actions for display
    const actionText = Object.entries(action)
      .filter(([_, value]) => value !== undefined)
      .map(([key, value]) => `${key}: ${Array.isArray(value) ? value.join(', ') : value}`)
      .join(', ');

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            filterId: response.data.id,
            criteria: criteriaText,
            action: actionText,
            message: 'Filter created successfully',
          }, null, 2),
        },
      ],
    };
  }

  /**
   * List all Gmail filters
   */
  async listFilters(args) {
    const { accessToken } = args._auth || {};

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.settings.filters.list({
        userId: 'me',
      })
    );

    const filters = response.data.filter || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            totalFilters: filters.length,
            filters: filters.map(filter => ({
              id: filter.id,
              criteria: filter.criteria,
              action: filter.action,
            })),
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Get a specific Gmail filter
   */
  async getFilter(args) {
    const { filterId } = args;
    const { accessToken } = args._auth || {};

    if (!filterId) {
      throw new Error('filterId is required');
    }

    const gmail = this.createGmailClient(accessToken);

    const response = await this.executeWithRetry(() =>
      gmail.users.settings.filters.get({
        userId: 'me',
        id: filterId,
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            id: response.data.id,
            criteria: response.data.criteria,
            action: response.data.action,
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Delete a Gmail filter
   */
  async deleteFilter(args) {
    const { filterId } = args;
    const { accessToken } = args._auth || {};

    if (!filterId) {
      throw new Error('filterId is required');
    }

    const gmail = this.createGmailClient(accessToken);

    await this.executeWithRetry(() =>
      gmail.users.settings.filters.delete({
        userId: 'me',
        id: filterId,
      })
    );

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            success: true,
            filterId: filterId,
            message: 'Filter deleted successfully',
          }, null, 2),
        },
      ],
    };
  }

  /**
   * Create filter from template
   */
  async createFilterFromTemplate(args) {
    const { template, parameters } = args;
    const { accessToken } = args._auth || {};

    if (!template) {
      throw new Error('Template name is required');
    }

    if (!parameters) {
      throw new Error('Template parameters are required');
    }

    // Get the template function
    const templateFn = FILTER_TEMPLATES[template];
    if (!templateFn) {
      throw new Error(`Unknown template: ${template}. Available templates: ${Object.keys(FILTER_TEMPLATES).join(', ')}`);
    }

    // Generate filter config based on template
    let filterConfig;
    switch (template) {
      case 'fromSender':
        if (!parameters.senderEmail) throw new Error('senderEmail is required for fromSender template');
        filterConfig = templateFn(parameters.senderEmail, parameters.labelIds, parameters.archive);
        break;
      case 'withSubject':
        if (!parameters.subjectText) throw new Error('subjectText is required for withSubject template');
        filterConfig = templateFn(parameters.subjectText, parameters.labelIds, parameters.markAsRead);
        break;
      case 'withAttachments':
        filterConfig = templateFn(parameters.labelIds);
        break;
      case 'largeEmails':
        filterConfig = templateFn(parameters.sizeInBytes, parameters.labelIds);
        break;
      case 'containingText':
        if (!parameters.searchText) throw new Error('searchText is required for containingText template');
        filterConfig = templateFn(parameters.searchText, parameters.labelIds, parameters.markImportant);
        break;
      case 'mailingList':
        if (!parameters.listIdentifier) throw new Error('listIdentifier is required for mailingList template');
        filterConfig = templateFn(parameters.listIdentifier, parameters.labelIds, parameters.archive);
        break;
      default:
        throw new Error(`Unknown template: ${template}`);
    }

    // Create the filter using the generated config
    return await this.createFilter({
      criteria: filterConfig.criteria,
      action: filterConfig.action,
      _auth: { accessToken },
    });
  }

  setupErrorHandling() {
    this.server.onerror = (error) => {
      console.error('[Gmail MCP Server Error]', error);
    };

    process.on('SIGINT', async () => {
      console.log('[Gmail MCP] Shutting down...');
      await this.server.close();
      process.exit(0);
    });

    process.on('SIGTERM', async () => {
      console.log('[Gmail MCP] Shutting down...');
      await this.server.close();
      process.exit(0);
    });

    process.on('uncaughtException', (error) => {
      console.error('[Gmail MCP] Uncaught exception:', error);
    });

    process.on('unhandledRejection', (reason, promise) => {
      console.error('[Gmail MCP] Unhandled rejection at:', promise, 'reason:', reason);
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