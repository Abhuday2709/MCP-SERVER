import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

class MCPClientManager {
  constructor() {
    this.clients = {
      gmail: null,
      teams: null,
    };
    this.transports = {
      gmail: null,
      teams: null,
    };
  }

  async connectGmail() {
    if (this.clients.gmail) {
      return this.clients.gmail;
    }

    try {
      const serverPath = path.join(__dirname, '../servers/gmail-server.js');
      
      console.log('Starting Gmail MCP server from:', serverPath);

      // Create transport - StdioClientTransport expects command and args
      const transport = new StdioClientTransport({
        command: 'node',
        args: [serverPath],
      });

      this.transports.gmail = transport;

      // Create client
      const client = new Client(
        {
          name: 'gmail-client',
          version: '1.0.0',
        },
        {
          capabilities: {},
        }
      );

      // Connect client to transport
      await client.connect(transport);
      this.clients.gmail = client;

      console.log('Gmail MCP client connected successfully');
      return client;
    } catch (error) {
      console.error('Error connecting Gmail client:', error);
      throw error;
    }
  }

  async connectTeams() {
    if (this.clients.teams) {
      return this.clients.teams;
    }

    try {
      const serverPath = path.join(__dirname, '../servers/teams-server.js');
      
      console.log('Starting Teams MCP server from:', serverPath);

      // Create transport - StdioClientTransport expects command and args
      const transport = new StdioClientTransport({
        command: 'node',
        args: [serverPath],
      });

      this.transports.teams = transport;

      // Create client
      const client = new Client(
        {
          name: 'teams-client',
          version: '1.0.0',
        },
        {
          capabilities: {},
        }
      );

      // Connect client to transport
      await client.connect(transport);
      this.clients.teams = client;

      console.log('Teams MCP client connected successfully');
      return client;
    } catch (error) {
      console.error('Error connecting Teams client:', error);
      throw error;
    }
  }

  async getAvailableTools(providers = []) {
    const tools = [];

    for (const provider of providers) {
      try {
        let client = null;

        if (provider === 'gmail') {
          client = await this.connectGmail();
        } else if (provider === 'teams') {
          client = await this.connectTeams();
        } else {
          console.log(`Skipping ${provider} - not implemented yet`);
          continue;
        }

        if (client) {
          console.log(`Requesting tools from ${provider}...`);
          const response = await client.listTools();
          
          console.log(`Received ${response.tools.length} tools from ${provider}`);
          tools.push(...response.tools);
        }
      } catch (error) {
        console.error(`Error getting tools from ${provider}:`, error.message);
      }
    }

    return tools;
  }

  async callTool(provider, toolName, args, accessToken) {
    try {
      let client = null;

      if (provider === 'gmail') {
        client = await this.connectGmail();
      } else if (provider === 'teams') {
        client = await this.connectTeams();
      } else {
        throw new Error(`Provider ${provider} not implemented yet`);
      }

      if (!client) {
        throw new Error(`Client not available for provider: ${provider}`);
      }

      console.log(`Calling tool: ${toolName} on ${provider}`);

      // Inject auth token into args
      const argsWithAuth = {
        ...args,
        _auth: { accessToken },
      };

      const response = await client.callTool({
        name: toolName,
        arguments: argsWithAuth,
      });

      return response;
    } catch (error) {
      console.error(`Error calling tool ${toolName}:`, error.message);
      throw error;
    }
  }

  async disconnect() {
    console.log('Disconnecting MCP clients...');
    
    for (const [provider, client] of Object.entries(this.clients)) {
      if (client) {
        try {
          await client.close();
          console.log(`${provider} client closed`);
        } catch (error) {
          console.error(`Error closing ${provider} client:`, error.message);
        }
      }
    }

    // Clean up transports
    for (const [provider, transport] of Object.entries(this.transports)) {
      if (transport) {
        try {
          await transport.close();
          console.log(`${provider} transport closed`);
        } catch (error) {
          console.error(`Error closing ${provider} transport:`, error.message);
        }
      }
    }

    this.clients = { gmail: null, teams: null };
    this.transports = { gmail: null, teams: null };
  }
}

export default new MCPClientManager();