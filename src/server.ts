import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import logger, { enableConsoleLogging } from './logger.js';
import { registerGraphTools } from './graph-tools.js';
import GraphClient from './graph-client.js';
import AuthManager from './auth.js';
import type { CommandOptions } from './cli.ts';


class MicrosoftGraphServer {
  private authManager: AuthManager;
  private options: CommandOptions;
  private graphClient: GraphClient;
  private server: McpServer | null;

  constructor(authManager: AuthManager, options: CommandOptions = {}) {
    this.authManager = authManager;
    this.options = options;
    this.graphClient = new GraphClient(authManager);
    this.server = null;
  }

  async initialize(version: string): Promise<void> {
    this.server = new McpServer({
      name: 'Microsoft365MCP',
      version,
    });

    registerGraphTools(
      this.server,
      this.graphClient,
      this.options.readOnly,
      this.options.enabledTools,
      this.options.orgMode
    );
  }

  async start(): Promise<void> {
    if (this.options.v) {
      enableConsoleLogging();
    }

    logger.info('Microsoft 365 MCP Server starting...');

    // Debug: Check if environment variables are loaded
    logger.info('Environment Variables Check:', {
      CLIENT_ID: process.env.MS365_MCP_CLIENT_ID
        ? `${process.env.MS365_MCP_CLIENT_ID.substring(0, 8)}...`
        : 'NOT SET',
      CLIENT_SECRET: process.env.MS365_MCP_CLIENT_SECRET
        ? `${process.env.MS365_MCP_CLIENT_SECRET.substring(0, 8)}...`
        : 'NOT SET',
      TENANT_ID: process.env.MS365_MCP_TENANT_ID || 'NOT SET',
      NODE_ENV: process.env.NODE_ENV || 'NOT SET',
    });

    if (this.options.readOnly) {
      logger.info('Server running in READ-ONLY mode. Write operations are disabled.');
    }

    const transport = new StdioServerTransport();
    await this.server!.connect(transport);
    logger.info('Server connected to stdio transport');
  }
}

export default MicrosoftGraphServer;
