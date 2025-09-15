import logger from './logger.js';
import AuthManager from './auth.js';
import { refreshAccessToken } from './lib/microsoft-auth.js';
import { gzip } from 'zlib';
import { promisify } from 'util';

interface GraphRequestOptions {
  headers?: Record<string, string>;
  method?: string;
  body?: string;
  rawResponse?: boolean;
  includeHeaders?: boolean;
  accessToken?: string;
  refreshToken?: string;
  queryParams?: Record<string, string>;
  compressionThreshold?: number;

  [key: string]: unknown;
}

interface ContentItem {
  type: 'text';
  text: string;

  [key: string]: unknown;
}

interface McpResponse {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

class GraphClient {
  private authManager: AuthManager;
  private accessToken: string | null = null;
  private refreshToken: string | null = null;
  private gzipAsync: (buffer: Buffer) => Promise<Buffer>;

  constructor(authManager: AuthManager) {
    this.authManager = authManager;
    this.gzipAsync = promisify(gzip);
  }

  setOAuthTokens(accessToken: string, refreshToken?: string): void {
    this.accessToken = accessToken;
    this.refreshToken = refreshToken || null;
  }

  async makeRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<unknown> {
    // Use OAuth tokens if available, otherwise fall back to authManager
    let accessToken =
      options.accessToken || this.accessToken || (await this.authManager.getToken());
    let refreshToken = options.refreshToken || this.refreshToken;

    if (!accessToken) {
      throw new Error('No access token available');
    }

    try {
      const response = await this.performRequest(endpoint, accessToken, options);

      if (response.status === 401 && refreshToken) {
        // Token expired, try to refresh
        await this.refreshAccessToken(refreshToken);

        // Update token for retry
        accessToken = this.accessToken || accessToken;
        if (!accessToken) {
          throw new Error('Failed to refresh access token');
        }

        // Retry the request with new token
        return this.performRequest(endpoint, accessToken, options);
      }

      if (response.status === 403) {
        const errorText = await response.text();
        if (errorText.includes('scope') || errorText.includes('permission')) {
          throw new Error(
            `Microsoft Graph API scope error: ${response.status} ${response.statusText} - ${errorText}. This tool requires organization mode. Please restart with --org-mode flag.`
          );
        }
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${errorText}`
        );
      }

      if (!response.ok) {
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${await response.text()}`
        );
      }

      const text = await response.text();
      let result: any;

      if (text === '') {
        result = { message: 'OK!' };
      } else {
        try {
          result = JSON.parse(text);
        } catch {
          result = { message: 'OK!', rawResponse: text };
        }
      }

      // If includeHeaders is requested, add response headers to the result
      if (options.includeHeaders) {
        const etag = response.headers.get('ETag') || response.headers.get('etag');

        // Simple approach: just add ETag to the result if it's an object
        if (result && typeof result === 'object' && !Array.isArray(result)) {
          return {
            ...result,
            _etag: etag || 'no-etag-found',
          };
        }
      }

      return result;
    } catch (error) {
      logger.error('Microsoft Graph API request failed:', error);
      throw error;
    }
  }

  private async refreshAccessToken(refreshToken: string): Promise<void> {
    const tenantId = process.env.MS365_MCP_TENANT_ID || 'common';
    const clientId = process.env.MS365_MCP_CLIENT_ID || '084a3e9f-a9f4-43f7-89f9-d229cf97853e';
    const clientSecret = process.env.MS365_MCP_CLIENT_SECRET;

    // Log these values for debugging (mask sensitive info)
    logger.info(`refreshAccessToken debug: ${JSON.stringify({
      tenantId,
      clientId: clientId ? `${clientId.substring(0, 8)}...` : 'NOT SET',
      clientSecret: clientSecret ? `${clientSecret.substring(0, 4)}...` : 'NOT SET',
    })}`);

    if (!clientSecret) {
      throw new Error('MS365_MCP_CLIENT_SECRET not configured');
    }

    const response = await refreshAccessToken(refreshToken, clientId, clientSecret, tenantId);
    this.accessToken = response.access_token;
    if (response.refresh_token) {
      this.refreshToken = response.refresh_token;
    }
  }

  private async performRequest(
    endpoint: string,
    accessToken: string,
    options: GraphRequestOptions
  ): Promise<Response> {
    const url = `https://graph.microsoft.com/v1.0${endpoint}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers,
    };

    return fetch(url, {
      method: options.method || 'GET',
      headers,
      body: options.body,
    });
  }

  async graphRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<McpResponse> {
    try {
      logger.info(`Calling ${endpoint} with options: ${JSON.stringify(options)}`);

      // Use new OAuth-aware request method
      const result = await this.makeRequest(endpoint, options);

      return await this.formatJsonResponse(result, options.rawResponse, options.compressionThreshold);
    } catch (error) {
      logger.error(`Error in Graph API request: ${error}`);
      return {
        content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
        isError: true,
      };
    }
  }

  private async compressJsonResponse(jsonString: string, originalSize: number): Promise<McpResponse> {
    try {
      // Use async gzip to avoid blocking the event loop
      const compressed = await this.gzipAsync(Buffer.from(jsonString, 'utf8'));
      const base64Compressed = compressed.toString('base64');
      
      const compressionRatio = ((originalSize - base64Compressed.length) / originalSize * 100).toFixed(2);
      logger.info(`Response compressed from ${originalSize} to ${base64Compressed.length} bytes (${compressionRatio}% reduction)`);

      const compressedResponse = {
        data: base64Compressed,
        is_compressed: true
      };

      return {
        content: [{ type: 'text', text: JSON.stringify(compressedResponse, null, 2) }],
        _meta: {
          originalSize,
          compressedSize: base64Compressed.length,
          compressionRatio: `${compressionRatio}%`
        }
      };
    } catch (compressionError) {
      logger.error(`Failed to compress response: ${compressionError}`);
      // Return the original uncompressed response instead of null
      return {
        content: [{ type: 'text', text: jsonString }],
      };
    }
  }

  private removeODataProperties(obj: Record<string, unknown>): void {
    if (typeof obj === 'object' && obj !== null) {
      Object.keys(obj).forEach((key) => {
        if (key.startsWith('@odata.')) {
          delete obj[key];
        } else if (typeof obj[key] === 'object') {
          this.removeODataProperties(obj[key] as Record<string, unknown>);
        }
      });
    }
  }

  async formatJsonResponse(data: unknown, rawResponse = false, compressionThreshold = 1024*1024*5): Promise<McpResponse> {
    // Handle the case where data includes headers metadata
    if (data && typeof data === 'object' && '_etag' in data) {
      const responseData = data as {
        [key: string]: unknown;
        _etag?: string;
      };

      const meta: Record<string, unknown> = {};
      if (responseData._etag) {
        meta.etag = responseData._etag;
      }

      // Remove _etag from the data before processing
      const { _etag, ...cleanData } = responseData;

      if (rawResponse) {
        return {
          content: [{ type: 'text', text: JSON.stringify(cleanData) }],
          _meta: meta,
        };
      }

      if (cleanData === null || cleanData === undefined) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
          _meta: meta,
        };
      }

      // Remove OData properties
      this.removeODataProperties(cleanData as Record<string, unknown>);

      const jsonString = JSON.stringify(cleanData, null, 2);
      const responseSizeBytes = Buffer.byteLength(jsonString, 'utf8');
      
      logger.info(`Response size: ${responseSizeBytes} bytes, compression threshold: ${compressionThreshold} bytes`);

      // Check if we should compress the response
      if (responseSizeBytes > compressionThreshold) {
        const compressedResponse = await this.compressJsonResponse(jsonString, responseSizeBytes);
        return {
          ...compressedResponse,
          _meta: meta,
        };
      }

      return {
        content: [{ type: 'text', text: jsonString }],
        _meta: meta,
      };
    }

    // Original handling for backward compatibility
    if (rawResponse) {
      return {
        content: [{ type: 'text', text: JSON.stringify(data) }],
      };
    }

    if (data === null || data === undefined) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
      };
    }

    // Remove OData properties
    this.removeODataProperties(data as Record<string, unknown>);

    const jsonString = JSON.stringify(data, null, 2);
    const responseSizeBytes = Buffer.byteLength(jsonString, 'utf8');
    
    logger.info(`Response size: ${responseSizeBytes} bytes, compression threshold: ${compressionThreshold} bytes`);

    // Check if we should compress the response
    if (responseSizeBytes > compressionThreshold) {
      return await this.compressJsonResponse(jsonString, responseSizeBytes);
    }

    return {
      content: [{ type: 'text', text: jsonString }],
    };
  }
}

export default GraphClient;
