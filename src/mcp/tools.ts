/**
 * MCP Tool definitions for SharePoint operations
 */

import { z } from 'zod';
import { sharepointClient } from '../sharepoint/client.js';
import { oauthManager } from '../auth/oauth-manager.js';
import { logger } from '../utils/logger.js';

/**
 * SharePoint Get Site Tool
 */
export const sharepointGetSiteTool = {
  name: 'sharepoint_get_site',
  definition: {
    title: 'Resolve a SharePoint site by hostname and path',
    description: 'Use this when you know a SharePoint hostname and site path and need the resolved site record, site ID, and `webUrl` before working with that site.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      hostname: z.string().describe('SharePoint hostname (e.g., "contoso.sharepoint.com")'),
      path: z.string().describe('Site path (e.g., "/sites/teamsite")'),
    },
    outputSchema: {
      site: z.object({
        id: z.string(),
        displayName: z.string(),
        name: z.string(),
        webUrl: z.string(),
        description: z.string().optional(),
        createdDateTime: z.string(),
      }),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, hostname, path } = args;

      logger.info('[Tool:sharepoint_get_site] Executing', { userId, hostname, path });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
      if (!accessToken) {
        return {
          content: [
            {
              type: 'text',
              text: 'Authentication required. Please authenticate with Microsoft first.',
            },
          ],
          isError: true,
        };
      }

      // Get site
      const site = await sharepointClient.getSite(accessToken, hostname, path);

      const output = {
        site: {
          id: site.id,
          displayName: site.displayName,
          name: site.name,
          webUrl: site.webUrl,
          description: site.description,
          createdDateTime: site.createdDateTime,
        },
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:sharepoint_get_site] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error getting site: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * SharePoint Search Tool
 */
export const sharepointSearchTool = {
  name: 'sharepoint_search',
  definition: {
    title: 'Search SharePoint/OneDrive documents by keyword',
    description: 'Use this to find SharePoint or OneDrive documents by keyword before fetching one. It returns each result\'s `id` and `webUrl`; use the `id` with `sharepoint_fetch` when you need the document content.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      query: z.string().describe('Search query keyword'),
      limit: z.number().min(1).max(500).default(20).describe('Maximum number of results to return'),
    },
    outputSchema: {
      results: z.array(
        z.object({
          id: z.string(),
          name: z.string(),
          webUrl: z.string(),
          summary: z.string().optional(),
          rank: z.number().optional(),
        })
      ),
      count: z.number(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, query, limit } = args;

      logger.info('[Tool:sharepoint_search] Executing', { userId, query, limit });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
      if (!accessToken) {
        return {
          content: [
            {
              type: 'text',
              text: 'Authentication required. Please authenticate with Microsoft first.',
            },
          ],
          isError: true,
        };
      }

      // Search documents
      const results = await sharepointClient.search(accessToken, query, limit);

      const output = {
        results: results.map((result) => ({
          id: result.id,
          name: result.name,
          webUrl: result.webUrl,
          summary: result.summary,
          rank: result.rank,
        })),
        count: results.length,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:sharepoint_search] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error searching: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * SharePoint List Recent Documents Tool
 */
export const sharepointRecentDocumentsTool = {
  name: 'sharepoint_list_recent_documents',
  definition: {
    title: 'Return recently accessed documents',
    description: 'Use this to browse recently accessed SharePoint or OneDrive documents before choosing one to read. The response includes each document\'s `id`, `webUrl`, and often a `downloadUrl`; the `id` or `downloadUrl` can be passed to `sharepoint_fetch`.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      limit: z.number().min(1).max(200).default(20).describe('Maximum number of documents to return'),
    },
    outputSchema: {
      documents: z.array(
        z.object({
          id: z.string(),
          name: z.string(),
          webUrl: z.string(),
          downloadUrl: z.string().optional(),
          createdDateTime: z.string(),
          lastModifiedDateTime: z.string(),
          size: z.number(),
        })
      ),
      count: z.number(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, limit } = args;

      logger.info('[Tool:sharepoint_list_recent_documents] Executing', { userId, limit });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
      if (!accessToken) {
        return {
          content: [
            {
              type: 'text',
              text: 'Authentication required. Please authenticate with Microsoft first.',
            },
          ],
          isError: true,
        };
      }

      // List recent documents
      const documents = await sharepointClient.listRecentDocuments(accessToken, limit);

      const output = {
        documents: documents.map((doc) => ({
          id: doc.id,
          name: doc.name,
          webUrl: doc.webUrl,
          downloadUrl: doc.downloadUrl,
          createdDateTime: doc.createdDateTime,
          lastModifiedDateTime: doc.lastModifiedDateTime,
          size: doc.size,
        })),
        count: documents.length,
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:sharepoint_list_recent_documents] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error listing recent documents: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * SharePoint Fetch File Content Tool
 */
export const sharepointFetchTool = {
  name: 'sharepoint_fetch',
  definition: {
    title: 'Fetch content from a Graph file download URL',
    description: 'Use this to read the contents of a specific SharePoint or OneDrive file after a search or recent-documents step. Prefer passing `itemId` from `sharepoint_search` or `sharepoint_list_recent_documents`; alternatively pass the `downloadUrl` returned by recent-document results.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
      itemId: z.string().optional().describe('SharePoint/OneDrive drive item ID returned by `sharepoint_search` or `sharepoint_list_recent_documents`.'),
      downloadUrl: z.string().optional().describe('Microsoft Graph download URL returned by `sharepoint_list_recent_documents`.'),
    },
    outputSchema: {
      success: z.boolean(),
      name: z.string().optional(),
      content: z.string(),
      mimeType: z.string().optional(),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId, itemId, downloadUrl } = args;

      logger.info('[Tool:sharepoint_fetch] Executing', { userId, itemId, downloadUrl });

      // Validate that either itemId or downloadUrl is provided
      if (!itemId && !downloadUrl) {
        return {
          content: [
            {
              type: 'text',
              text: 'Error: Either itemId or downloadUrl must be provided',
            },
          ],
          isError: true,
        };
      }

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
      if (!accessToken) {
        return {
          content: [
            {
              type: 'text',
              text: 'Authentication required. Please authenticate with Microsoft first.',
            },
          ],
          isError: true,
        };
      }

      let output;

      if (itemId) {
        // Fetch by item ID
        const fileData = await sharepointClient.getFileContentById(accessToken, itemId);
        output = {
          success: true,
          name: fileData.name,
          content: fileData.content,
          mimeType: fileData.mimeType,
        };
      } else if (downloadUrl) {
        // Fetch by download URL
        const content = await sharepointClient.fetchFileContent(accessToken, downloadUrl);
        output = {
          success: true,
          content,
        };
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:sharepoint_fetch] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error fetching file content: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * SharePoint Get User Profile Tool
 */
export const sharepointGetProfileTool = {
  name: 'sharepoint_get_profile',
  definition: {
    title: "Retrieve the current user's profile",
    description: 'Use this when you need the authenticated Microsoft 365 user profile for SharePoint context. It does not search or fetch documents.',
    inputSchema: {
      userId: z.string().describe('User ID for authentication'),
    },
    outputSchema: {
      profile: z.object({
        id: z.string(),
        displayName: z.string(),
        mail: z.string(),
        userPrincipalName: z.string(),
        jobTitle: z.string().optional(),
        department: z.string().optional(),
        officeLocation: z.string().optional(),
      }),
    },
  },
  handler: async (args: any, oauthManagerOverride?: any) => {
    try {
      const { userId } = args;

      logger.info('[Tool:sharepoint_get_profile] Executing', { userId });

      // Use provided oauth manager (for Cloudflare Workers) or default
      const manager = oauthManagerOverride || oauthManager;

      // Get valid access token
      const accessToken = await manager.getValidAccessToken(userId);
      if (!accessToken) {
        return {
          content: [
            {
              type: 'text',
              text: 'Authentication required. Please authenticate with Microsoft first.',
            },
          ],
          isError: true,
        };
      }

      // Get user profile
      const profile = await sharepointClient.getUserProfile(accessToken);

      const output = {
        profile: {
          id: profile.id,
          displayName: profile.displayName,
          mail: profile.mail,
          userPrincipalName: profile.userPrincipalName,
          jobTitle: profile.jobTitle,
          department: profile.department,
          officeLocation: profile.officeLocation,
        },
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(output, null, 2),
          },
        ],
        structuredContent: output,
      };
    } catch (error: any) {
      logger.error('[Tool:sharepoint_get_profile] Error:', error);
      return {
        content: [
          {
            type: 'text',
            text: `Error getting user profile: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  },
};

/**
 * Export all tools
 */
export const sharepointTools = [
  sharepointGetSiteTool,
  sharepointSearchTool,
  sharepointRecentDocumentsTool,
  sharepointFetchTool,
  sharepointGetProfileTool,
];






