/**
 * SharePoint/Microsoft Graph API Client wrapper
 */

import axios, { AxiosInstance } from 'axios';
import { logger } from '../utils/logger.js';
import { SharePointSite, SharePointDocument, SharePointSearchResult, UserProfile } from '../types/index.js';

export class SharePointClient {
  private graphApiUrl = 'https://graph.microsoft.com/v1.0';

  /**
   * Create axios instance with access token
   */
  private getClient(accessToken: string): AxiosInstance {
    return axios.create({
      baseURL: this.graphApiUrl,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });
  }

  /**
   * Get SharePoint site by hostname and path
   */
  async getSite(accessToken: string, hostname: string, path: string): Promise<SharePointSite> {
    try {
      logger.info('[SharePointClient] Getting site', { hostname, path });

      const client = this.getClient(accessToken);
      
      // Use the sites API to get site by hostname and path
      const response = await client.get(`/sites/${hostname}:${path}`);

      const site: SharePointSite = {
        id: response.data.id,
        displayName: response.data.displayName,
        name: response.data.name,
        webUrl: response.data.webUrl,
        description: response.data.description,
        createdDateTime: response.data.createdDateTime,
      };

      logger.info('[SharePointClient] Successfully retrieved site', { siteId: site.id });
      return site;
    } catch (error: any) {
      logger.error('[SharePointClient] Error getting site:', error.response?.data || error.message);
      throw new Error('Failed to get SharePoint site');
    }
  }

  /**
   * Search SharePoint/OneDrive documents by keyword
   */
  async search(accessToken: string, query: string, limit: number = 20): Promise<SharePointSearchResult[]> {
    try {
      logger.info('[SharePointClient] Searching documents', { query, limit });

      const client = this.getClient(accessToken);

      const searchRequest = {
        requests: [
          {
            entityTypes: ['driveItem'],
            query: {
              queryString: query,
            },
            from: 0,
            size: Math.min(limit, 500),
          },
        ],
      };

      const response = await client.post('/search/query', searchRequest);

      const hitsContainer = response.data.value?.[0]?.hitsContainers?.[0];
      const hits = hitsContainer?.hits || [];

      const results: SharePointSearchResult[] = hits.map((hit: any) => ({
        id: hit.resource?.id || hit.hitId,
        name: hit.resource?.name || '',
        webUrl: hit.resource?.webUrl || '',
        summary: hit.summary || '',
        resource: hit.resource,
        hitId: hit.hitId,
        rank: hit.rank,
      }));

      logger.info('[SharePointClient] Successfully searched documents', { count: results.length });
      return results;
    } catch (error: any) {
      logger.error('[SharePointClient] Error searching:', error.response?.data || error.message);
      throw new Error('Failed to search documents');
    }
  }

  /**
   * List recently accessed documents
   */
  async listRecentDocuments(accessToken: string, limit: number = 20): Promise<SharePointDocument[]> {
    try {
      logger.info('[SharePointClient] Listing recent documents', { limit });

      const client = this.getClient(accessToken);

      // Get recent files from the user's drive
      const response = await client.get('/me/drive/recent', {
        params: {
          $top: Math.min(limit, 200),
        },
      });

      const items = response.data.value || [];

      const documents: SharePointDocument[] = items.map((item: any) => ({
        id: item.id,
        name: item.name,
        webUrl: item.webUrl,
        downloadUrl: item['@microsoft.graph.downloadUrl'],
        createdDateTime: item.createdDateTime,
        lastModifiedDateTime: item.lastModifiedDateTime,
        size: item.size,
        folder: item.folder,
        file: item.file,
        createdBy: item.createdBy,
        lastModifiedBy: item.lastModifiedBy,
      }));

      logger.info('[SharePointClient] Successfully retrieved recent documents', { count: documents.length });
      return documents;
    } catch (error: any) {
      logger.error('[SharePointClient] Error listing recent documents:', error.response?.data || error.message);
      throw new Error('Failed to list recent documents');
    }
  }

  /**
   * Fetch content from a Graph file download URL
   */
  async fetchFileContent(accessToken: string, downloadUrl: string): Promise<string> {
    try {
      logger.info('[SharePointClient] Fetching file content');

      // Download URL from Graph API is a pre-authenticated URL
      const response = await axios.get(downloadUrl, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
        responseType: 'text',
      });

      logger.info('[SharePointClient] Successfully fetched file content');
      return response.data;
    } catch (error: any) {
      logger.error('[SharePointClient] Error fetching file content:', error.response?.data || error.message);
      throw new Error('Failed to fetch file content');
    }
  }

  /**
   * Get file by ID and fetch its content
   */
  async getFileContentById(accessToken: string, itemId: string): Promise<{ name: string; content: string; mimeType?: string }> {
    try {
      logger.info('[SharePointClient] Getting file content by ID', { itemId });

      const client = this.getClient(accessToken);

      // Get file metadata
      const metadataResponse = await client.get(`/me/drive/items/${itemId}`);
      const metadata = metadataResponse.data;

      // Get file content
      const contentResponse = await client.get(`/me/drive/items/${itemId}/content`, {
        responseType: 'text',
      });

      logger.info('[SharePointClient] Successfully retrieved file content');

      return {
        name: metadata.name,
        content: contentResponse.data,
        mimeType: metadata.file?.mimeType,
      };
    } catch (error: any) {
      logger.error('[SharePointClient] Error getting file content:', error.response?.data || error.message);
      throw new Error('Failed to get file content');
    }
  }

  /**
   * Get current user's profile
   */
  async getUserProfile(accessToken: string): Promise<UserProfile> {
    try {
      logger.info('[SharePointClient] Getting user profile');

      const client = this.getClient(accessToken);
      const response = await client.get('/me');

      const profile: UserProfile = {
        id: response.data.id,
        displayName: response.data.displayName,
        mail: response.data.mail,
        userPrincipalName: response.data.userPrincipalName,
        jobTitle: response.data.jobTitle,
        department: response.data.department,
        officeLocation: response.data.officeLocation,
        businessPhones: response.data.businessPhones,
        mobilePhone: response.data.mobilePhone,
      };

      logger.info('[SharePointClient] Successfully retrieved user profile', { userId: profile.id });
      return profile;
    } catch (error: any) {
      logger.error('[SharePointClient] Error getting user profile:', error.response?.data || error.message);
      throw new Error('Failed to get user profile');
    }
  }
}

// Singleton instance
export const sharepointClient = new SharePointClient();







