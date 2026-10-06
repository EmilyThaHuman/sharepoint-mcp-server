/**
 * Type definitions for SharePoint MCP Server
 */

export interface SharePointSite {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
  description?: string;
  createdDateTime: string;
}

export interface SharePointDocument {
  id: string;
  name: string;
  webUrl: string;
  downloadUrl?: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  size: number;
  folder?: {
    childCount: number;
  };
  file?: {
    mimeType: string;
  };
  createdBy?: {
    user?: {
      displayName: string;
      email: string;
    };
  };
  lastModifiedBy?: {
    user?: {
      displayName: string;
      email: string;
    };
  };
}

export interface SharePointSearchResult {
  id: string;
  name: string;
  webUrl: string;
  summary?: string;
  resource?: {
    '@odata.type': string;
    id: string;
    webUrl: string;
  };
  hitId?: string;
  rank?: number;
}

export interface UserProfile {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  businessPhones?: string[];
  mobilePhone?: string;
}

export interface OAuthTokens {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  token_type?: string;
  scope?: string;
}

export interface SessionData {
  userId: string;
  accessToken: string;
  refreshToken?: string;
  expiresAt?: Date;
  createdAt: Date;
}







