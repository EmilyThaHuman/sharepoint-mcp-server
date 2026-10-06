/**
 * OAuth Manager for handling Microsoft OAuth 2.0 authentication
 */

import axios from 'axios';
import { config } from '../config/index.js';
import { logger } from '../utils/logger.js';
import { OAuthTokens, SessionData } from '../types/index.js';

export class OAuthManager {
  private sessions: Map<string, SessionData>;

  constructor() {
    this.sessions = new Map();
  }

  /**
   * Generate OAuth authorization URL
   */
  getAuthorizationUrl(state: string): string {
    const params = new URLSearchParams({
      client_id: config.microsoft.clientId,
      response_type: 'code',
      redirect_uri: config.microsoft.redirectUri,
      response_mode: 'query',
      scope: config.sharepoint.scopes.join(' '),
      state,
      prompt: 'consent', // Force consent to get refresh token
    });

    const authUrl = `https://login.microsoftonline.com/${config.microsoft.tenantId}/oauth2/v2.0/authorize?${params.toString()}`;

    logger.info('[OAuthManager] Generated authorization URL', { state });
    return authUrl;
  }

  /**
   * Exchange authorization code for tokens
   */
  async exchangeCodeForTokens(code: string): Promise<OAuthTokens> {
    try {
      logger.info('[OAuthManager] Exchanging code for tokens');

      const params = new URLSearchParams({
        client_id: config.microsoft.clientId,
        client_secret: config.microsoft.clientSecret,
        code,
        redirect_uri: config.microsoft.redirectUri,
        grant_type: 'authorization_code',
        scope: config.sharepoint.scopes.join(' '),
      });

      const response = await axios.post(
        `https://login.microsoftonline.com/${config.microsoft.tenantId}/oauth2/v2.0/token`,
        params.toString(),
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        }
      );

      const tokens: OAuthTokens = {
        access_token: response.data.access_token,
        refresh_token: response.data.refresh_token,
        expires_in: response.data.expires_in,
        token_type: response.data.token_type,
        scope: response.data.scope,
      };

      logger.info('[OAuthManager] Successfully exchanged code for tokens');
      return tokens;
    } catch (error: any) {
      logger.error('[OAuthManager] Error exchanging code for tokens:', error.response?.data || error.message);
      throw new Error('Failed to exchange authorization code for tokens');
    }
  }

  /**
   * Refresh access token using refresh token
   */
  async refreshAccessToken(refreshToken: string): Promise<OAuthTokens> {
    try {
      logger.info('[OAuthManager] Refreshing access token');

      const params = new URLSearchParams({
        client_id: config.microsoft.clientId,
        client_secret: config.microsoft.clientSecret,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: config.sharepoint.scopes.join(' '),
      });

      const response = await axios.post(
        `https://login.microsoftonline.com/${config.microsoft.tenantId}/oauth2/v2.0/token`,
        params.toString(),
        {
          headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
          },
        }
      );

      const tokens: OAuthTokens = {
        access_token: response.data.access_token,
        refresh_token: response.data.refresh_token || refreshToken, // Some responses don't return new refresh token
        expires_in: response.data.expires_in,
        token_type: response.data.token_type,
        scope: response.data.scope,
      };

      logger.info('[OAuthManager] Successfully refreshed access token');
      return tokens;
    } catch (error: any) {
      logger.error('[OAuthManager] Error refreshing access token:', error.response?.data || error.message);
      throw new Error('Failed to refresh access token');
    }
  }

  /**
   * Store session data
   */
  storeSession(userId: string, tokens: OAuthTokens): void {
    const expiresAt = tokens.expires_in
      ? new Date(Date.now() + tokens.expires_in * 1000)
      : undefined;

    const sessionData: SessionData = {
      userId,
      accessToken: tokens.access_token,
      refreshToken: tokens.refresh_token,
      expiresAt,
      createdAt: new Date(),
    };

    this.sessions.set(userId, sessionData);
    logger.info('[OAuthManager] Stored session for user', { userId });
  }

  /**
   * Get session data
   */
  getSession(userId: string): SessionData | undefined {
    return this.sessions.get(userId);
  }

  /**
   * Check if session is valid (not expired)
   */
  isSessionValid(userId: string): boolean {
    const session = this.sessions.get(userId);
    if (!session) {
      return false;
    }

    if (!session.expiresAt) {
      return true; // No expiration set
    }

    return session.expiresAt > new Date();
  }

  /**
   * Get valid access token (refresh if needed)
   */
  async getValidAccessToken(userId: string): Promise<string | null> {
    const session = this.sessions.get(userId);
    if (!session) {
      logger.warn('[OAuthManager] No session found for user', { userId });
      return null;
    }

    // Check if token is still valid
    if (this.isSessionValid(userId)) {
      return session.accessToken;
    }

    // Token expired, try to refresh
    if (!session.refreshToken) {
      logger.warn('[OAuthManager] No refresh token available for user', { userId });
      return null;
    }

    try {
      const newTokens = await this.refreshAccessToken(session.refreshToken);
      this.storeSession(userId, newTokens);
      return newTokens.access_token;
    } catch (error) {
      logger.error('[OAuthManager] Failed to refresh token for user', { userId, error });
      return null;
    }
  }

  /**
   * Remove session
   */
  removeSession(userId: string): void {
    this.sessions.delete(userId);
    logger.info('[OAuthManager] Removed session for user', { userId });
  }

  /**
   * Get all active sessions
   */
  getActiveSessions(): string[] {
    return Array.from(this.sessions.keys());
  }
}

// Singleton instance
export const oauthManager = new OAuthManager();







