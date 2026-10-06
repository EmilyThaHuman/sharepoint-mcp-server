# SharePoint MCP Server Integration Guide

## Server Information

**Base URL**: `https://sharepoint-mcp-server.reed-b9b.workers.dev`

## Authentication Flow

### 1. Get Authorization URL

```javascript
const response = await fetch(
  'https://sharepoint-mcp-server.reed-b9b.workers.dev/oauth/authorize?userId=USER_ID&state=RANDOM_STATE'
);
const { authorizationUrl, state } = await response.json();

// Redirect user to authorizationUrl
window.location.href = authorizationUrl;
```

### 2. Handle OAuth Callback

After user authorizes, they will be redirected to:
```
https://zerotwo.ai/settings?oauth_success=true&provider=sharepoint
```

The tokens are automatically stored in the server's KV storage.

## Using MCP Tools

### Initialize MCP Client

```javascript
const mcpClient = {
  baseUrl: 'https://sharepoint-mcp-server.reed-b9b.workers.dev',
  
  async initialize() {
    const response = await fetch(`${this.baseUrl}/mcp`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        jsonrpc: '2.0',
        method: 'initialize',
        id: 1,
        params: {
          protocolVersion: '2024-11-05',
          capabilities: {},
          clientInfo: {
            name: 'zerotwo-client',
            version: '1.0.0'
          }
        }
      })
    });
    return response.json();
  },
  
  async listTools() {
    const response = await fetch(`${this.baseUrl}/mcp`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        jsonrpc: '2.0',
        method: 'tools/list',
        id: 2
      })
    });
    return response.json();
  },
  
  async callTool(toolName, args) {
    const response = await fetch(`${this.baseUrl}/mcp`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        jsonrpc: '2.0',
        method: 'tools/call',
        id: 3,
        params: {
          name: toolName,
          arguments: args
        }
      })
    });
    return response.json();
  }
};
```

## Tool Examples

### 1. Get SharePoint Site

```javascript
const result = await mcpClient.callTool('sharepoint_get_site', {
  userId: 'user123',
  hostname: 'contoso.sharepoint.com',
  path: '/sites/teamsite'
});

console.log(result.result.content[0].text);
// Output: Site information in JSON format
```

### 2. Search Documents

```javascript
const result = await mcpClient.callTool('sharepoint_search', {
  userId: 'user123',
  query: 'financial report Q4',
  limit: 20
});

const searchResults = JSON.parse(result.result.content[0].text);
console.log(`Found ${searchResults.count} documents`);
```

### 3. List Recent Documents

```javascript
const result = await mcpClient.callTool('sharepoint_list_recent_documents', {
  userId: 'user123',
  limit: 10
});

const documents = JSON.parse(result.result.content[0].text);
documents.documents.forEach(doc => {
  console.log(`${doc.name} - ${doc.webUrl}`);
});
```

### 4. Fetch File Content

```javascript
// By item ID
const result = await mcpClient.callTool('sharepoint_fetch', {
  userId: 'user123',
  itemId: 'file-id-from-search-or-recent'
});

// By download URL
const result2 = await mcpClient.callTool('sharepoint_fetch', {
  userId: 'user123',
  downloadUrl: 'https://graph.microsoft.com/...'
});

const fileData = JSON.parse(result.result.content[0].text);
console.log(`File: ${fileData.name}`);
console.log(`Content: ${fileData.content}`);
```

### 5. Get User Profile

```javascript
const result = await mcpClient.callTool('sharepoint_get_profile', {
  userId: 'user123'
});

const profile = JSON.parse(result.result.content[0].text);
console.log(`User: ${profile.profile.displayName}`);
console.log(`Email: ${profile.profile.mail}`);
```

## Error Handling

```javascript
try {
  const result = await mcpClient.callTool('sharepoint_search', {
    userId: 'user123',
    query: 'test'
  });
  
  if (result.result.isError) {
    console.error('Tool error:', result.result.content[0].text);
  } else {
    // Success
    const data = JSON.parse(result.result.content[0].text);
    console.log(data);
  }
} catch (error) {
  console.error('Request error:', error);
}
```

## Common Error Messages

1. **"Authentication required. Please authenticate with Microsoft first."**
   - User needs to complete OAuth flow
   - Call the `/oauth/authorize` endpoint to start authentication

2. **"User ID is required for all SharePoint operations."**
   - The `userId` parameter is missing from the tool call
   - Ensure you pass the correct user ID

3. **"Failed to get SharePoint site"**
   - Check that the hostname and path are correct
   - Verify user has access to the site

4. **"Failed to search documents"**
   - Verify the search query is valid
   - Check that user has the required permissions

## Token Management

### Refresh Token

Tokens are automatically refreshed by the server when they expire. If manual refresh is needed:

```javascript
const response = await fetch(
  'https://sharepoint-mcp-server.reed-b9b.workers.dev/oauth/refresh',
  {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      userId: 'user123',
      refreshToken: 'refresh_token_here'
    })
  }
);

const { accessToken, refreshToken, expiresIn } = await response.json();
```

### Disconnect User

```javascript
const response = await fetch(
  'https://sharepoint-mcp-server.reed-b9b.workers.dev/oauth/disconnect',
  {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      userId: 'user123'
    })
  }
);

const { success, message } = await response.json();
```

## Backend Token Sync

If you store OAuth tokens in your backend (e.g., in `profile.settings.oauth_tokens`), you can sync them to the MCP server:

```javascript
const response = await fetch(
  'https://sharepoint-mcp-server.reed-b9b.workers.dev/oauth/sync',
  {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      userId: 'user123',
      accessToken: 'access_token_here',
      refreshToken: 'refresh_token_here',
      expiresIn: 3600
    })
  }
);

const { success, message } = await response.json();
```

## Health Check

```javascript
const response = await fetch(
  'https://sharepoint-mcp-server.reed-b9b.workers.dev/health'
);
const health = await response.json();

console.log(`Server Status: ${health.status}`);
console.log(`Version: ${health.version}`);
```

## Rate Limiting

The server follows Microsoft Graph API rate limits:
- 10,000 requests per 10 minutes per app per tenant
- 2,000 requests per second per app per tenant

Handle rate limit errors appropriately and implement exponential backoff if needed.

## Support

For issues or questions:
- GitHub: https://github.com/zerotwo/sharepoint-mcp-server
- Documentation: See README.md

---

**Last Updated**: 2025-11-15







