# SharePoint MCP Server

A Model Context Protocol (MCP) server that provides SharePoint and OneDrive integration through Microsoft Graph API. This server enables AI assistants to interact with SharePoint sites, search documents, access recent files, and retrieve user profiles.

## Features

- **OAuth 2.0 Authentication** - Secure Microsoft authentication with token refresh
- **Site Resolution** - Get SharePoint sites by hostname and path
- **Document Search** - Search across SharePoint and OneDrive
- **Recent Documents** - List recently accessed documents
- **File Content Fetching** - Retrieve file content by ID or download URL
- **User Profile** - Get current user's profile information
- **Cloudflare Workers** - Deployed on Cloudflare for global edge performance
- **Session Management** - KV-based session storage with automatic token refresh

## Tools Available

1. **sharepoint_get_site** - Resolve a SharePoint site by hostname and path
2. **sharepoint_search** - Search SharePoint/OneDrive documents by keyword
3. **sharepoint_list_recent_documents** - Return recently accessed documents
4. **sharepoint_fetch** - Fetch content from a Graph file download URL or file ID
5. **sharepoint_get_profile** - Retrieve the current user's profile

## Required Scopes

- `Sites.Read.All` - Read SharePoint sites
- `Files.Read.All` - Read files across SharePoint and OneDrive
- `User.Read` - Read user profile
- `offline_access` - Get refresh tokens

## Setup

### 1. Microsoft Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Set redirect URI to your callback URL (e.g., `https://sharepoint-mcp.zerotwo.app/oauth/callback`)
5. Under **Certificates & secrets**, create a new client secret
6. Under **API permissions**, add the required Microsoft Graph permissions listed above
7. Grant admin consent for the permissions

### 2. Environment Variables

Copy `env.example` to `.env` and fill in your values:

```bash
cp env.example .env
```

Required variables:
- `MICROSOFT_CLIENT_ID` - Your Azure app client ID
- `MICROSOFT_CLIENT_SECRET` - Your Azure app client secret
- `MICROSOFT_TENANT_ID` - Your tenant ID (or "common" for multi-tenant)
- `MICROSOFT_REDIRECT_URI` - Your OAuth callback URL
- `FRONTEND_URL` - Your frontend application URL

### 3. KV Namespace Setup

Create a KV namespace for session storage:

```bash
npm run kv:create
```

Update the KV namespace ID in `wrangler.toml`.

### 4. Deploy to Cloudflare

```bash
npm run deploy
```

Set secrets:

```bash
wrangler secret put MICROSOFT_CLIENT_ID
wrangler secret put MICROSOFT_CLIENT_SECRET
wrangler secret put MICROSOFT_TENANT_ID
wrangler secret put MICROSOFT_REDIRECT_URI
wrangler secret put FRONTEND_URL
```

## Local Development

```bash
# Install dependencies
npm install

# Run locally with Node.js
npm run dev

# Or run with Cloudflare Workers locally
npm run dev:worker
```

## API Endpoints

### OAuth Endpoints

- `GET /oauth/authorize?userId={userId}&state={state}` - Get authorization URL
- `GET /oauth/callback` - OAuth callback handler
- `POST /oauth/refresh` - Refresh access token
- `POST /oauth/disconnect` - Disconnect user session
- `POST /oauth/sync` - Sync tokens from backend

### MCP Endpoints

- `POST /mcp` - MCP protocol JSON-RPC endpoint
- `GET /sse` - Server-Sent Events endpoint
- `GET /health` - Health check

## Usage Example

```typescript
// Initialize MCP client
const client = new MCPClient('https://sharepoint-mcp.zerotwo.app');

// Search documents
const result = await client.callTool('sharepoint_search', {
  userId: 'user123',
  query: 'financial report',
  limit: 10
});

// Get recent documents
const recent = await client.callTool('sharepoint_list_recent_documents', {
  userId: 'user123',
  limit: 20
});

// Fetch file content
const content = await client.callTool('sharepoint_fetch', {
  userId: 'user123',
  itemId: 'file-id-here'
});
```

## Architecture

- **Express Server** - Local development with Express and MCP SDK
- **Cloudflare Workers** - Production deployment on the edge
- **KV Storage** - Session and token management
- **Microsoft Graph API** - SharePoint and OneDrive integration

## License

MIT








---

## Powered by ZeroTwo

This SharePoint MCP connector is part of the [ZeroTwo AI platform](https://zerotwo.ai) — the all-in-one AI workspace that lets you search SharePoint sites, access documents, and manage your Microsoft 365 content through GPT-5, Claude, and Gemini.

| | |
|---|---|
| 🌐 **[ZeroTwo — All AI Models in One App](https://zerotwo.ai)** | Search and manage SharePoint sites with GPT-5, Claude, and Gemini — all in one place. |
| ✨ **[ZeroTwo Features](https://zerotwo.ai/features)** | AI document management, enterprise search, web search, and MCP-powered SharePoint tools. |
| 🤖 **[AI Models — GPT-5, Claude & Gemini](https://zerotwo.ai/zerotwo-models)** | Use the world's best AI to find documents, summarize content, and navigate SharePoint. |
| 🔌 **[ZeroTwo Connectors & Integrations](https://zerotwo.ai/connectors)** | Connect SharePoint, Teams, Outlook, OneDrive, and more to your AI workflow. |
| 💰 **[ZeroTwo Pricing](https://zerotwo.ai/pricing)** | One subscription that replaces ChatGPT Plus, Claude Pro, and Gemini Advanced. |
| 📝 **[ZeroTwo Blog](https://zerotwo.ai/blog)** | AI enterprise tools, SharePoint guides, and ZeroTwo product updates. |
| 🚀 **[Try ZeroTwo Free](https://app.zerotwo.ai/auth/login)** | Let AI navigate your SharePoint — get started free today. |

> **Built for ZeroTwo** — Use this SharePoint MCP server with [ZeroTwo's AI connector system](https://zerotwo.ai/connectors) to search sites, access documents, and retrieve user profiles through natural language in your AI assistant.
