# DEPRECATED — superseded by `sharepoint-mcp`

This SharePoint MCP server has been replaced by a rebuild on Cloudflare
Workers using the shared `@zerotwo/mcp-connector-kit` (JWT-based auth via
the central ZeroTwoApi JWKS).

- **Replacement worker:** `https://sharepoint-mcp.reed-b9b.workers.dev/mcp`
- **Rebuilt source:** `zerotwo-connectors/sharepoint-mcp/`
- **Refactor design doc:** `AppZeroTwo/docs/05-01-2026-connectors-mcp-refactor.md`

The legacy code in this repository is no longer deployed — production
traffic for SharePoint flows through the replacement worker URL above.
This repo is kept for git-history reference only and will be archived on
GitHub once this PR is merged.
