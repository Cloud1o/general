# general

## MCP server

`.mcp.json` registers a project-scoped MCP server named `database`. It is the
committed equivalent of:

```
claude mcp add --scope project --transport stdio database -- \
  npx -y mcp-remote <database_url>/mcp --header "Authorization: Bearer <API_KEY>"
```

Secrets are not stored in the file. Export them before starting Claude Code:

```
export DATABASE_URL="https://your-database.example.com"   # no trailing slash
export API_KEY="your-api-key"
claude
```

Claude Code asks you to approve project servers the first time it sees them.
Check the connection with `claude mcp list` or `/mcp`.

The header is passed as `Authorization:${AUTH_HEADER}` with no space after the
colon. This works around an `mcp-remote` problem where an argument containing a
space gets split on some platforms. The `Bearer ` prefix is in the `AUTH_HEADER`
env value instead.
