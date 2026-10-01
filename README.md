# general

## MCP server

`.mcp.json` registers a project-scoped MCP server named `odoo`, pointing at
`https://tester-clouded.odoo.com/mcp`. It is the
committed equivalent of:

```
claude mcp add --scope project --transport stdio odoo -- \
  npx -y mcp-remote https://tester-clouded.odoo.com/mcp --header "Authorization: Bearer <ODOO_API_KEY>"
```

The API key is not stored in the file. Export it before starting Claude Code:

```
export ODOO_API_KEY="your-odoo-api-key"
claude
```

Claude Code asks you to approve project servers the first time it sees them.
Check the connection with `claude mcp list` or `/mcp`.

The header is passed as `Authorization:${AUTH_HEADER}` with no space after the
colon. This works around an `mcp-remote` problem where an argument containing a
space gets split on some platforms. The `Bearer ` prefix is in the `AUTH_HEADER`
env value instead.
