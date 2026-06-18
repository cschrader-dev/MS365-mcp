# Microsoft Teams MCP Server (Python)

An MCP server that gives AI assistants (like Minimax) read/write access to Microsoft Teams via the Microsoft Graph API.

---

## Tools Provided

| Tool | Description |
|---|---|
| `list_teams` | List Teams the user belongs to |
| `list_channels` | List channels in a Team |
| `list_messages` | Fetch recent messages from a channel |
| `get_message` | Get a specific message by ID |
| `send_message` | Post a message to a channel |
| `reply_to_message` | Reply to an existing message |
| `create_channel` | Create a new channel in a Team |
| `list_members` | List members of a Team |

---

## Setup

### 1. Install dependencies
```bash
pip install mcp httpx msal
```

### 2. Register an Azure App
1. Go to [portal.azure.com](https://portal.azure.com) → Azure Active Directory → App registrations → New registration
2. Name it (e.g. "Teams MCP"), set it as single-tenant
3. Under **Certificates & secrets**, create a new client secret — save it
4. Under **API permissions**, add the following **Application** permissions (not Delegated):
   - `Team.ReadBasic.All`
   - `Channel.ReadBasic.All`
   - `ChannelMessage.Read.All`
   - `ChannelMessage.Send`
   - `TeamMember.Read.All`
   - `Channel.Create`
5. Click **Grant admin consent**

### 3. Set environment variables
```bash
export AZURE_TENANT_ID="your-tenant-id"
export AZURE_CLIENT_ID="your-client-id"
export AZURE_CLIENT_SECRET="your-client-secret"
export AZURE_USER_ID="user@yourdomain.com"  # UPN of the user to act on behalf of
```

### 4. Run the server
```bash
python server.py
```

---

## Connecting to Minimax

In your Minimax MCP config, point it to this server via stdio. Example config entry:

```json
{
  "mcpServers": {
    "teams": {
      "command": "python",
      "args": ["/path/to/server.py"],
      "env": {
        "AZURE_TENANT_ID": "...",
        "AZURE_CLIENT_ID": "...",
        "AZURE_CLIENT_SECRET": "...",
        "AZURE_USER_ID": "..."
      }
    }
  }
}
```

---

## Notes

- This uses the **client credentials** (app-only) flow via MSAL. Some Graph endpoints require delegated (user) auth — if you hit permission errors, you may need to switch to delegated flow with a refresh token.
- `ChannelMessage.Send` may require your tenant admin to approve it.
- Teams message bodies support HTML content.
