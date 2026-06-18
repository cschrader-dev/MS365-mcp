"""
Microsoft Teams MCP Server
Provides read/write access to Microsoft Teams via the Microsoft Graph API.
Designed for use with MCP-compatible AI assistants (e.g. Minimax).

Requirements:
    pip install mcp httpx msal

Authentication:
    Uses OAuth2 client credentials flow (app-only) or delegated (user) flow.
    Set the following environment variables:
        AZURE_TENANT_ID     - Your Azure AD tenant ID
        AZURE_CLIENT_ID     - Your app registration client ID
        AZURE_CLIENT_SECRET - Your app registration client secret
        AZURE_USER_ID       - (optional) User ID or UPN for delegated operations
"""

import os
import httpx
import asyncio
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent
import msal

# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
USER_ID = os.environ.get("AZURE_USER_ID", "me")  # UPN or object ID

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]


def get_access_token() -> str:
    """Acquire an access token via MSAL client credentials."""
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        raise RuntimeError(f"Failed to acquire token: {result.get('error_description')}")
    return result["access_token"]


def graph_headers() -> dict:
    token = get_access_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


# ---------------------------------------------------------------------------
# Graph helpers
# ---------------------------------------------------------------------------

async def graph_get(path: str) -> dict:
    async with httpx.AsyncClient() as client:
        r = await client.get(f"{GRAPH_BASE}{path}", headers=graph_headers(), timeout=30)
        r.raise_for_status()
        return r.json()


async def graph_post(path: str, body: dict) -> dict:
    async with httpx.AsyncClient() as client:
        r = await client.post(f"{GRAPH_BASE}{path}", headers=graph_headers(), json=body, timeout=30)
        r.raise_for_status()
        return r.json()


async def graph_patch(path: str, body: dict) -> dict:
    async with httpx.AsyncClient() as client:
        r = await client.patch(f"{GRAPH_BASE}{path}", headers=graph_headers(), json=body, timeout=30)
        r.raise_for_status()
        return r.json()


async def graph_delete(path: str) -> int:
    async with httpx.AsyncClient() as client:
        r = await client.delete(f"{GRAPH_BASE}{path}", headers=graph_headers(), timeout=30)
        r.raise_for_status()
        return r.status_code


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------

server = Server("teams-mcp")


@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name="list_teams",
            description="List all Microsoft Teams the authenticated user is a member of.",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        Tool(
            name="list_channels",
            description="List all channels in a specific Team.",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {"type": "string", "description": "The ID of the Team."},
                },
                "required": ["team_id"],
            },
        ),
        Tool(
            name="list_messages",
            description="List recent messages in a Teams channel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {"type": "string", "description": "The ID of the Team."},
                    "channel_id": {"type": "string", "description": "The ID of the channel."},
                    "top": {"type": "integer", "description": "Max number of messages to return (default 20)."},
                },
                "required": ["team_id", "channel_id"],
            },
        ),
        Tool(
            name="send_message",
            description="Send a message to a Teams channel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {"type": "string", "description": "The ID of the Team."},
                    "channel_id": {"type": "string", "description": "The ID of the channel."},
                    "content": {"type": "string", "description": "The message text (HTML supported)."},
                },
                "required": ["team_id", "channel_id", "content"],
            },
        ),
        Tool(
            name="reply_to_message",
            description="Reply to an existing message in a Teams channel.",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {"type": "string", "description": "The ID of the Team."},
                    "channel_id": {"type": "string", "description": "The ID of the channel."},
                    "message_id": {"type": "string", "description": "The ID of the message to reply to."},
                    "content": {"type": "string", "description": "The reply text (HTML supported)."},
                },
                "required": ["team_id", "channel_id", "message_id", "content"],
            },
        ),
        Tool(
            name="create_channel",
            description="Create a new channel in a Team.",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {"type": "string", "description": "The ID of the Team."},
                    "display_name": {"type": "string", "description": "Name of the new channel."},
                    "description": {"type": "string", "description": "Optional description."},
                },
                "required": ["team_id", "display_name"],
            },
        ),
        Tool(
            name="list_members",
            description="List members of a Team.",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {"type": "string", "description": "The ID of the Team."},
                },
                "required": ["team_id"],
            },
        ),
        Tool(
            name="get_message",
            description="Get a specific message from a Teams channel by its ID.",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {"type": "string", "description": "The ID of the Team."},
                    "channel_id": {"type": "string", "description": "The ID of the channel."},
                    "message_id": {"type": "string", "description": "The ID of the message."},
                },
                "required": ["team_id", "channel_id", "message_id"],
            },
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    try:
        result = await dispatch(name, arguments)
        return [TextContent(type="text", text=str(result))]
    except Exception as e:
        return [TextContent(type="text", text=f"Error: {e}")]


async def dispatch(name: str, args: dict):
    if name == "list_teams":
        data = await graph_get(f"/users/{USER_ID}/joinedTeams")
        teams = [{"id": t["id"], "displayName": t["displayName"]} for t in data.get("value", [])]
        return teams

    elif name == "list_channels":
        team_id = args["team_id"]
        data = await graph_get(f"/teams/{team_id}/channels")
        channels = [{"id": c["id"], "displayName": c["displayName"]} for c in data.get("value", [])]
        return channels

    elif name == "list_messages":
        team_id = args["team_id"]
        channel_id = args["channel_id"]
        top = args.get("top", 20)
        data = await graph_get(f"/teams/{team_id}/channels/{channel_id}/messages?$top={top}")
        messages = [
            {
                "id": m["id"],
                "from": m.get("from", {}).get("user", {}).get("displayName", "Unknown"),
                "createdDateTime": m.get("createdDateTime"),
                "body": m.get("body", {}).get("content", ""),
            }
            for m in data.get("value", [])
        ]
        return messages

    elif name == "send_message":
        team_id = args["team_id"]
        channel_id = args["channel_id"]
        body = {"body": {"contentType": "html", "content": args["content"]}}
        result = await graph_post(f"/teams/{team_id}/channels/{channel_id}/messages", body)
        return {"id": result.get("id"), "status": "sent"}

    elif name == "reply_to_message":
        team_id = args["team_id"]
        channel_id = args["channel_id"]
        message_id = args["message_id"]
        body = {"body": {"contentType": "html", "content": args["content"]}}
        result = await graph_post(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies", body
        )
        return {"id": result.get("id"), "status": "replied"}

    elif name == "create_channel":
        team_id = args["team_id"]
        body = {
            "displayName": args["display_name"],
            "description": args.get("description", ""),
            "membershipType": "standard",
        }
        result = await graph_post(f"/teams/{team_id}/channels", body)
        return {"id": result.get("id"), "displayName": result.get("displayName")}

    elif name == "list_members":
        team_id = args["team_id"]
        data = await graph_get(f"/teams/{team_id}/members")
        members = [
            {"id": m["id"], "displayName": m.get("displayName"), "email": m.get("email")}
            for m in data.get("value", [])
        ]
        return members

    elif name == "get_message":
        team_id = args["team_id"]
        channel_id = args["channel_id"]
        message_id = args["message_id"]
        m = await graph_get(f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}")
        return {
            "id": m["id"],
            "from": m.get("from", {}).get("user", {}).get("displayName", "Unknown"),
            "createdDateTime": m.get("createdDateTime"),
            "body": m.get("body", {}).get("content", ""),
        }

    else:
        raise ValueError(f"Unknown tool: {name}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(main())
