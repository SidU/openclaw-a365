# OpenClaw A365 Channel

Native Microsoft 365 Agents (A365) channel for OpenClaw with integrated Graph API tools.

## Features

- **Native Bot Framework Integration**: Receives and sends messages through Microsoft 365 Agents
- **Graph API Tools**: Built-in tools for calendar, email, and user operations
- **Federated Identity Credentials**: Uses T1/T2/User flow for secure Graph API access
- **Role-Based Access**: Distinguishes between Owner and Requester roles
- **Enterprise-Ready**: Supports single-tenant authentication, allowlists, and DM policies

## Quick Start

### 1. Prerequisites

- Azure Bot registration with Microsoft 365 Agents
- Azure AD app registration with:
  - Federated Identity Credential (FIC) configured
  - Graph API permissions: `Calendars.ReadWrite`, `Mail.Send`, `User.Read.All`
- OpenClaw installation

### 2. Configuration

Create or update your OpenClaw configuration (`~/.openclaw/openclaw.json`):

```json
{
  "channels": {
    "a365": {
      "enabled": true,
      "tenantId": "your-tenant-id",
      "appId": "your-bot-app-id",
      "appPassword": "your-bot-app-password",
      "webhook": {
        "port": 3978,
        "path": "/api/messages"
      },
      "graph": {
        "blueprintClientAppId": "your-app-id",
        "blueprintClientSecret": "your-secret",
        "aaInstanceId": "your-aa-instance-id"
      },
      "agentIdentity": "alto@contoso.com",
      "owner": "user@contoso.com",
      "ownerAadId": "user-aad-object-id"
    }
  }
}
```

### 3. Environment Variables

Alternatively, configure via environment variables:

```bash
export A365_APP_ID=your-bot-app-id
export A365_APP_PASSWORD=your-bot-app-password
export A365_TENANT_ID=your-tenant-id
export AA_INSTANCE_ID=your-aa-instance-id
export AGENT_IDENTITY=alto@contoso.com
export OWNER=user@contoso.com
export ANTHROPIC_API_KEY=your-anthropic-key
```

### 4. Start OpenClaw

```bash
openclaw gateway --channel a365
```

## Authentication

The A365 channel uses **Federated Identity Credentials (FIC)** to authenticate with Microsoft Graph API, matching the pattern used by the Microsoft 365 Agents SDK.

### T1/T2/User Token Flow

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   T1 Token      │────▶│   T2 Token      │────▶│  User Token     │
│ (client_creds   │     │ (jwt-bearer     │     │ (user_fic       │
│  + fmi_path)    │     │  assertion)     │     │  grant_type)    │
└─────────────────┘     └─────────────────┘     └─────────────────┘
```

1. **T1 Token**: Acquired using client credentials with `fmi_path` parameter
2. **T2 Token**: Exchanged from T1 using jwt-bearer assertion
3. **User Token**: Final token using `grant_type: user_fic` that can act on behalf of users

### Alternative: External Token Service

If you have an existing token service (e.g., the .NET AgenticScheduler), you can use it:

```json
{
  "tokenCallback": {
    "url": "http://your-service:8080/api/token/refresh",
    "token": "optional-auth-token"
  }
}
```

## Identity Configuration

| Property | Description |
|----------|-------------|
| `owner` | Email of the person this agent supports (the "principal") |
| `ownerAadId` | AAD Object ID of the owner (for role detection) |
| `agentIdentity` | Service account email used for Graph API calls |

When the owner interacts with the agent, they get `UserRole: Owner`. Others get `UserRole: Requester`.

## Graph API Tools

The following tools are available to Claude when Graph API is configured:

| Tool | Description |
|------|-------------|
| `get_calendar_events` | Get calendar events for a date range |
| `create_calendar_event` | Create a new calendar event |
| `update_calendar_event` | Update an existing event |
| `delete_calendar_event` | Delete a calendar event |
| `find_meeting_times` | Find available times for all attendees |
| `send_email` | Send an email via Microsoft Graph |
| `get_user_info` | Get user profile information |

## Configuration Reference

### Channel Configuration

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `enabled` | boolean | `true` | Enable/disable the channel |
| `appId` | string | - | Bot Framework App ID |
| `appPassword` | string | - | Bot Framework App Password |
| `tenantId` | string | - | Azure AD Tenant ID |
| `webhook.port` | number | `3978` | HTTP server port |
| `webhook.path` | string | `/api/messages` | Webhook endpoint path |
| `dmPolicy` | string | `pairing` | DM policy: `open`, `pairing`, `allowlist` |
| `allowFrom` | string[] | `[]` | Allowed user IDs |

### Graph API Configuration

| Option | Type | Description |
|--------|------|-------------|
| `graph.blueprintClientAppId` | string | Client ID for T1 token (usually same as appId) |
| `graph.blueprintClientSecret` | string | Client secret for T1 token |
| `graph.aaInstanceId` | string | Agent Application Instance ID (required for FIC) |
| `graph.scope` | string | OAuth scope (default: `https://graph.microsoft.com/.default`) |

### Identity Configuration

| Option | Type | Description |
|--------|------|-------------|
| `owner` | string | Email of the person this agent supports |
| `ownerAadId` | string | AAD Object ID of the owner |
| `agentIdentity` | string | Service account email for Graph API |
| `businessHours.start` | string | Business hours start (e.g., "08:00") |
| `businessHours.end` | string | Business hours end (e.g., "18:00") |
| `businessHours.timezone` | string | Timezone (e.g., "America/Los_Angeles") |

## Docker Deployment

### Using Docker Compose

1. Copy `.env.example` to `.env` and fill in your credentials
2. Run `docker-compose up -d`
3. Configure your A365 bot to point to `https://your-host:3978/api/messages`

## Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────────────┐
│ Microsoft Teams │───▶│  A365 Service   │───▶│    OpenClaw A365        │
│ Outlook/Email   │    │                 │    │    ┌───────────────┐    │
└─────────────────┘    └─────────────────┘    │    │ Claude Sonnet │    │
                                              │    │               │    │
        ┌─────────────────────────────────────│────│  Graph Tools  │    │
        │                                     │    └───────────────┘    │
        ▼                                     └─────────────────────────┘
   ┌─────────┐
   │ Graph   │  ◄── T1/T2/User Token Flow (FIC)
   │ API     │
   └─────────┘
```

## License

See the main OpenClaw license.
