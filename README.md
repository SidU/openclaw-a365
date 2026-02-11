# OpenClaw A365 Channel

Native Microsoft 365 Agents (A365) channel for OpenClaw with integrated Graph API tools.

## Features

- **Native Bot Framework Integration**: Receives and sends messages through Microsoft 365 Agents
- **Graph API Tools**: Built-in tools for calendar, email, and user operations
- **Agentic Blueprint**: Uses T1/T2/User token flow for secure Graph API access on behalf of users
- **Multi-Model Support**: Configure primary model and fallbacks (Anthropic, OpenAI, OpenRouter, Azure)
- **Role-Based Access**: Distinguishes between Owner and Requester roles
- **Enterprise-Ready**: Supports single-tenant authentication, allowlists, and DM policies

## Quick Start (Docker)

### 1. Prerequisites

- Docker and Docker Compose
- Azure Bot registration with Microsoft 365 Agents
- Autonomous Agent Instance ID (from M365 Agents registration)
- API key for at least one LLM provider (Anthropic, OpenAI, etc.)

### 2. Configure

```bash
cp .env.example .env
# Edit .env with your credentials
```

Required environment variables:

| Variable | Description |
|----------|-------------|
| `ANTHROPIC_API_KEY` | Anthropic API key (or use another provider) |
| `A365_APP_ID` | Bot Framework App ID |
| `A365_APP_PASSWORD` | Bot Framework App Password |
| `A365_TENANT_ID` | Azure AD Tenant ID |
| `AA_INSTANCE_ID` | Autonomous Agent Instance ID |
| `AGENT_IDENTITY` | Agent UPN (e.g., `agent@contoso.com`) |
| `OWNER` | Owner UPN (e.g., `user@contoso.com`) |
| `OWNER_AAD_ID` | Owner's AAD Object ID |

### 3. Run

```bash
docker-compose up -d
```

### 4. Configure A365

Point your A365 agent to `https://your-host:3978/api/messages`

## Model Configuration

Configure which LLM model to use via environment variables:

```bash
# Primary model (default: anthropic/claude-opus-4-6)
OPENCLAW_MODEL=anthropic/claude-sonnet-4-20250514

# Fallback models (comma-separated, tried in order if primary fails)
OPENCLAW_FALLBACK_MODELS=openai/gpt-4o,openrouter/anthropic/claude-3-haiku
```

Supported providers:
- **Anthropic**: `anthropic/claude-opus-4-6`, `anthropic/claude-sonnet-4-20250514`
- **OpenAI**: `openai/gpt-4o`, `openai/gpt-4-turbo`
- **OpenRouter**: `openrouter/anthropic/claude-3.5-sonnet`, etc.
- **Azure OpenAI**: `azure/gpt-4o` (requires `AZURE_OPENAI_*` config)

## Authentication

The A365 channel uses **Federated Identity Credentials (FIC)** via the Agentic Blueprint to authenticate with Microsoft Graph API on behalf of users.

### T1/T2/User Token Flow

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   T1 Token      │────▶│   T2 Token      │────▶│  User Token     │
│ (client_creds   │     │ (jwt-bearer     │     │ (user_fic       │
│  + fmi_path)    │     │  assertion)     │     │  grant_type)    │
└─────────────────┘     └─────────────────┘     └─────────────────┘
```

The Bot Framework credentials (`A365_APP_ID`, `A365_APP_PASSWORD`) are used for both:
1. Bot Framework message authentication
2. Graph API token acquisition (T1/T2 flow)

## Graph API Tools

The following tools are available to the LLM when Graph API is configured:

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

### Required Settings

| Variable | Description |
|----------|-------------|
| `A365_APP_ID` | Bot Framework App ID |
| `A365_APP_PASSWORD` | Bot Framework App Password |
| `A365_TENANT_ID` | Azure AD Tenant ID |
| `AA_INSTANCE_ID` | Autonomous Agent Instance ID (for FIC) |
| `AGENT_IDENTITY` | Agent service account UPN |
| `OWNER` | Owner's email address |
| `OWNER_AAD_ID` | Owner's AAD Object ID |

### Optional Settings

| Variable | Default | Description |
|----------|---------|-------------|
| `OPENCLAW_MODEL` | `anthropic/claude-opus-4-6` | Primary LLM model |
| `OPENCLAW_FALLBACK_MODELS` | - | Comma-separated fallback models |
| `BUSINESS_HOURS_START` | `08:00` | Business hours start |
| `BUSINESS_HOURS_END` | `18:00` | Business hours end |
| `TIMEZONE` | `America/Los_Angeles` | Timezone |
| `DM_POLICY` | `pairing` | DM policy: `open`, `pairing`, `closed` |

### API Keys (at least one required)

| Variable | Description |
|----------|-------------|
| `ANTHROPIC_API_KEY` | Anthropic API key |
| `OPENAI_API_KEY` | OpenAI API key |
| `OPENROUTER_API_KEY` | OpenRouter API key |
| `AZURE_OPENAI_API_KEY` | Azure OpenAI API key |
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI endpoint URL |

## Identity & Roles

| Property | Description |
|----------|-------------|
| `OWNER` | Email of the person this agent supports (the "principal") |
| `OWNER_AAD_ID` | AAD Object ID of the owner (for role detection) |
| `AGENT_IDENTITY` | Service account email used for Graph API calls |

When the owner interacts with the agent, they get `UserRole: Owner`. Others get `UserRole: Requester`.

## Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────────────┐
│ Microsoft Teams │───▶│  A365 Service   │───▶│    OpenClaw A365        │
│ Outlook/Email   │    │                 │    │    ┌───────────────┐    │
└─────────────────┘    └─────────────────┘    │    │  Claude/GPT   │    │
                                              │    │               │    │
        ┌─────────────────────────────────────│────│  Graph Tools  │    │
        │                                     │    └───────────────┘    │
        ▼                                     └─────────────────────────┘
   ┌─────────┐
   │ Graph   │  ◄── T1/T2/User Token Flow (Agentic Blueprint)
   │ API     │
   └─────────┘
```

## License

MIT
