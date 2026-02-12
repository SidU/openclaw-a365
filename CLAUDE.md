# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OpenClaw A365 is a Microsoft 365 Agents (A365) channel plugin for OpenClaw. It provides native integration with Microsoft 365 through the Bot Framework and Graph API, allowing agents to:
- Receive/send messages through Microsoft 365 Agents (Teams, Outlook)
- Perform calendar, email, and user operations via Microsoft Graph API
- Authenticate as an agentic identity (its own service account, not on-behalf-of-user)

## Build and Run Commands

**Development with Docker Compose (recommended):**
```bash
docker-compose up -d        # Build from source and run
docker-compose logs -f      # View logs
docker-compose down         # Stop
```

**Run pre-built image:**
```bash
docker run --cap-add=NET_ADMIN --env-file .env -p 3978:3978 ghcr.io/sidu/openclaw-a365:latest
```

**OpenClaw CLI commands (inside container or local dev):**
```bash
pnpm openclaw doctor --fix              # Validate/fix configuration
pnpm openclaw models set <model>        # Set primary model
pnpm openclaw models fallbacks add <m>  # Add fallback model
pnpm openclaw models status             # Check model configuration
pnpm openclaw gateway                   # Start the gateway
pnpm openclaw plugins install .         # Install plugin from current dir
pnpm openclaw plugins enable a365       # Enable A365 plugin
```

**Note:** There is no explicit test runner configured. Test files exist (`src/*.test.ts`) but are designed for the OpenClaw test framework.

## Architecture

### Core Components

| File | Purpose |
|------|---------|
| `index.ts` | Plugin entry point - exports `plugin` object that registers with OpenClaw |
| `src/channel.ts` | Channel plugin implementation - registers with OpenClaw, provides capabilities |
| `src/monitor.ts` | Bot Framework webhook listener - receives A365 messages on port 3978 |
| `src/token.ts` | Token management - T1/T2/Agent token flow, in-memory cache |
| `src/graph-tools.ts` | Graph API tools for LLM - calendar, email, user operations |
| `src/outbound.ts` | Sends messages back to Bot Framework |
| `src/conversation-store.ts` | Persists conversation references for proactive messaging |
| `src/types.ts` | TypeScript types and config schemas |

### Message Flow

1. **Inbound**: Teams/Outlook → A365 Service → POST `/api/messages` → `monitor.ts`
2. **Processing**: Bot activity → `extractMessageMetadata()` → LLM processes with Graph tools
3. **Graph API**: LLM uses tools from `graph-tools.ts` → authenticated via `token.ts`
4. **Outbound**: Response → `outbound.ts` → Bot Framework API → Teams/Outlook

### Key Patterns

**Thread-safe context (`src/graph-tools.ts`):**
Uses `AsyncLocalStorage` for isolating request context. Always wrap request handlers with `runWithGraphToolContext()` - this prevents cross-request data leakage.

**Token acquisition (`src/token.ts`):**
- T1 Token: client credentials + fmi_path
- T2 Token: JWT bearer assertion from T1
- Agent Token: User FIC for agent identity
- All tokens cached in-memory with 5-minute buffer before expiration

**Network policy (`scripts/network-policy.sh`):**
Enforces iptables rules based on `NETWORK_MODE` env var. Runs at container startup, resolves domains to IPs, and blocks all non-allowlisted outbound traffic.

**Proactive messaging (`src/conversation-store.ts`):**
Stores conversation references (serviceUrl, conversationId, botId, etc.) to a JSON file at `~/.openclaw/a365-conversations.json`. This enables sending messages back to conversations later (e.g., cron job results, async task completion). The store is automatically populated when messages are received and looked up when sending via `sendMessageA365()`.

## Configuration

Configuration is via environment variables. Copy `.env.example` to `.env` and fill in:

**Required:**
- `A365_APP_ID`, `A365_APP_PASSWORD`, `A365_TENANT_ID` - Bot Framework credentials
- `AA_INSTANCE_ID` - Agentic Blueprint for token flow
- `AGENT_IDENTITY` - Agent's UPN (e.g., `agent@contoso.com`)
- `OWNER`, `OWNER_AAD_ID` - Owner identity
- At least one LLM API key (`ANTHROPIC_API_KEY`, `OPENAI_API_KEY`, etc.)

**Network modes:**
- `unrestricted` (default): No network restrictions
- `restricted`: Only Microsoft + LLM provider domains
- `allowlist`: Essential domains + `NETWORK_ALLOWLIST` custom domains

## Key Concepts

**Agentic Identity**: The agent has its own Entra ID user account (`AGENT_IDENTITY`). It only accesses resources explicitly shared with it - this is different from traditional "on-behalf-of" OAuth flows. Users share calendars/resources with the agent like they would with a colleague.

**Role detection**: When `OWNER_AAD_ID` matches the message sender, they get `UserRole: Owner`. Others get `UserRole: Requester`.
