# A365 Proactive Messaging Architecture

## Overview

Proactive messaging allows the bot to send messages to users without a prior incoming message in the same request context. This is essential for:
- **Cron jobs**: Scheduled reminders, daily updates
- **Async task completion**: Notify user when a long-running task finishes
- **External triggers**: Webhooks, alerts from other systems

## Key Components

### 1. Conversation Reference Storage (`conversation-store.ts`)

When a user messages the bot, we store their conversation context:

```typescript
type StoredConversationReference = {
  conversationId: string;      // e.g., "19:399f383c-...@unq.gbl.spaces"
  serviceUrl: string;          // e.g., "https://smba.trafficmanager.net/amer/{tenantId}/"
  channelId: string;           // "msteams"
  botId: string;               // Bot's ID
  userId: string;              // User's ID
  userAadId?: string;          // User's Azure AD Object ID
  tenantId?: string;           // Azure AD Tenant ID
  isGroup: boolean;
  updatedAt: number;
}
```

**Storage location**: `~/.openclaw/a365-conversations.json`

### 2. Session Delivery Context

OpenClaw sessions track the last channel/target for delivery:

```typescript
// In sessions.json
{
  "agent:main:main": {
    "sessionId": "...",
    "lastChannel": "a365",
    "lastTo": "19:399f383c-...@unq.gbl.spaces",  // The conversation ID
    "deliveryContext": {
      "channel": "a365",
      "to": "19:399f383c-...@unq.gbl.spaces"
    }
  }
}
```

### 3. Outbound Adapter (`outbound.ts`)

The A365 outbound adapter handles sending messages:

```typescript
const a365Outbound: ChannelOutboundAdapter = {
  deliveryMode: "direct",

  resolveTarget: ({ to, allowFrom }) => {
    // Normalize the target (strip prefixes like "user:", "conversation:")
    // Return { ok: true, to: conversationId }
  },

  sendText: async ({ cfg, to, text }) => {
    // 1. Get Bot Framework token (via MSAL)
    // 2. Look up serviceUrl from conversation store
    // 3. POST to Bot Framework REST API
    return { ok: true, messageId, conversationId };
  }
}
```

## The Proactive Messaging Flow

### Step 1: User Messages Bot (Establish Context)

```
User → Teams → Bot Framework → A365 Plugin (monitor.ts)
                                    ↓
                              Extract metadata:
                              - conversationId
                              - serviceUrl
                              - tenantId
                                    ↓
                              Save to conversation-store.ts
                                    ↓
                              Update main session:
                              - lastChannel = "a365"
                              - lastTo = conversationId
```

### Step 2: Cron Job Fires (Proactive Send)

```
Cron Timer → Load main session
                 ↓
          Read lastChannel = "a365"
          Read lastTo = "19:...@unq.gbl.spaces"
                 ↓
          Resolve delivery via a365Outbound.resolveTarget()
                 ↓
          Agent generates response text
                 ↓
          Call a365Outbound.sendText({ to: conversationId, text })
                 ↓
          sendMessageA365():
            1. Get Bot Framework token (MSAL client credentials)
            2. Look up serviceUrl from conversation-store.json
            3. POST to: {serviceUrl}/v3/conversations/{conversationId}/activities
                 ↓
          Message appears in Teams
```

## Bot Framework REST API

To send a proactive message, we call:

```
POST {serviceUrl}/v3/conversations/{conversationId}/activities
Authorization: Bearer {botFrameworkToken}
Content-Type: application/json

{
  "type": "message",
  "text": "Hello from cron!"
}
```

### Required Information:
1. **serviceUrl**: The Bot Framework service endpoint (stored in conversation reference)
2. **conversationId**: The Teams conversation ID (stored in session's `lastTo`)
3. **Bot Framework Token**: Obtained via MSAL client credentials flow

## Token Acquisition

Using MSAL (same as Microsoft Agents SDK):

```typescript
const cca = new ConfidentialClientApplication({
  auth: {
    clientId: appId,           // Bot's App ID
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret: appPassword  // Bot's App Secret
  }
});

const token = await cca.acquireTokenByClientCredential({
  scopes: ["https://api.botframework.com/.default"]
});
```

## Investigation Progress

### Phase 1: Initial Issue - `sendText` Never Called

**Symptoms:**
- Cron fires, agent runs (`messageChannel=cron-event`)
- Typing indicator fails (no TurnContext)
- `a365Outbound.sendText()` never called
- Session shows `sessionKey=unknown`

**Resolution:** Added logging, discovered the outbound adapter WAS being called but token acquisition was failing.

---

### Phase 2: Token Acquisition Failure

**Symptoms:**
- `resolveTarget()` called ✓
- `sendText()` called ✓
- Manual MSAL token acquisition failing silently

**Approach tried:** Used MSAL `ConfidentialClientApplication` with client credentials:
```typescript
const token = await cca.acquireTokenByClientCredential({
  scopes: ["https://api.botframework.com/.default"]
});
```

**Key insight from user:** "Agents have to fetch tokens with their own identity, no?"

The Microsoft Agents SDK handles token acquisition differently - it uses the agent's identity, not just client credentials. The SDK's `adapter.continueConversation()` method handles this internally.

**Resolution:** Switched to using `adapter.continueConversation()` instead of manual REST API calls.

---

### Phase 3: Adapter-Based Approach

**Changes made:**
1. Created `adapter-store.ts` to store the CloudAdapter and auth config
2. Updated `monitor.ts` to store the adapter after initialization
3. Rewrote `sendMessageA365()` to use `adapter.continueConversation()`

**Code pattern (from SDK samples):**
```typescript
await adapter.continueConversation(
  authConfig.clientId,
  conversationReference,
  async (context) => {
    await context.sendActivity(text);
  }
);
```

**Result:** `continueConversation` is being called but still failing.

---

### Phase 4: Current Issue - Wrong Conversation ID

**Symptoms:**
- Adapter stored successfully ✓
- `continueConversation` called ✓
- But fails with unknown error

**Root cause discovered:** The `to` value passed to `sendText` is wrong.

**Session has correct value:**
```json
"lastTo": "19:399f383c-1109-4935-ab3b-92a4e02750d6_cc1d956d-e36b-473d-a802-243197616363@unq.gbl.spaces"
```

**But sendText receives:**
```
to=399f383c-1109-4935-ab3b-92a4e02750d6
```

This is just the userAadId, not the full Teams conversation ID. The conversation store lookup fails because the key doesn't match.

**Workaround added:** Fallback lookup by userAadId using `getConversationReferenceByUser()`. But this masks the real issue.

**Real question:** Why is cron delivery passing the userAadId instead of the session's `lastTo`?

---

## Current State

### What's Working
- [x] Conversation reference saved correctly (`~/.openclaw/a365-conversations.json`)
- [x] Main session updated with correct `lastTo` (verified in `sessions.json`)
- [x] `resolveTarget()` being called
- [x] `sendText()` being called
- [x] Adapter stored for proactive messaging
- [x] `continueConversation()` being invoked

### What's Not Working
- [ ] Cron delivery passes wrong `to` value (userAadId instead of full conversation ID)
- [ ] `continueConversation()` fails (error not yet captured in logs)

### Debug Checklist
- [x] Is conversation reference being saved? **YES** - correct format in store
- [x] Is main session updated? **YES** - correct `lastTo` value
- [x] Is `resolveTarget` being called? **YES** - with wrong value
- [x] Is `sendText` being called? **YES** - with wrong value
- [ ] Is `continueConversation` succeeding? **NO** - fails with unknown error
- [ ] Why does cron pass userAadId instead of lastTo? **UNKNOWN**

---

## Files Involved

| File | Purpose |
|------|---------|
| `monitor.ts` | Receives messages, saves conversation reference, updates session, stores adapter |
| `conversation-store.ts` | Persists conversation references (serviceUrl, conversationId) |
| `outbound.ts` | A365 outbound adapter with `resolveTarget` and `sendText` using `continueConversation` |
| `adapter-store.ts` | Stores CloudAdapter and auth config for proactive messaging |
| `token.ts` | Graph API token acquisition (T1/T2 flow) |

---

## Next Steps

### Option A: Fix Root Cause (Recommended)
Investigate why OpenClaw cron delivery passes the wrong `to` value:
1. Check how cron jobs store their delivery target when created
2. Check how cron delivery resolves `to` from session (is it using `lastTo` or something else?)
3. This might be a bug in OpenClaw core, not the A365 plugin

### Option B: Direct Testing
Create a test script to call `sendMessageA365()` directly with correct values:
1. Bypasses cron delivery system
2. Isolates whether `continueConversation` works when given correct values
3. If it works, confirms the issue is upstream in cron delivery

### Option C: Improve Error Visibility
The `continueConversation` error isn't being logged properly:
1. Latest build has improved error logging
2. Test again to capture actual error message
3. Error might reveal what's wrong (auth? conversation format? SDK issue?)

---

## Useful Commands

```bash
# Check conversation store
docker exec openclaw-a365-container cat /root/.openclaw/a365-conversations.json

# Check session lastTo
docker exec openclaw-a365-container grep -A 5 '"agent:main:main"' /root/.openclaw/agents/main/sessions/sessions.json

# Check recent logs for outbound activity
docker exec openclaw-a365-container cat /tmp/openclaw/openclaw-$(date +%Y-%m-%d).log | grep -E "sendText|resolveTarget|continueConversation" | tail -20

# Check cron jobs
docker exec openclaw-a365-container cat /root/.openclaw/cron/jobs.json
```
