import type { ChannelOutboundAdapter } from "openclaw/plugin-sdk";
import type { A365Config, A365MessageMetadata } from "./types.js";
import { resolveA365Credentials, getGraphToken } from "./token.js";
import { getA365Runtime } from "./runtime.js";
import { getConversationReference, getConversationReferenceByUser } from "./conversation-store.js";

/**
 * APX (Agent Platform Exchange) scope — the audience the Bot Framework
 * service URL expects for agentic identity tokens.
 */
const APX_SCOPE = "5a807f24-c9de-44ee-a3a7-329e88a00ffc/.default";

/**
 * Resolve conversation ID and service URL from the provided params and the
 * conversation store. Returns null if either value cannot be resolved.
 */
async function resolveConversation(
  to: string | undefined,
  serviceUrl: string | undefined,
  metadata: A365MessageMetadata | undefined,
  tenantId: string | undefined,
): Promise<{ conversationId: string; serviceUrl: string } | null> {
  const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });

  let conversationId = to || metadata?.conversationId;
  let conversationServiceUrl = serviceUrl || metadata?.serviceUrl;

  log.info(`Resolved conversation params: conversationId=${conversationId} serviceUrl=${conversationServiceUrl}`);

  if (!conversationServiceUrl && conversationId) {
    log.info(`Looking up stored conversation reference for conversationId=${conversationId}`);
    let storedRef = await getConversationReference(conversationId);

    // If direct lookup fails and conversationId looks like a userAadId (UUID without colons/@),
    // try looking up by userAadId as a fallback
    if (!storedRef && !conversationId.includes(":") && !conversationId.includes("@")) {
      log.info(`Direct lookup failed, trying userAadId lookup for ${conversationId}`);
      storedRef = await getConversationReferenceByUser(conversationId);
    }

    if (storedRef) {
      conversationServiceUrl = storedRef.serviceUrl;
      conversationId = storedRef.conversationId;
      log.info(`Found stored conversation reference: conversationId=${conversationId} serviceUrl=${conversationServiceUrl}`);
    }
  }

  // Fallback: Construct Teams service URL from tenant ID if still missing
  if (!conversationServiceUrl && tenantId) {
    conversationServiceUrl = `https://smba.trafficmanager.net/amer/${tenantId}/`;
    log.info("Using constructed Teams service URL", { serviceUrl: conversationServiceUrl });
  }

  if (!conversationServiceUrl || !conversationId) {
    return null;
  }

  return { conversationId, serviceUrl: conversationServiceUrl };
}

/**
 * Send a proactive message using the agent's own identity.
 *
 * Uses our manual T1/T2/User FIC token flow (which uses the correct `username`
 * field) with the APX scope, then sends via the SDK's ConnectorClient.
 *
 * Note: The SDK's built-in getAgenticUserToken() has a bug — it sends `user_id`
 * instead of `username` in the User FIC request, causing AADSTS50000 errors.
 * Our manual flow works around this.
 */
async function sendViaConnectorClient(params: {
  conversationId: string;
  serviceUrl: string;
  token: string;
  activity: Record<string, unknown>;
}): Promise<{ id?: string }> {
  const { conversationId, serviceUrl, token, activity } = params;
  const { ConnectorClient } = await import("@microsoft/agents-hosting");

  const client = ConnectorClient.createClientWithToken(serviceUrl, token);
  const result = await client.sendToConversation(conversationId, activity);
  return { id: (result as { id?: string })?.id };
}

/**
 * Send a message to a conversation using the agent's own identity.
 * Acquires an APX-scoped token via T1/T2/User FIC, then sends via ConnectorClient.
 */
export async function sendMessageA365(params: {
  cfg: unknown;
  to: string;
  text: string;
  serviceUrl?: string;
  metadata?: A365MessageMetadata;
}): Promise<{ ok: boolean; messageId?: string; conversationId?: string; error?: string }> {
  const { cfg, to, text, serviceUrl, metadata } = params;
  const a365Cfg = (cfg as { channels?: { a365?: A365Config } })?.channels?.a365;
  const creds = resolveA365Credentials(a365Cfg);

  const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });

  log.info(`sendMessageA365 called: to=${to} serviceUrl=${serviceUrl} hasMetadata=${!!metadata} hasA365Cfg=${!!a365Cfg} hasCreds=${!!creds}`);

  if (!creds) {
    log.error("A365 credentials not configured");
    return { ok: false, error: "A365 credentials not configured" };
  }

  const agentIdentity = a365Cfg?.agentIdentity || process.env.AGENT_IDENTITY;
  if (!agentIdentity) {
    log.error("Agent identity not configured");
    return { ok: false, error: "Agent identity not configured. Set AGENT_IDENTITY env var." };
  }

  const resolved = await resolveConversation(to, serviceUrl, metadata, creds.tenantId);
  if (!resolved) {
    return { ok: false, error: "Missing service URL or conversation ID. User must message the bot first." };
  }

  // Acquire APX-scoped token via T1/T2/User FIC
  log.info("Acquiring APX token via agent identity", {
    agentIdentity,
    scope: APX_SCOPE,
    conversationId: resolved.conversationId,
  });

  const token = await getGraphToken(a365Cfg, agentIdentity, APX_SCOPE);
  if (!token) {
    log.error("Failed to acquire APX token via agent identity");
    return { ok: false, error: "Failed to acquire APX token. Check agent identity and T1/T2 configuration." };
  }

  try {
    log.info("Sending proactive message via ConnectorClient", {
      conversationId: resolved.conversationId,
      textLength: text.length,
    });

    const result = await sendViaConnectorClient({
      conversationId: resolved.conversationId,
      serviceUrl: resolved.serviceUrl,
      token,
      activity: { type: "message", text },
    });

    log.info("Proactive message sent successfully", { messageId: result.id });

    return {
      ok: true,
      messageId: result.id,
      conversationId: resolved.conversationId,
    };
  } catch (err) {
    const axErr = err as { response?: { data?: unknown; status?: number }; message?: string };
    const detail = axErr.response?.data ? JSON.stringify(axErr.response.data) : (err instanceof Error ? err.message : String(err));
    log.error(`Proactive message failed: ${detail}`);
    return { ok: false, error: String(detail) };
  }
}

/**
 * Send an Adaptive Card to a conversation using the agent's own identity.
 */
export async function sendAdaptiveCardA365(params: {
  cfg: unknown;
  to: string;
  card: Record<string, unknown>;
  serviceUrl?: string;
  metadata?: A365MessageMetadata;
}): Promise<{ ok: boolean; messageId?: string; conversationId?: string; error?: string }> {
  const { cfg, to, card, serviceUrl, metadata } = params;
  const a365Cfg = (cfg as { channels?: { a365?: A365Config } })?.channels?.a365;
  const creds = resolveA365Credentials(a365Cfg);

  if (!creds) {
    return { ok: false, error: "A365 credentials not configured" };
  }

  const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });

  const agentIdentity = a365Cfg?.agentIdentity || process.env.AGENT_IDENTITY;
  if (!agentIdentity) {
    return { ok: false, error: "Agent identity not configured. Set AGENT_IDENTITY env var." };
  }

  const resolved = await resolveConversation(to, serviceUrl, metadata, creds.tenantId);
  if (!resolved) {
    return { ok: false, error: "Missing service URL or conversation ID. User must message the bot first." };
  }

  const token = await getGraphToken(a365Cfg, agentIdentity, APX_SCOPE);
  if (!token) {
    return { ok: false, error: "Failed to acquire APX token." };
  }

  try {
    const result = await sendViaConnectorClient({
      conversationId: resolved.conversationId,
      serviceUrl: resolved.serviceUrl,
      token,
      activity: {
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: card,
          },
        ],
      },
    });

    log.info("Card sent successfully", { messageId: result.id });

    return {
      ok: true,
      messageId: result.id,
      conversationId: resolved.conversationId,
    };
  } catch (err) {
    const axErr = err as { response?: { data?: unknown; status?: number }; message?: string };
    const detail = axErr.response?.data ? JSON.stringify(axErr.response.data) : (err instanceof Error ? err.message : String(err));
    log.error(`Card send failed: ${detail}`);
    return { ok: false, error: String(detail) };
  }
}

/**
 * Normalize an A365 target to a conversation ID.
 * Handles various formats that might come from session storage.
 */
function normalizeA365Target(to: string | undefined): string | undefined {
  if (!to) return undefined;
  const trimmed = to.trim();
  if (!trimmed) return undefined;

  // Strip common prefixes that might be stored in session lastTo
  // e.g., "user:xxx", "conversation:xxx", "a365:xxx"
  const prefixes = ["user:", "conversation:", "a365:", "a365:group:"];
  for (const prefix of prefixes) {
    if (trimmed.startsWith(prefix)) {
      return trimmed.slice(prefix.length);
    }
  }

  // Return as-is if no prefix (already a raw conversationId)
  return trimmed;
}

/**
 * A365 outbound adapter for sending messages.
 */
export const a365Outbound: ChannelOutboundAdapter = {
  deliveryMode: "direct",
  textChunkLimit: 4000,

  resolveTarget: ({ to, allowFrom }) => {
    const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });
    log.info("resolveTarget called", { to, allowFromCount: allowFrom?.length ?? 0 });

    const normalized = normalizeA365Target(to);

    if (normalized) {
      log.info("resolveTarget success", { normalized });
      return { ok: true, to: normalized };
    }

    // Fall back to first allowFrom entry if available
    const allowList = (allowFrom ?? []).map((entry) => String(entry).trim()).filter(Boolean);
    if (allowList.length > 0) {
      const fallback = normalizeA365Target(allowList[0]);
      if (fallback) {
        log.info("resolveTarget fallback", { fallback });
        return { ok: true, to: fallback };
      }
    }

    log.warn("resolveTarget failed - no target");
    return {
      ok: false,
      error: "No A365 conversation target specified. User must message the bot first.",
    };
  },

  sendText: async ({ cfg, to, text }) => {
    const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });
    log.info("sendText called", { to, textLength: text?.length ?? 0 });

    const result = await sendMessageA365({ cfg, to, text });
    if (!result.ok) {
      return {
        channel: "a365",
        ok: false,
        error: result.error,
      };
    }
    return {
      channel: "a365",
      ok: true,
      messageId: result.messageId,
      conversationId: result.conversationId,
    };
  },

  sendMedia: async ({ cfg, to, text, mediaUrl }) => {
    // TODO: Implement proper media attachment support via Bot Framework:
    // 1. Upload file to OneDrive/SharePoint using Graph API
    // 2. Create contentUrl reference
    // 3. Send as attachment with proper contentType
    // See: https://learn.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-add-media-attachments
    // For now, we just send the URL as a link.
    const messageText = mediaUrl ? `${text}\n\n${mediaUrl}` : text;
    const result = await sendMessageA365({ cfg, to, text: messageText });
    if (!result.ok) {
      return {
        channel: "a365",
        ok: false,
        error: result.error,
      };
    }
    return {
      channel: "a365",
      ok: true,
      messageId: result.messageId,
      conversationId: result.conversationId,
    };
  },
};

/**
 * Normalize A365 messaging target.
 */
export function normalizeA365MessagingTarget(raw: string): string | undefined {
  const trimmed = raw.trim();
  if (!trimmed) {
    return undefined;
  }

  // Handle conversation: prefix
  if (trimmed.toLowerCase().startsWith("conversation:")) {
    return trimmed.slice("conversation:".length).trim() || undefined;
  }

  // Handle user: prefix
  if (trimmed.toLowerCase().startsWith("user:")) {
    return `user:${trimmed.slice("user:".length).trim()}`;
  }

  // Return as-is if it looks like a conversation ID
  if (trimmed.includes("@") || trimmed.includes(":")) {
    return trimmed;
  }

  return trimmed;
}
