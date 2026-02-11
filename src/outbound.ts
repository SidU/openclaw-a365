import type { ChannelOutboundAdapter } from "openclaw/plugin-sdk";
import type { A365Config, A365MessageMetadata } from "./types.js";
import { resolveA365Credentials } from "./token.js";
import { getA365Runtime } from "./runtime.js";

/**
 * Cached Bot Framework token with expiration.
 */
type CachedBotToken = {
  accessToken: string;
  expiresAt: number;
};

/**
 * Bot Framework token cache.
 * Key format: "appId|tenantId"
 *
 * TODO: For multi-instance deployments (Kubernetes, etc.), consider using
 * a distributed cache (Redis) instead of in-memory storage.
 */
const botTokenCache = new Map<string, CachedBotToken>();

/**
 * Get an access token for Bot Framework using client credentials.
 * Implements caching with 5-minute expiration buffer.
 */
async function getBotFrameworkToken(
  appId: string,
  appPassword: string,
  tenantId: string,
): Promise<string | undefined> {
  const cacheKey = `${appId}|${tenantId}`;
  const cached = botTokenCache.get(cacheKey);

  // Return cached token if still valid (with 5 minute buffer)
  if (cached && cached.expiresAt > Date.now() + 5 * 60 * 1000) {
    return cached.accessToken;
  }

  const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: appId,
    client_secret: appPassword,
    scope: "https://api.botframework.com/.default",
  });

  try {
    const response = await fetch(tokenEndpoint, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!response.ok) {
      const errorText = await response.text();
      const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });
      log.error("Failed to get bot framework token", { status: response.status, error: errorText });
      return undefined;
    }

    const data = (await response.json()) as { access_token: string; expires_in?: number };

    // Cache the token
    const expiresIn = data.expires_in ?? 3600; // Default to 1 hour if not provided
    botTokenCache.set(cacheKey, {
      accessToken: data.access_token,
      expiresAt: Date.now() + expiresIn * 1000,
    });

    return data.access_token;
  } catch (err) {
    const log = getA365Runtime().logging.getChildLogger({ name: "a365-outbound" });
    log.error("Error getting bot framework token", { error: String(err) });
    return undefined;
  }
}

/**
 * Send a message to a conversation via Bot Framework REST API.
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

  if (!creds) {
    return { ok: false, error: "A365 credentials not configured" };
  }

  // If we have service URL and conversation ID from metadata, use proactive messaging
  const conversationServiceUrl = serviceUrl || metadata?.serviceUrl;
  const conversationId = to || metadata?.conversationId;

  if (!conversationServiceUrl || !conversationId) {
    return { ok: false, error: "Missing service URL or conversation ID for proactive message" };
  }

  try {
    // Get access token
    const token = await getBotFrameworkToken(creds.appId, creds.appPassword, creds.tenantId);
    if (!token) {
      return { ok: false, error: "Failed to get Bot Framework access token" };
    }

    // Send message via REST API
    const url = `${conversationServiceUrl.replace(/\/$/, "")}/v3/conversations/${encodeURIComponent(conversationId)}/activities`;

    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: "message",
        text,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return { ok: false, error: `Failed to send message: ${response.status} ${errorText}` };
    }

    const result = (await response.json()) as { id?: string };

    return {
      ok: true,
      messageId: result.id,
      conversationId,
    };
  } catch (err) {
    const runtime = getA365Runtime();
    runtime.logging.getChildLogger({ name: "a365" }).error("send failed", { error: String(err) });
    return { ok: false, error: String(err) };
  }
}

/**
 * Send an Adaptive Card to a conversation.
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

  const conversationServiceUrl = serviceUrl || metadata?.serviceUrl;
  const conversationId = to || metadata?.conversationId;

  if (!conversationServiceUrl || !conversationId) {
    return { ok: false, error: "Missing service URL or conversation ID" };
  }

  try {
    // Get access token
    const token = await getBotFrameworkToken(creds.appId, creds.appPassword, creds.tenantId);
    if (!token) {
      return { ok: false, error: "Failed to get Bot Framework access token" };
    }

    // Send card via REST API
    const url = `${conversationServiceUrl.replace(/\/$/, "")}/v3/conversations/${encodeURIComponent(conversationId)}/activities`;

    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: card,
          },
        ],
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return { ok: false, error: `Failed to send card: ${response.status} ${errorText}` };
    }

    const result = (await response.json()) as { id?: string };

    return {
      ok: true,
      messageId: result.id,
      conversationId,
    };
  } catch (err) {
    const runtime = getA365Runtime();
    runtime.logging.getChildLogger({ name: "a365" }).error("send card failed", { error: String(err) });
    return { ok: false, error: String(err) };
  }
}

/**
 * A365 outbound adapter for sending messages.
 */
export const a365Outbound: ChannelOutboundAdapter = {
  deliveryMode: "direct",
  textChunkLimit: 4000,

  sendText: async ({ cfg, to, text }) => {
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
