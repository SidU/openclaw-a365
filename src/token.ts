import type { A365Config } from "./types.js";
import { getA365Runtime } from "./runtime.js";

/**
 * Get the logger for token operations.
 * Returns a no-op logger if runtime is not yet initialized.
 */
function getLogger() {
  try {
    return getA365Runtime().logging.getChildLogger({ name: "a365-token" });
  } catch {
    // Runtime not initialized yet, return a no-op logger
    return {
      debug: () => {},
      info: () => {},
      warn: () => {},
      error: () => {},
    };
  }
}

export type A365Credentials = {
  appId: string;
  appPassword: string;
  tenantId: string;
};

export type GraphTokenConfig = {
  tenantId: string;
  blueprintClientAppId: string;
  blueprintClientSecret: string;
  aaInstanceId: string;
  scope?: string;
};

export type CachedToken = {
  accessToken: string;
  expiresAt: number;
};

export type TokenCallbackConfig = {
  callbackUrl: string;
  callbackToken?: string;
};

/**
 * Token cache for Graph API tokens: key = "username|scope"
 *
 * TODO: For multi-instance deployments (Kubernetes, load balancer), this in-memory
 * cache leads to redundant token fetches. Consider:
 * - Redis-based token cache with TTL
 * - Distributed cache (e.g., Azure Cache for Redis)
 * - Centralized token service
 */
const tokenCache = new Map<string, CachedToken>();

/**
 * Resolve A365 Bot Framework credentials from config or environment variables.
 */
export function resolveA365Credentials(cfg?: A365Config): A365Credentials | undefined {
  const appId = cfg?.appId?.trim() || process.env.A365_APP_ID?.trim();
  const appPassword = cfg?.appPassword?.trim() || process.env.A365_APP_PASSWORD?.trim();
  const tenantId = cfg?.tenantId?.trim() || process.env.A365_TENANT_ID?.trim();

  if (!appId || !appPassword || !tenantId) {
    return undefined;
  }

  return { appId, appPassword, tenantId };
}

/**
 * Resolve Graph API token configuration for T1/T2/User flow.
 */
export function resolveGraphTokenConfig(cfg?: A365Config): GraphTokenConfig | undefined {
  const tenantId = cfg?.tenantId?.trim() || process.env.A365_TENANT_ID?.trim();
  const blueprintClientAppId =
    cfg?.graph?.blueprintClientAppId?.trim() ||
    cfg?.appId?.trim() ||
    process.env.BLUEPRINT_CLIENT_APP_ID?.trim() ||
    process.env.A365_APP_ID?.trim();
  const blueprintClientSecret =
    cfg?.graph?.blueprintClientSecret?.trim() ||
    cfg?.appPassword?.trim() ||
    process.env.BLUEPRINT_CLIENT_SECRET?.trim() ||
    process.env.A365_APP_PASSWORD?.trim();
  const aaInstanceId =
    cfg?.graph?.aaInstanceId?.trim() || process.env.AA_INSTANCE_ID?.trim();
  const scope = cfg?.graph?.scope?.trim() || "https://graph.microsoft.com/.default";

  if (!tenantId || !blueprintClientAppId || !blueprintClientSecret || !aaInstanceId) {
    return undefined;
  }

  return { tenantId, blueprintClientAppId, blueprintClientSecret, aaInstanceId, scope };
}

/**
 * Resolve token callback configuration for external token service.
 */
export function resolveTokenCallbackConfig(cfg?: A365Config): TokenCallbackConfig | undefined {
  const callbackUrl =
    cfg?.tokenCallback?.url?.trim() || process.env.TOKEN_CALLBACK_URL?.trim();

  if (!callbackUrl) {
    return undefined;
  }

  return {
    callbackUrl,
    callbackToken: cfg?.tokenCallback?.token?.trim() || process.env.TOKEN_CALLBACK_TOKEN?.trim(),
  };
}

/**
 * Get a cached Graph API token or fetch a new one.
 * Supports two modes:
 * 1. T1/T2/User flow (Federated Identity Credentials) - for A365 Agents
 * 2. External callback - calls an external service for tokens
 *
 * Error handling: Returns undefined if no token could be acquired.
 * This allows callers to check `if (!token)` rather than try/catch,
 * providing a simpler API at the cost of losing error details.
 * Error details are logged for troubleshooting.
 */
export async function getGraphToken(
  cfg: A365Config | undefined,
  username: string,
  scope?: string,
): Promise<string | undefined> {
  const log = getLogger();
  const effectiveScope = scope || cfg?.graph?.scope || "https://graph.microsoft.com/.default";
  const cacheKey = `${username}|${effectiveScope}`;

  log.debug("getGraphToken called", { username, scope: effectiveScope });
  const cached = tokenCache.get(cacheKey);

  // Return cached token if still valid (with 5 minute buffer)
  if (cached && cached.expiresAt > Date.now() + 5 * 60 * 1000) {
    return cached.accessToken;
  }

  // Try external callback first (if configured)
  const callbackConfig = resolveTokenCallbackConfig(cfg);
  if (callbackConfig) {
    log.debug("Trying external token callback", { url: callbackConfig.callbackUrl });
    try {
      const token = await fetchTokenFromCallback(callbackConfig, username, effectiveScope);
      if (token) {
        log.debug("Token acquired from callback", { expiresAt: new Date(token.expiresAt).toISOString() });
        tokenCache.set(cacheKey, token);
        return token.accessToken;
      }
    } catch (err) {
      log.warn("Token callback failed, falling back to T1/T2 flow", { error: String(err) });
    }
  }

  // Fall back to T1/T2/User flow
  const tokenConfig = resolveGraphTokenConfig(cfg);
  log.debug("T1/T2 config resolved", {
    configured: !!tokenConfig,
    tenantId: tokenConfig?.tenantId,
    aaInstanceId: tokenConfig?.aaInstanceId,
  });

  if (tokenConfig) {
    try {
      log.debug("Attempting T1/T2/User flow", { username });
      const token = await fetchGraphTokenT1T2(tokenConfig, username, effectiveScope);
      if (token) {
        log.info("Token acquired via T1/T2 flow", { expiresAt: new Date(token.expiresAt).toISOString() });
        tokenCache.set(cacheKey, token);
        return token.accessToken;
      }
    } catch (err) {
      log.error(`T1/T2 token flow failed: ${String(err)}`);
    }
  } else {
    log.debug("T1/T2 not configured", {
      hasA365TenantId: !!process.env.A365_TENANT_ID,
      hasBlueprintSecret: !!process.env.BLUEPRINT_CLIENT_SECRET,
      hasAaInstanceId: !!process.env.AA_INSTANCE_ID,
    });
  }

  log.warn("No token acquired", { username });
  return undefined;
}

/**
 * Fetch token from external callback service.
 * This allows integration with existing .NET token services.
 */
async function fetchTokenFromCallback(
  config: TokenCallbackConfig,
  username: string,
  scope: string,
): Promise<CachedToken | undefined> {
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };

  if (config.callbackToken) {
    headers["Authorization"] = `Bearer ${config.callbackToken}`;
  }

  const response = await fetch(config.callbackUrl, {
    method: "POST",
    headers,
    body: JSON.stringify({
      username,
      scope,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Token callback failed: ${response.status} ${errorText}`);
  }

  const data = (await response.json()) as {
    access_token: string;
    expires_at?: string;
    expires_in?: number;
  };

  let expiresAt: number;
  if (data.expires_at) {
    expiresAt = new Date(data.expires_at).getTime();
  } else if (data.expires_in) {
    expiresAt = Date.now() + data.expires_in * 1000;
  } else {
    // Default to 1 hour
    expiresAt = Date.now() + 3600 * 1000;
  }

  return {
    accessToken: data.access_token,
    expiresAt,
  };
}

/**
 * Fetch Graph API token using T1/T2/User flow (Federated Identity Credentials).
 * This is the authentication pattern used by Microsoft 365 Agents SDK.
 *
 * Flow:
 * 1. T1 Token: Client credentials with fmi_path parameter
 * 2. T2 Token: Exchange T1 using jwt-bearer assertion
 * 3. User Token: Get token for specific user using user_fic grant type
 *
 * Error handling: This function throws on failure, which is caught by getGraphToken.
 * This keeps the multi-step flow logic clean while allowing getGraphToken to handle
 * retries and fallbacks centrally.
 */
async function fetchGraphTokenT1T2(
  config: GraphTokenConfig,
  username: string,
  scope: string,
): Promise<CachedToken | undefined> {
  const log = getLogger();
  const tokenEndpoint = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;

  log.info(`T1/T2 flow starting: username=${username} scope=${scope}`);

  // Step 1: Acquire T1 Token
  const t1Body = new URLSearchParams({
    scope: "api://AzureAdTokenExchange/.default",
    client_id: config.blueprintClientAppId,
    grant_type: "client_credentials",
    client_secret: config.blueprintClientSecret,
    fmi_path: config.aaInstanceId,
  });

  const t1Response = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: t1Body.toString(),
  });

  if (!t1Response.ok) {
    const errorText = await t1Response.text();
    throw new Error(`T1 token request failed (scope=${scope}): ${t1Response.status} ${errorText}`);
  }

  const t1Data = (await t1Response.json()) as { access_token: string };
  log.info("T1 token acquired successfully");

  // Step 2: Acquire T2 Token
  const t2Body = new URLSearchParams({
    scope: "api://AzureAdTokenExchange/.default",
    client_id: config.aaInstanceId,
    grant_type: "client_credentials",
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    client_assertion: t1Data.access_token,
  });

  const t2Response = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: t2Body.toString(),
  });

  if (!t2Response.ok) {
    const errorText = await t2Response.text();
    throw new Error(`T2 token request failed (scope=${scope}): ${t2Response.status} ${errorText}`);
  }

  const t2Data = (await t2Response.json()) as { access_token: string };
  log.info("T2 token acquired successfully");

  // Step 3: Acquire User Token using Federated Identity Credential
  const userBody = new URLSearchParams({
    scope,
    client_id: config.aaInstanceId,
    grant_type: "user_fic",
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    client_assertion: t1Data.access_token,
    username,
    user_federated_identity_credential: t2Data.access_token,
  });

  const userResponse = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: userBody.toString(),
  });

  if (!userResponse.ok) {
    const errorText = await userResponse.text();
    throw new Error(`User FIC token request failed (scope=${scope}): ${userResponse.status} ${errorText}`);
  }

  const userData = (await userResponse.json()) as {
    access_token: string;
    expires_in: number;
  };

  return {
    accessToken: userData.access_token,
    expiresAt: Date.now() + userData.expires_in * 1000,
  };
}

/**
 * Invalidate cached tokens for a user.
 */
export function invalidateGraphTokenCache(username?: string, scope?: string): void {
  if (username && scope) {
    const cacheKey = `${username}|${scope}`;
    tokenCache.delete(cacheKey);
  } else if (username) {
    // Clear all tokens for this user
    for (const key of tokenCache.keys()) {
      if (key.startsWith(`${username}|`)) {
        tokenCache.delete(key);
      }
    }
  } else {
    // Clear entire cache
    tokenCache.clear();
  }
}

/**
 * Get all cached tokens (for debugging).
 */
export function getTokenCacheStats(): { count: number; users: string[] } {
  const users = new Set<string>();
  for (const key of tokenCache.keys()) {
    const username = key.split("|")[0];
    if (username) {
      users.add(username);
    }
  }
  return {
    count: tokenCache.size,
    users: Array.from(users),
  };
}
