import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import {
  resolveA365Credentials,
  resolveGraphTokenConfig,
  getGraphToken,
  invalidateGraphTokenCache,
} from "./token.js";
import type { A365Config } from "./types.js";

describe("token", () => {
  const originalEnv = process.env;

  beforeEach(() => {
    process.env = { ...originalEnv };
    invalidateGraphTokenCache();
  });

  afterEach(() => {
    process.env = originalEnv;
    vi.restoreAllMocks();
  });

  describe("resolveA365Credentials", () => {
    it("returns undefined when no credentials are configured", () => {
      const result = resolveA365Credentials();
      expect(result).toBeUndefined();
    });

    it("resolves credentials from config", () => {
      const cfg: A365Config = {
        appId: "test-app-id",
        appPassword: "test-app-password",
        tenantId: "test-tenant-id",
      };
      const result = resolveA365Credentials(cfg);
      expect(result).toEqual({
        appId: "test-app-id",
        appPassword: "test-app-password",
        tenantId: "test-tenant-id",
      });
    });

    it("resolves credentials from environment variables", () => {
      process.env.A365_APP_ID = "env-app-id";
      process.env.A365_APP_PASSWORD = "env-app-password";
      process.env.A365_TENANT_ID = "env-tenant-id";

      const result = resolveA365Credentials();
      expect(result).toEqual({
        appId: "env-app-id",
        appPassword: "env-app-password",
        tenantId: "env-tenant-id",
      });
    });

    it("prefers config over environment variables", () => {
      process.env.A365_APP_ID = "env-app-id";
      process.env.A365_APP_PASSWORD = "env-app-password";
      process.env.A365_TENANT_ID = "env-tenant-id";

      const cfg: A365Config = {
        appId: "config-app-id",
        appPassword: "config-app-password",
        tenantId: "config-tenant-id",
      };

      const result = resolveA365Credentials(cfg);
      expect(result).toEqual({
        appId: "config-app-id",
        appPassword: "config-app-password",
        tenantId: "config-tenant-id",
      });
    });

    it("returns undefined when only partial credentials are provided", () => {
      const cfg: A365Config = {
        appId: "test-app-id",
        // Missing appPassword and tenantId
      };
      const result = resolveA365Credentials(cfg);
      expect(result).toBeUndefined();
    });
  });

  describe("resolveGraphTokenConfig", () => {
    it("returns undefined when no Graph config is provided", () => {
      const result = resolveGraphTokenConfig();
      expect(result).toBeUndefined();
    });

    it("resolves Graph config from config object", () => {
      const cfg: A365Config = {
        graph: {
          clientId: "graph-client-id",
          clientSecret: "graph-client-secret",
          tenantId: "graph-tenant-id",
        },
      };
      const result = resolveGraphTokenConfig(cfg);
      expect(result).toEqual({
        clientId: "graph-client-id",
        clientSecret: "graph-client-secret",
        tenantId: "graph-tenant-id",
        scope: "https://graph.microsoft.com/.default",
      });
    });

    it("uses main tenantId as fallback for Graph tenantId", () => {
      const cfg: A365Config = {
        tenantId: "main-tenant-id",
        graph: {
          clientId: "graph-client-id",
          clientSecret: "graph-client-secret",
        },
      };
      const result = resolveGraphTokenConfig(cfg);
      expect(result?.tenantId).toBe("main-tenant-id");
    });

    it("resolves Graph config from environment variables", () => {
      process.env.GRAPH_CLIENT_ID = "env-graph-client-id";
      process.env.GRAPH_CLIENT_SECRET = "env-graph-client-secret";
      process.env.A365_TENANT_ID = "env-tenant-id";

      const result = resolveGraphTokenConfig();
      expect(result).toEqual({
        clientId: "env-graph-client-id",
        clientSecret: "env-graph-client-secret",
        tenantId: "env-tenant-id",
        scope: "https://graph.microsoft.com/.default",
      });
    });

    it("uses custom scope when provided", () => {
      const cfg: A365Config = {
        graph: {
          clientId: "graph-client-id",
          clientSecret: "graph-client-secret",
          tenantId: "graph-tenant-id",
          scope: "https://custom.scope/.default",
        },
      };
      const result = resolveGraphTokenConfig(cfg);
      expect(result?.scope).toBe("https://custom.scope/.default");
    });
  });

  describe("getGraphToken", () => {
    it("returns undefined when Graph config is not available", async () => {
      const result = await getGraphToken();
      expect(result).toBeUndefined();
    });

    it("fetches token from Azure AD", async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({
          access_token: "test-access-token",
          expires_in: 3600,
        }),
      });
      global.fetch = mockFetch;

      const cfg: A365Config = {
        graph: {
          clientId: "graph-client-id",
          clientSecret: "graph-client-secret",
          tenantId: "graph-tenant-id",
        },
      };

      const result = await getGraphToken(cfg);
      expect(result).toBe("test-access-token");
      expect(mockFetch).toHaveBeenCalledWith(
        "https://login.microsoftonline.com/graph-tenant-id/oauth2/v2.0/token",
        expect.objectContaining({
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
        }),
      );
    });

    it("caches token and returns cached value on subsequent calls", async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: true,
        json: async () => ({
          access_token: "cached-token",
          expires_in: 3600,
        }),
      });
      global.fetch = mockFetch;

      const cfg: A365Config = {
        graph: {
          clientId: "graph-client-id",
          clientSecret: "graph-client-secret",
          tenantId: "graph-tenant-id",
        },
      };

      // First call
      const result1 = await getGraphToken(cfg);
      expect(result1).toBe("cached-token");
      expect(mockFetch).toHaveBeenCalledTimes(1);

      // Second call should use cache
      const result2 = await getGraphToken(cfg);
      expect(result2).toBe("cached-token");
      expect(mockFetch).toHaveBeenCalledTimes(1); // Still 1, not 2
    });

    it("returns undefined when token request fails", async () => {
      const mockFetch = vi.fn().mockResolvedValue({
        ok: false,
        status: 401,
        text: async () => "Unauthorized",
      });
      global.fetch = mockFetch;

      const cfg: A365Config = {
        graph: {
          clientId: "invalid-client-id",
          clientSecret: "invalid-secret",
          tenantId: "tenant-id",
        },
      };

      // Clear cache first
      invalidateGraphTokenCache(cfg);

      const result = await getGraphToken(cfg);
      expect(result).toBeUndefined();
    });
  });
});
