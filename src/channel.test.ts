import { describe, it, expect, vi, beforeEach } from "vitest";
import { a365Plugin } from "./channel.js";
import type { OpenClawConfig } from "openclaw/plugin-sdk";

describe("a365Plugin", () => {
  describe("metadata", () => {
    it("has correct id", () => {
      expect(a365Plugin.id).toBe("a365");
    });

    it("has correct meta information", () => {
      expect(a365Plugin.meta.id).toBe("a365");
      expect(a365Plugin.meta.label).toBe("Microsoft 365 Agents");
      expect(a365Plugin.meta.aliases).toContain("m365agents");
    });

    it("has correct capabilities", () => {
      expect(a365Plugin.capabilities.chatTypes).toContain("direct");
      expect(a365Plugin.capabilities.chatTypes).toContain("channel");
      expect(a365Plugin.capabilities.threads).toBe(true);
      expect(a365Plugin.capabilities.media).toBe(true);
    });
  });

  describe("config adapter", () => {
    it("lists single default account", () => {
      const cfg = {} as OpenClawConfig;
      const accountIds = a365Plugin.config.listAccountIds(cfg);
      expect(accountIds).toEqual(["default"]);
    });

    it("resolves account when a365 is not configured", () => {
      const cfg = {} as OpenClawConfig;
      const account = a365Plugin.config.resolveAccount(cfg);
      expect(account.accountId).toBe("default");
      expect(account.enabled).toBe(true);
      expect(account.configured).toBe(false);
    });

    it("resolves account when a365 is configured", () => {
      const cfg = {
        channels: {
          a365: {
            enabled: true,
            appId: "test-app-id",
            appPassword: "test-password",
            tenantId: "test-tenant",
            owner: "user@test.com",
          },
        },
      } as unknown as OpenClawConfig;
      const account = a365Plugin.config.resolveAccount(cfg);
      expect(account.accountId).toBe("default");
      expect(account.enabled).toBe(true);
      expect(account.configured).toBe(true);
      expect(account.owner).toBe("user@test.com");
    });

    it("respects enabled=false setting", () => {
      const cfg = {
        channels: {
          a365: {
            enabled: false,
            appId: "test-app-id",
            appPassword: "test-password",
            tenantId: "test-tenant",
          },
        },
      } as unknown as OpenClawConfig;
      const account = a365Plugin.config.resolveAccount(cfg);
      expect(account.enabled).toBe(false);
    });

    it("setAccountEnabled updates config", () => {
      const cfg = {
        channels: {
          a365: { enabled: true },
        },
      } as unknown as OpenClawConfig;
      const newCfg = a365Plugin.config.setAccountEnabled!({
        cfg,
        accountId: "default",
        enabled: false,
      });
      expect((newCfg.channels?.a365 as { enabled: boolean }).enabled).toBe(false);
    });

    it("deleteAccount removes a365 config", () => {
      const cfg = {
        channels: {
          a365: { enabled: true },
          telegram: { enabled: true },
        },
      } as unknown as OpenClawConfig;
      const newCfg = a365Plugin.config.deleteAccount!({ cfg, accountId: "default" });
      expect(newCfg.channels?.a365).toBeUndefined();
      expect(newCfg.channels?.telegram).toBeDefined();
    });

    it("resolveAllowFrom returns empty array when not configured", () => {
      const cfg = {} as OpenClawConfig;
      const allowFrom = a365Plugin.config.resolveAllowFrom!({ cfg });
      expect(allowFrom).toEqual([]);
    });

    it("resolveAllowFrom returns configured list", () => {
      const cfg = {
        channels: {
          a365: {
            allowFrom: ["user1", "user2"],
          },
        },
      } as unknown as OpenClawConfig;
      const allowFrom = a365Plugin.config.resolveAllowFrom!({ cfg });
      expect(allowFrom).toEqual(["user1", "user2"]);
    });
  });

  describe("security adapter", () => {
    it("returns pairing policy by default", () => {
      const cfg = {} as OpenClawConfig;
      const policy = a365Plugin.security!.resolveDmPolicy!({
        cfg,
        account: { accountId: "default", enabled: true, configured: false },
      });
      expect(policy?.policy).toBe("pairing");
    });

    it("returns configured dm policy", () => {
      const cfg = {
        channels: {
          a365: { dmPolicy: "open" },
        },
      } as unknown as OpenClawConfig;
      const policy = a365Plugin.security!.resolveDmPolicy!({
        cfg,
        account: { accountId: "default", enabled: true, configured: true },
      });
      expect(policy?.policy).toBe("open");
    });

    it("collects warnings for open group policy", () => {
      const cfg = {
        channels: {
          a365: { groupPolicy: "open" },
        },
      } as unknown as OpenClawConfig;
      const warnings = a365Plugin.security!.collectWarnings!({
        cfg,
        account: { accountId: "default", enabled: true, configured: true },
      });
      expect(warnings.length).toBeGreaterThan(0);
      expect(warnings[0]).toContain("groupPolicy");
    });
  });

  describe("messaging adapter", () => {
    it("normalizes conversation: prefix", () => {
      const result = a365Plugin.messaging!.normalizeTarget!("conversation:abc123");
      expect(result).toBe("abc123");
    });

    it("preserves user: prefix", () => {
      const result = a365Plugin.messaging!.normalizeTarget!("user:user123");
      expect(result).toBe("user:user123");
    });

    it("returns undefined for empty string", () => {
      const result = a365Plugin.messaging!.normalizeTarget!("");
      expect(result).toBeUndefined();
    });

    it("targetResolver identifies conversation IDs", () => {
      expect(a365Plugin.messaging!.targetResolver!.looksLikeId!("conversation:abc")).toBe(true);
      expect(a365Plugin.messaging!.targetResolver!.looksLikeId!("user:12345678-1234-1234-1234-123456789012")).toBe(true);
      expect(a365Plugin.messaging!.targetResolver!.looksLikeId!("abc@thread.tacv2")).toBe(true);
      expect(a365Plugin.messaging!.targetResolver!.looksLikeId!("random-text")).toBe(false);
    });
  });

  describe("agentTools", () => {
    it("returns empty array when Graph is not configured", () => {
      const tools = a365Plugin.agentTools!({ cfg: {} as OpenClawConfig });
      expect(tools).toEqual([]);
    });

    it("returns Graph tools when configured", () => {
      const cfg = {
        channels: {
          a365: {
            graph: {
              clientId: "client-id",
              clientSecret: "client-secret",
              tenantId: "tenant-id",
            },
          },
        },
      } as unknown as OpenClawConfig;
      const tools = a365Plugin.agentTools!({ cfg });
      expect(tools.length).toBeGreaterThan(0);
      expect(tools.map((t) => t.name)).toContain("get_calendar_events");
      expect(tools.map((t) => t.name)).toContain("create_calendar_event");
      expect(tools.map((t) => t.name)).toContain("send_email");
    });
  });

  describe("agentPrompt", () => {
    it("returns message tool hints", () => {
      const hints = a365Plugin.agentPrompt!.messageToolHints!({
        cfg: {} as OpenClawConfig,
      });
      expect(hints.length).toBeGreaterThan(0);
      expect(hints.some((h) => h.includes("A365"))).toBe(true);
    });

    it("includes calendar owner in hints when configured", () => {
      const cfg = {
        channels: {
          a365: {
            owner: "user@test.com",
          },
        },
      } as unknown as OpenClawConfig;
      const hints = a365Plugin.agentPrompt!.messageToolHints!({ cfg });
      expect(hints.some((h) => h.includes("user@test.com"))).toBe(true);
    });
  });

  describe("status adapter", () => {
    it("has default runtime status", () => {
      expect(a365Plugin.status!.defaultRuntime).toEqual({
        accountId: "default",
        running: false,
        lastStartAt: null,
        lastStopAt: null,
        lastError: null,
        port: null,
      });
    });

    it("builds account snapshot", async () => {
      const account = { accountId: "default", enabled: true, configured: true };
      const runtime = { running: true, port: 3978 };
      const snapshot = await a365Plugin.status!.buildAccountSnapshot!({
        account,
        cfg: {} as OpenClawConfig,
        runtime: runtime as any,
      });
      expect(snapshot.accountId).toBe("default");
      expect(snapshot.enabled).toBe(true);
      expect(snapshot.running).toBe(true);
      expect(snapshot.port).toBe(3978);
    });
  });
});
