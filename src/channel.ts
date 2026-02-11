import type { ChannelPlugin, OpenClawConfig } from "openclaw/plugin-sdk";
import { DEFAULT_ACCOUNT_ID } from "openclaw/plugin-sdk";
import type { A365Config, A365Probe, ResolvedA365Account } from "./types.js";
import { resolveA365Credentials, resolveGraphTokenConfig, resolveTokenCallbackConfig, getGraphToken } from "./token.js";
import { createGraphTools } from "./graph-tools.js";
import { a365Outbound, normalizeA365MessagingTarget } from "./outbound.js";

/**
 * Check if Graph API tools can be enabled.
 * Requires either T1/T2 flow config or external token callback.
 */
function isGraphConfigured(cfg?: A365Config): boolean {
  return Boolean(resolveGraphTokenConfig(cfg) || resolveTokenCallbackConfig(cfg));
}

const meta = {
  id: "a365",
  label: "Microsoft 365 Agents",
  selectionLabel: "Microsoft 365 Agents (A365)",
  docsPath: "/channels/a365",
  docsLabel: "a365",
  blurb: "Native A365 channel with Graph API tools for calendar and email.",
  aliases: ["m365agents", "agents365"],
  order: 55,
} as const;

/**
 * Probe A365 configuration to check if it's working.
 */
async function probeA365(cfg?: A365Config): Promise<A365Probe> {
  const creds = resolveA365Credentials(cfg);
  if (!creds) {
    return { ok: false, error: "Bot Framework credentials not configured" };
  }

  // Check if Graph API is configured
  const graphConfigured = isGraphConfigured(cfg);
  let graphConnected = false;

  if (graphConfigured && cfg?.agentIdentity) {
    try {
      const token = await getGraphToken(cfg, cfg.agentIdentity);
      graphConnected = Boolean(token);
    } catch {
      graphConnected = false;
    }
  }

  return {
    ok: true,
    botId: creds.appId,
    graphConnected,
    owner: cfg?.owner,
  };
}

/**
 * A365 Channel Plugin for OpenClaw.
 *
 * This channel enables:
 * - Receiving messages from Microsoft 365 Agents via Bot Framework
 * - Sending responses back through Bot Framework
 * - Native Graph API tools for calendar, email, and user operations
 */
export const a365Plugin: ChannelPlugin<ResolvedA365Account, A365Probe> = {
  id: "a365",
  meta: {
    ...meta,
  },
  pairing: {
    idLabel: "a365UserId",
    normalizeAllowEntry: (entry) => entry.replace(/^(a365|user):/i, ""),
  },
  capabilities: {
    chatTypes: ["direct", "channel", "thread"],
    threads: true,
    media: true,
  },
  agentPrompt: {
    messageToolHints: ({ cfg }) => {
      const a365Cfg = cfg?.channels?.a365 as A365Config | undefined;
      const timezone = a365Cfg?.businessHours?.timezone || "America/Los_Angeles";
      const now = new Date();
      const formatter = new Intl.DateTimeFormat("en-US", {
        timeZone: timezone,
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      });
      const currentDateTime = formatter.format(now);
      const dateOnly = new Date().toLocaleDateString("en-CA", { timeZone: timezone }); // YYYY-MM-DD format

      const hints = [
        "- A365 channel supports direct messages and channel conversations.",
        "- Use Graph API tools (get_calendar_events, create_calendar_event, etc.) for calendar operations.",
        `- Current date/time: ${currentDateTime} (${timezone}). Today's date in ISO format: ${dateOnly}.`,
      ];
      if (a365Cfg?.owner) {
        hints.push(`- Default calendar owner: ${a365Cfg.owner}`);
      }
      return hints;
    },
  },
  threading: {
    buildToolContext: ({ context, hasRepliedRef }) => ({
      currentChannelId: context.To?.trim() || undefined,
      currentThreadTs: context.ReplyToId,
      hasRepliedRef,
    }),
  },
  reload: { configPrefixes: ["channels.a365"] },
  config: {
    listAccountIds: () => [DEFAULT_ACCOUNT_ID],
    resolveAccount: (cfg) => {
      const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
      return {
        accountId: DEFAULT_ACCOUNT_ID,
        enabled: a365Cfg?.enabled !== false,
        configured: Boolean(resolveA365Credentials(a365Cfg)),
        owner: a365Cfg?.owner,
      };
    },
    defaultAccountId: () => DEFAULT_ACCOUNT_ID,
    setAccountEnabled: ({ cfg, enabled }) => ({
      ...cfg,
      channels: {
        ...cfg.channels,
        a365: {
          ...(cfg.channels?.a365 as A365Config),
          enabled,
        },
      },
    }),
    deleteAccount: ({ cfg }) => {
      const next = { ...cfg } as OpenClawConfig;
      const nextChannels = { ...cfg.channels };
      delete nextChannels.a365;
      if (Object.keys(nextChannels).length > 0) {
        next.channels = nextChannels;
      } else {
        delete next.channels;
      }
      return next;
    },
    isConfigured: (_account, cfg) => {
      const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
      return Boolean(resolveA365Credentials(a365Cfg));
    },
    describeAccount: (account) => ({
      accountId: account.accountId,
      enabled: account.enabled,
      configured: account.configured,
    }),
    resolveAllowFrom: ({ cfg }) => {
      const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
      return a365Cfg?.allowFrom?.map(String) ?? [];
    },
    formatAllowFrom: ({ allowFrom }) =>
      allowFrom
        .map((entry) => String(entry).trim())
        .filter(Boolean)
        .map((entry) => entry.toLowerCase()),
  },
  security: {
    resolveDmPolicy: ({ cfg }) => {
      const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
      return {
        policy: a365Cfg?.dmPolicy ?? "pairing",
        allowFrom: a365Cfg?.allowFrom ?? [],
        allowFromPath: "channels.a365.",
        approveHint: "Add user ID to channels.a365.allowFrom",
      };
    },
    collectWarnings: ({ cfg }) => {
      const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
      const groupPolicy = a365Cfg?.groupPolicy ?? "allowlist";
      if (groupPolicy !== "open") {
        return [];
      }
      return [
        `- A365 groups: groupPolicy="open" allows any member to trigger. Set channels.a365.groupPolicy="allowlist" + channels.a365.groupAllowFrom to restrict senders.`,
      ];
    },
  },
  setup: {
    resolveAccountId: () => DEFAULT_ACCOUNT_ID,
    applyAccountConfig: ({ cfg }) => ({
      ...cfg,
      channels: {
        ...cfg.channels,
        a365: {
          ...(cfg.channels?.a365 as A365Config),
          enabled: true,
        },
      },
    }),
  },
  messaging: {
    normalizeTarget: normalizeA365MessagingTarget,
    targetResolver: {
      looksLikeId: (raw) => {
        const trimmed = raw.trim();
        if (!trimmed) {
          return false;
        }
        if (/^conversation:/i.test(trimmed)) {
          return true;
        }
        if (/^user:/i.test(trimmed)) {
          const id = trimmed.slice("user:".length).trim();
          return /^[0-9a-fA-F-]{16,}$/.test(id);
        }
        return trimmed.includes("@thread") || trimmed.includes(":");
      },
      hint: "<conversationId|user:ID|conversation:ID>",
    },
  },
  directory: {
    self: async () => null,
    listPeers: async ({ cfg, query, limit }) => {
      const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
      const q = query?.trim().toLowerCase() || "";
      const ids = new Set<string>();
      for (const entry of a365Cfg?.allowFrom ?? []) {
        const trimmed = String(entry).trim();
        if (trimmed && trimmed !== "*") {
          ids.add(trimmed);
        }
      }
      return Array.from(ids)
        .filter((id) => (q ? id.toLowerCase().includes(q) : true))
        .slice(0, limit && limit > 0 ? limit : undefined)
        .map((id) => ({ kind: "user", id }) as const);
    },
    listGroups: async () => [],
  },
  // Register Graph API tools for agent use
  agentTools: ({ cfg }) => {
    const a365Cfg = cfg?.channels?.a365 as A365Config | undefined;
    // Only provide tools if Graph API is configured (T1/T2 flow or callback)
    if (!isGraphConfigured(a365Cfg)) {
      return [];
    }
    return createGraphTools(a365Cfg);
  },
  outbound: a365Outbound,
  status: {
    defaultRuntime: {
      accountId: DEFAULT_ACCOUNT_ID,
      running: false,
      lastStartAt: null,
      lastStopAt: null,
      lastError: null,
      port: null,
    },
    buildChannelSummary: ({ snapshot }) => ({
      configured: snapshot.configured ?? false,
      running: snapshot.running ?? false,
      lastStartAt: snapshot.lastStartAt ?? null,
      lastStopAt: snapshot.lastStopAt ?? null,
      lastError: snapshot.lastError ?? null,
      port: snapshot.port ?? null,
      probe: snapshot.probe,
      lastProbeAt: snapshot.lastProbeAt ?? null,
    }),
    probeAccount: async ({ cfg }) => {
      const a365Cfg = cfg.channels?.a365 as A365Config | undefined;
      return probeA365(a365Cfg);
    },
    buildAccountSnapshot: ({ account, runtime, probe }) => ({
      accountId: account.accountId,
      enabled: account.enabled,
      configured: account.configured,
      running: runtime?.running ?? false,
      lastStartAt: runtime?.lastStartAt ?? null,
      lastStopAt: runtime?.lastStopAt ?? null,
      lastError: runtime?.lastError ?? null,
      port: runtime?.port ?? null,
      probe,
    }),
  },
  gateway: {
    startAccount: async (ctx) => {
      const { monitorA365Provider } = await import("./monitor.js");
      const a365Cfg = ctx.cfg.channels?.a365 as A365Config | undefined;
      const port = a365Cfg?.webhook?.port ?? 3978;
      ctx.setStatus({ accountId: ctx.accountId, port });
      ctx.log?.info(`starting a365 provider (port ${port})`);
      return monitorA365Provider({
        cfg: ctx.cfg,
        runtime: ctx.runtime,
        abortSignal: ctx.abortSignal,
      });
    },
  },
};
