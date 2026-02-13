import type { OpenClawConfig, RuntimeEnv } from "openclaw/plugin-sdk";
import type { A365Config, A365MessageMetadata } from "./types.js";
import { getA365Runtime } from "./runtime.js";
import { runWithGraphToolContext } from "./graph-tools.js";
import { resolveA365Credentials } from "./token.js";
import { saveConversationReference } from "./conversation-store.js";
import { setAdapter, setBlueprintClientId } from "./adapter-store.js";

export type MonitorA365Opts = {
  cfg: OpenClawConfig;
  runtime?: RuntimeEnv;
  abortSignal?: AbortSignal;
};

export type MonitorA365Result = {
  app: unknown;
  shutdown: () => Promise<void>;
};

/**
 * Activity shape for metadata extraction.
 */
export type ActivityForMetadata = {
  from?: { id: string; name?: string; aadObjectId?: string };
  recipient?: { id: string; name?: string };
  conversation?: { id: string; isGroup?: boolean; tenantId?: string };
  serviceUrl?: string;
  id?: string;
  channelId?: string;
  locale?: string;
  channelData?: {
    tenant?: { id: string };
    team?: { id: string; name?: string };
    channel?: { id: string; name?: string };
  };
};

/**
 * Extract message metadata from an Agents SDK activity.
 */
export function extractMessageMetadata(activity: ActivityForMetadata): A365MessageMetadata {
  return {
    userId: activity.from?.id || "",
    userEmail: activity.from?.aadObjectId || activity.from?.id,
    userName: activity.from?.name,
    userAadId: activity.from?.aadObjectId,
    conversationId: activity.conversation?.id || "",
    isGroup: activity.conversation?.isGroup || false,
    tenantId: activity.conversation?.tenantId || activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl || "",
    activityId: activity.id,
    channelId: activity.channelId,
    teamId: activity.channelData?.team?.id,
    teamName: activity.channelData?.team?.name,
    channelName: activity.channelData?.channel?.name,
  };
}

/**
 * Build a StoredConversationReference from an activity for proactive messaging.
 */
export function buildConversationReference(activity: ActivityForMetadata): StoredConversationReference {
  return {
    conversationId: activity.conversation?.id || "",
    serviceUrl: activity.serviceUrl || "",
    channelId: activity.channelId || "msteams",
    botId: activity.recipient?.id || "",
    botName: activity.recipient?.name,
    userId: activity.from?.id || "",
    userName: activity.from?.name,
    userAadId: activity.from?.aadObjectId,
    tenantId: activity.conversation?.tenantId || activity.channelData?.tenant?.id,
    isGroup: activity.conversation?.isGroup || false,
    locale: activity.locale,
    updatedAt: Date.now(),
  };
}

/**
 * Start the A365 Microsoft Agents provider.
 */
export async function monitorA365Provider(opts: MonitorA365Opts): Promise<MonitorA365Result> {
  const core = getA365Runtime();
  const log = core.logging.getChildLogger({ name: "a365" });
  const cfg = opts.cfg;
  const a365Cfg = cfg.channels?.a365 as A365Config | undefined;

  if (!a365Cfg?.enabled) {
    log.debug("a365 provider disabled");
    return { app: null, shutdown: async () => {} };
  }

  const runtime: RuntimeEnv = opts.runtime ?? {
    log: console.log,
    error: console.error,
    exit: (code: number): never => {
      throw new Error(`exit ${code}`);
    },
  };

  const port = a365Cfg.webhook?.port ?? 3978;

  log.info(`starting a365 provider (port ${port})`);

  // Set environment variables for the Agents SDK
  // The SDK reads these for authentication configuration
  // Use resolveA365Credentials to get values from config or A365_* env vars
  const creds = resolveA365Credentials(a365Cfg);
  if (!creds) {
    log.error("A365 credentials not configured - set appId/appPassword/tenantId in config or A365_APP_ID/A365_APP_PASSWORD/A365_TENANT_ID env vars");
    return { app: null, shutdown: async () => {} };
  }

  // TODO: The Microsoft Agents SDK currently requires configuration via environment variables.
  // This is not ideal as it mutates global state and credentials may be logged/exposed.
  // Consider contributing to the SDK to support programmatic configuration, or wrapping
  // the SDK initialization in an isolated subprocess.
  // See: https://github.com/microsoft/agents/issues (check for config API feature requests)

  // Set new-style connection config for Agents SDK 1.x
  process.env["connections__serviceConnection__settings__clientId"] = creds.appId;
  process.env["connections__serviceConnection__settings__clientSecret"] = creds.appPassword;
  process.env["connections__serviceConnection__settings__tenantId"] = creds.tenantId;
  process.env["connectionsMap__0__connection"] = "serviceConnection";
  process.env["connectionsMap__0__serviceUrl"] = "*";

  // Also set legacy env vars for backwards compatibility
  process.env.MicrosoftAppId = creds.appId;
  process.env.MicrosoftAppPassword = creds.appPassword;
  process.env.MicrosoftAppTenantId = creds.tenantId;
  process.env.MicrosoftAppType = "SingleTenant";
  process.env.PORT = String(port);

  // Dynamic imports for Microsoft Agents SDK
  const { AgentApplication, MemoryStorage, TurnContext, TurnState } = await import(
    "@microsoft/agents-hosting"
  );
  const { ActivityTypes } = await import("@microsoft/agents-activity");

  // Create custom turn state type
  type ApplicationTurnState = typeof TurnState;

  // Create the Agent Application
  // Note: We use our own T1/T2/User token flow for Graph API access,
  // not the SDK's built-in authorization
  const storage = new MemoryStorage();
  const agentApp = new AgentApplication<ApplicationTurnState>({
    storage,
  });

  // Note: We use our own T1/T2/User token flow for Graph API access via token.ts,
  // rather than storing the agentApp reference globally (which would be insecure).

  // Handle welcome message (configurable via welcomeMessage setting)
  agentApp.onConversationUpdate("membersAdded", async (context: typeof TurnContext) => {
    log.debug("members added event");
    const welcomeMessage = a365Cfg?.welcomeMessage;
    // Only send if welcomeMessage is configured and not empty
    if (welcomeMessage !== undefined && welcomeMessage !== "") {
      await context.sendActivity(welcomeMessage);
    }
    // If welcomeMessage is undefined, skip sending (silent join)
  });

  // Handle all messages
  agentApp.onActivity(
    ActivityTypes.Message,
    async (context: typeof TurnContext, _state: ApplicationTurnState) => {
      const activity = context.activity;
      const text = activity.text?.trim();

      if (!text) {
        log.debug("skipping empty message");
        return;
      }

      const metadata = extractMessageMetadata(activity);
      log.info("received message", {
        from: metadata.userName || metadata.userId,
        isGroup: metadata.isGroup,
        textLength: text.length,
      });

      // Store conversation reference from this AU-based request for proactive messaging.
      // The SDK's getConversationReference() preserves agentic identity metadata
      // (recipient.role, agenticAppId, agenticUserId, etc.) which the SDK needs
      // to acquire the correct AU token when sending proactively.
      try {
        const convRef = activity.getConversationReference();
        log.info("Saving conversation reference", {
          conversationId: convRef.conversation?.id,
          serviceUrl: convRef.serviceUrl,
          agentRole: convRef.agent?.role,
        });
        await saveConversationReference(convRef, metadata.userAadId);
        log.info("Conversation reference saved successfully");
      } catch (err) {
        log.error(`Failed to save conversation reference: ${String(err)}`)
      }

      // Check allowlist if configured
      const allowFrom = a365Cfg?.allowFrom;
      if (allowFrom && allowFrom.length > 0 && !allowFrom.includes("*")) {
        const userAllowed =
          allowFrom.includes(metadata.userId) ||
          allowFrom.includes(metadata.userEmail || "") ||
          allowFrom.includes(metadata.userAadId || "");
        if (!userAllowed) {
          log.debug("user not in allowlist", { userId: metadata.userId });
          return;
        }
      }

      // Determine user role based on owner config
      const isOwner =
        (a365Cfg?.owner &&
          metadata.userEmail?.toLowerCase() === a365Cfg.owner.toLowerCase()) ||
        (a365Cfg?.ownerAadId && metadata.userAadId === a365Cfg.ownerAadId);
      const userRole = isOwner ? "Owner" : "Requester";

      // Run the message handler within the Graph tool context for thread-safety
      // This ensures each request has isolated context for token acquisition
      await runWithGraphToolContext(
        {
          agentIdentity: a365Cfg?.agentIdentity || a365Cfg?.owner,
          currentUserEmail: metadata.userEmail,
          currentUserAadId: metadata.userAadId,
          currentUserRole: userRole,
        },
        async () => {
          // Resolve the agent route
          const senderId = metadata.userAadId || metadata.userId;
          const conversationId = metadata.conversationId;
          const isDirectMessage = !metadata.isGroup;

          const route = core.channel.routing.resolveAgentRoute({
            cfg,
            channel: "a365",
            peer: {
              kind: isDirectMessage ? "dm" : "group",
              id: isDirectMessage ? senderId : conversationId,
            },
          });

          // Build inbound context using the standard API
          const a365From = isDirectMessage
            ? `a365:${senderId}`
            : `a365:group:${conversationId}`;
          const a365To = isDirectMessage ? `user:${senderId}` : `conversation:${conversationId}`;

          const ctxPayload = core.channel.reply.finalizeInboundContext({
            Body: text,
            RawBody: text,
            CommandBody: text,
            From: a365From,
            To: a365To,
            SessionKey: route.sessionKey,
            AccountId: route.accountId,
            ChatType: isDirectMessage ? "direct" : "group",
            ConversationLabel: metadata.userName || senderId,
            SenderName: metadata.userName || senderId,
            SenderId: senderId,
            Provider: "a365" as const,
            Surface: "a365" as const,
            MessageSid: metadata.activityId,
            Timestamp: Date.now(),
            WasMentioned: true, // Assume mentioned for DMs
            CommandAuthorized: isOwner,
            OriginatingChannel: "a365" as const,
            OriginatingTo: conversationId,
          });

          // Create a simple reply dispatcher that tracks pending sends
          // The Agents SDK context is only valid during the handler, so we must await all sends
          let replyCount = 0;
          const pendingSends: Promise<void>[] = [];

          const sendReply = async (replyText: string) => {
            try {
              log.debug("sendReply called", { length: replyText.length });
              const result = await context.sendActivity(replyText);
              replyCount++;
              log.debug("reply sent successfully", { replyCount, resultId: result?.id });
            } catch (sendErr) {
              const err = sendErr as Error;
              log.error("sendActivity failed", { error: err?.message });
              // Don't rethrow - we'll handle errors at the end
            }
          };

          // Simple dispatcher that tracks queued replies
          const queuedCounts = { tool: 0, block: 0, final: 0 };
          const dispatcher = {
            sendToolResult: (payload: { text?: string }) => {
              if (payload.text) {
                queuedCounts.tool++;
                pendingSends.push(sendReply(payload.text));
              }
              return true;
            },
            sendBlockReply: (payload: { text?: string }) => {
              if (payload.text) {
                queuedCounts.block++;
                pendingSends.push(sendReply(payload.text));
              }
              return true;
            },
            sendFinalReply: (payload: { text?: string }) => {
              if (payload.text) {
                queuedCounts.final++;
                pendingSends.push(sendReply(payload.text));
              }
              return true;
            },
            waitForIdle: async () => {
              // Wait for all pending sends to complete
              await Promise.all(pendingSends);
            },
            getQueuedCounts: () => queuedCounts,
          };

          const replyOptions = {
            onReplyStart: async () => {
              try {
                log.debug("sending typing indicator");
                await context.sendActivity({ type: "typing" });
                log.debug("typing indicator sent");
              } catch (typingErr) {
                const err = typingErr as Error;
                log.debug("typing indicator failed", {
                  error: String(err),
                  message: err?.message,
                  stack: err?.stack,
                });
              }
            },
            onTypingController: () => {},
            onTypingCleanup: () => {},
          };

          try {
            log.info("dispatching to agent", { sessionKey: route.sessionKey });

            const { queuedFinal, counts } = await core.channel.reply.dispatchReplyFromConfig({
              ctx: ctxPayload,
              cfg,
              dispatcher,
              replyOptions,
            });

            // Wait for all pending sends to complete before handler returns
            // The Agents SDK context is only valid during the handler
            await Promise.all(pendingSends);

            log.info("dispatch complete", { queuedFinal, textCount: counts?.text ?? 0, repliesSent: replyCount });

            // Update the main session's lastChannel/lastTo for cron delivery support.
            try {
              const storePath = core.channel.session.resolveStorePath(cfg.session?.store);
              const mainSessionKey = "agent:main:main";
              await core.channel.session.updateLastRoute({
                storePath,
                sessionKey: mainSessionKey,
                channel: "a365",
                to: conversationId,
              });
              log.info("Updated main session for cron delivery", { conversationId });
            } catch (updateErr) {
              log.error(`Failed to update main session: ${String(updateErr)}`);
            }
          } catch (err) {
            log.error("handler failed", { error: String(err) });
            runtime.error?.(`a365 handler failed: ${String(err)}`);

            // Send error message back to user
            try {
              await context.sendActivity(
                "I encountered an error processing your message. Please try again.",
              );
            } catch {
              // Ignore send failure
            }
          }
        },
      );
    },
  );

  // Store the adapter for proactive messaging
  const { CloudAdapter } = await import("@microsoft/agents-hosting");
  const adapter = (agentApp.adapter ?? new CloudAdapter()) as InstanceType<typeof CloudAdapter>;
  setAdapter(adapter);

  // Store the Blueprint Client App ID — continueConversation must be called with
  // this ID (not the bot's own app ID) so the SDK routes through the correct
  // T1/T2/AU token flow for agentic identity.
  const blueprintClientId =
    a365Cfg?.graph?.blueprintClientAppId?.trim() ||
    process.env.BLUEPRINT_CLIENT_APP_ID?.trim() ||
    creds.appId; // fallback to bot app ID if blueprint not separately configured
  setBlueprintClientId(blueprintClientId);
  log.info("Stored adapter and blueprint client ID for proactive messaging", { blueprintClientId });

  // Start the server using the Agents SDK
  const { startServer } = await import("@microsoft/agents-hosting-express");

  // startServer returns a promise that resolves when server is ready
  // It uses PORT env var for the port
  const serverPromise = startServer(agentApp);

  log.info(`a365 provider started on port ${port}`);

  const shutdown = async () => {
    log.info("shutting down a365 provider");
    // TODO: Implement graceful shutdown:
    // - Track pending message handlers and wait for completion
    // - Close database connections if any
    // - Flush any queued outbound messages
    // Note: The Agents SDK doesn't expose a clean shutdown method currently.
  };

  // Handle abort signal
  if (opts.abortSignal) {
    opts.abortSignal.addEventListener("abort", () => {
      void shutdown();
    });
  }

  return { app: agentApp, shutdown };
}
