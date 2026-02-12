/**
 * A365 channel configuration schema.
 */
export type A365Config = {
  enabled?: boolean;
  /** Bot Framework App ID */
  appId?: string;
  /** Bot Framework App Password */
  appPassword?: string;
  /** Azure AD Tenant ID */
  tenantId?: string;
  /** Webhook configuration */
  webhook?: {
    /** Port for the Bot Framework webhook server (default: 3978) */
    port?: number;
  };
  /** Graph API configuration for T1/T2/User flow (Federated Identity Credentials) */
  graph?: {
    /** Blueprint Client App ID (typically same as Bot App ID) */
    blueprintClientAppId?: string;
    /** Blueprint Client Secret (typically same as Bot App Password) */
    blueprintClientSecret?: string;
    /** AA Instance ID (Agent Application Instance ID) - required for user_fic flow */
    aaInstanceId?: string;
    /** OAuth scope for Graph API */
    scope?: string;
    /** @deprecated Use blueprintClientAppId instead */
    clientId?: string;
    /** @deprecated Use blueprintClientSecret instead */
    clientSecret?: string;
    /** @deprecated Use main tenantId instead */
    tenantId?: string;
  };
  /** External token callback service (alternative to T1/T2 flow) */
  tokenCallback?: {
    /** URL of the token service endpoint */
    url?: string;
    /** Bearer token for authenticating with the token service */
    token?: string;
  };
  /** Email of the person this agent supports (the "principal") */
  owner?: string;
  /** AAD Object ID of the owner (for role detection) */
  ownerAadId?: string;
  /** Agent identity - the service account email used for Graph API calls */
  agentIdentity?: string;
  /** Business hours configuration */
  businessHours?: {
    start?: string; // e.g., "08:00"
    end?: string;   // e.g., "18:00"
    timezone?: string; // e.g., "America/Los_Angeles"
  };
  /** Welcome message to send when a new conversation starts. Set to empty string to disable. */
  welcomeMessage?: string;
  /** DM policy: "open" | "pairing" | "allowlist" */
  dmPolicy?: string;
  /** Allowed users list */
  allowFrom?: Array<string | number>;
  /** Group allowlist */
  groupAllowFrom?: Array<string | number>;
  /** Group policy */
  groupPolicy?: string;
};

/**
 * Resolved A365 account information.
 */
export type ResolvedA365Account = {
  accountId: string;
  enabled: boolean;
  configured: boolean;
  owner?: string;
};

/**
 * A365 probe result for status checks.
 */
export type A365Probe = {
  ok: boolean;
  botId?: string;
  botName?: string;
  graphConnected?: boolean;
  owner?: string;
  error?: string;
};

/**
 * Activity from Bot Framework.
 */
export type BotActivity = {
  type: string;
  id?: string;
  timestamp?: string;
  localTimestamp?: string;
  serviceUrl?: string;
  channelId?: string;
  from?: {
    id: string;
    name?: string;
    aadObjectId?: string;
  };
  conversation?: {
    id: string;
    isGroup?: boolean;
    conversationType?: string;
    tenantId?: string;
    name?: string;
  };
  recipient?: {
    id: string;
    name?: string;
  };
  text?: string;
  textFormat?: string;
  attachments?: Array<{
    contentType: string;
    contentUrl?: string;
    content?: unknown;
    name?: string;
  }>;
  entities?: Array<{
    type: string;
    [key: string]: unknown;
  }>;
  channelData?: {
    tenant?: { id: string };
    team?: { id: string; name?: string };
    channel?: { id: string; name?: string };
    [key: string]: unknown;
  };
  value?: unknown;
  name?: string;
  locale?: string;
};

/**
 * Turn context for handling activities.
 */
export type A365TurnContext = {
  activity: BotActivity;
  sendActivity: (activity: string | Partial<BotActivity>) => Promise<{ id?: string }>;
  updateActivity?: (activity: Partial<BotActivity>) => Promise<void>;
  deleteActivity?: (activityId: string) => Promise<void>;
};

/**
 * Message metadata extracted from Bot Framework activity.
 */
export type A365MessageMetadata = {
  userId: string;
  userEmail?: string;
  userName?: string;
  userAadId?: string;
  conversationId: string;
  isGroup: boolean;
  tenantId?: string;
  serviceUrl: string;
  activityId?: string;
  channelId?: string;
  teamId?: string;
  teamName?: string;
  channelName?: string;
};

/**
 * Stored conversation reference for proactive messaging.
 * Contains all the information needed to send a message back to a conversation
 * after the original request has completed (e.g., for cron jobs, async tasks).
 */
export type StoredConversationReference = {
  /** The conversation ID */
  conversationId: string;
  /** Bot Framework service URL - required for sending proactive messages */
  serviceUrl: string;
  /** Channel ID (e.g., "msteams", "emulator") */
  channelId: string;
  /** The bot's ID in this channel */
  botId: string;
  /** The bot's display name */
  botName?: string;
  /** The user's ID in this channel */
  userId: string;
  /** The user's display name */
  userName?: string;
  /** The user's AAD Object ID */
  userAadId?: string;
  /** Azure AD tenant ID */
  tenantId?: string;
  /** Whether this is a group conversation */
  isGroup: boolean;
  /** User's locale */
  locale?: string;
  /** Timestamp when this reference was last updated */
  updatedAt: number;
};

/**
 * Graph API calendar event.
 */
export type GraphCalendarEvent = {
  id?: string;
  subject: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  attendees?: Array<{
    emailAddress: {
      address: string;
      name?: string;
    };
    type?: string;
    status?: {
      response?: string;
      time?: string;
    };
  }>;
  location?: {
    displayName?: string;
  };
  body?: {
    contentType?: string;
    content?: string;
  };
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
  organizer?: {
    emailAddress: {
      address: string;
      name?: string;
    };
  };
  showAs?: string;
  importance?: string;
  sensitivity?: string;
  isAllDay?: boolean;
  isCancelled?: boolean;
  responseRequested?: boolean;
};
