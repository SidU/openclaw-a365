import { AsyncLocalStorage } from "node:async_hooks";
import { Type, type TSchema } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import type { A365Config, GraphCalendarEvent } from "./types.js";
import { getGraphToken } from "./token.js";
import { getA365Runtime } from "./runtime.js";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const DEFAULT_TIMEZONE = "UTC";

/**
 * Get the default timezone from config, falling back to UTC.
 */
function getDefaultTimezone(cfg?: A365Config): string {
  return cfg?.businessHours?.timezone || DEFAULT_TIMEZONE;
}

/**
 * Get the logger for Graph API operations.
 * Returns a no-op logger if runtime is not yet initialized.
 */
function getLogger() {
  try {
    return getA365Runtime().logging.getChildLogger({ name: "a365-graph" });
  } catch {
    return {
      debug: () => {},
      info: () => {},
      warn: () => {},
      error: () => {},
    };
  }
}

/**
 * Context for Graph API tool execution.
 * This provides user information for token acquisition.
 */
export type GraphToolContext = {
  /** Username (email) of the agent service account */
  agentIdentity?: string;
  /** Username (email) of the current user interacting with the agent */
  currentUserEmail?: string;
  /** AAD Object ID of the current user */
  currentUserAadId?: string;
  /** Role of the current user */
  currentUserRole?: "Owner" | "Requester";
};

/**
 * AsyncLocalStorage for thread-safe tool context.
 * This ensures each request has its own isolated context,
 * preventing cross-request data leakage in concurrent scenarios.
 */
const toolContextStorage = new AsyncLocalStorage<GraphToolContext>();

/**
 * Run a function with the given tool context.
 * Use this to wrap request handlers to ensure proper context isolation.
 */
export function runWithGraphToolContext<T>(ctx: GraphToolContext, fn: () => T): T {
  return toolContextStorage.run(ctx, fn);
}

/**
 * Get the current tool context from AsyncLocalStorage.
 */
export function getGraphToolContext(): GraphToolContext | undefined {
  return toolContextStorage.getStore();
}

/**
 * @deprecated Use runWithGraphToolContext instead for thread-safe context management.
 * This function is kept for backwards compatibility but will set context globally
 * which is not safe in concurrent scenarios.
 */
export function setGraphToolContext(_ctx: GraphToolContext | undefined): void {
  // No-op: Context should be set via runWithGraphToolContext
  // This is kept for API compatibility but logs a warning
  getLogger().warn("setGraphToolContext is deprecated - use runWithGraphToolContext for thread-safe context");
}

/**
 * Make a request to the Microsoft Graph API.
 * Uses the agent username (service account) for token acquisition.
 *
 * TODO: Add retry logic with exponential backoff for transient failures (429, 503).
 */
async function graphRequest<T>(
  cfg: A365Config | undefined,
  method: string,
  path: string,
  body?: unknown,
): Promise<{ ok: true; data: T } | { ok: false; error: string; status?: number }> {
  const log = getLogger();

  // Get the username for token acquisition
  // Use agent username from context (thread-safe) or config
  const toolContext = getGraphToolContext();
  const agentIdentity =
    toolContext?.agentIdentity ||
    cfg?.agentIdentity ||
    cfg?.owner;

  if (!agentIdentity) {
    return { ok: false, error: "Agent username not configured. Set agentIdentity or owner in config." };
  }

  const token = await getGraphToken(cfg, agentIdentity);
  if (!token) {
    return { ok: false, error: "Graph API token not available. Check T1/T2/User flow configuration (blueprintClientAppId, blueprintClientSecret, aaInstanceId)." };
  }

  const url = `${GRAPH_BASE_URL}${path}`;
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  log.debug("Graph API request", { method, path, agentIdentity });

  try {
    const response = await fetch(url, {
      method,
      headers,
      body: body ? JSON.stringify(body) : undefined,
    });

    log.debug("Graph API response", { status: response.status });

    if (!response.ok) {
      const errorText = await response.text();
      log.warn("Graph API error", { status: response.status, error: errorText.slice(0, 200) });
      let errorMessage = `Graph API error: ${response.status}`;
      try {
        const errorJson = JSON.parse(errorText);
        errorMessage = errorJson.error?.message || errorMessage;
      } catch {
        errorMessage = errorText || errorMessage;
      }
      return { ok: false, error: errorMessage, status: response.status };
    }

    // Handle 204 No Content
    if (response.status === 204) {
      return { ok: true, data: {} as T };
    }

    const data = await response.json() as T;
    return { ok: true, data };
  } catch (err) {
    log.error("Graph API network error", { error: String(err) });
    return { ok: false, error: `Network error: ${String(err)}` };
  }
}

/**
 * Validate common tool parameters.
 */
function validateUserId(userId: string): { ok: true } | { ok: false; error: string } {
  if (!userId || !userId.trim()) {
    return { ok: false, error: "userId is required and cannot be empty" };
  }
  // Basic email format check (loose validation - Graph API will reject invalid IDs)
  if (!userId.includes("@") && !/^[0-9a-f-]{36}$/i.test(userId)) {
    return { ok: false, error: "userId should be an email address or a valid GUID" };
  }
  return { ok: true };
}

/**
 * Validate ISO datetime string format.
 */
function validateIsoDateTime(dateTime: string, fieldName: string): { ok: true } | { ok: false; error: string } {
  if (!dateTime || !dateTime.trim()) {
    return { ok: false, error: `${fieldName} is required` };
  }
  // Basic ISO format check (YYYY-MM-DDTHH:MM:SS)
  if (!/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2})?)?/.test(dateTime)) {
    return { ok: false, error: `${fieldName} should be in ISO format (e.g., 2024-01-15T14:00:00)` };
  }
  return { ok: true };
}

/**
 * Validate email addresses in an array.
 */
function validateEmails(emails: string[], fieldName: string): { ok: true } | { ok: false; error: string } {
  for (const email of emails) {
    if (!email.includes("@")) {
      return { ok: false, error: `Invalid email address in ${fieldName}: ${email}` };
    }
  }
  return { ok: true };
}

type ToolResult = AgentToolResult<unknown>;

/**
 * Get calendar events for a user within a date range.
 */
async function getCalendarEvents(
  cfg: A365Config | undefined,
  params: { userId: string; startDate: string; endDate: string },
): Promise<ToolResult> {
  const { userId, startDate, endDate } = params;

  // Validate inputs
  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  const startCheck = validateIsoDateTime(startDate, "startDate");
  if (!startCheck.ok) return { isError: true, content: [{ type: "text", text: startCheck.error }] };

  const endCheck = validateIsoDateTime(endDate, "endDate");
  if (!endCheck.ok) return { isError: true, content: [{ type: "text", text: endCheck.error }] };

  const path = `/users/${encodeURIComponent(userId)}/calendar/calendarView?startDateTime=${encodeURIComponent(startDate)}&endDateTime=${encodeURIComponent(endDate)}&$orderby=start/dateTime&$top=50`;

  const result = await graphRequest<{ value: GraphCalendarEvent[] }>(cfg, "GET", path);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  const events = result.data.value.map((event) => ({
    id: event.id,
    subject: event.subject,
    start: event.start,
    end: event.end,
    location: event.location?.displayName,
    attendees: event.attendees?.map((a) => ({
      email: a.emailAddress.address,
      name: a.emailAddress.name,
      response: a.status?.response,
    })),
    isOnlineMeeting: event.isOnlineMeeting,
    onlineMeetingUrl: event.onlineMeetingUrl,
    showAs: event.showAs,
    isCancelled: event.isCancelled,
  }));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ events, count: events.length }, null, 2),
      },
    ],
  };
}

/**
 * Create a calendar event for a user.
 */
async function createCalendarEvent(
  cfg: A365Config | undefined,
  params: {
    userId: string;
    subject: string;
    startDateTime: string;
    endDateTime: string;
    timeZone?: string;
    attendees?: string[];
    location?: string;
    body?: string;
    isOnlineMeeting?: boolean;
  },
): Promise<ToolResult> {
  const defaultTz = getDefaultTimezone(cfg);
  const {
    userId,
    subject,
    startDateTime,
    endDateTime,
    timeZone = defaultTz,
    attendees = [],
    location,
    body,
    isOnlineMeeting = false,
  } = params;

  // Validate inputs
  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  const startCheck = validateIsoDateTime(startDateTime, "startDateTime");
  if (!startCheck.ok) return { isError: true, content: [{ type: "text", text: startCheck.error }] };

  const endCheck = validateIsoDateTime(endDateTime, "endDateTime");
  if (!endCheck.ok) return { isError: true, content: [{ type: "text", text: endCheck.error }] };

  if (attendees.length > 0) {
    const emailsCheck = validateEmails(attendees, "attendees");
    if (!emailsCheck.ok) return { isError: true, content: [{ type: "text", text: emailsCheck.error }] };
  }

  const path = `/users/${encodeURIComponent(userId)}/calendar/events`;

  const eventBody: Partial<GraphCalendarEvent> = {
    subject,
    start: { dateTime: startDateTime, timeZone },
    end: { dateTime: endDateTime, timeZone },
    isOnlineMeeting,
  };

  if (attendees.length > 0) {
    eventBody.attendees = attendees.map((email) => ({
      emailAddress: { address: email },
      type: "required",
    }));
  }

  if (location) {
    eventBody.location = { displayName: location };
  }

  if (body) {
    eventBody.body = { contentType: "text", content: body };
  }

  const result = await graphRequest<GraphCalendarEvent>(cfg, "POST", path, eventBody);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            success: true,
            eventId: result.data.id,
            subject: result.data.subject,
            start: result.data.start,
            end: result.data.end,
            onlineMeetingUrl: result.data.onlineMeetingUrl,
          },
          null,
          2,
        ),
      },
    ],
  };
}

/**
 * Update an existing calendar event.
 */
async function updateCalendarEvent(
  cfg: A365Config | undefined,
  params: {
    userId: string;
    eventId: string;
    subject?: string;
    startDateTime?: string;
    endDateTime?: string;
    timeZone?: string;
    location?: string;
    body?: string;
  },
): Promise<ToolResult> {
  const defaultTz = getDefaultTimezone(cfg);
  const { userId, eventId, subject, startDateTime, endDateTime, timeZone, location, body } = params;

  const path = `/users/${encodeURIComponent(userId)}/calendar/events/${encodeURIComponent(eventId)}`;

  const eventBody: Partial<GraphCalendarEvent> = {};

  if (subject !== undefined) {
    eventBody.subject = subject;
  }
  if (startDateTime !== undefined) {
    eventBody.start = { dateTime: startDateTime, timeZone: timeZone || defaultTz };
  }
  if (endDateTime !== undefined) {
    eventBody.end = { dateTime: endDateTime, timeZone: timeZone || defaultTz };
  }
  if (location !== undefined) {
    eventBody.location = { displayName: location };
  }
  if (body !== undefined) {
    eventBody.body = { contentType: "text", content: body };
  }

  const result = await graphRequest<GraphCalendarEvent>(cfg, "PATCH", path, eventBody);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            success: true,
            eventId: result.data.id,
            subject: result.data.subject,
            start: result.data.start,
            end: result.data.end,
          },
          null,
          2,
        ),
      },
    ],
  };
}

/**
 * Delete a calendar event.
 */
async function deleteCalendarEvent(
  cfg: A365Config | undefined,
  params: { userId: string; eventId: string },
): Promise<ToolResult> {
  const { userId, eventId } = params;

  const path = `/users/${encodeURIComponent(userId)}/calendar/events/${encodeURIComponent(eventId)}`;

  const result = await graphRequest<unknown>(cfg, "DELETE", path);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({ success: true, deleted: eventId }, null, 2),
      },
    ],
  };
}

/**
 * Send an email using Microsoft Graph.
 */
async function sendEmail(
  cfg: A365Config | undefined,
  params: {
    userId: string;
    to: string[];
    subject: string;
    body: string;
    cc?: string[];
    bcc?: string[];
    importance?: "low" | "normal" | "high";
  },
): Promise<ToolResult> {
  const { userId, to, subject, body, cc = [], bcc = [], importance = "normal" } = params;

  // Validate inputs
  const userIdCheck = validateUserId(userId);
  if (!userIdCheck.ok) return { isError: true, content: [{ type: "text", text: userIdCheck.error }] };

  if (to.length === 0) {
    return { isError: true, content: [{ type: "text", text: "At least one recipient is required in 'to' field" }] };
  }

  const toCheck = validateEmails(to, "to");
  if (!toCheck.ok) return { isError: true, content: [{ type: "text", text: toCheck.error }] };

  if (cc.length > 0) {
    const ccCheck = validateEmails(cc, "cc");
    if (!ccCheck.ok) return { isError: true, content: [{ type: "text", text: ccCheck.error }] };
  }

  if (bcc.length > 0) {
    const bccCheck = validateEmails(bcc, "bcc");
    if (!bccCheck.ok) return { isError: true, content: [{ type: "text", text: bccCheck.error }] };
  }

  const path = `/users/${encodeURIComponent(userId)}/sendMail`;

  const mailBody = {
    message: {
      subject,
      body: {
        contentType: "text",
        content: body,
      },
      toRecipients: to.map((email) => ({ emailAddress: { address: email } })),
      ccRecipients: cc.map((email) => ({ emailAddress: { address: email } })),
      bccRecipients: bcc.map((email) => ({ emailAddress: { address: email } })),
      importance,
    },
    saveToSentItems: true,
  };

  const result = await graphRequest<unknown>(cfg, "POST", path, mailBody);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            success: true,
            message: `Email sent successfully to ${to.join(", ")}`,
          },
          null,
          2,
        ),
      },
    ],
  };
}

/**
 * Get user information from Microsoft Graph.
 */
async function getUserInfo(
  cfg: A365Config | undefined,
  params: { userId: string },
): Promise<ToolResult> {
  const { userId } = params;

  const path = `/users/${encodeURIComponent(userId)}?$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation`;

  const result = await graphRequest<{
    id: string;
    displayName: string;
    mail: string;
    userPrincipalName: string;
    jobTitle?: string;
    department?: string;
    officeLocation?: string;
  }>(cfg, "GET", path);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(result.data, null, 2),
      },
    ],
  };
}

/**
 * Find available meeting times using findMeetingTimes API.
 */
async function findMeetingTimes(
  cfg: A365Config | undefined,
  params: {
    userId: string;
    attendees: string[];
    durationMinutes: number;
    startDateTime: string;
    endDateTime: string;
    timeZone?: string;
  },
): Promise<ToolResult> {
  const defaultTz = getDefaultTimezone(cfg);
  const {
    userId,
    attendees,
    durationMinutes,
    startDateTime,
    endDateTime,
    timeZone = defaultTz,
  } = params;

  const path = `/users/${encodeURIComponent(userId)}/findMeetingTimes`;

  const body = {
    attendees: attendees.map((email) => ({
      emailAddress: { address: email },
      type: "required",
    })),
    timeConstraint: {
      activityDomain: "work",
      timeSlots: [
        {
          start: { dateTime: startDateTime, timeZone },
          end: { dateTime: endDateTime, timeZone },
        },
      ],
    },
    meetingDuration: `PT${durationMinutes}M`,
    maxCandidates: 5,
    isOrganizerOptional: false,
    returnSuggestionReasons: true,
  };

  const result = await graphRequest<{
    meetingTimeSuggestions: Array<{
      meetingTimeSlot: {
        start: { dateTime: string; timeZone: string };
        end: { dateTime: string; timeZone: string };
      };
      confidence: number;
      organizerAvailability: string;
      attendeeAvailability: Array<{
        attendee: { emailAddress: { address: string } };
        availability: string;
      }>;
      suggestionReason?: string;
    }>;
    emptySuggestionsReason?: string;
  }>(cfg, "POST", path, body);

  if (!result.ok) {
    return {
      isError: true,
      content: [{ type: "text", text: result.error }],
    };
  }

  const suggestions = result.data.meetingTimeSuggestions.map((s) => ({
    start: s.meetingTimeSlot.start,
    end: s.meetingTimeSlot.end,
    confidence: s.confidence,
    organizerAvailability: s.organizerAvailability,
    attendeeAvailability: s.attendeeAvailability?.map((a) => ({
      email: a.attendee.emailAddress.address,
      availability: a.availability,
    })),
    reason: s.suggestionReason,
  }));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            suggestions,
            count: suggestions.length,
            emptySuggestionsReason: result.data.emptySuggestionsReason,
          },
          null,
          2,
        ),
      },
    ],
  };
}

/**
 * Create the Graph API tools for the A365 channel.
 *
 * Note: The execute functions use type assertions (e.g., `params as Parameters<...>`)
 * because TypeBox validates parameters at runtime against the schema before execution.
 * This provides runtime type safety while keeping the tool definitions concise.
 */
export function createGraphTools(cfg?: A365Config): AgentTool<TSchema, unknown>[] {
  const owner = cfg?.owner;

  return [
    {
      name: "get_calendar_events",
      label: "Get Calendar Events",
      description: `Get calendar events for a user within a date range. ${owner ? `Default calendar owner: ${owner}` : "Requires userId parameter."}`,
      parameters: Type.Object({
        userId: Type.String({
          description: "User email or ID (use calendar owner's email for scheduling)",
        }),
        startDate: Type.String({
          description: "Start date/time in ISO format (e.g., 2024-01-15T00:00:00)",
        }),
        endDate: Type.String({
          description: "End date/time in ISO format (e.g., 2024-01-15T23:59:59)",
        }),
      }),
      execute: async (_toolCallId, params) => getCalendarEvents(cfg, params as Parameters<typeof getCalendarEvents>[1]),
    },
    {
      name: "create_calendar_event",
      label: "Create Calendar Event",
      description: `Create a new calendar event. ${owner ? `Default calendar owner: ${owner}` : "Requires userId parameter."}`,
      parameters: Type.Object({
        userId: Type.String({
          description: "User email or ID whose calendar to create event on",
        }),
        subject: Type.String({ description: "Event subject/title" }),
        startDateTime: Type.String({
          description: "Start date/time in ISO format (e.g., 2024-01-15T14:00:00)",
        }),
        endDateTime: Type.String({
          description: "End date/time in ISO format (e.g., 2024-01-15T15:00:00)",
        }),
        timeZone: Type.Optional(
          Type.String({ description: "Time zone (default: from config or UTC)" }),
        ),
        attendees: Type.Optional(
          Type.Array(Type.String(), { description: "List of attendee email addresses" }),
        ),
        location: Type.Optional(Type.String({ description: "Meeting location" })),
        body: Type.Optional(Type.String({ description: "Event body/description" })),
        isOnlineMeeting: Type.Optional(
          Type.Boolean({ description: "Create as Teams meeting (default: false)" }),
        ),
      }),
      execute: async (_toolCallId, params) =>
        createCalendarEvent(cfg, params as Parameters<typeof createCalendarEvent>[1]),
    },
    {
      name: "update_calendar_event",
      label: "Update Calendar Event",
      description: "Update an existing calendar event.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose calendar contains the event" }),
        eventId: Type.String({ description: "ID of the event to update" }),
        subject: Type.Optional(Type.String({ description: "New event subject/title" })),
        startDateTime: Type.Optional(
          Type.String({ description: "New start date/time in ISO format" }),
        ),
        endDateTime: Type.Optional(Type.String({ description: "New end date/time in ISO format" })),
        timeZone: Type.Optional(Type.String({ description: "Time zone for the new times" })),
        location: Type.Optional(Type.String({ description: "New meeting location" })),
        body: Type.Optional(Type.String({ description: "New event body/description" })),
      }),
      execute: async (_toolCallId, params) =>
        updateCalendarEvent(cfg, params as Parameters<typeof updateCalendarEvent>[1]),
    },
    {
      name: "delete_calendar_event",
      label: "Delete Calendar Event",
      description: "Delete a calendar event.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID whose calendar contains the event" }),
        eventId: Type.String({ description: "ID of the event to delete" }),
      }),
      execute: async (_toolCallId, params) =>
        deleteCalendarEvent(cfg, params as Parameters<typeof deleteCalendarEvent>[1]),
    },
    {
      name: "find_meeting_times",
      label: "Find Meeting Times",
      description:
        "Find available meeting times when all attendees are free. Uses Microsoft's scheduling assistant.",
      parameters: Type.Object({
        userId: Type.String({ description: "Organizer's email or ID" }),
        attendees: Type.Array(Type.String(), {
          description: "List of attendee email addresses",
        }),
        durationMinutes: Type.Number({ description: "Meeting duration in minutes" }),
        startDateTime: Type.String({
          description: "Search window start in ISO format",
        }),
        endDateTime: Type.String({
          description: "Search window end in ISO format",
        }),
        timeZone: Type.Optional(
          Type.String({ description: "Time zone (default: from config or UTC)" }),
        ),
      }),
      execute: async (_toolCallId, params) =>
        findMeetingTimes(cfg, params as Parameters<typeof findMeetingTimes>[1]),
    },
    {
      name: "send_email",
      label: "Send Email",
      description: "Send an email using Microsoft Graph.",
      parameters: Type.Object({
        userId: Type.String({ description: "Sender's email or ID (must have send permissions)" }),
        to: Type.Array(Type.String(), { description: "List of recipient email addresses" }),
        subject: Type.String({ description: "Email subject" }),
        body: Type.String({ description: "Email body content" }),
        cc: Type.Optional(Type.Array(Type.String(), { description: "CC recipients" })),
        bcc: Type.Optional(Type.Array(Type.String(), { description: "BCC recipients" })),
        importance: Type.Optional(
          Type.Union([Type.Literal("low"), Type.Literal("normal"), Type.Literal("high")], {
            description: "Email importance level",
          }),
        ),
      }),
      execute: async (_toolCallId, params) => sendEmail(cfg, params as Parameters<typeof sendEmail>[1]),
    },
    {
      name: "get_user_info",
      label: "Get User Info",
      description: "Get user profile information from Microsoft Graph.",
      parameters: Type.Object({
        userId: Type.String({ description: "User email or ID to look up" }),
      }),
      execute: async (_toolCallId, params) => getUserInfo(cfg, params as Parameters<typeof getUserInfo>[1]),
    },
  ];
}
