import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";
import { createGraphTools } from "./graph-tools.js";
import type { A365Config } from "./types.js";

describe("graph-tools", () => {
  const mockFetch = vi.fn();
  const originalFetch = global.fetch;

  beforeEach(() => {
    global.fetch = mockFetch;
    mockFetch.mockReset();
  });

  afterEach(() => {
    global.fetch = originalFetch;
  });

  const mockTokenResponse = () => {
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: async () => ({
        access_token: "test-token",
        expires_in: 3600,
      }),
    });
  };

  const cfg: A365Config = {
    graph: {
      clientId: "test-client-id",
      clientSecret: "test-client-secret",
      tenantId: "test-tenant-id",
    },
    owner: "owner@test.com",
  };

  describe("createGraphTools", () => {
    it("returns all expected tools", () => {
      const tools = createGraphTools(cfg);
      const toolNames = tools.map((t) => t.name);

      expect(toolNames).toContain("get_calendar_events");
      expect(toolNames).toContain("create_calendar_event");
      expect(toolNames).toContain("update_calendar_event");
      expect(toolNames).toContain("delete_calendar_event");
      expect(toolNames).toContain("find_meeting_times");
      expect(toolNames).toContain("send_email");
      expect(toolNames).toContain("get_user_info");
    });

    it("includes calendar owner in tool descriptions", () => {
      const tools = createGraphTools(cfg);
      const getEvents = tools.find((t) => t.name === "get_calendar_events");
      expect(getEvents?.description).toContain("owner@test.com");
    });
  });

  describe("get_calendar_events", () => {
    it("fetches calendar events successfully", async () => {
      mockTokenResponse();
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          value: [
            {
              id: "event1",
              subject: "Test Meeting",
              start: { dateTime: "2024-01-15T10:00:00", timeZone: "UTC" },
              end: { dateTime: "2024-01-15T11:00:00", timeZone: "UTC" },
              attendees: [],
            },
          ],
        }),
      });

      const tools = createGraphTools(cfg);
      const getTool = tools.find((t) => t.name === "get_calendar_events")!;

      const result = await getTool.execute({
        userId: "user@test.com",
        startDate: "2024-01-15T00:00:00",
        endDate: "2024-01-15T23:59:59",
      });

      expect(result.isError).toBeUndefined();
      expect(result.content[0].type).toBe("text");
      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.count).toBe(1);
      expect(parsed.events[0].subject).toBe("Test Meeting");
    });

    it("handles API errors gracefully", async () => {
      mockTokenResponse();
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 404,
        text: async () => JSON.stringify({ error: { message: "Calendar not found" } }),
      });

      const tools = createGraphTools(cfg);
      const getTool = tools.find((t) => t.name === "get_calendar_events")!;

      const result = await getTool.execute({
        userId: "nonexistent@test.com",
        startDate: "2024-01-15T00:00:00",
        endDate: "2024-01-15T23:59:59",
      });

      expect(result.isError).toBe(true);
      expect((result.content[0] as { text: string }).text).toContain("Calendar not found");
    });
  });

  describe("create_calendar_event", () => {
    it("creates a calendar event successfully", async () => {
      mockTokenResponse();
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          id: "new-event-id",
          subject: "New Meeting",
          start: { dateTime: "2024-01-15T14:00:00", timeZone: "Pacific Standard Time" },
          end: { dateTime: "2024-01-15T15:00:00", timeZone: "Pacific Standard Time" },
        }),
      });

      const tools = createGraphTools(cfg);
      const createTool = tools.find((t) => t.name === "create_calendar_event")!;

      const result = await createTool.execute({
        userId: "user@test.com",
        subject: "New Meeting",
        startDateTime: "2024-01-15T14:00:00",
        endDateTime: "2024-01-15T15:00:00",
      });

      expect(result.isError).toBeUndefined();
      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(true);
      expect(parsed.eventId).toBe("new-event-id");
    });

    it("creates event with attendees and online meeting", async () => {
      mockTokenResponse();
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          id: "online-event-id",
          subject: "Online Meeting",
          start: { dateTime: "2024-01-15T14:00:00", timeZone: "Pacific Standard Time" },
          end: { dateTime: "2024-01-15T15:00:00", timeZone: "Pacific Standard Time" },
          onlineMeetingUrl: "https://teams.microsoft.com/meet/123",
        }),
      });

      const tools = createGraphTools(cfg);
      const createTool = tools.find((t) => t.name === "create_calendar_event")!;

      const result = await createTool.execute({
        userId: "user@test.com",
        subject: "Online Meeting",
        startDateTime: "2024-01-15T14:00:00",
        endDateTime: "2024-01-15T15:00:00",
        attendees: ["attendee1@test.com", "attendee2@test.com"],
        isOnlineMeeting: true,
      });

      expect(result.isError).toBeUndefined();
      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.onlineMeetingUrl).toContain("teams.microsoft.com");

      // Verify the POST request body
      const postCall = mockFetch.mock.calls[1];
      const requestBody = JSON.parse(postCall[1].body);
      expect(requestBody.attendees).toHaveLength(2);
      expect(requestBody.isOnlineMeeting).toBe(true);
    });
  });

  describe("send_email", () => {
    it("sends an email successfully", async () => {
      mockTokenResponse();
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 202,
        json: async () => ({}),
      });

      const tools = createGraphTools(cfg);
      const sendTool = tools.find((t) => t.name === "send_email")!;

      const result = await sendTool.execute({
        userId: "sender@test.com",
        to: ["recipient@test.com"],
        subject: "Test Email",
        body: "This is a test email.",
      });

      expect(result.isError).toBeUndefined();
      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.success).toBe(true);
      expect(parsed.message).toContain("recipient@test.com");
    });
  });

  describe("find_meeting_times", () => {
    it("finds available meeting times", async () => {
      mockTokenResponse();
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          meetingTimeSuggestions: [
            {
              meetingTimeSlot: {
                start: { dateTime: "2024-01-15T10:00:00", timeZone: "Pacific Standard Time" },
                end: { dateTime: "2024-01-15T10:30:00", timeZone: "Pacific Standard Time" },
              },
              confidence: 100,
              organizerAvailability: "free",
              attendeeAvailability: [
                {
                  attendee: { emailAddress: { address: "attendee@test.com" } },
                  availability: "free",
                },
              ],
            },
          ],
        }),
      });

      const tools = createGraphTools(cfg);
      const findTool = tools.find((t) => t.name === "find_meeting_times")!;

      const result = await findTool.execute({
        userId: "organizer@test.com",
        attendees: ["attendee@test.com"],
        durationMinutes: 30,
        startDateTime: "2024-01-15T08:00:00",
        endDateTime: "2024-01-15T18:00:00",
      });

      expect(result.isError).toBeUndefined();
      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.count).toBe(1);
      expect(parsed.suggestions[0].confidence).toBe(100);
    });
  });

  describe("get_user_info", () => {
    it("retrieves user information", async () => {
      mockTokenResponse();
      mockFetch.mockResolvedValueOnce({
        ok: true,
        json: async () => ({
          id: "user-id",
          displayName: "Test User",
          mail: "user@test.com",
          userPrincipalName: "user@test.com",
          jobTitle: "Engineer",
          department: "Engineering",
        }),
      });

      const tools = createGraphTools(cfg);
      const getUserTool = tools.find((t) => t.name === "get_user_info")!;

      const result = await getUserTool.execute({
        userId: "user@test.com",
      });

      expect(result.isError).toBeUndefined();
      const parsed = JSON.parse((result.content[0] as { text: string }).text);
      expect(parsed.displayName).toBe("Test User");
      expect(parsed.jobTitle).toBe("Engineer");
    });
  });

  describe("error handling", () => {
    it("returns error when token is not available", async () => {
      const cfgNoGraph: A365Config = {};
      const tools = createGraphTools(cfgNoGraph);
      const getTool = tools.find((t) => t.name === "get_calendar_events")!;

      const result = await getTool.execute({
        userId: "user@test.com",
        startDate: "2024-01-15T00:00:00",
        endDate: "2024-01-15T23:59:59",
      });

      expect(result.isError).toBe(true);
      expect((result.content[0] as { text: string }).text).toContain("token not available");
    });

    it("handles network errors", async () => {
      mockTokenResponse();
      mockFetch.mockRejectedValueOnce(new Error("Network error"));

      const tools = createGraphTools(cfg);
      const getTool = tools.find((t) => t.name === "get_calendar_events")!;

      const result = await getTool.execute({
        userId: "user@test.com",
        startDate: "2024-01-15T00:00:00",
        endDate: "2024-01-15T23:59:59",
      });

      expect(result.isError).toBe(true);
      expect((result.content[0] as { text: string }).text).toContain("Network error");
    });
  });
});
