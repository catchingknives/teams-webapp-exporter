import { describe, it, expect, vi, afterEach } from "vitest";
import { fetchAllChats } from "../lib/graph";

// ---------------------------------------------------------------------------
// Mock fetch globally
// ---------------------------------------------------------------------------

function mockFetchResponse(conversations: any[]) {
  return vi.fn().mockResolvedValue({
    ok: true,
    json: async () => ({ conversations }),
    text: async () => JSON.stringify({ conversations }),
  });
}

function makeConversation(overrides: Record<string, any> = {}) {
  return {
    id: overrides.id ?? "19:abc@thread.v2",
    threadProperties: {
      threadType: overrides.threadType ?? "chat",
      topic: overrides.topic ?? "",
      uniquerosterthread: overrides.uniquerosterthread ?? "false",
      hidden: overrides.hidden ?? "false",
      createdat: overrides.createdat ?? "2025-06-01T00:00:00Z",
      ...overrides.threadProperties,
    },
    properties: {
      lastimreceivedtime: overrides.lastimreceivedtime ?? "2025-06-15T10:00:00Z",
      ...overrides.properties,
    },
    lastMessage: {
      fromDisplayNameInToken: overrides.fromDisplayName ?? "Bob",
      imdisplayname: overrides.imdisplayname ?? "Bob",
      ...overrides.lastMessage,
    },
  };
}

describe("fetchAllChats", () => {
  const originalFetch = globalThis.fetch;

  afterEach(() => {
    globalThis.fetch = originalFetch;
    vi.restoreAllMocks();
  });

  it("filters to exportable conversation types (chat, meeting, topic)", async () => {
    globalThis.fetch = mockFetchResponse([
      makeConversation({ id: "1", threadType: "chat" }),
      makeConversation({ id: "2", threadType: "meeting" }),
      makeConversation({ id: "3", threadType: "topic" }),
      makeConversation({ id: "4", threadType: "streamofconsciousness" }),
      makeConversation({ id: "5", threadType: "annotation" }),
    ]);

    const result = await fetchAllChats("fake-token", "emea");
    expect(result).toHaveLength(3);
    const types = result.map((r) => r.chatType);
    expect(types).toContain("chat");
    expect(types).toContain("meeting");
    expect(types).toContain("topic");
    expect(types).not.toContain("streamofconsciousness");
  });

  it("sorts by last activity (most recent first)", async () => {
    globalThis.fetch = mockFetchResponse([
      makeConversation({
        id: "old",
        lastimreceivedtime: "2025-01-01T00:00:00Z",
        fromDisplayName: "Old",
        topic: "Old Chat",
      }),
      makeConversation({
        id: "new",
        lastimreceivedtime: "2025-06-15T00:00:00Z",
        fromDisplayName: "New",
        topic: "New Chat",
      }),
    ]);

    const result = await fetchAllChats("fake-token", "emea");
    expect(result[0].name).toBe("New Chat");
    expect(result[1].name).toBe("Old Chat");
  });

  it("deduplicates names with (1), (2) suffixes", async () => {
    globalThis.fetch = mockFetchResponse([
      makeConversation({
        id: "a",
        topic: "Project",
        lastimreceivedtime: "2025-06-15T03:00:00Z",
      }),
      makeConversation({
        id: "b",
        topic: "Project",
        lastimreceivedtime: "2025-06-15T02:00:00Z",
      }),
      makeConversation({
        id: "c",
        topic: "Project",
        lastimreceivedtime: "2025-06-15T01:00:00Z",
      }),
    ]);

    const result = await fetchAllChats("fake-token", "emea");
    const names = result.map((r) => r.name);
    expect(names).toContain("Project (1)");
    expect(names).toContain("Project (2)");
    expect(names).toContain("Project (3)");
  });

  it("resolves 1:1 chat names from last message sender", async () => {
    globalThis.fetch = mockFetchResponse([
      makeConversation({
        id: "19:user@unq.gbl.spaces",
        threadType: "chat",
        topic: "",
        uniquerosterthread: "true",
        fromDisplayName: "Charlie",
      }),
    ]);

    const result = await fetchAllChats("fake-token", "emea");
    expect(result[0].name).toBe("Charlie");
  });

  it("detects hidden chats", async () => {
    globalThis.fetch = mockFetchResponse([
      makeConversation({ id: "visible", hidden: "false" }),
      makeConversation({ id: "hidden", hidden: "true" }),
    ]);

    const result = await fetchAllChats("fake-token", "emea");
    const hidden = result.find((r) => r.chatId === "hidden");
    const visible = result.find((r) => r.chatId === "visible");
    expect(hidden!.isHidden).toBe(true);
    expect(visible!.isHidden).toBe(false);
  });

  it("constructs web URLs from thread IDs", async () => {
    const threadId = "19:abc@thread.v2";
    globalThis.fetch = mockFetchResponse([
      makeConversation({ id: threadId }),
    ]);

    const result = await fetchAllChats("fake-token", "emea");
    expect(result[0].webUrl).toContain(encodeURIComponent(threadId));
  });

  it("throws on API error", async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: false,
      status: 401,
      text: async () => "Unauthorized",
    });

    await expect(fetchAllChats("bad-token", "emea")).rejects.toThrow(
      /chatsvc API error 401/
    );
  });

  it("falls back to Chat_N naming when no topic or sender", async () => {
    globalThis.fetch = mockFetchResponse([
      makeConversation({
        id: "19:group@thread.v2",
        topic: "",
        uniquerosterthread: "false",
        lastMessage: {
          fromDisplayNameInToken: "",
          imdisplayname: "",
        },
      }),
    ]);

    const result = await fetchAllChats("fake-token", "emea");
    expect(result[0].name).toMatch(/^Chat_\d+$/);
  });
});
