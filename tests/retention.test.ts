import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { mkdtemp, rm, writeFile } from "fs/promises";
import { join } from "path";
import { tmpdir } from "os";
import {
  analyzeRetention,
  formatRetentionSummary,
} from "../lib/retention";

// ---------------------------------------------------------------------------
// Fixtures — create markdown files that look like real exports
// ---------------------------------------------------------------------------

function chatMarkdown(
  name: string,
  timestamps: string[]
): string {
  const lines = [`# ${name}`, "", "Exported: 15.06.2025, 12:00", ""];
  for (const ts of timestamps) {
    lines.push(`**Alice** [${ts}]:`, "Hello", "");
  }
  return lines.join("\n");
}

describe("analyzeRetention", () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await mkdtemp(join(tmpdir(), "teams-retention-"));
  });

  afterEach(async () => {
    await rm(tmpDir, { recursive: true, force: true });
  });

  it("handles empty directory", async () => {
    const result = await analyzeRetention(tmpDir);
    expect(result.chatsAnalyzed).toBe(0);
    expect(result.chatsWithDates).toBe(0);
    expect(result.oldestMessage).toBeNull();
    expect(result.explanation).toContain("No chat files");
  });

  it("handles files with no parseable timestamps", async () => {
    await writeFile(
      join(tmpDir, "no_dates.md"),
      "# Chat\nJust text, no timestamps\n",
      "utf-8"
    );
    const result = await analyzeRetention(tmpDir);
    expect(result.chatsAnalyzed).toBe(1);
    expect(result.chatsWithDates).toBe(0);
    expect(result.oldestMessage).toBeNull();
  });

  it("parses ISO 8601 timestamps from markdown", async () => {
    await writeFile(
      join(tmpDir, "alice.md"),
      chatMarkdown("Alice", ["2025-06-15T10:30:00Z", "2025-06-15T11:00:00Z"]),
      "utf-8"
    );
    const result = await analyzeRetention(tmpDir);
    expect(result.chatsWithDates).toBe(1);
    expect(result.oldestMessage).not.toBeNull();
    expect(result.chatRanges[0].messageCount).toBe(2);
  });

  it("parses legacy German date timestamps from old exports", async () => {
    await writeFile(
      join(tmpDir, "legacy.md"),
      chatMarkdown("Legacy", ["15.06.2025, 10:30", "15.06.2025, 11:00"]),
      "utf-8"
    );
    const result = await analyzeRetention(tmpDir);
    expect(result.chatsWithDates).toBe(1);
    expect(result.oldestMessage).not.toBeNull();
    expect(result.chatRanges[0].messageCount).toBe(2);
  });

  it("identifies oldest and newest messages across chats", async () => {
    await writeFile(
      join(tmpDir, "old_chat.md"),
      chatMarkdown("Old Chat", ["2025-01-01T08:00:00Z"]),
      "utf-8"
    );
    await writeFile(
      join(tmpDir, "new_chat.md"),
      chatMarkdown("New Chat", ["2025-06-15T14:00:00Z"]),
      "utf-8"
    );
    const result = await analyzeRetention(tmpDir);
    expect(result.chatsWithDates).toBe(2);
    expect(result.oldestMessage!.chat).toBe("old chat");
    expect(result.newestMessage!.chat).toBe("new chat");
    expect(result.maxOldestAgeDays).toBeGreaterThan(result.medianOldestAgeDays);
  });

  it("sorts chatRanges by age descending", async () => {
    await writeFile(
      join(tmpDir, "recent.md"),
      chatMarkdown("Recent", ["2025-06-10T10:00:00Z"]),
      "utf-8"
    );
    await writeFile(
      join(tmpDir, "old.md"),
      chatMarkdown("Old", ["2025-01-01T10:00:00Z"]),
      "utf-8"
    );
    const result = await analyzeRetention(tmpDir);
    expect(result.chatRanges[0].ageDays).toBeGreaterThanOrEqual(
      result.chatRanges[1].ageDays
    );
  });

  it("finds nearest retention tier", async () => {
    // A chat from ~90 days ago should match the 90-day tier approximately
    const now = new Date();
    const ninetyAgo = new Date(now.getTime() - 90 * 86400000);
    const ts = ninetyAgo.toISOString();

    await writeFile(
      join(tmpDir, "tier_test.md"),
      chatMarkdown("Tier", [ts]),
      "utf-8"
    );
    const result = await analyzeRetention(tmpDir);
    expect(result.nearestRetentionTier).not.toBeNull();
    // Should be close to 90 days tier
    expect(result.nearestRetentionTier!.days).toBeLessThanOrEqual(180);
  });

  it("generates cluster explanation when 3+ chats cluster", async () => {
    // Create 4 chats all starting around the same date
    for (let i = 0; i < 4; i++) {
      await writeFile(
        join(tmpDir, `cluster_${i}.md`),
        chatMarkdown(`Cluster ${i}`, ["2025-01-01T10:00:00Z", "2025-06-15T12:00:00Z"]),
        "utf-8"
      );
    }
    const result = await analyzeRetention(tmpDir);
    expect(result.explanation).toContain("chats cluster");
  });
});

describe("formatRetentionSummary", () => {
  it("returns fallback for empty analysis", () => {
    const summary = formatRetentionSummary({
      chatsAnalyzed: 0,
      chatsWithDates: 0,
      oldestMessage: null,
      newestMessage: null,
      medianOldestAgeDays: 0,
      maxOldestAgeDays: 0,
      sidebarWindowDays: 0,
      explanation: "",
      nearestRetentionTier: null,
      chatRanges: [],
    });
    expect(summary).toContain("Could not analyze");
  });

  it("includes key sections in formatted output", async () => {
    const tmpDir = await mkdtemp(join(tmpdir(), "teams-fmt-"));
    try {
      await writeFile(
        join(tmpDir, "test.md"),
        chatMarkdown("Test", ["2025-03-01T10:00:00Z", "2025-06-15T14:00:00Z"]),
        "utf-8"
      );
      const analysis = await analyzeRetention(tmpDir);
      const summary = formatRetentionSummary(analysis);

      expect(summary).toContain("Chat Visibility & Retention Analysis");
      expect(summary).toContain("Chats exported:");
      expect(summary).toContain("Oldest message:");
      expect(summary).toContain("Sidebar window:");
      expect(summary).toContain("Top 5 oldest chats:");
      expect(summary).toContain("Microsoft Purview");
    } finally {
      await rm(tmpDir, { recursive: true, force: true });
    }
  });
});
