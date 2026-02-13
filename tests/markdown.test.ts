import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { mkdtemp, rm, readFile } from "fs/promises";
import { join } from "path";
import { tmpdir } from "os";
import {
  messagesToMarkdown,
  sanitizeFilename,
  getLatestTimestamp,
  saveChat,
} from "../lib/markdown";
import type { ExtractedMessage } from "../lib/extract";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function msg(
  overrides: Partial<ExtractedMessage> & { id?: number } = {}
): ExtractedMessage {
  return {
    id: overrides.id ?? 1,
    author: overrides.author ?? "Alice",
    timestamp: overrides.timestamp ?? "2025-06-15T10:30:00Z",
    contentHtml: overrides.contentHtml ?? "Hello",
  };
}

// ---------------------------------------------------------------------------
// htmlToText (tested indirectly through messagesToMarkdown)
// ---------------------------------------------------------------------------

describe("htmlToText (via messagesToMarkdown)", () => {
  it("converts <br> to newline", () => {
    const md = messagesToMarkdown([msg({ contentHtml: "line1<br>line2" })]);
    expect(md).toContain("line1\nline2");
  });

  it("converts <b> and <strong> to bold", () => {
    const md = messagesToMarkdown([
      msg({ contentHtml: "<b>bold</b> and <strong>strong</strong>" }),
    ]);
    expect(md).toContain("**bold**");
    expect(md).toContain("**strong**");
  });

  it("converts <i> and <em> to italic", () => {
    const md = messagesToMarkdown([
      msg({ contentHtml: "<i>italic</i> and <em>emph</em>" }),
    ]);
    expect(md).toContain("_italic_");
    expect(md).toContain("_emph_");
  });

  it("converts <code> to inline code", () => {
    const md = messagesToMarkdown([
      msg({ contentHtml: "use <code>npm test</code>" }),
    ]);
    expect(md).toContain("`npm test`");
  });

  it("converts links to markdown links", () => {
    const md = messagesToMarkdown([
      msg({
        contentHtml: '<a href="https://example.com">click here</a>',
      }),
    ]);
    expect(md).toContain("[click here](https://example.com)");
  });

  it("converts blockquotes", () => {
    const md = messagesToMarkdown([
      msg({ contentHtml: "<blockquote>quoted text</blockquote>" }),
    ]);
    expect(md).toContain("> quoted text");
  });

  it("decodes HTML entities", () => {
    const md = messagesToMarkdown([
      msg({
        contentHtml: "&amp; &lt; &gt; &quot; &#39; &nbsp;",
      }),
    ]);
    expect(md).toContain("& < > \" '");
  });

  it("strips unknown tags", () => {
    const md = messagesToMarkdown([
      msg({ contentHtml: "<span class='x'>plain</span>" }),
    ]);
    expect(md).toContain("plain");
    expect(md).not.toContain("<span");
  });
});

// ---------------------------------------------------------------------------
// messagesToMarkdown
// ---------------------------------------------------------------------------

describe("messagesToMarkdown", () => {
  it("returns empty string for empty array", () => {
    expect(messagesToMarkdown([])).toBe("");
  });

  it("renders a single message with date header and author", () => {
    const md = messagesToMarkdown([msg()]);
    expect(md).toContain("## ");
    expect(md).toContain("**Alice**");
    expect(md).toContain("Hello");
  });

  it("uses ISO 8601 timestamps in brackets", () => {
    const md = messagesToMarkdown([msg({ timestamp: "2025-06-15T10:30:00Z" })]);
    expect(md).toContain("[2025-06-15T10:30:00Z]");
  });

  it("groups consecutive messages by same author (time-only for subsequent)", () => {
    const md = messagesToMarkdown([
      msg({ id: 1, timestamp: "2025-06-15T10:30:00Z", contentHtml: "First" }),
      msg({ id: 2, timestamp: "2025-06-15T10:31:00Z", contentHtml: "Second" }),
    ]);
    // First message gets bold author, second gets italic time-only
    const boldCount = (md.match(/\*\*Alice\*\*/g) || []).length;
    expect(boldCount).toBe(1);
    expect(md).toContain("*[");
  });

  it("inserts date separator between different days", () => {
    const md = messagesToMarkdown([
      msg({ id: 1, timestamp: "2025-06-15T10:00:00Z" }),
      msg({ id: 2, timestamp: "2025-06-16T10:00:00Z" }),
    ]);
    expect(md).toContain("---");
    // Two date headers
    const headers = (md.match(/^## /gm) || []).length;
    expect(headers).toBe(2);
  });

  it("resets author on new day", () => {
    const md = messagesToMarkdown([
      msg({ id: 1, author: "Alice", timestamp: "2025-06-15T10:00:00Z" }),
      msg({ id: 2, author: "Alice", timestamp: "2025-06-17T10:00:00Z" }),
    ]);
    // Alice should appear bold twice (once per day)
    const boldCount = (md.match(/\*\*Alice\*\*/g) || []).length;
    expect(boldCount).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// sanitizeFilename
// ---------------------------------------------------------------------------

describe("sanitizeFilename", () => {
  it("replaces special characters with underscores", () => {
    expect(sanitizeFilename('foo/bar\\baz:qux*"<>|')).toBe("foo_bar_baz_qux");
  });

  it("collapses multiple spaces and underscores", () => {
    expect(sanitizeFilename("hello   world__test")).toBe("hello_world_test");
  });

  it("trims leading and trailing underscores", () => {
    expect(sanitizeFilename("_hello_")).toBe("hello");
  });

  it("truncates to 200 characters", () => {
    const long = "a".repeat(300);
    expect(sanitizeFilename(long).length).toBe(200);
  });

  it("handles empty-ish input", () => {
    expect(sanitizeFilename("***")).toBe("");
  });
});

// ---------------------------------------------------------------------------
// getLatestTimestamp
// ---------------------------------------------------------------------------

describe("getLatestTimestamp", () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await mkdtemp(join(tmpdir(), "teams-test-"));
  });

  afterEach(async () => {
    await rm(tmpDir, { recursive: true, force: true });
  });

  it("returns null for nonexistent file", async () => {
    const result = await getLatestTimestamp(join(tmpDir, "nope.md"));
    expect(result).toBeNull();
  });

  it("returns null for file without metadata comment", async () => {
    const filePath = join(tmpDir, "no-meta.md");
    const { writeFile } = await import("fs/promises");
    await writeFile(filePath, "# Chat\nHello world\n", "utf-8");
    const result = await getLatestTimestamp(filePath);
    expect(result).toBeNull();
  });

  it("parses metadata comment timestamp", async () => {
    const filePath = join(tmpDir, "with-meta.md");
    const { writeFile } = await import("fs/promises");
    await writeFile(
      filePath,
      "# Chat\nHello\n<!-- last-message: 2025-06-15T14:30:00.000Z -->\n",
      "utf-8"
    );
    const result = await getLatestTimestamp(filePath);
    expect(result).toBeInstanceOf(Date);
    expect(result!.toISOString()).toBe("2025-06-15T14:30:00.000Z");
  });
});

// ---------------------------------------------------------------------------
// saveChat
// ---------------------------------------------------------------------------

describe("saveChat", () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await mkdtemp(join(tmpdir(), "teams-save-"));
  });

  afterEach(async () => {
    await rm(tmpDir, { recursive: true, force: true });
  });

  it("creates a fresh export file with header and metadata", async () => {
    const messages = [msg({ contentHtml: "Test message" })];
    const count = await saveChat("Test Chat", messages, tmpDir);

    expect(count).toBe(1);
    const content = await readFile(join(tmpDir, "Test_Chat.md"), "utf-8");
    expect(content).toContain("# Test Chat");
    expect(content).toContain("Test message");
    expect(content).toContain("<!-- last-message:");
  });

  it("appends new messages to existing export", async () => {
    // First export
    const messages1 = [
      msg({
        id: 1,
        timestamp: "2025-06-15T10:00:00.000Z",
        contentHtml: "First",
      }),
    ];
    await saveChat("Append Test", messages1, tmpDir);

    // Second export with newer message
    const messages2 = [
      msg({
        id: 1,
        timestamp: "2025-06-15T10:00:00.000Z",
        contentHtml: "First",
      }),
      msg({
        id: 2,
        timestamp: "2025-06-15T11:00:00.000Z",
        contentHtml: "Second",
      }),
    ];
    const count = await saveChat("Append Test", messages2, tmpDir);

    expect(count).toBe(1); // Only the new message
    const content = await readFile(join(tmpDir, "Append_Test.md"), "utf-8");
    expect(content).toContain("First");
    expect(content).toContain("Second");
    expect(content).toContain("Export appended:");
  });

  it("returns 0 when no new messages", async () => {
    const messages = [
      msg({ timestamp: "2025-06-15T10:00:00.000Z", contentHtml: "Only" }),
    ];
    await saveChat("No New", messages, tmpDir);
    const count = await saveChat("No New", messages, tmpDir);
    expect(count).toBe(0);
  });
});
