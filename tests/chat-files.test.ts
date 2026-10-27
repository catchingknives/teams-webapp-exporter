import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { mkdtemp, rm, writeFile, readFile } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";
import {
  ChatFileIndex,
  familyOf,
  chatIdSuffix,
  parseExportFile,
  stampChatId,
} from "../lib/chat-files";
import { saveChat } from "../lib/markdown";

function legacyFile(title: string, lastMessage: string): string {
  return `# ${title}\n\nExported: 2026-06-01T00:00:00Z\n\n**Bob** [${lastMessage}]:\nhi\n\n<!-- last-message: ${lastMessage} -->\n`;
}

describe("familyOf", () => {
  it("groups dedup counters and ID suffixes under one family", () => {
    expect(familyOf("Project (3)")).toBe("Project");
    expect(familyOf("Project_(3).md")).toBe("Project");
    expect(familyOf("Project.md")).toBe("Project");
    expect(familyOf(`Project__${chatIdSuffix("19:x@thread.v2")}.md`)).toBe("Project");
  });

  it("keeps different names apart", () => {
    expect(familyOf("Doe Jane (ACME OU-A) (5)")).not.toBe(familyOf("Doe Jane (ACME OU-B) (5)"));
  });
});

describe("parseExportFile", () => {
  it("reads the chat-id and last-message markers", () => {
    const f = parseExportFile("A.md", "# A\n<!-- chat-id: 19:a@thread.v2 -->\n\n<!-- last-message: 2026-01-01T10:00:00.000Z -->\n");
    expect(f.chatId).toBe("19:a@thread.v2");
    expect(f.lastMessage?.toISOString()).toBe("2026-01-01T10:00:00.000Z");
  });

  it("leaves chatId undefined on files written before chat-ids", () => {
    expect(parseExportFile("A.md", legacyFile("A", "2026-01-01T10:00:00.000Z")).chatId).toBeUndefined();
  });
});

describe("ChatFileIndex", () => {
  const files = [
    parseExportFile("Project_(1).md", legacyFile("Project (1)", "2026-03-01T10:00:00.000Z")),
    parseExportFile("Project_(2).md", legacyFile("Project (2)", "2026-02-01T10:00:00.000Z")),
    parseExportFile("Other.md", legacyFile("Other", "2026-03-01T10:00:00.000Z")),
    parseExportFile("Bound.md", "# Bound\n<!-- chat-id: 19:bound -->\n<!-- last-message: 2026-01-01T00:00:00.000Z -->\n"),
  ];

  it("finds a bound file by thread ID", () => {
    expect(new ChatFileIndex(files).boundFile("19:bound")).toBe("Bound.md");
  });

  it("offers only unbound files of the same family as candidates", () => {
    const names = new ChatFileIndex(files).candidates("Project (7)").map((f) => f.filename).sort();
    expect(names).toEqual(["Project_(1).md", "Project_(2).md"]);
  });

  it("matches the file whose last message the chat actually contains, whatever its number", () => {
    // Today this chat is called "Project (1)", but its history is in Project_(2).md
    const match = new ChatFileIndex(files).matchCandidate("Project (1)", [
      "2026-01-15T09:00:00.000Z",
      "2026-02-01T10:00:00.000Z",
      "2026-02-20T09:00:00.000Z",
    ]);
    expect(match?.filename).toBe("Project_(2).md");
  });

  it("matches nothing when no candidate's last message is in the chat", () => {
    expect(new ChatFileIndex(files).matchCandidate("Project (1)", ["2026-02-01T10:00:00.001Z"])).toBeUndefined();
  });

  it("never matches a file from another family, even on an identical timestamp", () => {
    const match = new ChatFileIndex(files).matchCandidate("Unrelated", ["2026-03-01T10:00:00.000Z"]);
    expect(match).toBeUndefined();
  });

  it("prefers the candidate with the newest marker when several match", () => {
    const match = new ChatFileIndex(files).matchCandidate("Project", [
      "2026-02-01T10:00:00.000Z",
      "2026-03-01T10:00:00.000Z",
    ]);
    expect(match?.filename).toBe("Project_(1).md");
  });

  it("removes a file from the candidates once bound", () => {
    const index = new ChatFileIndex(files);
    index.bind("19:p", "Project_(1).md");
    expect(index.boundFile("19:p")).toBe("Project_(1).md");
    expect(index.candidates("Project").map((f) => f.filename)).toEqual(["Project_(2).md"]);
  });

  it("never hands out an existing filename for a new chat", () => {
    const index = new ChatFileIndex(files);
    expect(index.newFilename("Fresh", "19:f")).toBe("Fresh.md");
    expect(index.newFilename("Project (1)", "19:f")).toBe(`Project_(1)__${chatIdSuffix("19:f")}.md`);
  });
});

describe("binding on disk", () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await mkdtemp(join(tmpdir(), "chat-files-"));
  });

  afterEach(async () => {
    await rm(tmpDir, { recursive: true });
  });

  it("stamps the chat-id below the title of a legacy file, once", async () => {
    const path = join(tmpDir, "A.md");
    await writeFile(path, legacyFile("A", "2026-01-01T10:00:00.000Z"));
    await stampChatId(path, "19:a");
    await stampChatId(path, "19:a");
    const content = await readFile(path, "utf-8");
    expect(content.startsWith("# A\n<!-- chat-id: 19:a -->\n\nExported:")).toBe(true);
    expect(content.match(/chat-id/g)).toHaveLength(1);
  });

  it("writes new files with their chat-id, and the index loads it back", async () => {
    await saveChat("Fresh", [{ id: 1, author: "Bob", timestamp: "2026-01-01T10:00:00.000Z", content: "hi" }], tmpDir, {
      filename: "Fresh__abc.md",
      chatId: "19:fresh",
    });
    const index = await ChatFileIndex.load(tmpDir);
    expect(index.boundFile("19:fresh")).toBe("Fresh__abc.md");
  });

  it("appends to the bound file, not to the file the chat's current name points at", async () => {
    await writeFile(join(tmpDir, "Project_(1).md"), legacyFile("Project (1)", "2026-03-01T10:00:00.000Z"));
    await writeFile(join(tmpDir, "Project_(2).md"), legacyFile("Project (2)", "2026-02-01T10:00:00.000Z"));
    await saveChat(
      "Project (1)",
      [{ id: 1, author: "Bob", timestamp: "2026-04-01T10:00:00.000Z", content: "new" }],
      tmpDir,
      { filename: "Project_(2).md", chatId: "19:p2" }
    );
    expect(await readFile(join(tmpDir, "Project_(1).md"), "utf-8")).toBe(legacyFile("Project (1)", "2026-03-01T10:00:00.000Z"));
    expect(await readFile(join(tmpDir, "Project_(2).md"), "utf-8")).toContain("new");
  });
});
