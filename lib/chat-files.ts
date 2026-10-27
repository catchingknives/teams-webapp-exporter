/**
 * Binds each conversation to its export file by thread ID.
 *
 * Deduplicated names like "Project (3)" are handed out by recency, so the
 * same name points at a different conversation from one run to the next.
 * Choosing the file by name therefore appended one chat's messages into
 * another chat's file. Each file now carries a `<!-- chat-id: ... -->` line,
 * and the file is looked up by that ID instead.
 *
 * Files written before this change have no chat-id. A conversation adopts
 * one only when the file's last-message timestamp is one of the messages
 * actually seen in that conversation: an exact millisecond match between two
 * different chats does not happen in practice, so this cannot misroute.
 */

import { readdir, readFile, writeFile } from "fs/promises";
import { existsSync } from "fs";
import { join } from "path";
import { createHash } from "crypto";
import { sanitizeFilename } from "./markdown";

export interface ExportFile {
  filename: string;
  chatId?: string;
  lastMessage: Date | null;
  family: string;
}

const CHAT_ID_RE = /^<!-- chat-id: (\S+) -->$/m;
const LAST_MESSAGE_RE = /<!-- last-message: (\d{4}-\d{2}-\d{2}T[^\s]+) -->/;

/**
 * A name with its dedup counter and ID suffix removed, so that
 * "Project_(3).md", "Project.md" and "Project__1a2b3c4d.md" share one family.
 */
export function familyOf(filenameOrName: string): string {
  return sanitizeFilename(filenameOrName.replace(/\.md$/, "").replace(/__[0-9a-f]{8}$/, ""))
    .replace(/_\(\d+\)$/, "");
}

export function chatIdSuffix(chatId: string): string {
  return createHash("sha1").update(chatId).digest("hex").slice(0, 8);
}

export function parseExportFile(filename: string, content: string): ExportFile {
  const id = content.match(CHAT_ID_RE);
  const last = content.match(LAST_MESSAGE_RE);
  return {
    filename,
    chatId: id ? id[1] : undefined,
    lastMessage: last ? new Date(last[1]) : null,
    family: familyOf(filename),
  };
}

export class ChatFileIndex {
  private byChatId = new Map<string, string>();
  private unbound: ExportFile[] = [];
  private taken = new Set<string>();

  constructor(files: ExportFile[]) {
    for (const f of files) {
      this.taken.add(f.filename);
      if (f.chatId) this.byChatId.set(f.chatId, f.filename);
      else if (f.lastMessage) this.unbound.push(f);
    }
  }

  static async load(outputDir: string): Promise<ChatFileIndex> {
    if (!existsSync(outputDir)) return new ChatFileIndex([]);
    const names = (await readdir(outputDir)).filter((n) => n.endsWith(".md"));
    const files = await Promise.all(
      names.map(async (n) => parseExportFile(n, await readFile(join(outputDir, n), "utf-8")))
    );
    return new ChatFileIndex(files);
  }

  /** The file already bound to this conversation, if any. */
  boundFile(chatId: string): string | undefined {
    return this.byChatId.get(chatId);
  }

  /** Unbound files of the same name family that this conversation might own. */
  candidates(chatName: string): ExportFile[] {
    const family = familyOf(chatName);
    return this.unbound.filter((f) => f.family === family);
  }

  /**
   * The unbound candidate whose last message is one of the given timestamps.
   * If several match (an older file and a newer one of the same chat), the
   * one with the newest marker wins, since it holds the most history.
   */
  matchCandidate(chatName: string, timestamps: string[]): ExportFile | undefined {
    const seen = new Set(timestamps.map((t) => new Date(t).getTime()));
    return this.candidates(chatName)
      .filter((f) => seen.has(f.lastMessage!.getTime()))
      .sort((a, b) => b.lastMessage!.getTime() - a.lastMessage!.getTime())[0];
  }

  /** A filename for a conversation that has no file yet; never an existing one. */
  newFilename(chatName: string, chatId: string): string {
    const plain = sanitizeFilename(chatName) + ".md";
    if (!this.taken.has(plain)) return plain;
    return sanitizeFilename(chatName).slice(0, 190) + "__" + chatIdSuffix(chatId) + ".md";
  }

  /** Record a binding made during this run. */
  bind(chatId: string, filename: string): void {
    this.byChatId.set(chatId, filename);
    this.taken.add(filename);
    this.unbound = this.unbound.filter((f) => f.filename !== filename);
  }
}

/** Write the chat-id line into a file that lacks one, right below its title. */
export async function stampChatId(filePath: string, chatId: string): Promise<void> {
  const content = await readFile(filePath, "utf-8");
  if (CHAT_ID_RE.test(content)) return;
  const nl = content.indexOf("\n");
  const stamped = nl < 0
    ? `${content}\n<!-- chat-id: ${chatId} -->\n`
    : `${content.slice(0, nl + 1)}<!-- chat-id: ${chatId} -->\n${content.slice(nl + 1)}`;
  await writeFile(filePath, stamped, "utf-8");
}
