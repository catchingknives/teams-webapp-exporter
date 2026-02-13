/**
 * Server-side Markdown conversion and file append logic.
 * Converts structured message data to Markdown and handles incremental exports.
 */

import { readFile, writeFile, mkdir } from "fs/promises";
import { existsSync } from "fs";
import { join } from "path";
import type { ExtractedMessage } from "./extract";

/**
 * Convert a single message's HTML content to plain-text Markdown.
 * Strips tags, preserves basic formatting.
 */
function htmlToText(html: string): string {
  return html
    // Convert <br> and block elements to newlines
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/?(p|div|li|tr)>/gi, "\n")
    // Bold
    .replace(/<b>(.*?)<\/b>/gi, "**$1**")
    .replace(/<strong>(.*?)<\/strong>/gi, "**$1**")
    // Italic
    .replace(/<i>(.*?)<\/i>/gi, "_$1_")
    .replace(/<em>(.*?)<\/em>/gi, "_$1_")
    // Inline code
    .replace(/<code>(.*?)<\/code>/gi, "`$1`")
    // Links
    .replace(/<a[^>]+href="([^"]*)"[^>]*>(.*?)<\/a>/gi, "[$2]($1)")
    // Blockquote content
    .replace(/<blockquote>(.*?)<\/blockquote>/gis, (_, content) => {
      return content
        .replace(/<[^>]+>/g, "")
        .trim()
        .split("\n")
        .map((l: string) => `> ${l}`)
        .join("\n");
    })
    // Strip remaining HTML tags
    .replace(/<[^>]+>/g, "")
    // Decode common entities
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ")
    // Clean up excessive newlines
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/**
 * Format a timestamp for display in brackets.
 * Uses ISO 8601 so retention analysis can parse it regardless of locale.
 */
function formatTimestamp(iso: string): string {
  return iso;
}

/**
 * Format a date for section headers.
 * Uses the system locale for human-readable display.
 */
function formatDate(iso: string): string {
  const d = new Date(iso);
  return d.toLocaleDateString(undefined, {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
  });
}

/**
 * Convert an array of extracted messages to Markdown.
 */
export function messagesToMarkdown(messages: ExtractedMessage[]): string {
  if (messages.length === 0) return "";

  const lines: string[] = [];
  let lastAuthor = "";
  let lastDateStr = "";

  for (const msg of messages) {
    const dateStr = new Date(msg.timestamp).toDateString();
    const content = htmlToText(msg.contentHtml);

    // Date separator
    if (dateStr !== lastDateStr) {
      if (lastDateStr !== "") lines.push("", "---", "");
      lines.push(`## ${formatDate(msg.timestamp)}`, "");
      lastDateStr = dateStr;
      lastAuthor = ""; // Reset author on new day
    }

    // Author header (only when author changes)
    if (msg.author !== lastAuthor) {
      lines.push(`**${msg.author}** [${formatTimestamp(msg.timestamp)}]:`);
      lastAuthor = msg.author;
    } else {
      // Same author, just show time
      lines.push(`*[${formatTimestamp(msg.timestamp)}]:*`);
    }

    // Message content — indent continuation lines for readability
    const contentLines = content.split("\n");
    for (const line of contentLines) {
      lines.push(line);
    }
    lines.push(""); // blank line between messages
  }

  return lines.join("\n");
}

/**
 * Sanitize a chat name for use as a filename.
 */
export function sanitizeFilename(name: string): string {
  return name
    .replace(/[\/\\:*?"<>|]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_|_$/g, "")
    .slice(0, 200); // Prevent overly long filenames
}

/**
 * Read an existing export file and return the latest timestamp found,
 * or null if the file doesn't exist or has no timestamps.
 */
export async function getLatestTimestamp(
  filePath: string
): Promise<Date | null> {
  if (!existsSync(filePath)) return null;

  const content = await readFile(filePath, "utf-8");
  // Match ISO timestamps in the metadata comments we embed
  const timestampRegex = /<!-- last-message: (\d{4}-\d{2}-\d{2}T[^\s]+) -->/;
  const match = content.match(timestampRegex);
  if (match) {
    return new Date(match[1]);
  }
  return null;
}

/**
 * Save messages to a Markdown file with append logic.
 * Returns the number of new messages written.
 */
export async function saveChat(
  chatName: string,
  messages: ExtractedMessage[],
  outputDir: string
): Promise<number> {
  await mkdir(outputDir, { recursive: true });

  const filename = sanitizeFilename(chatName) + ".md";
  const filePath = join(outputDir, filename);

  // Check for existing export
  const lastTimestamp = await getLatestTimestamp(filePath);

  // Filter to only new messages
  let newMessages = messages;
  if (lastTimestamp) {
    newMessages = messages.filter(
      (m) => new Date(m.timestamp) > lastTimestamp
    );
  }

  if (newMessages.length === 0) return 0;

  const markdown = messagesToMarkdown(newMessages);
  const lastMsg = newMessages[newMessages.length - 1];
  const metadata = `<!-- last-message: ${lastMsg.timestamp} -->`;

  if (existsSync(filePath) && lastTimestamp) {
    // Append mode: add separator and new messages
    const existing = await readFile(filePath, "utf-8");
    // Remove old metadata line
    const cleaned = existing.replace(
      /<!-- last-message: [^\s]+ -->\n?$/,
      ""
    );
    const appended =
      cleaned.trimEnd() +
      "\n\n---\n\n" +
      `### Export appended: ${new Date().toISOString()}\n\n` +
      markdown +
      "\n" +
      metadata +
      "\n";
    await writeFile(filePath, appended, "utf-8");
  } else {
    // Fresh export
    const header = `# ${chatName}\n\nExported: ${new Date().toISOString()}\n\n`;
    const content = header + markdown + "\n" + metadata + "\n";
    await writeFile(filePath, content, "utf-8");
  }

  return newMessages.length;
}
