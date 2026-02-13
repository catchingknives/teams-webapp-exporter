#!/usr/bin/env npx tsx
/**
 * export-all.ts — Export all MS Teams chats to Markdown via Playwright.
 *
 * Connects to an existing Chrome instance via CDP, navigates Teams,
 * iterates through all chats, and exports each to a Markdown file.
 *
 * Usage:
 *   npx tsx export-all.ts [options]
 *
 * Options:
 *   --days <n>       How far back to scroll per chat (default: 0 = all history)
 *   --output <dir>   Target directory (default: ./output)
 *   --chat <name>    Export specific chat only
 *   --dry-run        List chats without exporting
 *   --port <n>       Chrome debugging port (default: 9222)
 *   --timeout <s>    Per-chat extraction timeout in seconds (default: 120)
 *   --scroll-down    Scroll down from current position instead of up
 *   --region <r>     Teams chatsvc region (default: emea)
 *   --user <name>    Your display name (auto-detected if omitted)
 *
 * Prerequisites:
 *   Start Chrome with remote debugging:
 *   /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \
 *     --remote-debugging-port=9222 \
 *     --remote-allow-origins='*' \
 *     --user-data-dir="$HOME/Library/Application Support/Google/Chrome-Debug"
 */

import { chromium, type Page, type BrowserContext } from "playwright-core";
import { parseArgs } from "util";
import { getExtractionScript, getScrollDownExtractionScript, type ExtractionResult } from "./lib/extract";
import { saveChat, sanitizeFilename, getLatestTimestamp } from "./lib/markdown";
import { join } from "path";
import { analyzeRetention, formatRetentionSummary } from "./lib/retention";
import { getToken, fetchAllChats } from "./lib/graph";

// ── CLI argument parsing ──────────────────────────────────────────────
const { values: args } = parseArgs({
  args: process.argv.slice(2),
  options: {
    days: { type: "string", default: "0" },
    output: {
      type: "string",
      default: "./output",
    },
    chat: { type: "string" },
    "dry-run": { type: "boolean", default: false },
    port: { type: "string", default: "9222" },
    timeout: { type: "string", default: "120" },
    "scroll-down": { type: "boolean", default: false },
    region: { type: "string", default: "emea" },
    user: { type: "string" },
  },
});

const DAYS = parseInt(args.days!, 10);
const OUTPUT_DIR = args.output!;
const CHAT_FILTER = args.chat;
const DRY_RUN = args["dry-run"]!;
const PORT = parseInt(args.port!, 10);
const TIMEOUT_MS = parseInt(args.timeout!, 10) * 1000;
const SCROLL_DOWN = args["scroll-down"]!;
const REGION = args.region!;

// ── Helpers ───────────────────────────────────────────────────────────

function log(msg: string) {
  console.log(`[teams-export] ${msg}`);
}

function warn(msg: string) {
  console.warn(`[teams-export] ⚠ ${msg}`);
}

async function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ── Auto-detect current user ──────────────────────────────────────────

/**
 * Detect the logged-in user's display name from the Teams DOM.
 * Tries several selectors (profile button, me-control, settings).
 * Returns the display name or null if detection fails.
 */
async function detectCurrentUser(page: Page): Promise<string | null> {
  const name = await page.evaluate(`(function() {
    // Strategy 1: me-control / profile button aria-label
    // e.g. aria-label="Profile, John Doe, Online"
    var selectors = [
      '[data-tid="me-control"]',
      '#me-control-button',
      'button[aria-label*="Profile"]',
      '[data-tid="app-bar-me-control"]',
    ];
    for (var i = 0; i < selectors.length; i++) {
      var el = document.querySelector(selectors[i]);
      if (!el) continue;
      var label = el.getAttribute('aria-label') || '';
      // aria-label is typically "Profile, Display Name, Status"
      var parts = label.split(',');
      if (parts.length >= 2) {
        var name = parts[1].trim();
        if (name && name.length > 1) return name;
      }
    }

    // Strategy 2: Profile card / display name element
    var nameEl = document.querySelector('[data-tid="me-control-displayname"]');
    if (nameEl && nameEl.textContent) {
      var text = nameEl.textContent.trim();
      if (text.length > 1) return text;
    }

    return null;
  })()`);

  return name as string | null;
}

// ── Connect to Chrome ─────────────────────────────────────────────────

async function connectToChrome(): Promise<{
  context: BrowserContext;
  page: Page;
}> {
  log(`Connecting to Chrome on port ${PORT}...`);

  let browser;
  try {
    browser = await chromium.connectOverCDP(`http://localhost:${PORT}`);
  } catch (e: any) {
    console.error(`\nPlaywright error: ${e.message}\n`);
    console.error(
      `\nFailed to connect to Chrome on port ${PORT}.\n` +
        `Chrome must be started with remote debugging ENABLED.\n\n` +
        `Step 1: Fully quit Chrome (Cmd+Q), then verify no processes remain:\n` +
        `  pkill -f "Google Chrome"\n\n` +
        `Step 2: Wait 2-3 seconds, then launch with debugging:\n` +
        `  /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome \\\n` +
        `    --remote-debugging-port=${PORT} \\\n` +
        `    --remote-allow-origins='*' \\\n` +
        `    --user-data-dir="$HOME/Library/Application Support/Google/Chrome-Debug" &\n\n` +
        `Step 3: Navigate to https://teams.microsoft.com and sign in.\n` +
        `Step 4: Re-run this script.\n`
    );
    process.exit(1);
  }

  const contexts = browser.contexts();
  if (contexts.length === 0) {
    console.error("No browser contexts found.");
    process.exit(1);
  }

  const context = contexts[0];

  // Find a Teams tab
  const pages = context.pages();
  let teamsPage: Page | undefined;

  for (const p of pages) {
    const url = p.url();
    if (url.includes("teams.microsoft.com") || url.includes("teams.live.com")) {
      teamsPage = p;
      break;
    }
  }

  if (!teamsPage) {
    // Try navigating the first page to Teams
    warn("No Teams tab found. Looking for any Microsoft Teams page...");
    for (const p of pages) {
      const title = await p.title();
      if (title.toLowerCase().includes("teams")) {
        teamsPage = p;
        break;
      }
    }
  }

  if (!teamsPage) {
    console.error(
      "No Microsoft Teams tab found.\n" +
        "Please open https://teams.microsoft.com in Chrome and sign in first.\n"
    );
    process.exit(1);
  }

  log(`Found Teams tab: ${teamsPage.url()}`);
  return { context, page: teamsPage };
}

// ── Chat List Discovery ───────────────────────────────────────────────

interface ChatEntry {
  name: string;
  index: number;
  chatId?: string;  // Thread ID from chatsvc API
  webUrl?: string;  // Deep link URL to open chat
  chatType?: string;
  isHidden?: boolean;
  value?: string;   // Legacy: data-fui-tree-item-value for sidebar fallback
}

// Set of thread IDs known to exist in the sidebar (populated during discovery)
const sidebarThreadIds = new Set<string>();

// Current user's display name (detected or provided via --user flag)
let currentUserName: string | null = null;

async function discoverChats(page: Page): Promise<ChatEntry[]> {
  log("Discovering chats via Teams internal API...");

  try {
    // Step 1: Get ic3 token from localStorage
    log("Extracting auth token from Teams session...");
    const token = await getToken(page);
    log(`Token found (${token.length} chars)`);

    // Step 2: Fetch ALL conversations via chatsvc API (including hidden)
    log("Fetching all conversations from Teams chatsvc API...");
    const chats = await fetchAllChats(token, REGION);

    const chatCount = chats.filter((c) => c.chatType === "chat").length;
    const meetingCount = chats.filter((c) => c.chatType === "meeting").length;
    const topicCount = chats.filter((c) => c.chatType === "topic").length;
    const hiddenCount = chats.filter((c) => c.isHidden).length;

    log(`Found ${chats.length} conversations (${chatCount} chats, ${meetingCount} meetings, ${topicCount} topics, ${hiddenCount} hidden)`);

    // Step 3: Resolve names from sidebar for chats with ambiguous names
    await resolveNamesFromSidebar(page, chats);

    return chats;
  } catch (e: any) {
    warn(`chatsvc API discovery failed: ${e.message}`);
    warn("Falling back to sidebar scrolling method...");
    return discoverChatsSidebar(page);
  }
}

/**
 * Scroll the sidebar to collect threadId → displayName mappings,
 * then fix names for API-discovered chats that show the user's own name.
 */
async function resolveNamesFromSidebar(page: Page, chats: ChatEntry[]): Promise<void> {
  log("Resolving chat names from sidebar...");

  // Collect sidebar tree items: array of { value, name } for matching
  const sidebarItems: Array<{ value: string; name: string }> = await page.evaluate(`(async function() {
    var sleep = function(ms) { return new Promise(function(r) { setTimeout(r, ms); }); };
    var items = [];
    var seenValues = {};

    var tree = document.querySelector("[role='tree']");
    if (!tree) return items;

    var scrollContainer = null;
    var el = tree;
    while (el && el !== document.documentElement) {
      var style = getComputedStyle(el);
      if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && el.scrollHeight > el.clientHeight) {
        scrollContainer = el;
        break;
      }
      el = el.parentElement;
    }
    if (!scrollContainer && tree.scrollHeight > tree.clientHeight) {
      scrollContainer = tree;
    }

    var collect = function() {
      var els = document.querySelectorAll("[data-fui-tree-item-value]");
      for (var i = 0; i < els.length; i++) {
        var val = els[i].getAttribute("data-fui-tree-item-value") || "";
        if (!val || seenValues[val]) continue;
        seenValues[val] = true;

        // Get display name from first meaningful span
        var spans = els[i].querySelectorAll("span");
        var name = "";
        for (var j = 0; j < spans.length; j++) {
          var text = (spans[j].textContent || "").trim();
          if (text.length > 1 && !text.match(/^\\d+:\\d+\\s*(AM|PM)?$/i) && !text.match(/^\\d+\\/\\d+$/)) {
            name = text;
            break;
          }
        }
        if (name) items.push({ value: val, name: name });
      }
    };

    collect();

    if (scrollContainer) {
      var noNew = 0, prev = items.length;
      for (var i = 0; i < 200; i++) {
        scrollContainer.scrollTop += scrollContainer.clientHeight * 0.8;
        await sleep(300);
        collect();
        if (items.length === prev) { noNew++; if (noNew >= 5) break; }
        else noNew = 0;
        prev = items.length;
      }
      scrollContainer.scrollTop = 0;
    }

    return items;
  })()`);

  log(`Collected ${sidebarItems.length} names from sidebar`);

  // Build set of thread IDs found in sidebar (for navigation optimization)
  for (const chat of chats) {
    if (!chat.chatId) continue;
    const inSidebar = sidebarItems.some(s => s.value.includes(chat.chatId!));
    if (inSidebar) sidebarThreadIds.add(chat.chatId);
  }
  log(`${sidebarThreadIds.size} chats navigable via sidebar`);

  // Fix chat names using sidebar data (match by threadId substring, like navigateToChat)
  let fixed = 0;

  for (const chat of chats) {
    if (!chat.chatId) continue;

    const currentLower = chat.name.toLowerCase();
    if (!currentUserName || (!currentLower.includes(currentUserName.toLowerCase()) && !currentLower.startsWith("chat_"))) continue;

    // Find sidebar item whose value contains this thread ID
    const match = sidebarItems.find(s => s.value.includes(chat.chatId!));
    if (match) {
      chat.name = match.name;
      fixed++;
    }
  }

  // Re-deduplicate names after corrections
  if (fixed > 0) {
    const allNames = new Map<string, ChatEntry[]>();
    for (const chat of chats) {
      // Strip existing dedup suffixes like " (1)", " (2)"
      const baseName = chat.name.replace(/ \(\d+\)$/, "");
      if (!allNames.has(baseName)) allNames.set(baseName, []);
      allNames.get(baseName)!.push(chat);
    }
    for (const [baseName, entries] of allNames) {
      if (entries.length > 1) {
        entries.forEach((e, i) => { e.name = `${baseName} (${i + 1})`; });
      } else {
        entries[0].name = baseName;
      }
    }
    log(`Fixed ${fixed} chat names from sidebar`);
  }
}

/**
 * Legacy sidebar scrolling discovery (fallback if Graph API fails).
 */
async function discoverChatsSidebar(page: Page): Promise<ChatEntry[]> {
  // Make sure we're on the Chat view
  try {
    const chatNavButton = page.locator(
      '[data-tid="app-bar-chat-button"], [data-tid="chat-tab"], button[aria-label*="Chat" i]'
    );
    if (await chatNavButton.first().isVisible({ timeout: 3000 })) {
      await chatNavButton.first().click();
      await sleep(2000);
    }
  } catch {
    // May already be on chat view
  }

  await page
    .waitForSelector("[role='tree']", { timeout: 15000 })
    .catch(() => {
      warn("Could not find chat tree container.");
    });

  log("Scrolling chat list to load all chats...");

  const chats: ChatEntry[] = await page.evaluate(`(async () => {
    var sleep = function(ms) { return new Promise(function(r) { setTimeout(r, ms); }); };

    var tree = document.querySelector("[role='tree']");
    if (!tree) return [];

    var scrollContainer = null;
    var el = tree;
    while (el && el !== document.documentElement) {
      var style = getComputedStyle(el);
      var ov = style.overflowY;
      if ((ov === 'auto' || ov === 'scroll') && el.scrollHeight > el.clientHeight) {
        scrollContainer = el;
        break;
      }
      el = el.parentElement;
    }
    if (!scrollContainer && tree.scrollHeight > tree.clientHeight) {
      scrollContainer = tree;
    }

    var seen = new Map();

    var collectVisible = function() {
      var items = document.querySelectorAll("[role='tree'] [role='group'] > [role='treeitem']");
      items.forEach(function(item) {
        var value = item.getAttribute('data-fui-tree-item-value') || '';
        if (!value.includes('Conversation')) return;
        if (seen.has(value)) return;

        var spans = item.querySelectorAll('span');
        var name = '';
        for (var i = 0; i < spans.length; i++) {
          var text = spans[i].textContent ? spans[i].textContent.trim() : '';
          if (text.length > 1 && !text.match(/^\\d+:\\d+\\s*(AM|PM)?$/i) && !text.match(/^\\d+\\/\\d+$/)) {
            name = text;
            break;
          }
        }
        if (!name) {
          var firstText = item.textContent ? item.textContent.trim().split('\\n')[0].trim() : '';
          name = firstText || ('Chat_' + (seen.size + 1));
        }
        seen.set(value, { name: name, value: value });
      });
    };

    collectVisible();

    if (scrollContainer) {
      var noNewCount = 0;
      var prevCount = seen.size;

      for (var i = 0; i < 200; i++) {
        scrollContainer.scrollTop += scrollContainer.clientHeight * 0.8;
        await sleep(500);
        collectVisible();

        if (seen.size === prevCount) {
          noNewCount++;
          if (noNewCount >= 5) break;
        } else {
          noNewCount = 0;
        }
        prevCount = seen.size;
      }

      scrollContainer.scrollTop = 0;
    }

    var results = [];
    var idx = 0;
    seen.forEach(function(chat) {
      results.push({ name: chat.name, index: idx++, value: chat.value });
    });
    return results;
  })()`) as ChatEntry[];

  log(`Found ${chats.length} chats (sidebar method)`);
  return chats;
}

// ── Click into a specific chat ────────────────────────────────────────

/**
 * Navigate to a chat and return { success, sidebarName }.
 * sidebarName is the display name shown in the sidebar tree item
 * (useful for resolving 1:1 chats where the API shows the user's own name).
 */
async function navigateToChat(page: Page, chat: ChatEntry): Promise<{ success: boolean; sidebarName?: string }> {
  try {
    // Method 1: Find sidebar tree item by thread ID (works for chatsvc API discovery)
    if (chat.chatId) {
      const result = await page.evaluate((threadId: string) => {
        const allItems = document.querySelectorAll("[data-fui-tree-item-value]");
        for (const item of allItems) {
          const value = item.getAttribute("data-fui-tree-item-value") || "";
          if (value.includes(threadId)) {
            // Extract display name from first meaningful span
            let name = "";
            const spans = item.querySelectorAll("span");
            for (const span of spans) {
              const text = (span.textContent || "").trim();
              if (text.length > 1 && !/^\d+:\d+\s*(AM|PM)?$/i.test(text) && !/^\d+\/\d+$/.test(text)) {
                name = text;
                break;
              }
            }
            (item as HTMLElement).scrollIntoView({ block: "center" });
            (item as HTMLElement).click();
            return { clicked: true, name };
          }
        }
        return { clicked: false, name: "" };
      }, chat.chatId);

      if (result.clicked) {
        await page
          .waitForSelector("#chat-pane-list", { timeout: 10000 })
          .catch(() => null);
        await sleep(2000);
        return { success: true, sidebarName: result.name || undefined };
      }

      // Tree item not found in currently visible items.
      // If we already know this chat isn't in the sidebar, skip the expensive scroll.
      if (chat.chatId && sidebarThreadIds.size > 0 && !sidebarThreadIds.has(chat.chatId)) {
        return { success: false };
      }

      // Try scrolling the sidebar to load more items, then retry.
      const scrollResult = await page.evaluate(`(async function() {
        var sleep = function(ms) { return new Promise(function(r) { setTimeout(r, ms); }); };
        var threadId = ${JSON.stringify(chat.chatId)};

        var tree = document.querySelector("[role='tree']");
        if (!tree) return { clicked: false, name: "" };

        var scrollContainer = null;
        var el = tree;
        while (el && el !== document.documentElement) {
          var style = getComputedStyle(el);
          if ((style.overflowY === 'auto' || style.overflowY === 'scroll') && el.scrollHeight > el.clientHeight) {
            scrollContainer = el;
            break;
          }
          el = el.parentElement;
        }
        if (!scrollContainer) return { clicked: false, name: "" };

        for (var i = 0; i < 100; i++) {
          scrollContainer.scrollTop += scrollContainer.clientHeight * 0.8;
          await sleep(300);

          var items = document.querySelectorAll("[data-fui-tree-item-value]");
          for (var j = 0; j < items.length; j++) {
            var value = items[j].getAttribute("data-fui-tree-item-value") || "";
            if (value.includes(threadId)) {
              // Extract display name
              var name = "";
              var spans = items[j].querySelectorAll("span");
              for (var k = 0; k < spans.length; k++) {
                var text = (spans[k].textContent || "").trim();
                if (text.length > 1 && !/^\\d+:\\d+\\s*(AM|PM)?$/i.test(text) && !/^\\d+\\/\\d+$/.test(text)) {
                  name = text;
                  break;
                }
              }
              items[j].scrollIntoView({ block: "center" });
              items[j].click();
              scrollContainer.scrollTop = 0;
              return { clicked: true, name: name };
            }
          }
        }

        scrollContainer.scrollTop = 0;
        return { clicked: false, name: "" };
      })()`) as { clicked: boolean; name: string };

      if (scrollResult.clicked) {
        await page
          .waitForSelector("#chat-pane-list", { timeout: 10000 })
          .catch(() => null);
        await sleep(2000);
        return { success: true, sidebarName: scrollResult.name || undefined };
      }
    }

    // Method 2: Legacy sidebar click by tree item value
    if (chat.value) {
      const clicked = await page.evaluate(`(function() {
        var value = ${JSON.stringify(chat.value)};
        var item = document.querySelector('[data-fui-tree-item-value="' + CSS.escape(value) + '"]');
        if (item) {
          item.scrollIntoView({ block: 'center' });
          item.click();
          return true;
        }
        return false;
      })()`);

      if (clicked) {
        await page
          .waitForSelector("#chat-pane-list", { timeout: 10000 })
          .catch(() => null);
        await sleep(2000);
        return { success: true };
      }
    }

    return { success: false };
  } catch (e: any) {
    warn(`Failed to navigate to chat "${chat.name}": ${e.message}`);
    return { success: false };
  }
}

// ── Extract messages from current chat ────────────────────────────────

async function extractCurrentChat(
  page: Page,
  days: number,
  timeoutMs: number = 120000,
  scrollDown: boolean = false,
  sinceTimestamp?: string
): Promise<ExtractionResult> {
  const script = scrollDown ? getScrollDownExtractionScript() : getExtractionScript(days, sinceTimestamp);

  // Forward browser console.log to terminal for progress visibility
  const consoleHandler = (msg: any) => {
    const text = msg.text();
    if (text.startsWith("[extract]") || text.startsWith("[scroll-down]")) {
      log(text);
    }
  };
  page.on("console", consoleHandler);

  try {
    // Race the extraction against a timeout
    const result = await Promise.race([
      page.evaluate(script) as Promise<ExtractionResult>,
      new Promise<ExtractionResult>((_, reject) =>
        setTimeout(() => reject(new Error(`Extraction timed out after ${timeoutMs / 1000}s`)), timeoutMs)
      ),
    ]);
    return result;
  } catch (e: any) {
    return { messages: [], error: `Evaluation failed: ${e.message}` };
  } finally {
    page.removeListener("console", consoleHandler);
  }
}

// ── Quick-check: skip unchanged chats without full extraction ─────────

/**
 * Read the newest visible message timestamp from the currently open chat.
 * Teams renders recent messages at the bottom by default — no scrolling needed.
 * Returns an ISO timestamp string, or null if none found.
 */
async function getNewestVisibleTimestamp(page: Page): Promise<string | null> {
  return page.evaluate(`(function() {
    var list = document.getElementById('chat-pane-list');
    if (!list) return null;
    var timestamps = list.querySelectorAll('[id^="timestamp-"]');
    if (timestamps.length === 0) return null;
    var newest = null;
    for (var i = 0; i < timestamps.length; i++) {
      var dt = timestamps[i].getAttribute('datetime');
      if (dt && (!newest || dt > newest)) newest = dt;
    }
    return newest;
  })()`);
}

/**
 * Quick-check whether a chat has new messages since last export.
 * Compares the newest visible DOM timestamp with the file's last-message marker.
 * Returns true if the chat can be skipped (no new messages).
 */
async function canSkipChat(page: Page, chatName: string, outputDir: string): Promise<boolean> {
  const filename = sanitizeFilename(chatName) + ".md";
  const filePath = join(outputDir, filename);

  const lastExported = await getLatestTimestamp(filePath);
  if (!lastExported) return false; // No existing file — must extract

  const newestVisible = await getNewestVisibleTimestamp(page);
  if (!newestVisible) return false; // Can't determine — extract to be safe

  const newestDate = new Date(newestVisible);
  // Skip if the newest visible message is at or before the last exported one
  return newestDate <= lastExported;
}

// ── Main ──────────────────────────────────────────────────────────────

async function main() {
  log("MS Teams Chat Exporter");
  log("=".repeat(50));
  log(`Output:   ${OUTPUT_DIR}`);
  log(`Days:     ${DAYS === 0 ? "all history" : DAYS}`);
  log(`Scroll:   ${SCROLL_DOWN ? "DOWN (from current position)" : "UP (default)"}`);
  log(`Region:   ${REGION}`);
  log(`Dry run:  ${DRY_RUN}`);
  if (CHAT_FILTER) log(`Filter:   "${CHAT_FILTER}"`);
  log("");

  const { page } = await connectToChrome();

  // Detect current user's display name (for name-fix filter)
  if (args.user) {
    currentUserName = args.user;
    log(`User (from --user flag): ${currentUserName}`);
  } else {
    currentUserName = await detectCurrentUser(page);
    if (currentUserName) {
      log(`User (auto-detected): ${currentUserName}`);
    } else {
      warn("Could not auto-detect user name. Use --user flag to enable name-fix filter.");
    }
  }

  // Discover all chats
  let chats = await discoverChats(page);

  if (chats.length === 0) {
    console.error(
      "No chats found in the sidebar.\n" +
        "Make sure you're on the Teams Chat view with chats visible."
    );
    process.exit(1);
  }

  // Apply filter if specified
  if (CHAT_FILTER) {
    const filter = CHAT_FILTER.toLowerCase();
    chats = chats.filter((c) => c.name.toLowerCase().includes(filter));
    if (chats.length === 0) {
      console.error(`No chats matching "${CHAT_FILTER}" found.`);
      process.exit(1);
    }
    log(`Filtered to ${chats.length} chats matching "${CHAT_FILTER}"`);
  }

  // Dry run: just list chats
  if (DRY_RUN) {
    log("\nChats found:");
    for (const chat of chats) {
      const filename = sanitizeFilename(chat.name) + ".md";
      const flags = [
        chat.chatType || "",
        chat.isHidden ? "HIDDEN" : "",
      ].filter(Boolean).join(", ");
      const suffix = flags ? `  (${flags})` : "";
      console.log(`  ${chat.index + 1}. ${chat.name}  ->  ${filename}${suffix}`);
    }
    const hiddenCount = chats.filter((c) => c.isHidden).length;
    log(`\nTotal: ${chats.length} chats${hiddenCount ? ` (${hiddenCount} hidden in sidebar)` : ""}`);

    // Also show retention analysis on dry-run (from existing exports)
    const retention = await analyzeRetention(OUTPUT_DIR);
    if (retention.chatsWithDates > 0) {
      log("");
      log("=".repeat(50));
      log(formatRetentionSummary(retention));
    }
    return;
  }

  // Export each chat
  let exported = 0;
  let skipped = 0;
  let quickSkipped = 0;
  let failed = 0;
  const totalChats = chats.length;

  for (let i = 0; i < totalChats; i++) {
    const chat = chats[i];
    const progress = `[${i + 1}/${totalChats}]`;
    log(`${progress} Exporting: ${chat.name}...`);

    // Navigate to the chat
    const navResult = await navigateToChat(page, chat);
    if (!navResult.success) {
      warn(`${progress} Skipping "${chat.name}" — could not navigate`);
      failed++;
      continue;
    }

    // Fix chat name from sidebar if it shows the user's own name
    if (navResult.sidebarName && currentUserName) {
      const nameLower = chat.name.toLowerCase();
      if (nameLower.includes(currentUserName.toLowerCase()) || nameLower.startsWith("chat_")) {
        log(`${progress} Name resolved: "${chat.name}" -> "${navResult.sidebarName}"`);
        chat.name = navResult.sidebarName;
      }
    }

    // Quick-check: skip unchanged chats without expensive full extraction
    if (await canSkipChat(page, chat.name, OUTPUT_DIR)) {
      log(`${progress} "${chat.name}": unchanged (quick-check), skipping`);
      quickSkipped++;
      continue;
    }

    // Look up last-exported timestamp to limit scrolling to only new messages
    const filePath = join(OUTPUT_DIR, sanitizeFilename(chat.name) + ".md");
    const lastExported = await getLatestTimestamp(filePath);
    const sinceTs = lastExported ? lastExported.toISOString() : undefined;

    // Extract messages (scrolls only to sinceTs, not full history)
    const result = await extractCurrentChat(page, DAYS, TIMEOUT_MS, SCROLL_DOWN, sinceTs);

    if (!result) {
      warn(`${progress} "${chat.name}": extraction returned null`);
      failed++;
      continue;
    }

    if (result.error) {
      warn(`${progress} "${chat.name}": ${result.error}`);
      if (result.messages.length === 0) {
        failed++;
        continue;
      }
    }

    if (result.messages.length === 0) {
      log(`${progress} "${chat.name}": no messages found, skipping`);
      skipped++;
      continue;
    }

    // Save to file
    const newCount = await saveChat(chat.name, result.messages, OUTPUT_DIR);

    if (newCount === 0) {
      log(`${progress} "${chat.name}": ${result.messages.length} messages, all already exported`);
      skipped++;
    } else {
      log(
        `${progress} "${chat.name}": ${newCount} new messages saved (${result.messages.length} total)`
      );
      exported++;
    }

    // Small delay between chats to avoid overwhelming Teams
    await sleep(1000);
  }

  // Summary
  log("");
  log("=".repeat(50));
  log("Export complete!");
  log(`  Exported: ${exported} chats with new messages`);
  log(`  Skipped:  ${skipped} chats (no new messages after extraction)`);
  log(`  Quick-skipped: ${quickSkipped} chats (unchanged, no extraction needed)`);
  log(`  Failed:   ${failed} chats`);
  log(`  Output:   ${OUTPUT_DIR}`);

  // Retention policy analysis
  log("");
  log("=".repeat(50));
  log("Analyzing retention policy...");
  const retention = await analyzeRetention(OUTPUT_DIR);
  log(formatRetentionSummary(retention));
}

main().catch((e) => {
  console.error("Fatal error:", e.message);
  process.exit(1);
});
