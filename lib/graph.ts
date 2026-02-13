/**
 * Teams chat discovery via the internal chatsvc API.
 *
 * Uses the ic3.teams.office.com access token from the MSAL cache in
 * localStorage to call the Teams internal chat service, which returns
 * ALL conversations including hidden sidebar ones.
 *
 * API: GET /api/chatsvc/{region}/v1/users/ME/conversations?view=mychats
 * Auth: Bearer token for audience https://ic3.teams.office.com
 *
 * This replaces the previous Graph API approach which required
 * Chat.ReadBasic scope (not pre-authorized for the Teams first-party app).
 */

import type { Page } from "playwright-core";

export interface ChatEntry {
  name: string;
  index: number;
  chatId: string;
  webUrl: string;
  chatType: "chat" | "meeting" | "topic" | "space" | string;
  isHidden: boolean;
}

/**
 * Extract the ic3.teams.office.com Bearer token from localStorage.
 */
export async function getToken(page: Page): Promise<string> {
  const token = await page.evaluate(() => {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i) || "";
      const val = localStorage.getItem(key) || "";
      if (!key.includes("accesstoken")) continue;

      try {
        const parsed = JSON.parse(val);
        const secret = parsed.secret || "";
        if (!secret.startsWith("eyJ")) continue;

        // Decode JWT to check audience
        const payloadB64 = secret.split(".")[1];
        const payload = JSON.parse(
          atob(payloadB64.replace(/-/g, "+").replace(/_/g, "/"))
        );

        if (payload.aud === "https://ic3.teams.office.com") {
          // Check token is not expired
          if (payload.exp && payload.exp * 1000 > Date.now()) {
            return secret;
          }
        }
      } catch {}
    }
    return "";
  });

  if (!token) {
    throw new Error(
      "No valid ic3.teams.office.com token found in localStorage. " +
      "Make sure Teams is open and you are signed in."
    );
  }

  return token;
}

/**
 * Fetch all conversations from the Teams internal chatsvc API.
 *
 * Returns conversations filtered to chat, meeting, and topic types
 * (excludes internal streams like notes, annotations, call logs, etc).
 *
 * @param token  Bearer token for ic3.teams.office.com
 * @param region chatsvc region (e.g. "emea", "amer", "apac")
 */
export async function fetchAllChats(token: string, region: string): Promise<ChatEntry[]> {
  const url = `https://teams.microsoft.com/api/chatsvc/${region}/v1/users/ME/conversations?view=mychats&pageSize=500`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`chatsvc API error ${response.status}: ${text.slice(0, 200)}`);
  }

  const data = await response.json();
  const conversations = data.conversations || [];

  // Filter to exportable conversation types
  const exportableTypes = new Set(["chat", "meeting", "topic"]);
  const filtered = conversations.filter((conv: any) => {
    const threadType = conv.threadProperties?.threadType || "";
    return exportableTypes.has(threadType);
  });

  // Sort by last activity (most recent first)
  filtered.sort((a: any, b: any) => {
    const aTime = a.properties?.lastimreceivedtime || a.threadProperties?.createdat || "";
    const bTime = b.properties?.lastimreceivedtime || b.threadProperties?.createdat || "";
    return bTime.localeCompare(aTime);
  });

  // Convert to ChatEntry format with deduplicated names
  const nameCount = new Map<string, number>();

  const entries = filtered.map((conv: any, index: number) => {
    const threadType = conv.threadProperties?.threadType || "chat";
    const topic = conv.threadProperties?.topic || "";
    const threadId = conv.id || "";
    // Detect 1:1 chats: case-insensitive flag OR thread ID pattern
    const uniqueFlag = (conv.threadProperties?.uniquerosterthread || "").toLowerCase() === "true";
    const unqThreadId = threadId.includes("@unq.gbl.spaces");
    const isOneOnOne = uniqueFlag || unqThreadId;

    // Resolve display name
    let name = "";

    if (topic) {
      name = topic;
    } else if (isOneOnOne) {
      // 1:1 chat: use the other person's display name
      name = conv.lastMessage?.fromDisplayNameInToken
        || conv.lastMessage?.imdisplayname
        || `Chat_${index + 1}`;
    } else if (conv.lastMessage?.fromDisplayNameInToken) {
      // Group chat without topic: show last sender with "Group" prefix
      name = `Group - ${conv.lastMessage.fromDisplayNameInToken}`;
    } else {
      name = `Chat_${index + 1}`;
    }

    // Deduplicate names by appending a counter
    const count = nameCount.get(name) || 0;
    nameCount.set(name, count + 1);
    if (count > 0) {
      name = `${name} (${count + 1})`;
    }

    // Determine if hidden (meetings have explicit `hidden` threadProperty)
    const isHidden = conv.threadProperties?.hidden === "true" || false;

    // Construct Teams deep link for navigation (fallback — sidebar click preferred)
    const webUrl = `https://teams.microsoft.com/l/chat/${encodeURIComponent(threadId)}/0`;

    return {
      name,
      index,
      chatId: threadId,
      webUrl,
      chatType: threadType,
      isHidden,
    };
  });

  // Fix first occurrences of duplicated names — add (1) suffix retroactively
  for (const [baseName, count] of nameCount) {
    if (count > 1) {
      const first = entries.find((e) => e.name === baseName);
      if (first) {
        first.name = `${baseName} (1)`;
      }
    }
  }

  return entries;
}
