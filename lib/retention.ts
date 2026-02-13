/**
 * Chat visibility and retention analysis for exported Teams chats.
 *
 * IMPORTANT DISTINCTION:
 * Teams has TWO separate mechanisms that limit what messages you see:
 *
 * 1. SIDEBAR VISIBILITY — Teams dynamically hides inactive chats from the
 *    sidebar. There is no documented hard limit; visibility is based on
 *    "recent activity, pinned chats (up to 15), and usage patterns."
 *    Hidden chats still exist and can be found via search.
 *    Ref: https://learn.microsoft.com/en-us/answers/questions/5496974
 *
 * 2. RETENTION POLICY — Org-configured via Microsoft Purview. Actually
 *    deletes messages after the retention period. Default is keep forever.
 *    Common tiers: 30, 60, 90, 180, 365, 730 days, or 1/3/5/7 years.
 *    Ref: https://learn.microsoft.com/en-us/purview/retention-policies-teams
 *
 * This module analyzes EXPORTED data only, which is limited to sidebar-visible
 * chats. It CANNOT determine the actual retention policy from this data alone,
 * because the sidebar visibility cutoff may hide older chats before retention
 * would delete them. The analysis reports what we CAN observe and flags what
 * we CANNOT determine.
 */

import { readFile } from "fs/promises";
import { join } from "path";
import { globSync } from "fs";

// Known Microsoft Teams retention policy tiers (in days)
const MS_RETENTION_TIERS = [
  { days: 30, label: "30 days" },
  { days: 60, label: "60 days" },
  { days: 90, label: "90 days" },
  { days: 120, label: "120 days" },
  { days: 180, label: "180 days" },
  { days: 365, label: "1 year (365 days)" },
  { days: 730, label: "2 years (730 days)" },
  { days: 1095, label: "3 years" },
  { days: 1825, label: "5 years" },
  { days: 2555, label: "7 years" },
];

export interface ChatDateRange {
  name: string;
  oldest: Date;
  newest: Date;
  ageDays: number;
  messageCount: number;
}

export interface VisibilityAnalysis {
  chatsAnalyzed: number;
  chatsWithDates: number;
  oldestMessage: { date: Date; chat: string } | null;
  newestMessage: { date: Date; chat: string } | null;
  medianOldestAgeDays: number;
  maxOldestAgeDays: number;
  sidebarWindowDays: number;
  explanation: string;
  nearestRetentionTier: { days: number; label: string } | null;
  chatRanges: ChatDateRange[];
}

/**
 * Parse DD.MM.YYYY from the markdown timestamp format: [DD.MM.YYYY, HH:MM]
 */
function parseGermanDate(dateStr: string): Date | null {
  const m = dateStr.match(/(\d{2})\.(\d{2})\.(\d{4}),\s*(\d{2}):(\d{2})/);
  if (!m) return null;
  const [, day, month, year, hour, minute] = m;
  return new Date(
    parseInt(year),
    parseInt(month) - 1,
    parseInt(day),
    parseInt(hour),
    parseInt(minute)
  );
}

/**
 * Extract the oldest and newest message timestamps from a chat Markdown file.
 */
async function extractDateRange(
  filePath: string
): Promise<{ oldest: Date; newest: Date; count: number } | null> {
  const content = await readFile(filePath, "utf-8");

  const timestampRegex = /\[(\d{2}\.\d{2}\.\d{4},\s*\d{2}:\d{2})\]/g;
  let oldest: Date | null = null;
  let newest: Date | null = null;
  let count = 0;

  let match;
  while ((match = timestampRegex.exec(content)) !== null) {
    const date = parseGermanDate(match[1]);
    if (!date || isNaN(date.getTime())) continue;
    count++;
    if (!oldest || date < oldest) oldest = date;
    if (!newest || date > newest) newest = date;
  }

  if (!oldest || !newest || count === 0) return null;
  return { oldest, newest, count };
}

/**
 * Find the nearest MS retention tier to a given number of days.
 */
function findNearestTier(
  days: number
): { days: number; label: string } | null {
  let best: { days: number; label: string; distance: number } | null = null;
  for (const tier of MS_RETENTION_TIERS) {
    const distance = Math.abs(tier.days - days);
    if (!best || distance < best.distance) {
      best = { ...tier, distance };
    }
  }
  return best ? { days: best.days, label: best.label } : null;
}

/**
 * Analyze all exported chat files for visibility and date range statistics.
 */
export async function analyzeRetention(
  outputDir: string
): Promise<VisibilityAnalysis> {
  const files = globSync(join(outputDir, "*.md"));
  const now = new Date();
  const chatRanges: ChatDateRange[] = [];

  for (const filePath of files) {
    const name = filePath
      .split("/")
      .pop()!
      .replace(/\.md$/, "")
      .replace(/_/g, " ");
    const range = await extractDateRange(filePath);
    if (!range) continue;

    const ageDays = Math.floor(
      (now.getTime() - range.oldest.getTime()) / (1000 * 60 * 60 * 24)
    );
    chatRanges.push({
      name,
      oldest: range.oldest,
      newest: range.newest,
      ageDays,
      messageCount: range.count,
    });
  }

  chatRanges.sort((a, b) => b.ageDays - a.ageDays);

  if (chatRanges.length === 0) {
    return {
      chatsAnalyzed: files.length,
      chatsWithDates: 0,
      oldestMessage: null,
      newestMessage: null,
      medianOldestAgeDays: 0,
      maxOldestAgeDays: 0,
      sidebarWindowDays: 0,
      explanation: "No chat files with parseable timestamps found.",
      nearestRetentionTier: null,
      chatRanges: [],
    };
  }

  const ages = chatRanges.map((c) => c.ageDays);
  const maxAge = ages[0];
  const medianAge = ages[Math.floor(ages.length / 2)];

  const oldestChat = chatRanges[0];
  const newestChat = chatRanges.reduce((a, b) =>
    a.newest > b.newest ? a : b
  );

  const nearest = findNearestTier(maxAge);

  // Build explanation acknowledging the two-layer limitation
  const clusterThresholdDays = 14;
  const oldestCluster = chatRanges.filter(
    (c) => c.ageDays >= maxAge - clusterThresholdDays
  );

  let explanation: string;

  if (oldestCluster.length >= 3) {
    const clusterDate = oldestCluster[0].oldest.toLocaleDateString("de-DE", {
      year: "numeric",
      month: "long",
      day: "numeric",
    });
    explanation =
      `${oldestCluster.length} chats cluster around ${clusterDate} (~${maxAge} days ago), ` +
      `suggesting this is your account start date or the Teams sidebar visibility window. ` +
      `Teams hides inactive chats from the sidebar — older conversations may exist ` +
      `but weren't exported because they weren't visible. ` +
      `To find hidden chats: search for a contact name in Teams, then re-export.`;
  } else {
    explanation =
      `Oldest exported messages are ${maxAge} days old. ` +
      `This reflects what was visible in the Teams sidebar at export time. ` +
      `Teams hides inactive chats — there may be older conversations accessible via search.`;
  }

  return {
    chatsAnalyzed: files.length,
    chatsWithDates: chatRanges.length,
    oldestMessage: { date: oldestChat.oldest, chat: oldestChat.name },
    newestMessage: { date: newestChat.newest, chat: newestChat.name },
    medianOldestAgeDays: medianAge,
    maxOldestAgeDays: maxAge,
    sidebarWindowDays: maxAge,
    explanation,
    nearestRetentionTier: nearest,
    chatRanges,
  };
}

/**
 * Format the analysis as a human-readable summary for terminal output.
 */
export function formatRetentionSummary(analysis: VisibilityAnalysis): string {
  if (!analysis.oldestMessage) {
    return "  Could not analyze (no timestamps found)";
  }

  const lines: string[] = [];
  lines.push("");
  lines.push("  Chat Visibility & Retention Analysis");
  lines.push("  " + "-".repeat(44));
  lines.push(`  Chats exported:     ${analysis.chatsWithDates}/${analysis.chatsAnalyzed}`);
  lines.push(
    `  Oldest message:     ${analysis.oldestMessage.date.toLocaleDateString("de-DE")} (${analysis.maxOldestAgeDays} days ago)`
  );
  lines.push(`    in chat:          ${analysis.oldestMessage.chat}`);
  lines.push(`  Median oldest age:  ${analysis.medianOldestAgeDays} days`);
  lines.push(`  Sidebar window:     ~${analysis.sidebarWindowDays} days`);
  lines.push("");
  lines.push(`  ${analysis.explanation}`);

  // Top 5 oldest chats
  lines.push("");
  lines.push("  Top 5 oldest chats:");
  for (const chat of analysis.chatRanges.slice(0, 5)) {
    lines.push(
      `    ${chat.ageDays}d  ${chat.oldest.toLocaleDateString("de-DE")}  ${chat.name.slice(0, 50)}`
    );
  }

  // Retention policy guidance
  lines.push("");
  lines.push("  NOTE: This analysis covers sidebar-visible chats only.");
  lines.push("  Teams hides inactive chats — search for contacts to resurface them.");
  lines.push("  Actual retention policy is set in Microsoft Purview admin center.");
  lines.push("  Ref: https://learn.microsoft.com/en-us/purview/retention-policies-teams");

  return lines.join("\n");
}
