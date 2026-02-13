# Teams Webapp Exporter

Bulk export Microsoft Teams chats to Markdown files via Playwright browser automation.

Based on [microsoft-teams-chat-extractor](https://github.com/ingo/microsoft-teams-chat-extractor) by Ingo Muschenetz. Rewritten as a Playwright-based CLI for automated bulk export with incremental sync, quick-check optimization, and scroll cutoff.

## Features

- **Bulk export** — exports all chats (1:1, group, meetings) in one run
- **Incremental sync** — only downloads new messages since last export
- **Quick-check** — skips unchanged chats without full DOM extraction
- **Scroll cutoff** — stops scrolling when it reaches already-exported messages
- **Hidden chat discovery** — finds chats hidden from the Teams sidebar via internal API
- **Retention analysis** — reports chat age distribution and sidebar visibility window

## Prerequisites

- **Node.js** (v18+)
- **Google Chrome** launched with remote debugging enabled

### Launch Chrome for debugging

Chrome must be started with three flags:

```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \
  --remote-debugging-port=9222 \
  --remote-allow-origins='*' \
  --user-data-dir="$HOME/Library/Application Support/Google/Chrome-Debug" &
```

Then navigate to https://teams.microsoft.com and sign in.

> **Note:** `--user-data-dir` must point to a separate profile directory — Chrome refuses to enable DevTools on the default profile. The `--remote-allow-origins='*'` flag must be quoted in zsh.

## Installation

```bash
git clone <this-repo>
cd teams-webapp-exporter
npm install
```

## Usage

```bash
npx tsx export-all.ts [options]
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--days <n>` | `0` (all) | How far back to scroll per chat (0 = full history) |
| `--output <dir>` | `./output` | Target directory for Markdown files |
| `--chat <name>` | — | Export only chats matching this substring |
| `--dry-run` | `false` | List all chats without exporting |
| `--port <n>` | `9222` | Chrome remote debugging port |
| `--timeout <s>` | `120` | Per-chat extraction timeout in seconds |
| `--scroll-down` | `false` | Scroll down from current position instead of up |
| `--region <r>` | `emea` | Teams chatsvc API region (`emea`, `amer`, `apac`) |
| `--user <name>` | auto-detect | Your display name (for resolving 1:1 chat names) |

### Examples

List all chats without exporting:

```bash
npx tsx export-all.ts --dry-run
```

Export all chats from the last 30 days:

```bash
npx tsx export-all.ts --days 30
```

Export a specific chat:

```bash
npx tsx export-all.ts --chat "Project Updates"
```

Export to a custom directory with a different region:

```bash
npx tsx export-all.ts --output ./exports --region amer
```

## How incremental export works

Each Markdown file ends with a `<!-- last-message: ... -->` comment containing the timestamp of the newest exported message. On subsequent runs:

1. **Quick-check** — reads the newest visible DOM timestamp and compares it to the file marker. If unchanged, the chat is skipped entirely (no scrolling).
2. **Scroll cutoff** — if new messages exist, scrolling stops when it reaches the last-exported timestamp instead of loading full history.
3. **Append** — only new messages are appended to the existing file with a separator.

## Output format

Each chat produces a Markdown file named after the chat (sanitized). Messages are grouped by date with author headers:

```markdown
# Chat Name

Exported: 2026-02-13T15:00:00.000Z

## Thursday, 13 February 2026

**Alice** [2026-02-13T09:15:00Z]:
Good morning! Here's the update...

**Bob** [2026-02-13T09:20:00Z]:
Thanks, looks good.
```

## Testing & Linting

```bash
npm test          # Run all tests (Vitest)
npm run lint      # Run ESLint
```

53 tests across 4 modules: message extraction, Markdown conversion, retention analysis, and Graph API chat discovery. Tests use temp directories for file I/O and mock `fetch` for API calls — no network or browser needed.

## License

MIT — see [LICENSE](LICENSE). Original work copyright (c) 2021-2026 Ingo Muschenetz.
