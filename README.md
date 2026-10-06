# Hey365

**日本語版は [README.ja.md](README.ja.md) をご覧ください。**

Free for personal and noncommercial use. **Commercial use requires prior consent and ¥150 per seat** — see [License](#license).

Hey365 is a Model Context Protocol (MCP) server that answers one question:

> **"What do I actually owe people right now?"**

Say **"Hey365"** in any MCP client and it scans the last 36 hours of your Outlook mail, Teams chats, Teams channels and calendar, works out which conversations are *still waiting on you*, explains why, and drafts the replies. Nothing is ever sent unless you explicitly say so.

```
Hey365 チェック完了

過去36時間で、あなたの対応が必要と思われるものが 2件あります。

🔴 1. 田中 太郎 / Teams Channel
00:48（1日前）

内容:
承知しました。調整頂けましたら会議リンクの送付をお願いいたします！

返信が必要な理由:
「調整頂けましたら会議リンクの送付をお願いいたします！」のように、依頼を受けています。

返信案:
ご連絡ありがとうございます。10/6（火）10:00-10:30 で確定しましたので、
本日中に会議リンクをお送りします。

------------------------------------------------

送信したいものがあれば、「1を送って」「全部送って」のように指示してください。
```

---

## Why it exists

Inbox rules and unread counts do not know the difference between "FYI, sharing this" and "I am blocked waiting for your approval". Hey365 answers that by reconstructing every conversation thread and asking three questions in order:

1. **Did I already answer?** If my last message is newer than their last message, the thread is done. No keyword matching — pure thread chronology.
2. **Is this for me?** Automated notifications, bots, mails where I am only on CC, pure FYI broadcasts, reaction-only replies and closed conversations are dropped.
3. **What is being asked?** Approval requests, decisions, direct questions, scheduling, follow-up nudges and assigned actions are detected in both Japanese and English, then weighted and scored.

What survives all three is shown to you, newest signal first, with the exact sentence that triggered it.

## How it works

```mermaid
flowchart LR
    A[Outlook mail<br/>Teams chats<br/>Teams channels<br/>Calendar] --> B[Work IQ CLI<br/>Microsoft Graph]
    B --> C[Thread reconstruction]
    C --> D[Triage<br/>signals + scoring]
    D --> E[Dedupe<br/>same topic across channels]
    E --> F[Reply drafts<br/>M365 Copilot]
    F --> G[You approve]
    G --> H[Send<br/>reply in the same thread]
```

Hey365 talks to Microsoft 365 through the [Work IQ CLI](https://www.npmjs.com/package/@microsoft/workiq), which is an authenticated Microsoft Graph proxy. Reply drafts are written by Microsoft 365 Copilot via `workiq ask`, so **Hey365 needs no API keys of its own**.

## Requirements

| Requirement | Notes |
|---|---|
| Node.js 20+ | `node --version` |
| A Microsoft 365 work account | Outlook + Teams |
| Microsoft 365 Copilot license | Required for reply drafting and session summaries |
| `@microsoft/workiq` | Installed and authenticated automatically by `hey365 setup` |

## Install

```bash
git clone https://github.com/monoqlo78/hey365.git
cd hey365
npm install
npm run build
```

Then connect Microsoft 365 (installs Work IQ if missing, disables broker auth, opens a browser sign-in):

```bash
node dist/index.js setup
node dist/index.js health     # should print "hey365": "ok"
```

## Register with your MCP client

One command writes the correct config for you, merging into any existing file:

```bash
node dist/index.js install copilot-cli        # GitHub Copilot CLI
node dist/index.js install vscode             # VS Code (GitHub Copilot)
node dist/index.js install claude-code        # Claude Code
node dist/index.js install claude-desktop     # Claude Desktop
node dist/index.js install codex              # OpenAI Codex CLI
node dist/index.js install cursor             # Cursor
node dist/index.js install windsurf           # Windsurf
node dist/index.js install scout              # Scout

node dist/index.js install --all              # every client above
node dist/index.js install codex --dry-run    # preview without writing
node dist/index.js install cursor --path ./.cursor/mcp.json
```

<details>
<summary>Manual configuration</summary>

Most clients (`~/.copilot/mcp-config.json`, `~/.claude.json`, `~/.cursor/mcp.json`, …):

```json
{
  "mcpServers": {
    "hey365": {
      "command": "node",
      "args": ["/absolute/path/to/hey365/dist/index.js", "mcp"]
    }
  }
}
```

VS Code (`.vscode/mcp.json`):

```json
{
  "servers": {
    "hey365": {
      "type": "stdio",
      "command": "node",
      "args": ["/absolute/path/to/hey365/dist/index.js", "mcp"]
    }
  }
}
```

Codex CLI (`~/.codex/config.toml`):

```toml
[mcp_servers.hey365]
command = "node"
args = ["/absolute/path/to/hey365/dist/index.js", "mcp"]
```

</details>

## Usage

Just talk to your assistant.

| You say | What happens |
|---|---|
| `Hey365` | Triage the last 36 hours and draft replies |
| `Hey365 24時間` / `Hey365 last 12 hours` | Change the window |
| `Hey365 for the last 3 business days` | Cover three business days, counting today |
| `Hey365 as of September 18, 3 business days` | Reach back from a past date — useful after leave |
| `Hey365 メールだけ` / `Teams only` | Limit the sources |
| `What's my day look like?` / `morning digest` | Today's meetings, deadlines, stale items and unanswered threads on one screen |
| `Where did the X discussion end up?` | Search Outlook and Teams together and show whose turn it is |
| `1をもう少し柔らかく` | Rewrite draft #1 in a softer tone |
| `3を英語で` | Rewrite draft #3 in English |
| `1を送って` / `1と3を送って` / `全部送って` | Send those replies **(only when you say so)** |
| `1 can wait until tomorrow` | Hide it until the next business day |
| `1 is already handled` | Hide it for good — until a newer message arrives |
| `Give me the digest every morning at 8` | Register a scheduled run (shown for review first) |
| `Hey365 session <会議名>` | Summarise a meeting or conversation with decisions, action items and open questions |

The numbering survives client restarts, so you can triage now and send later.

### Coming back from leave

```
Hey365, as of September 18, three business days
```

`asOf` sets the reference point and `businessDays` sets how far back to reach. The reference day counts as the first day, and weekends and public holidays do not consume one. This is how you decide where to start reading again after a long break.

### Stale alerts

Anything left unanswered past three business days (`HEY365_STALE_AFTER_DAYS`) is marked, promoted in importance and sorted to the top.

### From the terminal

```bash
node dist/index.js triage --hours 36 --limit 10
node dist/index.js triage --business-days 3 --as-of 2026-09-18
node dist/index.js digest                      # today's meetings, deadlines, unanswered
node dist/index.js digest --out ~/hey365.txt   # write to a file, for scheduled runs
node dist/index.js find "Fabric migration"     # search people and topics across both
node dist/index.js mutes                       # what is currently hidden
node dist/index.js schedule add --at 08:30     # print the command only
node dist/index.js schedule add --at 08:30 --apply
node dist/index.js session "Fabric 定例"
node dist/index.js health --deep
node dist/index.js reconnect                   # re-establish the Work IQ sign-in
```

## Tools

| Tool | Input | Purpose |
|---|---|---|
| `hey365` | `hours`, `businessDays`, `asOf`, `sources`, `limit`, `includeDrafts` | Alias of `hey365_triage`, the entry point for "Hey365" |
| `hey365_triage` | `hours` (default 36), `businessDays`, `asOf`, `sources` (`outlook`/`teams`/`meetings`), `limit`, `includeDrafts` | Scan, triage and draft |
| `hey365_digest` | `asOf`, `businessDays`, `limit` | Today's meetings, due items, stale items and unanswered threads on one screen |
| `hey365_find` | `keyword`, `days`, `sources`, `limit` | Search Outlook and Teams by person or topic and show whose turn it is |
| `hey365_session` | `sessionId`, `includeDraft` | Summarise one meeting/thread: decisions, action items, open questions |
| `hey365_draft` | `itemId`, `instruction` | Rewrite a draft ("softer", "shorter", "in English") |
| `hey365_send` | `itemIds`, `confirm` | Send approved replies into the original thread |
| `hey365_snooze` | `itemIds`, `action` (`snooze`/`done`/`unmute`), `businessDays`, `note` | Hide for a while, mark handled, or bring it back |
| `hey365_mutes` | – | List everything currently hidden |
| `hey365_schedule` | `action` (`list`/`add`/`remove`), `job`, `at`, `weekdaysOnly`, `apply` | Register, remove or list scheduled digest runs |
| `hey365_health` | `deep` | Connection, authentication, read/write status |
| `hey365_reconnect` | `force`, `browser` | Re-establish the Work IQ sign-in. Called automatically when a session expires |
| `hey365_setup` | `interactive` | Install + authenticate Work IQ |
| `hey365_install` | `clients`, `dryRun` | Register Hey365 with other MCP clients |

## Safety

- **Send is a separate, explicit step.** Triage and drafting never send anything.
- **What you saw is what is sent.** Each draft carries a hash; `hey365_send` refuses to send text you have not seen.
- **Replies stay in the thread.** Sending uses `reply`, so Hey365 never starts a new conversation or adds new recipients.
- **Drafts never invent facts.** If the thread lacks the information needed to answer, the draft says so and asks you to decide.
- **"Handled" is reversible.** A conversation you marked done comes back automatically once a newer message arrives, because that was only ever a statement about what you had already read.
- **Scheduling does not touch your machine silently.** `hey365_schedule` prints the exact command it would register and only writes it when you pass `apply=true`.
- **Nothing leaves your machine except through Microsoft Graph.** No third-party API, no telemetry. Logs go to stderr with tokens redacted, and local state (`~/.hey365/state.json`) is written with `0600`.

## Environment variables

| Variable | Default | Purpose |
|---|---|---|
| `HEY365_TIMEZONE` | `Asia/Tokyo` | Display timezone |
| `HEY365_VIP` | – | Comma-separated addresses that get an importance boost |
| `HEY365_MY_NAMES` | – | Extra spellings of your name (e.g. `曽我部,Sogabe`) used to detect when someone addresses you by name |
| `HEY365_BUSINESS_DAYS` | `on` | `off` counts the window in wall-clock hours instead of working time |
| `HEY365_STALE_AFTER_DAYS` | `3` | Business days after which an unanswered item is flagged as stale |
| `HEY365_HOLIDAY_CALENDAR` | auto | `jp` forces the Japanese public-holiday calendar, `none` disables it. Defaults to `jp` when the timezone is `Asia/Tokyo` |
| `HEY365_HOLIDAYS` | – | Extra non-working days as comma-separated `YYYY-MM-DD` (company holidays, your own leave) |
| `HEY365_WORKIQ_COMMAND` | auto | Override how the Work IQ CLI is launched |
| `HEY365_WORKIQ_ACCOUNT` | – | Account to use when several are signed in |
| `HEY365_WORKIQ_TIMEOUT_MS` | `120000` | Per-call timeout |
| `HEY365_AUTO_RECONNECT` | `on` | Set to `off`/`0` to disable automatic reconnection when the session expires |
| `HEY365_RECONNECT_TIMEOUT_MS` | `120000` | Time budget for an automatic `workiq auth login` |
| `HEY365_RECONNECT_COOLDOWN_MS` | `60000` | How long to wait before retrying after a failed reconnect |
| `HEY365_LOG_LEVEL` | `info` | `silent`/`error`/`warn`/`info`/`debug` |
| `HEY365_STATE_FILE` | `~/.hey365/state.json` | Where drafts are persisted |

## Troubleshooting

Run `node dist/index.js health` first. Every failure maps to a code with a concrete next step:

| Code | Fix |
|---|---|
| `WORKIQ_NOT_INSTALLED` | `node dist/index.js setup`, or `npm i -g @microsoft/workiq` |
| `WORKIQ_NOT_AUTHENTICATED` / `WORKIQ_AUTH_EXPIRED` | Usually recovers on its own. If it persists, run `node dist/index.js reconnect` (or tell your assistant "reconnect to Work IQ") |
| `WORKIQ_EULA_REQUIRED` | `npx @microsoft/workiq accept-eula` |
| `WORKIQ_ADMIN_CONSENT_REQUIRED` | Ask a tenant admin to run `npx @microsoft/workiq auth consent` |
| `WORKIQ_CONNECTION_ERROR` | Check the network, then retry |
| `WORKIQ_PERMISSION_DENIED` | Some sources are not readable; Hey365 continues with the rest |
| `WORKIQ_WRITE_DISABLED` | Triage and drafts work, sending is blocked by policy |
| `SESSION_NOT_FOUND` | Re-run `hey365_session` with a subject keyword or date |
| `DRAFT_MISMATCH` | The draft changed after you saw it — review and send again |

On Windows, sign-in can fail inside the broker (WAM). `setup` disables it automatically (`workiq config set disableBrokeredAuth=true`) and falls back to browser authentication.

### When the session expires

Work IQ tokens expire after hours or days. Hey365 **handles that itself** instead of asking you to re-authenticate:

1. When a tool hits an authentication failure, Hey365 runs `workiq auth login` right there. If the token cache is still usable this completes immediately, with no browser.
2. Once reconnected it **re-runs the original tool with the same arguments** and prefixes the result with a note that it reconnected.
3. Only if reconnecting fails does it return an error — carrying a machine-readable `recovery` hint telling the assistant to call `hey365_reconnect`.

Concurrent failures share a single `auth login` (single-flight), and after a failed attempt Hey365 waits `HEY365_RECONNECT_COOLDOWN_MS` (60s by default) before trying again.

To do it manually, tell your assistant "**reconnect to Work IQ**" or run:

```bash
node dist/index.js reconnect              # no-op when already connected
node dist/index.js reconnect --force      # re-establish even when connected
node dist/index.js reconnect --no-browser # cache only, never open a browser
```

Set `HEY365_AUTO_RECONNECT=0` to turn automatic reconnection off.

## Development

```bash
npm run typecheck     # tsc --noEmit
npm test              # vitest
npm run build         # tsc
npm run smoke         # MCP handshake against the built server
```

## License

Hey365 is **source-available**, not open source.

| Use | Licence | Cost |
| --- | --- | --- |
| Personal, hobby, study, research | [PolyForm Noncommercial 1.0.0](LICENSE) | Free |
| Charities, schools, public research, government | [PolyForm Noncommercial 1.0.0](LICENSE) | Free |
| **Any commercial or business use** | [Commercial licence](COMMERCIAL-LICENSE.md) | **Prior consent + ¥150 per seat** |

### Commercial use

Commercial use is **not** granted by the noncommercial licence. To use Hey365 at
work, do both of these **before** you start:

1. **Email [monoqlo78@gmail.com](mailto:monoqlo78@gmail.com)** with your
   organisation, the number of seats, and what you plan to use it for, and wait
   for the licensor's consent.
2. **Pay in advance** — either **¥150 per seat** by bank transfer, or the
   **equivalent of USD 1 per seat** in Bitcoin.

Full terms, bank details and the Bitcoin address are in
**[COMMERCIAL-LICENSE.md](COMMERCIAL-LICENSE.md)**.

Redistribution, resale, or offering Hey365 to third parties as a service is not
covered — contact the licensor separately.

Dependencies and Microsoft Work IQ carry their own licences and terms.

> Releases up to and including v0.1.0 were published under the MIT License, kept
> in [LICENSE-MIT-legacy.txt](LICENSE-MIT-legacy.txt). That grant is not revoked
> for those versions.
