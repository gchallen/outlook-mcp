# Setup (local MCP server for Claude Code)

This server runs locally over stdio and connects Claude Code to a single user's
Outlook mailbox, calendar, and contacts via the Microsoft Graph API. It uses the
**device-code OAuth flow** — you sign in once in a browser and the refresh token
is cached locally, so the server can refresh silently afterward.

These instructions are precise enough to follow start-to-finish (including by a
coding agent). Commands assume macOS/Linux with Node.js 18+.

## 1. Prerequisites

- Node.js 18 or higher (`node --version`).
- The [Claude Code CLI](https://claude.com/claude-code) (`claude --version`).
- A browser available to complete a one-time sign-in (can be on any device).

## 2. Clone and build

```bash
git clone git@github.com:gchallen/outlook-mcp.git
cd outlook-mcp
npm install
npm run build
```

## 3. Configure credentials

Create a `.env` file in the project root:

```bash
AZURE_CLIENT_ID=<application-client-id>
AZURE_TENANT_ID=<directory-tenant-id>
```

Notes:

- These are the **client ID** and **tenant ID** of the Azure app registration
  (single-tenant, University of Illinois). They are **not secrets** — a public
  client device-code app has no client secret. The only sensitive artifact is
  the token cache produced in the next step, which stays on this machine and is
  gitignored.
- `GRAPH_SCOPES` is optional; it defaults to
  `Mail.ReadWrite,Mail.Send,Calendars.ReadWrite,Contacts.ReadWrite,People.Read,User.Read`,
  plus the Teams scopes
  `Chat.ReadWrite,ChatMessage.Send,Team.ReadBasic.All,Channel.ReadBasic.All,ChannelMessage.Send`
  at login. `EXTRA_GRAPH_SCOPES` appends more at login, e.g.
  `ChannelMessage.Read.All` for reading channel posts (needs admin consent).
- The Azure app must have **"Allow public client flows" = Yes** and the above
  delegated Graph permissions. (The existing registration already does. If you
  are setting up a brand-new app instead, see `INSTRUCTIONS.md`.)

## 4. Sign in once

```bash
npm run login
```

This prints a URL (https://microsoft.com/devicelogin) and a short code to
**stderr**. Open the URL in any browser, enter the code, sign in as the mailbox
owner, and approve the requested permissions. On success it writes
`.token-cache.json` to the project root (mode 600, gitignored). You should see:

```
Signed in as <you>@illinois.edu. Token cache saved.
```

If you instead hit a "needs admin approval" screen, the tenant requires IT
consent for these scopes — request it and retry.

## 5. Register the server with Claude Code

Use **absolute paths** for this machine. From the project directory:

```bash
claude mcp add outlook --scope user \
  --env AZURE_CLIENT_ID=<application-client-id> \
  --env AZURE_TENANT_ID=<directory-tenant-id> \
  --env OUTLOOK_TOKEN_CACHE="$(pwd)/.token-cache.json" \
  -- node "$(pwd)/dist/index.js"
```

Passing the env vars (and an absolute `OUTLOOK_TOKEN_CACHE`) here makes the
server independent of the working directory Claude Code launches it from.

## 6. Verify

```bash
claude mcp list          # 'outlook' should show ✔ Connected
```

Restart Claude Code, then run `/mcp` — the `outlook` tools should be listed.
Ask it something like "what's the most recent message in my inbox?" to confirm a
live call works.

## Maintenance

- **Re-authenticate** (token revoked/expired, or "Not authenticated" errors):
  `npm run login` again.
- **After editing `src/`**: `npm run build` (the registered server runs the
  compiled `dist/`). Restart Claude Code to pick up changes.
- **Quality gates**: `npm test`, `npm run lint`, `npm run format:check`,
  `npm run build` should all pass.

## What stays local (never commit)

- `.env` — credentials/config
- `.token-cache.json` — cached refresh tokens

Both are gitignored. Everything else (including this machine's absolute paths in
the `claude mcp add` command) is specific to the local install.
