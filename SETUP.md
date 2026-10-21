# Setup From Source (Development Install)

Most people should use the [guided setup](./docs/GUIDED-SETUP.md) or the
`npx` quick start in the [README](./README.md). This page covers running the
server from a clone, for development. Every step can be copied and pasted,
including by a coding agent. Commands assume macOS or Linux with Node.js 18+.

## 1. Prerequisites

- Node.js 18 or higher (`node --version`).
- An MCP client: Claude Code (`claude --version`), Codex (`codex --version`),
  or another.
- An Entra app registration (single tenant, **Allow public client flows =
  Yes**). See [INSTRUCTIONS.md](./INSTRUCTIONS.md).

## 2. Clone and Build

```bash
git clone git@github.com:gchallen/outlook-mcp.git
cd outlook-mcp
npm install
npm run build
```

## 3. Configure and Sign In

```bash
npm run setup -- --email you@university.edu --client-id <application-client-id>
```

This finds the tenant from the email domain, writes
`~/.outlook-mcp/config.json`, and runs the device-code sign-in. It prints a
URL and a short code to **stderr**. Open the URL in any browser, enter the
code, sign in as the mailbox owner, and approve the permissions. The token
cache is written to `~/.outlook-mcp/token-cache.json` (mode 600).

Options: `--tenant-id` to skip the lookup, `--no-teams` to skip the Teams
permissions (for a "Need admin approval" screen), and `--no-login` to save
settings only and sign in from the first tool call instead.

To sign in again later (expired or revoked token): `npm run login`.

## 4. Register the Built Server

Point the client at the compiled entry point with an **absolute path**:

```bash
claude mcp add outlook --scope user -- node "$(pwd)/dist/cli.js"
codex mcp add outlook -- node "$(pwd)/dist/cli.js"
```

No environment is needed: the server reads `~/.outlook-mcp/config.json` no
matter which directory the client starts it from. (Older installs that pass
`AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, and `OUTLOOK_TOKEN_CACHE` as env vars and
run `dist/index.js` keep working: env vars override the file.)

## 5. Verify

```bash
claude mcp list          # 'outlook' should show ✔ Connected
```

Restart the client and ask it to "check my Outlook connection". This calls
`outlook_status`, which reports the signed-in account and signature.

## Maintenance

- **After editing `src/`**: `npm run build` (the registered server runs the
  compiled `dist/`). Restart the client to pick up changes.
- **Quality gates**: `npm test`, `npm run lint`, `npm run format:check`,
  `npm run build` should all pass.
- **Claude Desktop bundle**: `npm run bundle` builds `build/outlook-mcp.mcpb`
  (production dependencies only) and validates its manifest.

## What Stays Local (Never Commit)

- `~/.outlook-mcp/` holds the config and the cached refresh token.
- `.env`, if you use one, is gitignored.
