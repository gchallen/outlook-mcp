# Outlook MCP Server

A Model Context Protocol (MCP) server that provides access to Microsoft Outlook via the Microsoft Graph API. This server enables AI assistants to interact with your Outlook mailbox, including sending emails, managing folders, and organizing messages.

## Features

- **Message Management**: List, read, create, update, delete, send, reply, and forward messages
- **Calendar**: List calendars, list/view events, create, update, delete events, and RSVP to invites
- **Contacts**: List, search, read, create, update, and delete contacts
- **Teams**: Read, send, and search Teams chats; list teams and channels; post and reply in channels
- **Folder Operations**: Create, manage, and organize mail folders
- **Search Functionality**: Search through messages with advanced filtering
- **Attachment Support**: List, add, and manage email attachments
- **Message Rules**: Create and manage inbox rules for automatic message processing
- **Well-Known Folders**: Access standard folders like Inbox, Sent Items, Drafts, etc.
- **Automatic Signature**: Outgoing mail, drafts, replies, and forwards get the configured signature appended automatically
- **Automatic Pagination**: All list operations include built-in pagination to prevent excessive context generation

## Message formatting (lists & signatures)

Email bodies take a `body.contentType` of either `"text"` or `"html"`. Every
message goes out as HTML, in the same markup the Outlook web composer writes for
typed text, so authored and hand-typed paragraphs are indistinguishable when a
draft is edited in OWA:

- **A `"text"` body** is escaped, its bare URLs are hyperlinked, and each line
  becomes a font-bearing `<div>`; a blank line becomes `<div><br></div>`,
  exactly as OWA represents one. A leading `-` renders as a literal dash — it
  does **not** become a bullet.
- **Use `"html"` for any bulleted or numbered lists or other rich formatting.**
  Wrap items in `<ul><li>…</li></ul>` (or `<ol><li>…</li></ol>`) and paragraphs
  in `<p>`. Each `<p>` is rewritten as the same OWA-style `<div>` block, with a
  blank-line block between paragraphs (and after a list) instead of margins.
  Author margins on a `<p>` are dropped; other inline styles are kept.
- **Signatures:** `create_message`, `send_mail`, and the reply/forward tools
  append the styled HTML signature automatically. It opens with its own blank
  line, so never end a body with one. A body that already contains the
  signature (detected by its `id="Signature"` wrapper) is left untouched, so it
  is never doubled; a spelled-out plain-text signature is replaced by the
  styled one.
- **`update_message`** applies the same font and paragraph normalization to a
  replacement body but does **not** inject a signature. When you replace a
  draft's body, include the full HTML and the signature yourself. Note: editing
  a draft can rotate its message ID, so re-fetch the draft (via
  `list_messages`/`get_message`) before a follow-up `update_message`.

## Getting Started

**Not technical? Start with the [guided setup](./docs/GUIDED-SETUP.md).** Paste
its prompt into your AI app (Claude Desktop, Claude Code, Codex, Cursor, …) and
the assistant walks you through every step, usually with no help needed from
your IT department.

Each person registers their own app in their organization's Microsoft Entra
directory (free, and allowed by default in most organizations). The server then
runs on their own computer and signs in as them. There is no shared app and no
hosted service.

### Quick Start (Technical Users)

1. Register a single-tenant app in Entra with **Allow public client flows = Yes**
   ([INSTRUCTIONS.md](./INSTRUCTIONS.md)).
2. Run setup. It finds your tenant from your email, saves
   `~/.outlook-mcp/config.json`, and signs you in once:

   ```bash
   npx -y @gchallen/outlook-mcp setup --email you@university.edu --client-id <application-client-id>
   ```

3. Add the server to your AI app:

   | App                 | How                                                                        |
   | ------------------- | -------------------------------------------------------------------------- |
   | Claude Desktop      | Double-click `outlook-mcp.mcpb` (from Releases; `npm run bundle` builds it) |
   | Claude Code         | `claude mcp add outlook --scope user -- npx -y @gchallen/outlook-mcp`     |
   | Codex               | `codex mcp add outlook -- npx -y @gchallen/outlook-mcp`                   |
   | Cursor, VS Code, …  | `{"mcpServers": {"outlook": {"command": "npx", "args": ["-y", "@gchallen/outlook-mcp"]}}}` |

   ChatGPT and claude.ai on the web only take remote (hosted) MCP servers, so
   they can't use this local server.

You can skip the sign-in during setup (`--no-login`). The first tool call then
returns a Microsoft sign-in link and code, so you can sign in from inside any AI
app with no terminal. The `outlook_status` tool does this on purpose.

For a development install from source, see [SETUP.md](./SETUP.md).

## Configuration

Settings live in `~/.outlook-mcp/config.json` (or under `$OUTLOOK_MCP_HOME`):

```json
{
  "email": "you@university.edu",
  "clientId": "1a2b3c4d-…",
  "tenantId": "…",
  "teams": true,
  "signature": "Jane Doe\nProfessor of Physics\nhttps://example.edu/~jane",
  "extraScopes": []
}
```

- `tenantId` is optional when `email` is set. It's looked up from the email's
  domain.
- `teams: false` skips the Teams permissions at sign-in, for organizations that
  require admin approval for them.
- `signature` is plain text (one line per line, URLs linked) or an HTML
  fragment. Leave it unset for no signature. The `set_signature` tool updates it,
  so you can just ask your AI app to "use the signature from my last sent email".
- The sign-in token is cached in `~/.outlook-mcp/token-cache.json` (mode 600).

Environment variables override the file: `AZURE_CLIENT_ID`, `AZURE_TENANT_ID`,
`OUTLOOK_EMAIL`, `OUTLOOK_TEAMS`, `GRAPH_SCOPES`, `EXTRA_GRAPH_SCOPES`, and
`OUTLOOK_TOKEN_CACHE`. The exception is `OUTLOOK_SIGNATURE`, which applies only
when the file has no `signature`. A `.env` file in the working directory is
also read.

### Available Tools

#### Setup Tools

- `outlook_status` - Show the signed-in account and signature; starts a sign-in if needed
- `get_signature` - Show the signature appended to outgoing mail
- `set_signature` - Set or clear that signature (saved to the config file)

#### Message Tools

- `list_messages` - List messages in mailbox or specific folder (paginated, default: 25, max: 100)
- `get_message` - Get a specific message by ID
- `create_message` - Create a draft message
- `update_message` - Update message properties
- `delete_message` - Delete a message
- `send_message` - Send a draft message
- `send_mail` - Send an email directly
- `reply_to_message` - Reply to a message
- `reply_all_to_message` - Reply all to a message
- `forward_message` - Forward a message
- `reply_to_message_draft` - Create a reply draft without sending
- `reply_all_to_message_draft` - Create a reply-all draft without sending
- `forward_message_draft` - Create a forward draft without sending
- `move_message` - Move a message to another folder
- `copy_message` - Copy a message to another folder
- `search_messages` - Search for messages (paginated, default: 25, max: 50)
- `list_attachments` - List message attachments (paginated, default: 10, max: 25)
- `add_attachment` - Attach a local file to a message or draft (handles >3 MB via upload session)
- `save_attachment` - Download an attachment to a local path (decoded and written server-side)

#### Folder Tools

- `list_mail_folders` - List mail folders (paginated, default: 20, max: 50)
- `get_mail_folder` - Get a specific mail folder
- `create_mail_folder` - Create a new mail folder
- `update_mail_folder` - Update mail folder properties
- `delete_mail_folder` - Delete a mail folder
- `move_mail_folder` - Move a mail folder
- `list_child_folders` - List child folders (paginated, default: 20, max: 50)
- `get_well_known_folder` - Get well-known folders (inbox, sentitems, etc.)

#### Rule Tools

- `list_message_rules` - List inbox message rules (paginated, default: 10, max: 25)
- `create_message_rule` - Create a new inbox message rule
- `delete_message_rule` - Delete an inbox message rule

#### Calendar Tools

- `list_calendars` - List the user's calendars (paginated, default: 20, max: 50)
- `list_events` - List events from the default or a specific calendar (paginated, default: 25, max: 100)
- `get_calendar_view` - List events within a date/time range, expanding recurring series into occurrences
- `get_event` - Get a specific event by ID
- `create_event` - Create a calendar event (supports attendees, location, online meeting)
- `update_event` - Update an event's properties
- `delete_event` - Delete an event
- `respond_to_event` - Accept, decline, or tentatively accept a meeting invitation

#### Contacts Tools

- `list_contacts` - List contacts (paginated, default: 25, max: 100)
- `search_contacts` - Search contacts (paginated, default: 25, max: 100)
- `get_contact` - Get a specific contact by ID
- `create_contact` - Create a new contact
- `update_contact` - Update a contact's properties
- `delete_contact` - Delete a contact

#### Teams Tools

- `list_teams_chats` - List chats, most recent first, with an unread flag (paginated, default: 20, max: 50)
- `get_teams_chat_messages` - Read recent messages in a chat (paginated, default: 20, max: 50)
- `send_teams_chat_message` - Send a message to a chat
- `create_teams_chat` - Start (or find) a one-on-one or group chat by email address
- `mark_teams_chat_read` - Mark a chat as read
- `react_to_teams_chat_message` - React to a chat message
- `search_teams_messages` - Full-text search across Teams messages (max: 25)
- `list_joined_teams` - List the teams you belong to
- `list_teams_channels` - List a team's channels
- `list_teams_channel_messages` - List top-level channel posts (needs `ChannelMessage.Read.All`)
- `get_teams_channel_message_replies` - Read a channel thread (needs `ChannelMessage.Read.All`)
- `send_teams_channel_message` - Post a new channel message
- `reply_to_teams_channel_message` - Reply in a channel thread

## Examples

### List Recent Messages

```typescript
// Lists the 10 most recent messages with pagination info
const result = await messageTools.listMessages({
  top: 10,
  orderBy: 'receivedDateTime desc',
});

// Result includes pagination metadata:
// {
//   value: [...messages...],
//   '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$top=10&$skip=10',
//   '@odata.count': 150
// }
```

### Paginate Through Messages

```typescript
// Get first page of messages
let result = await messageTools.listMessages({ top: 25 });
console.log(`Found ${result.value.length} messages`);

// Get next page if available
if (result['@odata.nextLink']) {
  const nextResult = await messageTools.listMessages({
    top: 25,
    skip: 25,
  });
  console.log(`Next page has ${nextResult.value.length} messages`);
}
```

### Send an Email

```typescript
await messageTools.sendMail({
  message: {
    subject: 'Hello from MCP!',
    body: {
      contentType: 'html',
      content: '<p>This email was sent via the Outlook MCP server.</p>',
    },
    toRecipients: [
      {
        emailAddress: {
          address: 'recipient@example.com',
          name: 'Recipient Name',
        },
      },
    ],
  },
});
```

### Create a Mail Folder

```typescript
await folderTools.createMailFolder({
  displayName: 'Important Projects',
  parentFolderId: 'inbox', // Optional: create as subfolder
});
```

### Search Messages

```typescript
const results = await messageTools.searchMessages('from:boss@company.com subject:urgent');
```

## Pagination

All list operations in this MCP server include automatic pagination to prevent excessive context generation:

### Default Limits

- **Messages**: 25 per page (max: 100)
- **Search Results**: 25 per page (max: 50)
- **Folders**: 20 per page (max: 50)
- **Rules**: 10 per page (max: 25)
- **Attachments**: 10 per page (max: 25)

### Response Format

All paginated operations return an object with:

```typescript
{
  value: [...items...],           // Array of actual results
  '@odata.nextLink'?: string,     // URL for next page (if available)
  '@odata.count'?: number         // Total count (if available)
}
```

### Pagination Parameters

Use `top` and `skip` parameters to control pagination:

- `top`: Number of items to return (respects maximum limits)
- `skip`: Number of items to skip (for getting subsequent pages)

## Development

### Building

```bash
npm run build
```

### Testing

```bash
npm test
```

### Linting

```bash
npm run lint
npm run format
```

### Type Checking

```bash
npm run typecheck
```

## Authentication

This server uses the **device-code flow** (a delegated, public-client OAuth
flow, with no client secret). The first sign-in (from `setup`, `outlook-mcp login`, or
any tool call) shows a URL and a short code. You sign in through a browser, and a
refresh token is cached locally. After that the server refreshes access tokens
silently. When the refresh token expires, the next tool call starts a new
sign-in.

Sign-in failures are explained in plain language: admin consent required,
public client flows disabled, wrong client ID, wrong account, and so on.

## Permissions

The following delegated Microsoft Graph permissions are used:

- `Mail.ReadWrite` - Read and write access to user mail
- `Mail.Send` - Send mail as the user
- `Calendars.ReadWrite` - Read and write calendar events
- `Contacts.ReadWrite` - Read and write contacts
- `User.Read` - Read the signed-in user's profile
- `People.Read` - Search relevant people and the directory
- `Chat.ReadWrite`, `ChatMessage.Send` - Read, send, and mark Teams chats read
- `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Send` - List teams and channels and post to them
- `offline_access` - Obtain refresh tokens for silent re-authentication
- `ChannelMessage.Read.All` (optional, admin consent) - Read channel posts; add it
  with `"extraScopes": ["ChannelMessage.Read.All"]` and sign in again

The Teams scopes are requested only at sign-in, and `teams: false` drops them.
If they were never granted, the mail, calendar, and contacts tools keep working
and the Teams tools return a permission error.

## Security

- The server runs locally and uses delegated permissions only: it can do what
  the signed-in user can do in their own mailbox, and nothing else.
- There are no client secrets. The sensitive file is the token cache
  (`~/.outlook-mcp/token-cache.json`, mode 600). Delete it to sign out.
- Deleting the app registration in Entra revokes access immediately.

## License

ISC

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass
6. Submit a pull request

## Support

For issues and questions:

1. Check the [guided setup](./docs/GUIDED-SETUP.md) troubleshooting table
2. Review the Microsoft Graph documentation
3. File an issue on the repository
