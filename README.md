# Outlook MCP Server

A Model Context Protocol (MCP) server that provides access to Microsoft Outlook via the Microsoft Graph API. This server enables AI assistants to interact with your Outlook mailbox, including sending emails, managing folders, and organizing messages.

## Features

- **Message Management**: List, read, create, update, delete, send, reply, and forward messages
- **Calendar**: List calendars, list/view events, create, update, delete events, and RSVP to invites
- **Contacts**: List, search, read, create, update, and delete contacts
- **Folder Operations**: Create, manage, and organize mail folders
- **Search Functionality**: Search through messages with advanced filtering
- **Attachment Support**: List, add, and manage email attachments
- **Message Rules**: Create and manage inbox rules for automatic message processing
- **Well-Known Folders**: Access standard folders like Inbox, Sent Items, Drafts, etc.
- **Automatic Signature**: Outgoing mail, drafts, replies, and forwards get the configured signature appended automatically
- **Automatic Pagination**: All list operations include built-in pagination to prevent excessive context generation

## Message formatting (lists & signatures)

Email bodies take a `body.contentType` of either `"text"` or `"html"`:

- **Use `"html"` for any bulleted or numbered lists or other rich formatting.**
  Wrap items in `<ul><li>…</li></ul>` (or `<ol><li>…</li></ol>`) and paragraphs
  in `<p>`. A plain `"text"` body renders a leading `-` as a literal dash — it
  does **not** become a bullet.
- **Signatures:** `create_message`, `send_mail`, and the reply/forward tools
  append the configured signature automatically — the styled HTML signature for
  `"html"` bodies, the plain-text signature for `"text"` bodies. A body that
  already contains the signature (detected by the `geoffreychallen.com` marker)
  is left untouched, so it is never doubled.
- **`update_message` writes the body verbatim** and does **not** inject a
  signature. When you replace a draft's body, include the full HTML and the
  signature yourself. Note: editing a draft can rotate its message ID, so
  re-fetch the draft (via `list_messages`/`get_message`) before a follow-up
  `update_message`.

## Installation

### Prerequisites

- Node.js 18 or higher
- An Azure App Registration with appropriate permissions
- Microsoft Graph API access

### Install via npm

```bash
npm install -g outlook-mcp
```

### Install from source

```bash
git clone <repository-url>
cd outlook-mcp
npm install
npm run build
```

## Configuration

> **Setting this up locally for Claude Code?** Follow [SETUP.md](./SETUP.md) for
> exact, copy-paste install steps (clone → build → sign in → register). The
> notes below are the reference for the configuration it uses.

### Environment Variables

Create a `.env` file in your project root:

```bash
AZURE_CLIENT_ID=your_client_id_here
AZURE_TENANT_ID=your_tenant_id_here
# Optional; defaults to Mail.ReadWrite,Mail.Send,Calendars.ReadWrite,Contacts.ReadWrite,User.Read
# GRAPH_SCOPES=Mail.ReadWrite,Mail.Send,Calendars.ReadWrite,Contacts.ReadWrite,User.Read
```

Authentication uses the **device-code flow** (no client secret). The Azure app
registration must have **"Allow public client flows" = Yes** and the matching
delegated Graph permissions. See [INSTRUCTIONS.md](./INSTRUCTIONS.md) for how to
create such a registration.

You authenticate once with `npm run login`, which caches a refresh token to
`.token-cache.json` (gitignored). An optional `OUTLOOK_TOKEN_CACHE` env var
overrides where that cache lives — useful when Claude Code launches the server
from a different working directory.

## Usage

### As an MCP Server

The server is designed to be used with MCP-compatible AI assistants. Start the server using:

```bash
npx outlook-mcp
```

### Available Tools

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
- `move_message` - Move a message to another folder
- `copy_message` - Copy a message to another folder
- `search_messages` - Search for messages (paginated, default: 25, max: 50)
- `list_attachments` - List message attachments (paginated, default: 10, max: 25)

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
flow). Run `npm run login` once: it prints a URL and a short code, you sign in
through a browser, and a refresh token is cached to `.token-cache.json`. After
that the server refreshes access tokens silently — no client secret involved.

Delegated (`/me`) Graph endpoints require a user token, which is exactly what
device-code provides. A client-secret/application token would not work for the
`/me/...` calls this server makes.

## Permissions

The following delegated Microsoft Graph permissions are used:

- `Mail.ReadWrite` - Read and write access to user mail
- `Mail.Send` - Send mail as the user
- `Calendars.ReadWrite` - Read and write calendar events
- `Contacts.ReadWrite` - Read and write contacts
- `User.Read` - Read the signed-in user's profile
- `offline_access` - Obtain refresh tokens for silent re-authentication

## Security

- Never commit your `.env` file or expose your client secrets
- Use principle of least privilege when configuring Azure App permissions
- Regularly rotate your client secrets
- Monitor your application's usage in the Azure portal

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

1. Check the [INSTRUCTIONS.md](./INSTRUCTIONS.md) for setup help
2. Review the Microsoft Graph documentation
3. File an issue on the repository
