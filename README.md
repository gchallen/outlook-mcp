# Outlook MCP Server

A Model Context Protocol (MCP) server that provides access to Microsoft Outlook via the Microsoft Graph API. This server enables AI assistants to interact with your Outlook mailbox, including sending emails, managing folders, and organizing messages.

## Features

- **Message Management**: List, read, create, update, delete, send, reply, and forward messages
- **Folder Operations**: Create, manage, and organize mail folders
- **Search Functionality**: Search through messages with advanced filtering
- **Attachment Support**: List, add, and manage email attachments
- **Message Rules**: Create and manage inbox rules for automatic message processing
- **Well-Known Folders**: Access standard folders like Inbox, Sent Items, Drafts, etc.

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

### Environment Variables

Create a `.env` file in your project root or set the following environment variables:

```bash
AZURE_CLIENT_ID=your_client_id_here
AZURE_TENANT_ID=your_tenant_id_here
AZURE_CLIENT_SECRET=your_client_secret_here
GRAPH_SCOPES=Mail.ReadWrite,Mail.Send,MailboxSettings.ReadWrite
```

See [INSTRUCTIONS.md](./INSTRUCTIONS.md) for detailed instructions on obtaining these credentials.

## Usage

### As an MCP Server

The server is designed to be used with MCP-compatible AI assistants. Start the server using:

```bash
npx outlook-mcp
```

### Available Tools

#### Message Tools

- `list_messages` - List messages in mailbox or specific folder
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
- `search_messages` - Search for messages

#### Folder Tools

- `list_mail_folders` - List mail folders
- `get_mail_folder` - Get a specific mail folder
- `create_mail_folder` - Create a new mail folder
- `update_mail_folder` - Update mail folder properties
- `delete_mail_folder` - Delete a mail folder
- `move_mail_folder` - Move a mail folder
- `get_well_known_folder` - Get well-known folders (inbox, sentitems, etc.)

#### Rule Tools

- `list_message_rules` - List inbox message rules
- `create_message_rule` - Create a new inbox message rule
- `delete_message_rule` - Delete an inbox message rule

## Examples

### List Recent Messages

```typescript
// Lists the 10 most recent messages
await messageTools.listMessages({
  top: 10,
  orderBy: 'receivedDateTime desc'
});
```

### Send an Email

```typescript
await messageTools.sendMail({
  message: {
    subject: 'Hello from MCP!',
    body: {
      contentType: 'html',
      content: '<p>This email was sent via the Outlook MCP server.</p>'
    },
    toRecipients: [{
      emailAddress: {
        address: 'recipient@example.com',
        name: 'Recipient Name'
      }
    }]
  }
});
```

### Create a Mail Folder

```typescript
await folderTools.createMailFolder({
  displayName: 'Important Projects',
  parentFolderId: 'inbox' // Optional: create as subfolder
});
```

### Search Messages

```typescript
const results = await messageTools.searchMessages('from:boss@company.com subject:urgent');
```

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

This server supports two authentication methods:

1. **Client Secret Flow** (recommended for automated scenarios)
2. **Device Code Flow** (for interactive authentication)

The authentication method is automatically determined based on whether `AZURE_CLIENT_SECRET` is provided.

## Permissions

The following Microsoft Graph permissions are required:

- `Mail.ReadWrite` - Read and write access to user mail
- `Mail.Send` - Send mail as the user
- `MailboxSettings.ReadWrite` - Read and write mailbox settings

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