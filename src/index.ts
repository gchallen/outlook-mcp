#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
} from '@modelcontextprotocol/sdk/types.js';
import { AuthRequiredError, GraphAuthProvider, explainAuthError } from './auth.js';
import {
  GUIDE_URL,
  PACKAGE_NAME,
  discoverTenantId,
  resolveConfig,
  writeConfigFile,
} from './config.js';
import { hasSignature, setSignature, signatureHtml, signatureText } from './signature.js';
import { MessageTools } from './tools/message-tools.js';
import { FolderTools } from './tools/folder-tools.js';
import { RuleTools } from './tools/rule-tools.js';
import { CalendarTools } from './tools/calendar-tools.js';
import { ContactsTools } from './tools/contacts-tools.js';
import { PeopleTools } from './tools/people-tools.js';
import { TeamsTools } from './tools/teams-tools.js';
import * as dotenv from 'dotenv';
import 'isomorphic-fetch';

// quiet: true suppresses dotenv's startup banner. On an MCP stdio server,
// anything written to stdout that isn't JSON-RPC corrupts the protocol stream
// and breaks the connection.
dotenv.config({ quiet: true });

const config = resolveConfig();
setSignature(config.signature);

// Tool descriptions say what happens to sign-offs, which depends on whether
// this user has a signature configured. Without one, telling the model "do not
// write a sign-off, the server adds it" would send unsigned mail. These are
// functions because set_signature can change the answer mid-session.
const signatureAppended = () =>
  hasSignature()
    ? 'the styled HTML signature is appended'
    : 'no signature is appended (none is configured)';
const signatureRule = () =>
  hasSignature()
    ? 'Do not write your own signature or closing sign-off such as "Best, <name>": the server appends the signature that closes the message (a typed-out plain signature is stripped and replaced by the styled one).'
    : 'No signature is configured, so end the body with whatever sign-off the user wants (set_signature can configure one to append automatically).';
const replySignatureRule = () =>
  hasSignature()
    ? 'Do not write your own signature — a typed-out plain signature is stripped and replaced by the styled one.'
    : 'No signature is configured, so include any sign-off the user wants in the comment.';

const SETUP_HELP = `Outlook is not set up yet: no Azure app client ID is configured. Follow the guided setup at ${GUIDE_URL}, or run \`npx -y ${PACKAGE_NAME} setup\` in a terminal.`;

const server = new Server(
  {
    name: 'outlook-mcp',
    version: '2026.10.0',
  },
  {
    capabilities: {
      tools: { listChanged: true },
    },
  },
);

let messageTools: MessageTools;
let folderTools: FolderTools;
let ruleTools: RuleTools;
let calendarTools: CalendarTools;
let contactsTools: ContactsTools;
let peopleTools: PeopleTools;
let teamsTools: TeamsTools;

// Trim a draft message down to the fields worth surfacing after creating a
// reply draft, avoiding a dump of the full quoted-thread HTML body.
function summarizeDraft(draft: any) {
  return {
    id: draft?.id,
    subject: draft?.subject,
    isDraft: draft?.isDraft,
    webLink: draft?.webLink,
    toRecipients: draft?.toRecipients,
    ccRecipients: draft?.ccRecipients,
    bodyPreview: draft?.bodyPreview,
  };
}

let authProvider: GraphAuthProvider | undefined;
let configError: string | undefined;

async function initializeGraphClient() {
  if (!config.clientId) {
    configError = SETUP_HELP;
    return;
  }
  let tenantId = config.tenantId;
  if (!tenantId && config.email) {
    try {
      tenantId = await discoverTenantId(config.email);
    } catch (error) {
      configError = `Could not look up the Microsoft 365 organization for ${config.email}: ${(error as Error).message}`;
      return;
    }
  }
  if (!tenantId) {
    configError = `Outlook is not fully set up: neither a tenant ID nor an email address is configured. Re-run setup (\`npx -y ${PACKAGE_NAME} setup\`).`;
    return;
  }

  authProvider = new GraphAuthProvider({
    clientId: config.clientId,
    tenantId,
    scopes: config.scopes,
    loginScopes: config.loginScopes,
    cachePath: config.cachePath,
  });

  const client = authProvider.getClient();
  messageTools = new MessageTools(client);
  folderTools = new FolderTools(client);
  ruleTools = new RuleTools(client);
  calendarTools = new CalendarTools(client);
  contactsTools = new ContactsTools(client);
  peopleTools = new PeopleTools(client);
  teamsTools = new TeamsTools(client);
}

// Built per request so descriptions track the current signature.
const buildTools = (): Tool[] => [
  // Setup Tools
  {
    name: 'outlook_status',
    description:
      'Check the Outlook connection: whether setup is complete, which account is signed in, and whether a signature is configured. If the user is not signed in yet, this starts a Microsoft sign-in and returns a link and code for the user to enter in their browser; after they finish, call it again to confirm. Use this first when setting up, or when another tool reports that sign-in is needed.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'get_signature',
    description:
      'Show the email signature that is appended automatically to new mail, replies, and forwards (empty if none is configured).',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'set_signature',
    description:
      "Set the email signature appended automatically to new mail, replies, and forwards. Pass plain text (one line per line; URLs become links) or an HTML fragment, or an empty string to remove the signature. To copy the user's existing signature, read one of their recent sent messages and pass its signature block, then show the user the result of get_signature to confirm. Saved to the user's config file, and takes effect immediately. Tool descriptions that mention the signature update after the AI app is restarted.",
    inputSchema: {
      type: 'object',
      properties: {
        signature: {
          type: 'string',
          description: 'Signature as plain text or an HTML fragment; empty string removes it',
        },
      },
      required: ['signature'],
    },
  },

  // Message Tools
  {
    name: 'list_messages',
    description: 'List messages in mailbox or specific folder with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        filter: { type: 'string', description: 'OData filter query' },
        select: { type: 'string', description: 'Comma-separated list of fields to return' },
        top: {
          type: 'number',
          description: 'Number of messages to return (default: 25, max: 100)',
          minimum: 1,
          maximum: 100,
        },
        skip: {
          type: 'number',
          description: 'Number of messages to skip for pagination',
          minimum: 0,
        },
        orderBy: { type: 'string', description: 'Field to order by' },
        folderId: { type: 'string', description: 'Folder ID to list messages from' },
      },
    },
  },
  {
    name: 'get_message',
    description: 'Get a specific message by ID',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID' },
        select: { type: 'string', description: 'Comma-separated list of fields to return' },
        expand: { type: 'string', description: 'Related entities to expand' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'create_message',
    description: `Create a draft message. Every body goes out as HTML: a "text" body is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved, "- " and "1. " lines made into lists) and ${signatureAppended()} either way, so plain prose can simply be sent as "text". Lines starting "- " or "1. " in a "text" body become real bulleted or numbered lists. For other rich formatting, set body.contentType to "html" and use <p> for paragraphs. Bare URLs in HTML bodies are also auto-hyperlinked (existing <a> tags are left alone). ${signatureRule()} Use Title Case for the subject line.`,
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string' },
        body: {
          type: 'object',
          properties: {
            contentType: { type: 'string', enum: ['text', 'html'] },
            content: {
              type: 'string',
              description:
                'Body content. In a "text" body, lines starting "- " or "1. " become real lists; use contentType "html" for other formatting.',
            },
          },
        },
        toRecipients: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              emailAddress: {
                type: 'object',
                properties: {
                  address: { type: 'string' },
                  name: { type: 'string' },
                },
              },
            },
          },
        },
        ccRecipients: { type: 'array', items: { type: 'object' } },
        bccRecipients: { type: 'array', items: { type: 'object' } },
        importance: { type: 'string', enum: ['low', 'normal', 'high'] },
      },
    },
  },
  {
    name: 'update_message',
    description:
      'Update message properties (PATCH). Replacement bodies always go out as HTML: a "text" body is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved, "- " and "1. " lines made into lists), and an HTML body gets the outgoing font, paragraph spacing, and URL auto-linking stamped on it — authored <p>/<ul>/<li> markup needs no inline font styling, and a body that already declares its own font-family is left alone. The signature is NOT added — a PATCH usually replaces the whole body, so include the signature yourself (a typed-out plain signature is upgraded to the styled hyperlinked one). In a "text" body, lines starting "- " or "1. " become real lists. Note: editing a draft can rotate its message ID, so re-fetch the draft (list/get) before a follow-up update.',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID' },
        updates: { type: 'object', description: 'Properties to update' },
      },
      required: ['messageId', 'updates'],
    },
  },
  {
    name: 'delete_message',
    description: 'Delete a message (moves to Deleted Items)',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'send_message',
    description: 'Send a previously created draft message',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Draft message ID' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'send_mail',
    description: `Send a new email directly. Same formatting rules as create_message: every body goes out as HTML — a "text" body is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved, "- " and "1. " lines made into lists) and ${signatureAppended()} either way. Lines starting "- " or "1. " in a "text" body become real lists. ${signatureRule()} Use Title Case for the subject line.`,
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string' },
        body: {
          type: 'object',
          properties: {
            contentType: { type: 'string', enum: ['text', 'html'] },
            content: { type: 'string' },
          },
        },
        toRecipients: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              emailAddress: {
                type: 'object',
                properties: {
                  address: { type: 'string' },
                  name: { type: 'string' },
                },
              },
            },
          },
        },
        saveToSentItems: { type: 'boolean', default: true },
      },
      required: ['subject', 'body', 'toRecipients'],
    },
  },
  {
    name: 'reply_to_message',
    description: `Reply to a message and send immediately. The comment is treated as plain text and converted to HTML automatically (escaped, bare URLs hyperlinked, line breaks preserved, "- " and "1. " lines made into lists) with ${hasSignature() ? 'the styled signature appended' : 'no signature appended'}. ${replySignatureRule()}`,
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID to reply to' },
        comment: { type: 'string', description: 'Reply comment' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'reply_all_to_message',
    description: `Reply all to a message and send immediately. The comment is treated as plain text and converted to HTML automatically (escaped, bare URLs hyperlinked, line breaks preserved, "- " and "1. " lines made into lists) with ${hasSignature() ? 'the styled signature appended' : 'no signature appended'}. ${replySignatureRule()}`,
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID to reply to' },
        comment: { type: 'string', description: 'Reply comment' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'reply_to_message_draft',
    description: `Create a reply draft (to sender only) without sending it. Returns the draft so it can be reviewed and sent later. The reply always goes out as HTML (${signatureAppended()}): a "text" comment is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved, "- " and "1. " lines made into lists). ${replySignatureRule()}`,
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID to reply to' },
        comment: { type: 'string', description: 'Reply body' },
        contentType: {
          type: 'string',
          enum: ['text', 'html'],
          description:
            'How to treat comment (default text). Use html only for rich formatting such as <ul>/<li> lists; plain prose can stay text — it is converted to HTML with URLs auto-linked either way.',
        },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'reply_all_to_message_draft',
    description: `Create a reply-all draft without sending it. Returns the draft so it can be reviewed and sent later. The reply always goes out as HTML (${signatureAppended()}): a "text" comment is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved, "- " and "1. " lines made into lists). ${replySignatureRule()}`,
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID to reply to' },
        comment: { type: 'string', description: 'Reply body' },
        contentType: {
          type: 'string',
          enum: ['text', 'html'],
          description:
            'How to treat comment (default text). Use html only for rich formatting such as <ul>/<li> lists; plain prose can stay text — it is converted to HTML with URLs auto-linked either way.',
        },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'forward_message',
    description: `Forward a message and send immediately. The comment is treated as plain text and converted to HTML automatically (escaped, bare URLs hyperlinked, line breaks preserved, "- " and "1. " lines made into lists) with ${hasSignature() ? 'the styled signature appended' : 'no signature appended'}. ${replySignatureRule()}`,
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID to forward' },
        toRecipients: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              emailAddress: {
                type: 'object',
                properties: {
                  address: { type: 'string' },
                },
              },
            },
          },
        },
        comment: { type: 'string', description: 'Forward comment' },
      },
      required: ['messageId', 'toRecipients'],
    },
  },
  {
    name: 'forward_message_draft',
    description: `Create a forward draft without sending it. Returns the draft so it can be reviewed and sent later. The forward always goes out as HTML (${signatureAppended()}, above the quoted original): a "text" comment is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved, "- " and "1. " lines made into lists). ${replySignatureRule()}`,
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID to forward' },
        toRecipients: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              emailAddress: {
                type: 'object',
                properties: {
                  address: { type: 'string' },
                },
              },
            },
          },
        },
        comment: { type: 'string', description: 'Forward comment' },
        contentType: {
          type: 'string',
          enum: ['text', 'html'],
          description:
            'How to treat comment (default text). Use html only for rich formatting such as <ul>/<li> lists; plain prose can stay text — it is converted to HTML with URLs auto-linked either way.',
        },
      },
      required: ['messageId', 'toRecipients'],
    },
  },
  {
    name: 'move_message',
    description: 'Move a message to another folder',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID' },
        destinationId: { type: 'string', description: 'Destination folder ID' },
      },
      required: ['messageId', 'destinationId'],
    },
  },
  {
    name: 'move_folder_messages',
    description:
      'Bulk-move messages from a source folder into a destination folder using Graph\'s $batch endpoint (up to 20 moves per request). Built for retiring a folder of hundreds of messages, or archiving a dated subset, that the web UI can\'t move in one action. Pass an optional OData `filter` to move only the matching subset — e.g. filter "receivedDateTime lt 2025-08-20T00:00:00Z" with sourceFolderId inbox and destinationId the Archive folder to archive everything older than that date; omit it to move the whole folder. Do NOT add an $orderby — Graph rejects filter+orderby on this collection. Returns a summary { moved, failed, batches }; does not delete the emptied source folder.',
    inputSchema: {
      type: 'object',
      properties: {
        sourceFolderId: { type: 'string', description: 'Folder to move messages out of' },
        destinationId: { type: 'string', description: 'Destination folder ID' },
        filter: {
          type: 'string',
          description:
            'Optional OData $filter to move only matching messages (e.g. "receivedDateTime lt 2025-08-20T00:00:00Z"). Omit to move the entire folder.',
        },
        limit: {
          type: 'number',
          description: 'Optional cap on how many messages to move (default: all)',
          minimum: 1,
        },
        batchSize: {
          type: 'number',
          description: 'Moves per $batch request (default and max: 20)',
          minimum: 1,
          maximum: 20,
        },
      },
      required: ['sourceFolderId', 'destinationId'],
    },
  },
  {
    name: 'copy_message',
    description: 'Copy a message to another folder',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID' },
        destinationId: { type: 'string', description: 'Destination folder ID' },
      },
      required: ['messageId', 'destinationId'],
    },
  },
  {
    name: 'search_messages',
    description:
      'Full-text search across mail using Microsoft Graph $search (KQL). Prefer property-scoped queries over bare keywords for much better precision. Supported properties (combine space-separated for implicit AND, or with explicit AND/OR/NOT): from:alice@x.com, to:, cc:, subject:budget, body:, participants: (any of from/to/cc), hasAttachment:true, attachment:report.pdf, received>=2024-01-01, received:2024-01-01..2024-01-31, sent>=, importance:high, isRead:false. Do NOT quote simple single-token values like an email address (from:alice@x.com, not from:"alice@x.com"). Only put double quotes around multi-word phrase values: subject:"quarterly report". Do NOT wrap the whole query in outer quotes — the tool adds the outer quoting Graph requires. Examples: from:dean@example.edu subject:budget received>=2024-09-01 / hasAttachment:true subject:"final grades". Note: $search results are relevance-ranked and CANNOT be date-sorted or combined with an OData $filter (a Graph limitation, not a tool one); when you need strict newest-first ordering or a hard field filter, use list_messages with filter/orderBy instead. Scope to one folder with folderId (get IDs from list_mail_folders / get_well_known_folder).',
    inputSchema: {
      type: 'object',
      properties: {
        searchTerm: {
          type: 'string',
          description:
            'KQL query. Use property restrictions (from:, subject:, body:, received>=, hasAttachment:true, etc.) rather than bare keywords when possible. Leave simple values unquoted; quote only multi-word phrases (subject:"quarterly report"). Do not add outer quotes around the whole query — the tool wraps it.',
        },
        folderId: {
          type: 'string',
          description:
            'Optional mail folder ID to scope the search to a single folder (e.g. inbox, or a well-known folder ID). Omit to search all folders.',
        },
        select: { type: 'string', description: 'Fields to return' },
        top: {
          type: 'number',
          description: 'Number of results (default: 25, max: 50)',
          minimum: 1,
          maximum: 50,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
      },
      required: ['searchTerm'],
    },
  },
  // Folder Tools
  {
    name: 'list_mail_folders',
    description: 'List mail folders with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        filter: { type: 'string', description: 'OData filter' },
        select: { type: 'string', description: 'Fields to return' },
        top: {
          type: 'number',
          description: 'Number of folders (default: 20, max: 50)',
          minimum: 1,
          maximum: 50,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
      },
    },
  },
  {
    name: 'get_mail_folder',
    description: 'Get a specific mail folder',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: { type: 'string', description: 'Folder ID' },
        select: { type: 'string', description: 'Fields to return' },
      },
      required: ['folderId'],
    },
  },
  {
    name: 'create_mail_folder',
    description: 'Create a new mail folder',
    inputSchema: {
      type: 'object',
      properties: {
        displayName: { type: 'string', description: 'Folder name' },
        parentFolderId: { type: 'string', description: 'Parent folder ID (optional)' },
      },
      required: ['displayName'],
    },
  },
  {
    name: 'update_mail_folder',
    description: 'Update mail folder properties',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: { type: 'string', description: 'Folder ID' },
        displayName: { type: 'string', description: 'New folder name' },
      },
      required: ['folderId', 'displayName'],
    },
  },
  {
    name: 'delete_mail_folder',
    description: 'Delete a mail folder',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: { type: 'string', description: 'Folder ID' },
      },
      required: ['folderId'],
    },
  },
  {
    name: 'move_mail_folder',
    description: 'Move a mail folder',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: { type: 'string', description: 'Folder ID to move' },
        destinationId: { type: 'string', description: 'Destination parent folder ID' },
      },
      required: ['folderId', 'destinationId'],
    },
  },
  {
    name: 'get_well_known_folder',
    description:
      'Get well-known folder (inbox, sentitems, deleteditems, drafts, outbox, junkemail)',
    inputSchema: {
      type: 'object',
      properties: {
        folderName: {
          type: 'string',
          enum: ['inbox', 'sentitems', 'deleteditems', 'drafts', 'outbox', 'junkemail'],
        },
      },
      required: ['folderName'],
    },
  },
  {
    name: 'list_attachments',
    description: 'List attachments for a message with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID' },
        top: {
          type: 'number',
          description: 'Number of attachments (default: 10, max: 25)',
          minimum: 1,
          maximum: 25,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'add_attachment',
    description:
      'Attach a local file to a message or draft by absolute path. Transparently handles large files (>3 MB) via a Microsoft Graph upload session.',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message or draft ID to attach the file to' },
        filePath: { type: 'string', description: 'Absolute path to a local file' },
      },
      required: ['messageId', 'filePath'],
    },
  },
  {
    name: 'save_attachment',
    description:
      'Download a message attachment to a local path. The file is decoded and written by the server, so its bytes never pass through the conversation - use this instead of list_attachments when the goal is to get a file onto disk. If destination is an existing directory (or ends with a separator), the attachment keeps its own filename inside it; otherwise destination is the full target path, and missing parent directories are created. Only file attachments can be saved (item and reference attachments carry no file content). Get attachmentId from list_attachments.',
    inputSchema: {
      type: 'object',
      properties: {
        messageId: { type: 'string', description: 'Message ID' },
        attachmentId: {
          type: 'string',
          description: 'Attachment ID, as returned by list_attachments',
        },
        destination: {
          type: 'string',
          description:
            'Absolute path to write to: either the target file path, or an existing directory in which the attachment keeps its own filename',
        },
      },
      required: ['messageId', 'attachmentId', 'destination'],
    },
  },
  {
    name: 'list_child_folders',
    description: 'List child folders of a parent folder with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: { type: 'string', description: 'Parent folder ID' },
        filter: { type: 'string', description: 'OData filter' },
        select: { type: 'string', description: 'Fields to return' },
        top: {
          type: 'number',
          description: 'Number of folders (default: 20, max: 50)',
          minimum: 1,
          maximum: 50,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
      },
      required: ['folderId'],
    },
  },
  // Rule Tools
  {
    name: 'list_message_rules',
    description: 'List inbox message rules with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        top: {
          type: 'number',
          description: 'Number of rules (default: 10, max: 25)',
          minimum: 1,
          maximum: 25,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
      },
    },
  },
  {
    name: 'create_message_rule',
    description: 'Create a new inbox message rule',
    inputSchema: {
      type: 'object',
      properties: {
        displayName: { type: 'string', description: 'Rule name' },
        sequence: { type: 'number', description: 'Order of execution' },
        isEnabled: { type: 'boolean', description: 'Whether rule is active' },
        conditions: { type: 'object', description: 'Rule conditions' },
        actions: { type: 'object', description: 'Rule actions' },
      },
      required: ['displayName'],
    },
  },
  {
    name: 'delete_message_rule',
    description: 'Delete an inbox message rule',
    inputSchema: {
      type: 'object',
      properties: {
        ruleId: { type: 'string', description: 'Rule ID' },
      },
      required: ['ruleId'],
    },
  },
  // Calendar Tools
  {
    name: 'list_calendars',
    description: 'List the user calendars with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        select: { type: 'string', description: 'Fields to return' },
        top: {
          type: 'number',
          description: 'Number of calendars (default: 20, max: 50)',
          minimum: 1,
          maximum: 50,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
      },
    },
  },
  {
    name: 'list_events',
    description: 'List events from the default or a specific calendar with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        calendarId: {
          type: 'string',
          description: 'Calendar ID (optional; defaults to primary calendar)',
        },
        filter: { type: 'string', description: 'OData filter query' },
        select: { type: 'string', description: 'Fields to return' },
        top: {
          type: 'number',
          description: 'Number of events (default: 25, max: 100)',
          minimum: 1,
          maximum: 100,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
        orderBy: { type: 'string', description: 'Field to order by (default: start/dateTime)' },
      },
    },
  },
  {
    name: 'get_calendar_view',
    description: 'Get events within a date/time range, expanding recurring series into occurrences',
    inputSchema: {
      type: 'object',
      properties: {
        startDateTime: {
          type: 'string',
          description: 'Range start, ISO 8601 (e.g. 2026-06-26T00:00:00)',
        },
        endDateTime: {
          type: 'string',
          description: 'Range end, ISO 8601 (e.g. 2026-07-03T00:00:00)',
        },
        calendarId: {
          type: 'string',
          description: 'Calendar ID (optional; defaults to primary calendar)',
        },
        select: { type: 'string', description: 'Fields to return' },
        top: {
          type: 'number',
          description: 'Number of events (default: 25, max: 100)',
          minimum: 1,
          maximum: 100,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
      },
      required: ['startDateTime', 'endDateTime'],
    },
  },
  {
    name: 'get_event',
    description: 'Get a specific calendar event by ID',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'Event ID' },
        select: { type: 'string', description: 'Fields to return' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'create_event',
    description: 'Create a calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        subject: { type: 'string', description: 'Event title' },
        body: {
          type: 'object',
          properties: {
            contentType: { type: 'string', enum: ['text', 'html'] },
            content: { type: 'string' },
          },
        },
        start: {
          type: 'object',
          description: 'Start time',
          properties: {
            dateTime: {
              type: 'string',
              description: 'ISO 8601 local datetime (e.g. 2026-06-26T14:00:00)',
            },
            timeZone: {
              type: 'string',
              description: 'IANA/Windows time zone (e.g. America/Chicago)',
            },
          },
        },
        end: {
          type: 'object',
          description: 'End time',
          properties: {
            dateTime: { type: 'string', description: 'ISO 8601 local datetime' },
            timeZone: { type: 'string', description: 'IANA/Windows time zone' },
          },
        },
        location: {
          type: 'object',
          properties: { displayName: { type: 'string' } },
        },
        attendees: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              emailAddress: {
                type: 'object',
                properties: {
                  address: { type: 'string' },
                  name: { type: 'string' },
                },
              },
              type: { type: 'string', enum: ['required', 'optional', 'resource'] },
            },
          },
        },
        isOnlineMeeting: { type: 'boolean', description: 'Create a Teams online meeting' },
        calendarId: {
          type: 'string',
          description: 'Calendar ID (optional; defaults to primary calendar)',
        },
      },
      required: ['subject', 'start', 'end'],
    },
  },
  {
    name: 'update_event',
    description: 'Update properties of a calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'Event ID' },
        updates: {
          type: 'object',
          description: 'Properties to update (subject, body, start, end, location, attendees, ...)',
        },
      },
      required: ['eventId', 'updates'],
    },
  },
  {
    name: 'delete_event',
    description: 'Delete a calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'Event ID' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'respond_to_event',
    description: 'Accept, decline, or tentatively accept a meeting invitation',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: { type: 'string', description: 'Event ID' },
        response: { type: 'string', enum: ['accept', 'decline', 'tentativelyAccept'] },
        comment: { type: 'string', description: 'Optional comment to include in the response' },
        sendResponse: {
          type: 'boolean',
          description: 'Send a response to the organizer (default: true)',
        },
      },
      required: ['eventId', 'response'],
    },
  },
  // Contacts Tools
  {
    name: 'list_contacts',
    description: 'List contacts with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        filter: { type: 'string', description: 'OData filter query' },
        select: { type: 'string', description: 'Fields to return' },
        top: {
          type: 'number',
          description: 'Number of contacts (default: 25, max: 100)',
          minimum: 1,
          maximum: 100,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
        orderBy: { type: 'string', description: 'Field to order by' },
      },
    },
  },
  {
    name: 'search_contacts',
    description: 'Search contacts with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        searchTerm: { type: 'string', description: 'Search term' },
        select: { type: 'string', description: 'Fields to return' },
        top: {
          type: 'number',
          description: 'Number of contacts (default: 25, max: 100)',
          minimum: 1,
          maximum: 100,
        },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
      },
      required: ['searchTerm'],
    },
  },
  {
    name: 'search_people',
    description:
      'Search the organization directory and relevant people via /me/people. Covers directory colleagues and frequently-contacted people, so this is the right tool when search_contacts (the personal Contacts folder) returns nothing.',
    inputSchema: {
      type: 'object',
      properties: {
        searchTerm: { type: 'string', description: 'Search term' },
        top: {
          type: 'number',
          description: 'Number of people to return (default: 10, max: 25)',
          minimum: 1,
          maximum: 25,
        },
      },
      required: ['searchTerm'],
    },
  },
  {
    name: 'get_contact',
    description: 'Get a specific contact by ID',
    inputSchema: {
      type: 'object',
      properties: {
        contactId: { type: 'string', description: 'Contact ID' },
        select: { type: 'string', description: 'Fields to return' },
      },
      required: ['contactId'],
    },
  },
  {
    name: 'create_contact',
    description: 'Create a new contact',
    inputSchema: {
      type: 'object',
      properties: {
        givenName: { type: 'string' },
        surname: { type: 'string' },
        emailAddresses: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              address: { type: 'string' },
              name: { type: 'string' },
            },
          },
        },
        businessPhones: { type: 'array', items: { type: 'string' } },
        mobilePhone: { type: 'string' },
        companyName: { type: 'string' },
        jobTitle: { type: 'string' },
      },
    },
  },
  {
    name: 'update_contact',
    description: 'Update properties of a contact',
    inputSchema: {
      type: 'object',
      properties: {
        contactId: { type: 'string', description: 'Contact ID' },
        updates: { type: 'object', description: 'Properties to update' },
      },
      required: ['contactId', 'updates'],
    },
  },
  {
    name: 'delete_contact',
    description: 'Delete a contact',
    inputSchema: {
      type: 'object',
      properties: {
        contactId: { type: 'string', description: 'Contact ID' },
      },
      required: ['contactId'],
    },
  },

  // Teams Tools
  {
    name: 'list_teams_chats',
    description:
      'List Microsoft Teams chats (one-on-one, group, and meeting chats), most recently active first, with members, a last-message preview, and an "unread" flag. Use unreadOnly to see just the chats waiting on a reply. Read a chat with get_teams_chat_messages.',
    inputSchema: {
      type: 'object',
      properties: {
        top: {
          type: 'number',
          description: 'Number of chats to return (default: 20, max: 50)',
          minimum: 1,
          maximum: 50,
        },
        unreadOnly: {
          type: 'boolean',
          description: 'Only return chats whose last message from someone else is unread',
        },
      },
    },
  },
  {
    name: 'get_teams_chat_messages',
    description:
      'Get the most recent messages in a Teams chat, returned oldest first with HTML flattened to text. System events (members added, etc.) are omitted.',
    inputSchema: {
      type: 'object',
      properties: {
        chatId: { type: 'string', description: 'Chat ID (from list_teams_chats)' },
        top: {
          type: 'number',
          description: 'Number of recent messages to return (default: 20, max: 50)',
          minimum: 1,
          maximum: 50,
        },
      },
      required: ['chatId'],
    },
  },
  {
    name: 'send_teams_chat_message',
    description:
      'Send a message to a Teams chat immediately. A "text" message is escaped with line breaks kept and bare URLs hyperlinked; use "html" for richer formatting. Teams chats have no threaded replies, so this is also how to respond in a chat. No signature is added.',
    inputSchema: {
      type: 'object',
      properties: {
        chatId: { type: 'string', description: 'Chat ID' },
        content: { type: 'string', description: 'Message content' },
        contentType: { type: 'string', enum: ['text', 'html'], description: 'Default: text' },
      },
      required: ['chatId', 'content'],
    },
  },
  {
    name: 'create_teams_chat',
    description:
      'Start a Teams chat with one or more people by email address. A single address gives a one-on-one chat (the existing chat is returned if there already is one); several give a new group chat. Then send with send_teams_chat_message.',
    inputSchema: {
      type: 'object',
      properties: {
        emails: {
          type: 'array',
          items: { type: 'string' },
          description: 'Email addresses (UPNs) of the other participants',
        },
        topic: { type: 'string', description: 'Optional group chat name' },
      },
      required: ['emails'],
    },
  },
  {
    name: 'mark_teams_chat_read',
    description: 'Mark a Teams chat as read.',
    inputSchema: {
      type: 'object',
      properties: {
        chatId: { type: 'string', description: 'Chat ID' },
      },
      required: ['chatId'],
    },
  },
  {
    name: 'react_to_teams_chat_message',
    description: 'Add a reaction to a Teams chat message, e.g. "👍", "❤️", "😆", "😮", "😢", "😡".',
    inputSchema: {
      type: 'object',
      properties: {
        chatId: { type: 'string', description: 'Chat ID' },
        messageId: { type: 'string', description: 'Message ID' },
        reactionType: { type: 'string', description: 'Reaction emoji' },
      },
      required: ['chatId', 'messageId', 'reactionType'],
    },
  },
  {
    name: 'list_joined_teams',
    description: 'List the Teams the user is a member of.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'list_teams_channels',
    description: 'List the channels in a team.',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'Team ID (from list_joined_teams)' },
      },
      required: ['teamId'],
    },
  },
  {
    name: 'list_teams_channel_messages',
    description:
      'List recent top-level posts in a Teams channel, newest first. Read a thread with get_teams_channel_message_replies. Needs the ChannelMessage.Read.All permission, which requires tenant admin consent.',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'Team ID' },
        channelId: { type: 'string', description: 'Channel ID' },
        top: {
          type: 'number',
          description: 'Number of posts to return (default: 20, max: 50)',
          minimum: 1,
          maximum: 50,
        },
      },
      required: ['teamId', 'channelId'],
    },
  },
  {
    name: 'get_teams_channel_message_replies',
    description:
      'Get the replies in a Teams channel thread, oldest first. Needs ChannelMessage.Read.All (admin consent).',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'Team ID' },
        channelId: { type: 'string', description: 'Channel ID' },
        messageId: { type: 'string', description: 'ID of the top-level post' },
        top: {
          type: 'number',
          description: 'Number of replies to return (default: 20, max: 50)',
          minimum: 1,
          maximum: 50,
        },
      },
      required: ['teamId', 'channelId', 'messageId'],
    },
  },
  {
    name: 'send_teams_channel_message',
    description:
      'Post a new top-level message in a Teams channel immediately. Same formatting rules as send_teams_chat_message.',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'Team ID' },
        channelId: { type: 'string', description: 'Channel ID' },
        content: { type: 'string', description: 'Message content' },
        contentType: { type: 'string', enum: ['text', 'html'], description: 'Default: text' },
        subject: { type: 'string', description: 'Optional post subject' },
      },
      required: ['teamId', 'channelId', 'content'],
    },
  },
  {
    name: 'reply_to_teams_channel_message',
    description:
      'Reply in a Teams channel thread immediately. Same formatting rules as send_teams_chat_message.',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: { type: 'string', description: 'Team ID' },
        channelId: { type: 'string', description: 'Channel ID' },
        messageId: { type: 'string', description: 'ID of the top-level post to reply to' },
        content: { type: 'string', description: 'Reply content' },
        contentType: { type: 'string', enum: ['text', 'html'], description: 'Default: text' },
      },
      required: ['teamId', 'channelId', 'messageId', 'content'],
    },
  },
  {
    name: 'search_teams_messages',
    description:
      'Search Teams chat messages, e.g. "budget from:alice". Uses Graph search when the tenant grants ChannelMessage.Read.All; otherwise (the usual case) it scans the latest 50 messages of the 50 most recent chats, matching every term case-insensitively (from:name matches the sender), and says so in a "note". Each hit carries the chatId to open with get_teams_chat_messages.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Search query' },
        top: {
          type: 'number',
          description: 'Number of hits to return (default: 25, max: 25)',
          minimum: 1,
          maximum: 25,
        },
      },
      required: ['query'],
    },
  },
];

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: buildTools(),
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name } = request.params;
  const args = request.params.arguments ?? {};

  try {
    // Signature tools work without a Microsoft sign-in.
    if (name === 'get_signature') return textResult(describeSignature());
    if (name === 'set_signature') {
      const signature = String(args.signature ?? '');
      writeConfigFile(config.configPath, { signature });
      setSignature(signature);
      // Tool descriptions mention the signature; tell clients to re-fetch them.
      await server.sendToolListChanged().catch(() => undefined);
      return textResult(`Signature saved to ${config.configPath}.\n\n${describeSignature()}`);
    }

    if (configError || !authProvider) return textResult(configError ?? SETUP_HELP);

    // Check for a usable sign-in before calling Graph. Without one, start a
    // device-code sign-in and hand its instructions back as the tool result,
    // so a user can sign in from any MCP client without a terminal.
    try {
      await authProvider.getAccessToken();
    } catch (error) {
      if (!(error instanceof AuthRequiredError)) throw error;
      return textResult(await signInInstructions(authProvider, error.message));
    }

    switch (name) {
      case 'outlook_status': {
        const username = await authProvider.getUsername();
        return textResult(
          [
            `Connected to Outlook as ${username}.`,
            `Config file: ${config.configPath}`,
            hasSignature()
              ? `Signature: configured (${signatureText().split('\n').join(' / ')})`
              : 'Signature: none configured (set_signature can add one).',
          ].join('\n'),
        );
      }

      // Message operations
      case 'list_messages':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await messageTools.listMessages(args as any)) },
          ],
        };
      case 'get_message':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await messageTools.getMessage(args.messageId as string, args as any),
              ),
            },
          ],
        };
      case 'create_message':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await messageTools.createMessage(args as any)) },
          ],
        };
      case 'update_message':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await messageTools.updateMessage(args.messageId as string, args.updates as any),
              ),
            },
          ],
        };
      case 'delete_message':
        await messageTools.deleteMessage(args.messageId as string);
        return { content: [{ type: 'text', text: 'Message deleted successfully' }] };
      case 'send_message':
        await messageTools.sendMessage(args.messageId as string);
        return { content: [{ type: 'text', text: 'Message sent successfully' }] };
      case 'send_mail':
        await messageTools.sendMail({
          message: {
            subject: args.subject as string,
            body: args.body as any,
            toRecipients: args.toRecipients as any[],
          },
          saveToSentItems: (args.saveToSentItems as boolean) ?? true,
        });
        return { content: [{ type: 'text', text: 'Mail sent successfully' }] };
      case 'reply_to_message':
        await messageTools.replyToMessage(args.messageId as string, args.comment as string);
        return { content: [{ type: 'text', text: 'Reply sent successfully' }] };
      case 'reply_all_to_message':
        await messageTools.replyAllToMessage(args.messageId as string, args.comment as string);
        return { content: [{ type: 'text', text: 'Reply all sent successfully' }] };
      case 'reply_to_message_draft': {
        const draft = await messageTools.replyToMessageDraft(
          args.messageId as string,
          args.comment as string,
          (args.contentType as string) === 'html',
        );
        return { content: [{ type: 'text', text: JSON.stringify(summarizeDraft(draft)) }] };
      }
      case 'reply_all_to_message_draft': {
        const draft = await messageTools.replyAllToMessageDraft(
          args.messageId as string,
          args.comment as string,
          (args.contentType as string) === 'html',
        );
        return { content: [{ type: 'text', text: JSON.stringify(summarizeDraft(draft)) }] };
      }
      case 'forward_message':
        await messageTools.forwardMessage(
          args.messageId as string,
          args.toRecipients as any[],
          args.comment as string,
        );
        return { content: [{ type: 'text', text: 'Message forwarded successfully' }] };
      case 'forward_message_draft': {
        const draft = await messageTools.forwardMessageDraft(
          args.messageId as string,
          args.toRecipients as any[],
          args.comment as string,
          (args.contentType as string) === 'html',
        );
        return { content: [{ type: 'text', text: JSON.stringify(summarizeDraft(draft)) }] };
      }
      case 'move_message':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await messageTools.moveMessage(
                  args.messageId as string,
                  args.destinationId as string,
                ),
              ),
            },
          ],
        };
      case 'move_folder_messages':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await messageTools.moveFolderMessages(
                  args.sourceFolderId as string,
                  args.destinationId as string,
                  {
                    limit: args.limit as number | undefined,
                    batchSize: args.batchSize as number | undefined,
                    filter: args.filter as string | undefined,
                  },
                ),
              ),
            },
          ],
        };
      case 'copy_message':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await messageTools.copyMessage(
                  args.messageId as string,
                  args.destinationId as string,
                ),
              ),
            },
          ],
        };
      case 'search_messages':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await messageTools.searchMessages(args.searchTerm as string, args as any),
              ),
            },
          ],
        };

      // Folder operations
      case 'list_mail_folders':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await folderTools.listMailFolders(args as any)) },
          ],
        };
      case 'get_mail_folder':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await folderTools.getMailFolder(args.folderId as string, args as any),
              ),
            },
          ],
        };
      case 'create_mail_folder':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await folderTools.createMailFolder(args as any)) },
          ],
        };
      case 'update_mail_folder':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await folderTools.updateMailFolder(args.folderId as string, {
                  displayName: args.displayName as string,
                }),
              ),
            },
          ],
        };
      case 'delete_mail_folder':
        await folderTools.deleteMailFolder(args.folderId as string);
        return { content: [{ type: 'text', text: 'Folder deleted successfully' }] };
      case 'move_mail_folder':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await folderTools.moveMailFolder(
                  args.folderId as string,
                  args.destinationId as string,
                ),
              ),
            },
          ],
        };
      case 'get_well_known_folder':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await folderTools.getWellKnownFolder(args.folderName as any)),
            },
          ],
        };

      // Attachment operations
      case 'list_attachments':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await messageTools.listAttachments(args.messageId as string, args as any),
              ),
            },
          ],
        };
      case 'add_attachment':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await messageTools.addFileAttachment(
                  args.messageId as string,
                  args.filePath as string,
                ),
              ),
            },
          ],
        };
      case 'save_attachment':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await messageTools.saveAttachment(
                  args.messageId as string,
                  args.attachmentId as string,
                  args.destination as string,
                ),
              ),
            },
          ],
        };

      // Child folder operations
      case 'list_child_folders':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await folderTools.listChildFolders(args.folderId as string, args as any),
              ),
            },
          ],
        };

      // Rule operations
      case 'list_message_rules':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await ruleTools.listMessageRules(args as any)) },
          ],
        };
      case 'create_message_rule':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await ruleTools.createMessageRule(args as any)) },
          ],
        };
      case 'delete_message_rule':
        await ruleTools.deleteMessageRule(args.ruleId as string);
        return { content: [{ type: 'text', text: 'Rule deleted successfully' }] };

      // Calendar operations
      case 'list_calendars':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await calendarTools.listCalendars(args as any)) },
          ],
        };
      case 'list_events':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await calendarTools.listEvents(args as any)) },
          ],
        };
      case 'get_calendar_view':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await calendarTools.getCalendarView(args as any)),
            },
          ],
        };
      case 'get_event':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await calendarTools.getEvent(args.eventId as string, args as any),
              ),
            },
          ],
        };
      case 'create_event': {
        const { calendarId, ...event } = args as any;
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await calendarTools.createEvent(event, calendarId)),
            },
          ],
        };
      }
      case 'update_event':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await calendarTools.updateEvent(args.eventId as string, args.updates as any),
              ),
            },
          ],
        };
      case 'delete_event':
        await calendarTools.deleteEvent(args.eventId as string);
        return { content: [{ type: 'text', text: 'Event deleted successfully' }] };
      case 'respond_to_event':
        await calendarTools.respondToEvent(
          args.eventId as string,
          args.response as any,
          args.comment as string,
          (args.sendResponse as boolean) ?? true,
        );
        return { content: [{ type: 'text', text: `Event ${args.response} sent successfully` }] };

      // Contacts operations
      case 'list_contacts':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await contactsTools.listContacts(args as any)) },
          ],
        };
      case 'search_contacts':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await contactsTools.searchContacts(args.searchTerm as string, args as any),
              ),
            },
          ],
        };
      case 'search_people':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await peopleTools.searchPeople(args.searchTerm as string, args as any),
              ),
            },
          ],
        };
      case 'get_contact':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await contactsTools.getContact(args.contactId as string, args as any),
              ),
            },
          ],
        };
      case 'create_contact':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await contactsTools.createContact(args as any)) },
          ],
        };
      case 'update_contact':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await contactsTools.updateContact(args.contactId as string, args.updates as any),
              ),
            },
          ],
        };
      case 'delete_contact':
        await contactsTools.deleteContact(args.contactId as string);
        return { content: [{ type: 'text', text: 'Contact deleted successfully' }] };

      // Teams operations
      case 'list_teams_chats':
        return {
          content: [
            { type: 'text', text: JSON.stringify(await teamsTools.listChats(args as any)) },
          ],
        };
      case 'get_teams_chat_messages':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await teamsTools.getChatMessages(args.chatId as string, args as any),
              ),
            },
          ],
        };
      case 'send_teams_chat_message':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await teamsTools.sendChatMessage(
                  args.chatId as string,
                  args.content as string,
                  args.contentType as any,
                ),
              ),
            },
          ],
        };
      case 'create_teams_chat':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await teamsTools.createChat(args.emails as string[], args.topic as string),
              ),
            },
          ],
        };
      case 'mark_teams_chat_read':
        await teamsTools.markChatRead(args.chatId as string);
        return { content: [{ type: 'text', text: 'Chat marked as read' }] };
      case 'react_to_teams_chat_message':
        await teamsTools.setChatReaction(
          args.chatId as string,
          args.messageId as string,
          args.reactionType as string,
        );
        return { content: [{ type: 'text', text: 'Reaction added' }] };
      case 'list_joined_teams':
        return {
          content: [{ type: 'text', text: JSON.stringify(await teamsTools.listJoinedTeams()) }],
        };
      case 'list_teams_channels':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(await teamsTools.listChannels(args.teamId as string)),
            },
          ],
        };
      case 'list_teams_channel_messages':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await teamsTools.listChannelMessages(
                  args.teamId as string,
                  args.channelId as string,
                  args as any,
                ),
              ),
            },
          ],
        };
      case 'get_teams_channel_message_replies':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await teamsTools.getChannelMessageReplies(
                  args.teamId as string,
                  args.channelId as string,
                  args.messageId as string,
                  args as any,
                ),
              ),
            },
          ],
        };
      case 'send_teams_channel_message':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await teamsTools.sendChannelMessage(
                  args.teamId as string,
                  args.channelId as string,
                  args.content as string,
                  { contentType: args.contentType as any, subject: args.subject as string },
                ),
              ),
            },
          ],
        };
      case 'reply_to_teams_channel_message':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await teamsTools.replyToChannelMessage(
                  args.teamId as string,
                  args.channelId as string,
                  args.messageId as string,
                  args.content as string,
                  args.contentType as any,
                ),
              ),
            },
          ],
        };
      case 'search_teams_messages':
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                await teamsTools.searchMessages(args.query as string, args as any),
              ),
            },
          ],
        };

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    return textResult(
      `Error: ${error instanceof Error ? error.message : String(error)}${permissionHint(name, error)}`,
    );
  }
});

function textResult(text: string) {
  return { content: [{ type: 'text', text }] };
}

function describeSignature(): string {
  if (!hasSignature()) return 'No signature is configured. Nothing is appended to outgoing mail.';
  return `Current signature (plain text):\n${signatureText()}\n\nHTML appended to messages:\n${signatureHtml()}`;
}

async function signInInstructions(auth: GraphAuthProvider, reason: string): Promise<string> {
  const failure = auth.takeLoginError();
  const previous = failure
    ? `The previous sign-in attempt failed: ${explainAuthError(failure)}\n\n`
    : '';
  try {
    const instructions = await auth.beginDeviceLogin();
    return (
      previous +
      `${reason} To connect Outlook, the user needs to sign in: ${instructions}\n\n` +
      'Show the user this link and code, and have them sign in with their work or school account and accept the permissions. Once they say they are done, call this tool again (or call outlook_status to confirm).'
    );
  } catch (error) {
    return `${previous}Could not start a Microsoft sign-in: ${explainAuthError(error)}`;
  }
}

// A 403 from Graph almost always means a permission (scope) this user never
// consented to, often because their organization requires admin approval for
// it. Say so instead of leaving a bare "Forbidden".
function permissionHint(tool: string, error: unknown): string {
  const status = (error as { statusCode?: number })?.statusCode;
  if (status !== 403) return '';
  const area = tool.includes('teams') ? 'Teams' : 'this';
  return `\n\nMicrosoft denied access. This usually means the permission ${area} feature needs was not granted when the user signed in, often because their organization requires an administrator to approve it. Other Outlook tools may still work.`;
}

async function main() {
  await initializeGraphClient();

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Outlook MCP Server running...');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
