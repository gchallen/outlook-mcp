#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
} from '@modelcontextprotocol/sdk/types.js';
import { GraphAuthProvider } from './auth.js';
import { MessageTools } from './tools/message-tools.js';
import { FolderTools } from './tools/folder-tools.js';
import { RuleTools } from './tools/rule-tools.js';
import { CalendarTools } from './tools/calendar-tools.js';
import { ContactsTools } from './tools/contacts-tools.js';
import { PeopleTools } from './tools/people-tools.js';
import * as dotenv from 'dotenv';
import 'isomorphic-fetch';

// quiet: true suppresses dotenv's startup banner. On an MCP stdio server,
// anything written to stdout that isn't JSON-RPC corrupts the protocol stream
// and breaks the connection.
dotenv.config({ quiet: true });

const server = new Server(
  {
    name: 'outlook-mcp',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  },
);

let messageTools: MessageTools;
let folderTools: FolderTools;
let ruleTools: RuleTools;
let calendarTools: CalendarTools;
let contactsTools: ContactsTools;
let peopleTools: PeopleTools;

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

async function initializeGraphClient() {
  const authProvider = new GraphAuthProvider({
    clientId: process.env.AZURE_CLIENT_ID || '',
    tenantId: process.env.AZURE_TENANT_ID,
    scopes: process.env.GRAPH_SCOPES?.split(','),
    cachePath: process.env.OUTLOOK_TOKEN_CACHE,
  });

  const client = authProvider.getClient();
  messageTools = new MessageTools(client);
  folderTools = new FolderTools(client);
  ruleTools = new RuleTools(client);
  calendarTools = new CalendarTools(client);
  contactsTools = new ContactsTools(client);
  peopleTools = new PeopleTools(client);
}

const tools: Tool[] = [
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
    description:
      'Create a draft message. Every body goes out as HTML: a "text" body is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved) and the styled HTML signature is appended either way, so plain prose can simply be sent as "text". For bulleted or numbered lists (or any rich formatting), set body.contentType to "html" and use <ul><li>...</li></ul> or <ol><li>...</li></ol> with <p> for paragraphs — markdown-style "-" bullets do NOT render. Bare URLs in HTML bodies are also auto-hyperlinked (existing <a> tags are left alone). Do not write your own signature or closing sign-off such as "Best, <name>": the server appends the signature that closes the message (a typed-out plain signature is stripped and replaced by the styled one). Use Title Case for the subject line.',
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
                'Body content. For lists/formatting use contentType "html" with <ul>/<li> or <ol>/<li>; markdown-style "-" bullets do NOT render in a "text" body.',
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
      'Update message properties (PATCH). Replacement bodies always go out as HTML: a "text" body is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved), and an HTML body gets the outgoing font, paragraph spacing, and URL auto-linking stamped on it — authored <p>/<ul>/<li> markup needs no inline font styling, and a body that already declares its own font-family is left alone. The signature is NOT added — a PATCH usually replaces the whole body, so include the signature yourself (a typed-out plain signature is upgraded to the styled hyperlinked one). Use contentType "html" with <ul>/<li> for lists. Note: editing a draft can rotate its message ID, so re-fetch the draft (list/get) before a follow-up update.',
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
    description:
      'Send a new email directly. Same formatting rules as create_message: every body goes out as HTML — a "text" body is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved) and the styled signature is appended either way. For bulleted/numbered lists set body.contentType to "html" and use <ul>/<li> or <ol>/<li> — markdown-style "-" bullets do NOT render. Do not write your own signature or closing sign-off such as "Best, <name>": the server appends the signature that closes the message (a typed-out plain signature is stripped and replaced by the styled one). Use Title Case for the subject line.',
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
    description:
      'Reply to a message and send immediately. The comment is treated as plain text and converted to HTML automatically (escaped, bare URLs hyperlinked, line breaks preserved) with the styled signature appended — do not write your own signature.',
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
    description:
      'Reply all to a message and send immediately. The comment is treated as plain text and converted to HTML automatically (escaped, bare URLs hyperlinked, line breaks preserved) with the styled signature appended — do not write your own signature.',
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
    description:
      'Create a reply draft (to sender only) without sending it. Returns the draft so it can be reviewed and sent later. The reply always goes out as HTML with the styled signature appended: a "text" comment is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved). Do not write your own signature — a typed-out plain signature is stripped and replaced by the styled one.',
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
    description:
      'Create a reply-all draft without sending it. Returns the draft so it can be reviewed and sent later. The reply always goes out as HTML with the styled signature appended: a "text" comment is converted automatically (escaped, bare URLs hyperlinked, line breaks preserved). Do not write your own signature — a typed-out plain signature is stripped and replaced by the styled one.',
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
    description:
      'Forward a message and send immediately. The comment is treated as plain text and converted to HTML automatically (escaped, bare URLs hyperlinked, line breaks preserved) with the styled signature appended — do not write your own signature.',
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
      'Full-text search across mail using Microsoft Graph $search (KQL). Prefer property-scoped queries over bare keywords for much better precision. Supported properties (combine space-separated for implicit AND, or with explicit AND/OR/NOT): from:alice@x.com, to:, cc:, subject:budget, body:, participants: (any of from/to/cc), hasAttachment:true, attachment:report.pdf, received>=2024-01-01, received:2024-01-01..2024-01-31, sent>=, importance:high, isRead:false. Do NOT quote simple single-token values like an email address (from:alice@x.com, not from:"alice@x.com"). Only put double quotes around multi-word phrase values: subject:"quarterly report". Do NOT wrap the whole query in outer quotes — the tool adds the outer quoting Graph requires. Examples: from:dean@illinois.edu subject:budget received>=2024-09-01 / hasAttachment:true subject:"final grades". Note: $search results are relevance-ranked and CANNOT be date-sorted or combined with an OData $filter (a Graph limitation, not a tool one); when you need strict newest-first ordering or a hard field filter, use list_messages with filter/orderBy instead. Scope to one folder with folderId (get IDs from list_mail_folders / get_well_known_folder).',
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
];

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools,
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    if (!args) {
      throw new Error('No arguments provided');
    }

    switch (name) {
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

      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${error instanceof Error ? error.message : String(error)}`,
        },
      ],
    };
  }
});

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
