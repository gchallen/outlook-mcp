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
import * as dotenv from 'dotenv';
import 'isomorphic-fetch';

dotenv.config();

// Pagination limits to prevent excessive context generation
const PAGINATION_LIMITS = {
  MESSAGES: { default: 25, max: 100 },
  SEARCH: { default: 25, max: 50 },
  FOLDERS: { default: 20, max: 50 },
  RULES: { default: 10, max: 25 },
  ATTACHMENTS: { default: 10, max: 25 },
} as const;

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

async function initializeGraphClient() {
  const authProvider = new GraphAuthProvider({
    clientId: process.env.AZURE_CLIENT_ID || '',
    tenantId: process.env.AZURE_TENANT_ID,
    clientSecret: process.env.AZURE_CLIENT_SECRET,
    scopes: process.env.GRAPH_SCOPES?.split(',') || [
      'Mail.ReadWrite',
      'Mail.Send',
      'MailboxSettings.ReadWrite',
    ],
  });

  const client = authProvider.getClient();
  messageTools = new MessageTools(client);
  folderTools = new FolderTools(client);
  ruleTools = new RuleTools(client);
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
        top: { type: 'number', description: 'Number of messages to return (default: 25, max: 100)', minimum: 1, maximum: 100 },
        skip: { type: 'number', description: 'Number of messages to skip for pagination', minimum: 0 },
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
    description: 'Create a draft message',
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
        ccRecipients: { type: 'array', items: { type: 'object' } },
        bccRecipients: { type: 'array', items: { type: 'object' } },
        importance: { type: 'string', enum: ['low', 'normal', 'high'] },
      },
    },
  },
  {
    name: 'update_message',
    description: 'Update message properties',
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
    description: 'Send a new email directly',
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
    description: 'Reply to a message',
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
    description: 'Reply all to a message',
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
    name: 'forward_message',
    description: 'Forward a message',
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
    description: 'Search for messages with pagination',
    inputSchema: {
      type: 'object',
      properties: {
        searchTerm: { type: 'string', description: 'Search term' },
        select: { type: 'string', description: 'Fields to return' },
        top: { type: 'number', description: 'Number of results (default: 25, max: 50)', minimum: 1, maximum: 50 },
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
        top: { type: 'number', description: 'Number of folders (default: 20, max: 50)', minimum: 1, maximum: 50 },
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
    description: 'Get well-known folder (inbox, sentitems, deleteditems, drafts, outbox, junkemail)',
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
        top: { type: 'number', description: 'Number of attachments (default: 10, max: 25)', minimum: 1, maximum: 25 },
        skip: { type: 'number', description: 'Number to skip for pagination', minimum: 0 },
      },
      required: ['messageId'],
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
        top: { type: 'number', description: 'Number of folders (default: 20, max: 50)', minimum: 1, maximum: 50 },
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
        top: { type: 'number', description: 'Number of rules (default: 10, max: 25)', minimum: 1, maximum: 25 },
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
        return { content: [{ type: 'text', text: JSON.stringify(await messageTools.listMessages(args as any)) }] };
      case 'get_message':
        return { content: [{ type: 'text', text: JSON.stringify(await messageTools.getMessage(args.messageId as string, args as any)) }] };
      case 'create_message':
        return { content: [{ type: 'text', text: JSON.stringify(await messageTools.createMessage(args as any)) }] };
      case 'update_message':
        return { content: [{ type: 'text', text: JSON.stringify(await messageTools.updateMessage(args.messageId as string, args.updates as any)) }] };
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
      case 'forward_message':
        await messageTools.forwardMessage(args.messageId as string, args.toRecipients as any[], args.comment as string);
        return { content: [{ type: 'text', text: 'Message forwarded successfully' }] };
      case 'move_message':
        return { content: [{ type: 'text', text: JSON.stringify(await messageTools.moveMessage(args.messageId as string, args.destinationId as string)) }] };
      case 'copy_message':
        return { content: [{ type: 'text', text: JSON.stringify(await messageTools.copyMessage(args.messageId as string, args.destinationId as string)) }] };
      case 'search_messages':
        return { content: [{ type: 'text', text: JSON.stringify(await messageTools.searchMessages(args.searchTerm as string, args as any)) }] };

      // Folder operations
      case 'list_mail_folders':
        return { content: [{ type: 'text', text: JSON.stringify(await folderTools.listMailFolders(args as any)) }] };
      case 'get_mail_folder':
        return { content: [{ type: 'text', text: JSON.stringify(await folderTools.getMailFolder(args.folderId as string, args as any)) }] };
      case 'create_mail_folder':
        return { content: [{ type: 'text', text: JSON.stringify(await folderTools.createMailFolder(args as any)) }] };
      case 'update_mail_folder':
        return { content: [{ type: 'text', text: JSON.stringify(await folderTools.updateMailFolder(args.folderId as string, { displayName: args.displayName as string })) }] };
      case 'delete_mail_folder':
        await folderTools.deleteMailFolder(args.folderId as string);
        return { content: [{ type: 'text', text: 'Folder deleted successfully' }] };
      case 'move_mail_folder':
        return { content: [{ type: 'text', text: JSON.stringify(await folderTools.moveMailFolder(args.folderId as string, args.destinationId as string)) }] };
      case 'get_well_known_folder':
        return { content: [{ type: 'text', text: JSON.stringify(await folderTools.getWellKnownFolder(args.folderName as any)) }] };

      // Attachment operations
      case 'list_attachments':
        return { content: [{ type: 'text', text: JSON.stringify(await messageTools.listAttachments(args.messageId as string, args as any)) }] };

      // Child folder operations
      case 'list_child_folders':
        return { content: [{ type: 'text', text: JSON.stringify(await folderTools.listChildFolders(args.folderId as string, args as any)) }] };

      // Rule operations
      case 'list_message_rules':
        return { content: [{ type: 'text', text: JSON.stringify(await ruleTools.listMessageRules(args as any)) }] };
      case 'create_message_rule':
        return { content: [{ type: 'text', text: JSON.stringify(await ruleTools.createMessageRule(args as any)) }] };
      case 'delete_message_rule':
        await ruleTools.deleteMessageRule(args.ruleId as string);
        return { content: [{ type: 'text', text: 'Rule deleted successfully' }] };

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