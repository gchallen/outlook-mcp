# Guided Setup: Connect Your AI Assistant to Outlook

This page sets up the Outlook connector with an AI assistant guiding you through each step.
It connects an AI app (Claude, Codex, Cursor, and others) to **your own** Microsoft 365
email, calendar, contacts, and Teams chats, usually without any help from your IT
department.

## For You: Start Here

Open your AI app and paste this message:

> Please help me connect you to my Outlook email. Follow the guided setup instructions at
> https://github.com/gchallen/outlook-mcp/blob/main/docs/GUIDED-SETUP.md step by step.
> I'm not very technical, so explain each step simply and wait for me before moving on.

If your AI app can't open web pages, download this file and attach it to the chat instead.

**What you need:** a work or school Microsoft 365 account (the one you use for Outlook),
about 15 minutes, and a web browser. The setup is free.

**What this does, and doesn't do:** the connector runs on your computer, and your sign-in
stays on your computer. Your AI app reads or sends email only when you ask it to, and you
can undo the whole thing at any time (see [Removing It](#removing-it)).

---

## For the AI Assistant: Instructions

You are helping a person, most likely an educator who may not be technical, set up the
`outlook-mcp` server. Follow these rules throughout:

- **One step at a time.** Explain what the step is for in a sentence or two, give exact
  instructions, then wait for the person to confirm before moving on.
- **Plain language.** Say "the app you registered", not "the service principal". Spell
  out exactly what to click. Expect the person not to know what a terminal is.
- **Check, don't assume.** After each step, check that it worked (ask what they see, or
  run the check given below). If something doesn't match, use
  [Troubleshooting](#troubleshooting) before going on.
- **Never ask for passwords or sign-in codes.** The only value the person needs to give
  you is the *Application (client) ID*, which is not a secret. If they offer a password,
  tell them not to share it.
- **Be honest about dead ends.** Some organizations block self-service setup. If you hit
  one of the dead ends below, say so plainly, don't hunt for workarounds, and offer to
  draft a short note to their IT department.
- **Keep track of the step you're on.** Restarting the AI app (Step 5) may clear this
  conversation. Before asking for a restart, tell the person exactly what to paste
  afterward to pick up where they left off.

### Step 1: Figure Out the Setup Path

Work out which AI app you're running in and whether you can run terminal commands:

| You are running in…                                                         | Path                     |
| --------------------------------------------------------------------------- | ------------------------ |
| Claude Desktop (the Claude app for Mac or Windows), in a regular chat       | **Desktop path**         |
| Claude Code, Codex (CLI or app), Cursor, or another agent that can run commands | **Terminal path**    |
| Anything else (ChatGPT, claude.ai in a browser, a phone app)                | See the note below       |

**Note:** ChatGPT, claude.ai in a web browser, and phone apps can only use connectors
hosted on the internet, and this one runs on the person's own computer. Suggest they
install Claude Desktop (https://claude.ai/download) and start over there, or use one of
the apps listed above.

Also ask which email address they use for Outlook (for example `jane@university.edu`).
You need it in Steps 2 and 4.

### Step 2: Check the Organization Allows Self-Service

**Terminal path:** run this, replacing the domain with the part of their email after the
`@`:

```bash
curl -s https://login.microsoftonline.com/university.edu/v2.0/.well-known/openid-configuration | head -c 300
```

If the output contains `"issuer"`, the organization uses Microsoft 365 and you can
continue. An error means the address isn't a Microsoft 365 account, so this connector
can't be used with it.

**Desktop path:** skip this check. A non-Microsoft 365 address shows up as an error in
Step 6.

Whether the organization lets people register their own apps only shows up in Step 3.

### Step 3: Register an App in Microsoft Entra

This creates a small "app registration" in the organization's Microsoft directory. It
lets the connector ask for permission to reach the person's mailbox. It belongs to the
person, it's free, and nobody else can use it without their sign-in.

Guide them through these steps. The Microsoft site changes now and then, so if a label
doesn't match exactly, ask what they see and help them find the closest match.

1. Open **https://entra.microsoft.com** and sign in with the work or school account.
   - If they see "You don't have access" or are sent to a page about a free trial or
     subscription: try **https://portal.azure.com** instead, search for **App
     registrations** at the top, and continue from step 3 below.
2. In the left-hand menu, open **Entra ID** (on older versions: **Identity** →
   **Applications**), then **App registrations**.
3. Click **+ New registration**.
   - **Dead end:** if the button is greyed out, or a message says they don't have
     permission to register applications, the organization blocks self-service app
     registration. Stop here and offer to draft a note to IT (see
     [Asking IT for Help](#asking-it-for-help)).
4. Fill in the form:
   - **Name:** `Outlook for AI` (any name works)
   - **Supported account types:** *Accounts in this organizational directory only*
     (the first option, "Single tenant")
   - **Redirect URI:** leave it empty
   - Click **Register**.
5. They now see the app's **Overview** page. Ask them to copy the
   **Application (client) ID** (it looks like `1a2b3c4d-1234-5678-9abc-1234567890ab`)
   and paste it to you. **Not** the "Object ID" or the "Directory (tenant) ID".
   - Check: it should be 36 characters, with dashes in those positions.
6. In the app's left-hand menu, click **Authentication**. Find **Allow public client
   flows** (under *Advanced settings*, or on a *Settings* tab in newer versions), set
   **Enable the following mobile and desktop flows** to **Yes**, and click **Save**.
   - This is the step people most often miss. It lets the connector sign in with a
     short code instead of a password.
7. Recommended: in the left-hand menu, click **API permissions** → **+ Add a
   permission** → **Microsoft Graph** → **Delegated permissions**. Search for and tick
   each of these, then click **Add permissions**:
   `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Contacts.ReadWrite`,
   `People.Read`, `User.Read`, `offline_access`, and for Teams: `Chat.ReadWrite`,
   `ChatMessage.Send`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`,
   `ChannelMessage.Send`.
   - They should **not** click "Grant admin consent". They'll approve the permissions
     themselves when they sign in.
   - If this step is confusing, it's fine to skip it: the sign-in screen asks for the
     same permissions either way.

### Step 4: Install the Connector

#### Desktop Path (Claude Desktop)

1. Download `outlook-mcp.mcpb` from
   https://github.com/gchallen/outlook-mcp/releases/latest.
2. Double-click the downloaded file. Claude Desktop opens an install screen. (Or: in
   Claude Desktop, open **Settings** → **Extensions**, and drag the file in.)
3. Click **Install**, then fill in:
   - **Work or school email address**: their Outlook address
   - **Application (client) ID**: the ID from Step 3
   - **Email signature**: optional, leave it empty for now (Step 7 can copy their
     existing one)
   - **Include Microsoft Teams**: leave it on
4. Make sure the extension is switched on. Then go to Step 6. Claude Desktop doesn't
   usually need a restart after installing an extension, but if the Outlook tools don't
   show up, quit Claude Desktop completely and reopen it.

#### Terminal Path

1. Check that Node.js 18 or newer is installed:

   ```bash
   node --version
   ```

   If the command isn't found or the version is below 18, ask them to install the
   **LTS** version from https://nodejs.org (a normal installer: click through with the
   defaults), then open a new terminal or restart the AI app, and check again.

2. Save their settings. Use `--no-login`: sign-in happens later, inside the AI app,
   where the person can see the sign-in code.

   ```bash
   npx -y @gchallen/outlook-mcp setup --email jane@university.edu --client-id 1a2b3c4d-... --no-login
   ```

   It should print "Found your Microsoft 365 organization" and "Settings saved."

3. Register the connector with the AI app you're running in:

   - **Claude Code:** `claude mcp add outlook --scope user -- npx -y @gchallen/outlook-mcp`
   - **Codex:** `codex mcp add outlook -- npx -y @gchallen/outlook-mcp`
   - **Cursor:** add this to `~/.cursor/mcp.json` (create the file if it doesn't exist,
     and merge it with any existing `mcpServers` entries):

     ```json
     { "mcpServers": { "outlook": { "command": "npx", "args": ["-y", "@gchallen/outlook-mcp"] } } }
     ```

   - **Other apps:** most take the same `mcpServers` JSON. Check the app's
     documentation for where the file goes.
   - **Windows:** if the app can't start `npx`, use `"command": "cmd"` and
     `"args": ["/c", "npx", "-y", "@gchallen/outlook-mcp"]`.

### Step 5: Restart the AI App

**Terminal path only.** AI apps load connectors when they start. Before asking for a
restart, give the person this message to paste afterward:

> I've installed the Outlook connector. Please continue the guided setup at
> https://github.com/gchallen/outlook-mcp/blob/main/docs/GUIDED-SETUP.md from Step 6.

Then ask them to quit the AI app completely and start it again. (Claude Code: type
`/exit`, then run `claude` again. Codex: exit and run `codex` again.)

### Step 6: Sign In to Microsoft

1. Call the **`outlook_status`** tool.
2. The first time, it returns a link (https://microsoft.com/devicelogin or
   https://login.microsoft.com/device) and a short code. Show both clearly, and ask the
   person to:
   - open the link in their browser,
   - type in the code,
   - sign in with their work or school account,
   - review the list of permissions and click **Accept**.
3. When they say they're done, call `outlook_status` again. It should say
   "Connected to Outlook as …".

**If they see "Need admin approval"** instead of an Accept button, the organization
requires IT approval for some permissions. Teams permissions are the most common cause.
Try again without Teams:

- **Terminal path:** run
  `npx -y @gchallen/outlook-mcp setup --email … --client-id … --no-teams --no-login`,
  restart the AI app (Step 5), and sign in again.
- **Desktop path:** in **Settings** → **Extensions** → Outlook, turn off **Include
  Microsoft Teams**, then call `outlook_status` again.

If it still asks for admin approval, that's a dead end (see
[Asking IT for Help](#asking-it-for-help)).

If `outlook_status` reports an error about the previous attempt, it explains what went
wrong. Common causes are in [Troubleshooting](#troubleshooting).

### Step 7: Set Up Their Signature (Optional)

The connector can add their signature to every email it writes. Ask whether they'd like
that. If so, offer to copy the one they already use:

1. Use `search_messages` or `list_messages` on the Sent Items folder
   (`get_well_known_folder` with `sentitems`) to find a recent email they wrote.
2. Read it with `get_message` and find the signature block at the end of their text
   (name, title, department, and so on, above any quoted earlier messages).
3. Show them the signature you found and ask whether it's right.
4. Save it with `set_signature`, as plain text with one line per line. Then show them
   the result of `get_signature`.

If they'd rather type one, or don't want one, that's fine. They can change it any time
by asking you.

### Step 8: Try It Out

1. List the subjects of their 5 most recent emails (`list_messages` with `top: 5` and
   `select: "subject,from,receivedDateTime"`) and ask them to check these match their
   inbox.
2. Create a **draft** (not a sent email) to their own address, with the subject
   "Test From My AI Assistant" and a one-line body, using `create_message`. Ask them to
   find it in their Outlook **Drafts** folder and check that it looks right. They can
   delete it afterward.
3. Tell them setup is finished, and suggest a few things to try:
   - "Summarize the unread email in my inbox from today."
   - "Draft a reply to the latest email from [colleague], saying [what to say]."
   - "What's on my calendar this week?"
   - "Find the email about [topic] from last month."

   Also suggest a good habit: have the assistant write **drafts** to review in Outlook,
   rather than sending email directly.

---

## Troubleshooting

| What you see                                                                | What it means and what to do                                                                                                                                                                                      |
| --------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| "New registration" greyed out, or "you do not have permission to register applications" | The organization blocks self-service app registration. A dead end without IT. See [Asking IT for Help](#asking-it-for-help).                                                                     |
| "Need admin approval" when signing in                                       | The organization requires IT approval for some permissions. Try without Teams (Step 6). If that still fails, it's a dead end.                                                                                    |
| AADSTS7000218, or a message about "client_assertion"                        | "Allow public client flows" isn't on. Go back to Step 3, item 6, and save.                                                                                                                                        |
| AADSTS700016, "application was not found in the directory"                  | The client ID is wrong, or the app was registered with a different account. Check for the **Application (client) ID** on the app's Overview page, then run setup again (Terminal path) or update the extension settings (Desktop path). |
| AADSTS50020, or "account does not exist in tenant"                          | They signed in with a different account (often a personal one). Sign in with the same work or school account they used in Step 3.                                                                                |
| The sign-in code expired                                                    | Codes last 15 minutes. Call `outlook_status` again for a new code.                                                                                                                                               |
| The Outlook tools don't show up in the AI app                               | Restart the AI app completely. Terminal path: check that Node 18+ is installed and `npx -y @gchallen/outlook-mcp setup --help` runs.                                                                            |
| "Outlook is not set up yet"                                                 | No client ID is saved. Redo Step 4.                                                                                                                                                                                |
| A Teams tool fails with "Microsoft denied access"                           | The Teams permissions weren't granted. Email, calendar, and contacts still work. Reading Teams channel posts always needs IT approval.                                                                           |
| "Your Microsoft 365 sign-in has expired"                                    | Normal after long breaks or password changes. Follow the link and code it shows to sign in again.                                                                                                                 |

## Asking IT for Help

If the organization blocks self-service, the person can ask IT. Offer to draft a note
along these lines, adjusted to fit:

> **Subject: Request to Register a Personal Outlook App for an AI Assistant**
>
> I'd like to connect an AI assistant to my own Outlook mailbox using the open-source
> outlook-mcp connector (https://github.com/gchallen/outlook-mcp). It runs on my
> computer, signs in as me with delegated permissions only (no application
> permissions, and no access to anyone else's mailbox), and stores its sign-in token
> locally. Could you either allow me to register a single-tenant app with
> "Allow public client flows" enabled, or approve these delegated Microsoft Graph
> permissions for it: Mail.ReadWrite, Mail.Send, Calendars.ReadWrite,
> Contacts.ReadWrite, People.Read, User.Read, offline_access (and optionally
> Chat.ReadWrite, ChatMessage.Send, Team.ReadBasic.All, Channel.ReadBasic.All,
> ChannelMessage.Send for Teams)?

## Removing It

- **Remove the connector:** Claude Desktop: **Settings** → **Extensions** → remove
  Outlook. Claude Code: `claude mcp remove outlook`. Codex: `codex mcp remove outlook`.
  Other apps: delete the `outlook` entry from their config file.
- **Delete the saved settings and sign-in:** delete the `.outlook-mcp` folder in the
  home folder.
- **Revoke access on Microsoft's side:** delete the app registration at
  https://entra.microsoft.com (**App registrations** → the app → **Delete**). This
  immediately ends the connector's access to the mailbox.
