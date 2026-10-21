# Registering the Entra App

Each person using this server registers their own app in their organization's
Microsoft Entra directory. It's free, and most organizations allow it without
an administrator. The app has no client secret: it uses the device-code
sign-in flow, with delegated permissions only.

For a step-by-step version written for non-technical users, including what to
do when an organization blocks self-service, see
[docs/GUIDED-SETUP.md](./docs/GUIDED-SETUP.md).

## Steps

1. Sign in at https://entra.microsoft.com with your work or school account
   (or use https://portal.azure.com and search for **App registrations**).
2. Go to **Entra ID** → **App registrations** → **+ New registration**.
   - **Name:** anything, e.g. `Outlook for AI`.
   - **Supported account types:** *Accounts in this organizational directory
     only* (single tenant).
   - **Redirect URI:** none.
3. On the **Overview** page, copy the **Application (client) ID**. You don't
   need the Directory (tenant) ID: setup looks it up from your email domain.
4. Under **Authentication**, set **Allow public client flows** to **Yes** and
   save. Without this, sign-in fails with `AADSTS7000218`.
5. Optional: under **API permissions** → **Add a permission** → **Microsoft
   Graph** → **Delegated permissions**, add:
   - `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `Contacts.ReadWrite`,
     `People.Read`, `User.Read`, `offline_access`
   - Teams: `Chat.ReadWrite`, `ChatMessage.Send`, `Team.ReadBasic.All`,
     `Channel.ReadBasic.All`, `ChannelMessage.Send`

   Listing them is optional: the sign-in requests these scopes either way and
   you approve them yourself. Don't click "Grant admin consent". It isn't
   needed for your own mailbox.

Then run setup with the client ID:

```bash
npx -y @gchallen/outlook-mcp setup --email you@university.edu --client-id <application-client-id>
```

## When the Organization Blocks It

- **New registration is greyed out**, or there's a "you do not have permission"
  message: the tenant restricts app registration to administrators.
- **"Need admin approval" at sign-in** (`AADSTS65001`, `AADSTS90094`): the
  tenant restricts user consent. Try `setup --no-teams` first, since Teams
  scopes are the most commonly restricted. If it still fails, user consent is
  blocked for this app.

Either way, IT has to help: by registering the app, or by granting admin
consent for the delegated scopes above. `ChannelMessage.Read.All` (reading Teams
channel posts) always needs admin consent. Add it with
`"extraScopes": ["ChannelMessage.Read.All"]` in `~/.outlook-mcp/config.json`
once it's approved.
