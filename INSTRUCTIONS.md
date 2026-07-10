# Setup Instructions for Outlook MCP Server

This guide will walk you through setting up the necessary Azure App Registration and permissions to use the Outlook MCP Server.

## Email Conventions

- Email subjects should use Title Case.

## Prerequisites

- An Azure account with access to Azure Active Directory
- Administrative permissions or the ability to register applications in your Azure tenant
- A Microsoft 365 account with Outlook/Exchange Online

## Step 1: Create an Azure App Registration

1. **Sign in to the Azure Portal**
   - Go to [https://portal.azure.com](https://portal.azure.com)
   - Sign in with your Azure account

2. **Navigate to Azure Active Directory**
   - In the left sidebar, click "Azure Active Directory"
   - Or search for "Azure Active Directory" in the search bar

3. **Register a New Application**
   - Click "App registrations" in the left menu
   - Click "+ New registration"
   - Fill in the application details:
     - **Name**: "Outlook MCP Server" (or your preferred name)
     - **Supported account types**:
       - Choose "Accounts in this organizational directory only" for single tenant
       - Choose "Accounts in any organizational directory" for multi-tenant
       - Choose "Accounts in any organizational directory and personal Microsoft accounts" for the broadest access
     - **Redirect URI**: Leave blank for now
   - Click "Register"

4. **Note Your Application Details**
   - After registration, you'll see the "Overview" page
   - Copy and save these values (you'll need them later):
     - **Application (client) ID**
     - **Directory (tenant) ID**

## Step 2: Create a Client Secret

1. **Navigate to Certificates & Secrets**
   - In your app registration, click "Certificates & secrets" in the left menu

2. **Create a New Client Secret**
   - Click "+ New client secret"
   - Add a description: "Outlook MCP Server Secret"
   - Choose an expiration period (recommendation: 24 months for production, shorter for development)
   - Click "Add"

3. **Copy the Secret Value**
   - **IMPORTANT**: Copy the secret value immediately and store it securely
   - This value will not be shown again after you navigate away
   - Save this as your `AZURE_CLIENT_SECRET`

## Step 3: Configure API Permissions

1. **Navigate to API Permissions**
   - Click "API permissions" in the left menu

2. **Add Microsoft Graph Permissions**
   - Click "+ Add a permission"
   - Select "Microsoft Graph"
   - Choose "Delegated permissions"
   - Search for and select the following permissions:
     - `Mail.ReadWrite` - Read and write access to user mail
     - `Mail.Send` - Send mail as the user
     - `MailboxSettings.ReadWrite` - Read and write mailbox settings

3. **Grant Admin Consent** (if required)
   - If you're an admin, click "Grant admin consent for [your organization]"
   - If you're not an admin, request admin consent from your IT administrator
   - The permissions should show a green checkmark when properly consented

## Step 4: Configure Authentication (Optional)

For enhanced security and specific redirect scenarios:

1. **Add Redirect URIs** (if using web-based auth flows)
   - Click "Authentication" in the left menu
   - Click "+ Add a platform"
   - Choose the appropriate platform type
   - Add your redirect URIs

2. **Configure Advanced Settings**
   - Enable "Allow public client flows" if you plan to use device code authentication
   - Configure logout URLs if needed

## Step 5: Set Up Environment Variables

Create a `.env` file in your project root with the following values:

```bash
# Required: Azure App Registration Details
AZURE_CLIENT_ID=your_application_client_id_here
AZURE_TENANT_ID=your_directory_tenant_id_here
AZURE_CLIENT_SECRET=your_client_secret_here

# Optional: Custom Graph API scopes (defaults shown)
GRAPH_SCOPES=Mail.ReadWrite,Mail.Send,MailboxSettings.ReadWrite
```

### Finding Your Values

- **AZURE_CLIENT_ID**: Found on the "Overview" page of your app registration as "Application (client) ID"
- **AZURE_TENANT_ID**: Found on the "Overview" page of your app registration as "Directory (tenant) ID"
- **AZURE_CLIENT_SECRET**: The secret value you copied when creating the client secret

## Step 6: Test Your Configuration

1. **Build the Project**

   ```bash
   npm run build
   ```

2. **Run the Server**

   ```bash
   npx outlook-mcp
   ```

3. **Verify Authentication**
   - The server should start without authentication errors
   - If using device code flow, you'll see a login prompt
   - If using client secret flow, authentication should happen automatically

## Authentication Methods

### Client Secret Flow (Recommended for Automation)

- Uses the client secret for automated authentication
- No user interaction required
- Best for server-to-server scenarios
- Requires `AZURE_CLIENT_SECRET` to be set

### Device Code Flow (Interactive)

- Uses device code for interactive authentication
- User must complete authentication in a browser
- Better for development and testing
- Used when `AZURE_CLIENT_SECRET` is not provided

## Common Issues and Solutions

### Issue: "Insufficient privileges to complete the operation"

**Solution**: Ensure admin consent has been granted for the required permissions.

### Issue: "The application 'XXX' is not registered in directory 'YYY'"

**Solution**:

- Verify your `AZURE_TENANT_ID` is correct
- Ensure the app is registered in the correct tenant
- Check if you need to use "common" as the tenant ID for multi-tenant apps

### Issue: "Invalid client secret"

**Solution**:

- Verify the `AZURE_CLIENT_SECRET` is copied correctly
- Check if the secret has expired
- Create a new client secret if needed

### Issue: Authentication timeout or device code issues

**Solution**:

- Try using client secret authentication instead
- Check your network connectivity
- Verify the tenant ID is correct

## Security Best Practices

1. **Protect Your Secrets**
   - Never commit `.env` files to version control
   - Use Azure Key Vault for production deployments
   - Rotate client secrets regularly

2. **Principle of Least Privilege**
   - Only request the minimum permissions needed
   - Consider using application permissions vs. delegated permissions based on your use case

3. **Monitor Usage**
   - Regularly review the "Sign-in logs" in Azure AD
   - Monitor API usage in the Azure portal
   - Set up alerts for unusual activity

4. **Regular Maintenance**
   - Keep client secrets up to date
   - Review and remove unused permissions
   - Update the application registration as needed

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Azure App Registration Documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Microsoft Graph Permissions Reference](https://docs.microsoft.com/en-us/graph/permissions-reference)
- [Azure Identity SDK Documentation](https://docs.microsoft.com/en-us/javascript/api/overview/azure/identity-readme)

## Support

If you encounter issues:

1. Check the Azure portal for any error messages in the "Sign-in logs"
2. Verify all environment variables are set correctly
3. Ensure the application has proper permissions and admin consent
4. Review the application logs for specific error details
