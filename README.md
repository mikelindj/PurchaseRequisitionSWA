# Purchase Requisition Form

A static web application for processing purchase requisitions up to $6000, hosted on Azure Static Web Apps with Azure AD authentication and Power Automate integration.

## Features

- ✅ Automatic Azure AD authentication via Azure Static Web Apps Easy Auth (Microsoft organization users only)
- ✅ No sign-in required - automatically authenticates M365 users
- ✅ Form validation based on purchase value
- ✅ File upload support (up to 10 files, 10MB each)
- ✅ Integration with Power Automate HTTP trigger
- ✅ Responsive, modern UI
- ✅ Automated CI/CD via GitHub Actions

## Setup Instructions

### 1. Update Configuration Files

#### Update `app.js`:
- Replace `YOUR_POWER_AUTOMATE_HTTP_URL` with your Power Automate HTTP trigger URL

### 2. Power Automate Setup

1. Create a new Power Automate flow
2. Add an HTTP trigger (When an HTTP request is received)
3. Set Request Body JSON Schema (optional, for validation):
```json
{
  "type": "object",
  "properties": {
    "timestamp": {"type": "string"},
    "user": {
      "type": "object",
      "properties": {
        "name": {"type": "string"},
        "email": {"type": "string"},
        "id": {"type": "string"}
      }
    },
    "purchaseValue": {"type": "string"},
    "budget": {"type": "string"},
    "items": {"type": "string"},
    "totalCost": {"type": "number"},
    "remarks": {"type": "string"},
    "documents": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "name": {"type": "string"},
          "type": {"type": "string"},
          "size": {"type": "number"},
          "content": {"type": "string"}
        }
      }
    }
  }
}
```
4. Copy the HTTP POST URL
5. Add actions to process the data (e.g., send email, save to SharePoint, etc.)

### 3. Azure Static Web App Setup

1. Go to [Azure Portal](https://portal.azure.com) → Create a resource → Static Web App
2. Fill in the details:
   - Subscription: Your subscription
   - Resource group: Create new or use existing
   - Name: Your app name
   - Plan type: Free or Standard
   - Region: Choose your region
   - Source: GitHub
   - Sign in to GitHub and authorize
   - Organization: Your GitHub organization
   - Repository: Your repository
   - Branch: main
   - Build Presets: Custom
   - App location: `/`
   - Api location: (leave empty)
   - Output location: (leave empty)
3. Click "Review + create" then "Create"
4. After creation, go to the Static Web App → **Authentication** tab
5. Click "Add identity provider"
6. Select **"Microsoft"** (Azure Active Directory)
7. Choose **"Create new app registration"** or select an existing one
8. For the new registration:
   - Name: "Purchase Requisition Form" (or your preferred name)
   - Supported account types: **"Current tenant - Single tenant"** (restricts to your organization only)
9. Click "Add" to save
10. Copy the **Deployment token** from the Overview page

**Important**: The authentication is automatically configured to restrict access to users from your Microsoft organization only. Users will be automatically authenticated when they visit the site - no manual sign-in required.

### 4. GitHub Secrets

1. Go to your GitHub repository → Settings → Secrets and variables → Actions
2. Add a new secret:
   - Name: `AZURE_STATIC_WEB_APPS_API_TOKEN`
   - Value: The deployment token from Azure Static Web App

### 5. Deploy

1. Commit and push your changes to the `main` branch
2. GitHub Actions will automatically build and deploy to Azure Static Web Apps
3. The app will be available at `https://your-app.azurestaticapps.net`

## Form Validation Rules

- **Reimbursement Request (<$100)**: Requires at least 1 document
- **<=$1000**: Requires at least 1 quote
- **$1001-$6000**: Requires at least 2 quotations
- **Total Cost**: Must be less than 6001
- **File Limits**: Maximum 10 files, 10MB per file
- **Allowed File Types**: Word, Excel, PPT, PDF, Image, Video, Audio

## File Structure

```
PurchaseRequisitionSWA/
├── index.html              # Main HTML form
├── styles.css              # Styling
├── app.js                  # JavaScript logic (auth, form handling)
├── staticwebapp.config.json # Azure Static Web Apps configuration
├── .github/
│   └── workflows/
│       └── azure-static-web-apps.yml # CI/CD workflow
└── README.md               # This file
```

## How Authentication Works

The app uses **Azure Static Web Apps Easy Auth** which:
- Automatically authenticates users from your Microsoft organization
- No sign-in button required - users are redirected to Microsoft login if not already authenticated
- Restricts access to your organization only (configured via Azure AD app registration)
- User information is automatically retrieved from `/.auth/me` endpoint
- Users remain authenticated across sessions

## Security Notes

- The app uses Azure Static Web Apps Easy Auth to restrict access to your Microsoft organization only
- All unauthenticated requests are automatically redirected to Microsoft login
- All form submissions are sent to your Power Automate endpoint
- Files are converted to base64 for transmission
- Ensure your Power Automate flow has proper authentication/authorization

## Troubleshooting

- **Authentication not working**: 
  - Verify Azure AD app registration is configured correctly in Azure Static Web App → Authentication
  - Ensure "Current tenant - Single tenant" is selected to restrict to your organization
  - Check that the app registration supports the correct account types
- **Users not being auto-authenticated**: 
  - Clear browser cache and cookies
  - Verify the `staticwebapp.config.json` has the correct route configuration
- **Form submission fails**: Check Power Automate URL is correct and the flow is enabled
- **Files not uploading**: Verify file size and type restrictions
- **Deployment fails**: Check GitHub Actions logs and verify the deployment token is set correctly

## License

This project is for internal use only.

