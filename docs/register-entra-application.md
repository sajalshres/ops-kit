# Setting up Azure Entra ID App Registration for SharePoint CLI (CI Pipeline)

## 1\. Purpose

This document describes how to create and configure an Azure Entra ID (Azure AD) **application registration** that the SharePoint CLI (`spcli.py`) will use to upload files to a SharePoint Online document library from a CI/CD pipeline.

The app will:

-   Use **client credentials (client ID + client secret)**.
    
-   Call **Microsoft Graph** with **app-only permissions**.
    
-   Have **write access only to specific SharePoint sites** (recommended) or all sites (for initial testing).
    

---

## 2\. Prerequisites

-   Global Administrator or Application Administrator rights in Azure Entra ID.
    
-   SharePoint Online Administrator rights (for per-site permission grants).
    
-   Access to the CI/CD system where the CLI will run (to configure secrets/variables).
    
-   The following information about the target SharePoint site:
    
    -   Site URL, e.g.  
        `https://<tenant>.sharepoint.com/sites/Engineering`
        
    -   Document library name, e.g. `Documents` or `Shared Documents`.
        

---

## 3\. Register the Application in Azure Entra ID

1.  Sign in to the **Azure portal**:
    
    -   Go to **Azure portal → Azure Active Directory (Entra ID)**.
        
2.  Navigate to **App registrations**:
    
    -   Left menu → **Manage → App registrations**.
        
    -   Click **New registration**.
        
3.  Fill out the registration form:
    
    -   **Name**:  
        e.g. `sp-cli-sharepoint-upload` (or your internal naming convention).
        
    -   **Supported account types**:
        
        -   Typically select **Accounts in this organizational directory only**.
            
    -   **Redirect URI**:
        
        -   Leave empty for this scenario (client credentials don’t need it).
            
4.  Click **Register**.
    
5.  After registration, note these values (you’ll need them for CI/CD and the CLI):
    
    -   **Application (client) ID**
        
    -   **Directory (tenant) ID**
        

---

## 4\. Configure Microsoft Graph API Permissions

We’ll set up **application permissions** (no user sign-in required).

### 4.1 Open the API permissions blade

1.  In your app registration, go to:
    
    -   Left menu → **Manage → API permissions**.
        
    -   Click **Add a permission**.
        
    -   Select **Microsoft Graph**.
        
    -   Choose **Application permissions**.
        

### 4.2 Choose permissions

You have two options:

#### Recommended for production (least privilege)

-   Add:
    
    -   `Sites.Selected`
        

This tells Graph that the app *can* be granted access to specific sites, but by default has **no access** until you explicitly grant it per-site.

#### Simpler (but broader) for testing / non-production

-   Add:
    
    -   `Sites.ReadWrite.All`
        

This allows the app to read/write all SharePoint sites in the tenant. Only use this if security is acceptable or temporarily for testing.

### 4.3 Grant admin consent

1.  Still in **API permissions**, click:
    
    -   **Grant admin consent for <Your Organization>**.
        
2.  Confirm when prompted.
    

You should now see the permissions with a green check mark under **Status**.

---

## 5\. (Recommended) Restrict the App to Specific SharePoint Sites (Sites.Selected)

> Skip this section if you chose `Sites.ReadWrite.All` and are OK with tenant-wide access (not recommended for production).

When using `Sites.Selected`, you must now grant the app explicit permissions to each site it should access. The CLI needs **write** access to the target site(s).

The simplest way is via **PnP PowerShell** (run by a SharePoint admin):

1.  Install PnP PowerShell (if not already installed):
    
    ```powershell
    Install-Module PnP.PowerShell -Scope CurrentUser
    ```
    
2.  Connect to your SharePoint admin center:
    
    ```powershell
    # Replace with your tenant name
    Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -Interactive
    ```
    
3.  Grant the app write access to the site (example site URL):
    
    ```powershell
    $appId = "<YOUR-APP-CLIENT-ID>"   # from app registration
    $siteUrl = "https://<tenant>.sharepoint.com/sites/Engineering"
    
    Grant-PnPAzureADAppSitePermission `
        -AppId $appId `
        -Site $siteUrl `
        -DisplayName "sp-cli-sharepoint-upload" `
        -Permissions Write
    ```
    
    Valid `-Permissions` values include `Read` or `Write`. The CLI needs `Write` to upload files.
    
4.  (Optional) Verify the grant:
    
    ```powershell
    Get-PnPAzureADAppSitePermission -Site $siteUrl
    ```
    

You should see an entry for your app with **Write** permissions.

---

## 6\. Create a Client Secret

The CLI uses **client credentials** (client ID + client secret + tenant ID).

1.  In the app registration, go to:
    
    -   Left menu → **Manage → Certificates & secrets**.
        
    -   Under **Client secrets**, click **New client secret**.
        
2.  Configure the secret:
    
    -   **Description**:  
        e.g. `sp-cli-ci-secret-2025-01`.
        
    -   **Expires**:  
        Choose a sensible lifetime (e.g., 6 or 12 months). Shorter is more secure but requires more frequent rotation.
        
3.  Click **Add**.
    
4.  Immediately copy the **Value** of the secret and store it securely:
    
    -   You will **not** be able to see this value again later.
        
    -   This value is what the CLI uses as `SP_CLIENT_SECRET`.
        

> 🔐 Treat this like a password — don’t share it in plain text, and never commit it to source control.

---

## 7\. Configure CI/CD Pipeline Variables

Your pipeline will call the CLI, which expects configuration via **CLI options** and/or **environment variables**.

### 7.1 Environment variables expected by the CLI

Our CLI supports these environment variables:

-   `SP_TENANT_ID` – Tenant ID (Directory ID).
    
-   `SP_CLIENT_ID` – Application (client) ID.
    
-   `SP_CLIENT_SECRET` – Client secret **value**.
    
-   `SP_SITE_URL` – SharePoint site URL, e.g.  
    `https://<tenant>.sharepoint.com/sites/Engineering`
    
-   `SP_LIBRARY` – Document library display name (default: `Documents`).
    
-   `SP_TARGET_FOLDER` – Folder path inside the document library (optional), e.g. `Uploads/BuildArtifacts`.
    

Example mapping in *any* CI system:

-   Create **secured / secret variables**:
    
    -   `SP_TENANT_ID = <Directory (tenant) ID>`
        
    -   `SP_CLIENT_ID = <Application (client) ID>`
        
    -   `SP_CLIENT_SECRET = <client secret value>`
        
    -   `SP_SITE_URL = https://<tenant>.sharepoint.com/sites/Engineering`
        
    -   `SP_LIBRARY = Documents`
        
    -   `SP_TARGET_FOLDER = Uploads/AppName/$(Build.BuildNumber)` (or similar)
        

Then ensure the job that runs the CLI sees these variables as environment variables.

### 7.2 Example: running the CLI in a pipeline step

Assuming the repo contains `spcli.py` and you’ve installed the dependencies (`click`, `requests`, `msal`):

```bash
# In your CI job step
python spcli.py upload \
  --local-folder "./dist" \
  --small-upload-max $((4 * 1024 * 1024)) \
  --chunk-size $((8 * 1024 * 1024))
```

Because the auth and site values are taken from env vars, you don’t need to pass them as CLI args.

---

## 8\. Local Testing (Optional but Recommended)

Before wiring into CI, you can test locally using the same values.

1.  Set environment variables:
    
    ```bash
    export SP_TENANT_ID="<tenant-id>"
    export SP_CLIENT_ID="<client-id>"
    export SP_CLIENT_SECRET="<client-secret-value>"
    export SP_SITE_URL="https://<tenant>.sharepoint.com/sites/Engineering"
    export SP_LIBRARY="Documents"
    export SP_TARGET_FOLDER="Uploads/TestRun"
    ```
    
2.  Run the CLI with `--dry-run` first:
    
    ```bash
    python spcli.py upload --local-folder ./dist --dry-run -v
    ```
    
    This will show what would be uploaded without actually sending anything.
    
3.  If that looks correct, remove `--dry-run`:
    
    ```bash
    python spcli.py upload --local-folder ./dist -v
    ```
    

Check the target SharePoint library to confirm files appear as expected.

---

## 9\. Security & Maintenance Guidelines

-   **Least privilege**:
    
    -   Prefer `Sites.Selected` + site-level grants rather than `Sites.ReadWrite.All`.
        
    -   Only grant **Write** to the specific site(s) needed by the pipeline.
        
-   **Secrets management**:
    
    -   Store client secrets only in secure secret stores / pipeline secret variables.
        
    -   Never commit secrets to Git or share them in plain text.
        
-   **Secret rotation**:
    
    -   Set a calendar reminder before the client secret expiration.
        
    -   Before the old secret expires:
        
        -   Create a **new** client secret.
            
        -   Update the CI/CD secret variable.
            
        -   Re-run a test pipeline.
            
        -   Once confirmed, you may delete the old secret.
            
-   **App usage monitoring**:
    
    -   Periodically review **Sign-in logs** and **Audit logs** for the app ID in Entra ID.
        
-   **Change management**:
    
    -   If the CLI or permissions change, update this doc and inform both:
        
        -   Azure admins (for permissions & secrets).
            
        -   DevOps team (for pipeline variables & usage patterns).
            

---

If you want, I can also draft a “copy-paste” ticket template for your Azure/Entra admin that includes all fields they need to fill out (names, permissions, target site URLs, etc.).