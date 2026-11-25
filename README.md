# Entra App Cleanup

**Sign in with your M365 admin account and clean up Entra ID app registrations: last sign-in, owners, active secrets.**

One HTML page, 100% client-side, read-only. No backend, no telemetry: the data goes from Microsoft Graph straight to your browser.

**Try it with sample data:** `https://BrahAgou.github.io/entra-app-cleanup/` → *Try with sample data*

## What you get

- **Overview**: number of app registrations, candidates to delete, unused apps that still hold a valid secret, apps without an active owner, apps to review. Distribution of last sign-in, owners and credentials, and a priority list.
- **App registrations**: every app with its verdict, last sign-in, owners (disabled owners flagged), valid and expired secrets and certificates. Search, filters, sort, CSV export. Select a row for the details: sign-in by type (delegated or app-only, as client or as resource), each credential with its expiry and **last use**, and what to do.
- **Cleanup plan**: four lists in order: (1) remove secrets on unused apps, (2) delete unused apps, (3) give an owner, (4) remove expired credentials. Generates a PowerShell script where every delete line is **commented out**.
- English and French, light and dark themes.

## Verdict

| Verdict | Rule |
| --- | --- |
| Delete | No sign-in of any kind for more than *N* days (default 90) and the app is older than *N* days |
| Review | Used, but no active owner, expired credentials, a credential never used, a credential expiring soon or a secret valid over 2 years |
| Keep | Used, owned, clean |
| Unknown | Sign-in activity could not be read |

"Valid secret on unused app" is flagged on its own: it is the real exposure, remove those secrets first.

## Setup (once per tenant)

1. Entra admin center → **App registrations** → **New registration**, single tenant.
2. **Authentication** → **Add a platform** → **Single-page application** → redirect URI: the exact URL where you open the page, for example `https://brahagou.github.io/entra-app-cleanup/` (the page shows the value to copy).
3. **API permissions** → Microsoft Graph → **Delegated**: `Application.Read.All`, `AuditLog.Read.All` → **Grant admin consent**.
4. Open the page, enter the tenant ID and the client ID, sign in.

No client secret: the page uses the authorization code flow with PKCE.

### Or with Microsoft Graph PowerShell

```powershell
# Sign in with the right to create an app registration and grant consent
Connect-MgGraph -Scopes 'Application.ReadWrite.All','DelegatedPermissionGrant.ReadWrite.All'

# Create the app registration with an SPA redirect URI (change it if you host the page elsewhere)
$app = New-MgApplication -DisplayName 'Entra App Cleanup' -SignInAudience AzureADMyOrg `
  -Spa @{ RedirectUris = @('https://brahagou.github.io/entra-app-cleanup/') } `
  -RequiredResourceAccess @{
    ResourceAppId  = '00000003-0000-0000-c000-000000000000'   # Microsoft Graph
    ResourceAccess = @(
      @{ Id = 'c79f8feb-a9db-4090-85f9-90d820caa0eb'; Type = 'Scope' }  # Application.Read.All (delegated)
      @{ Id = 'e4c9e354-4dc5-45b8-9e7c-e1393b0b1a20'; Type = 'Scope' }  # AuditLog.Read.All (delegated)
    ) }

# Create its service principal, then grant tenant-wide admin consent for the two scopes
$sp    = New-MgServicePrincipal -AppId $app.AppId
$graph = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
New-MgOauth2PermissionGrant -ClientId $sp.Id -ConsentType AllPrincipals -ResourceId $graph.Id `
  -Scope 'Application.Read.All AuditLog.Read.All' | Out-Null

# Values to paste into the page
"Client ID: $($app.AppId)"; "Tenant ID: $((Get-MgContext).TenantId)"
```

## Requirements

- A role that can read applications and sign-in reports: **Global Reader**, **Reports Reader** or **Security Reader**.
- **Entra ID P1 or P2** for sign-in activity (`/beta/reports/servicePrincipalSignInActivities` and `/beta/reports/appCredentialSignInActivities`). Without it, the page still lists owners and credentials, and the verdict shows *Unknown*.

## Run locally

```bash
cd docs && python -m http.server 8080
# then open http://localhost:8080/ and register http://localhost:8080/ as SPA redirect URI
```

## Host on GitHub Pages

Settings → Pages → *Deploy from a branch* → `main` / `docs`.

## Limits

- Sign-in activity is kept for a limited time: an app used once a year can look inactive. Confirm with the owner before deleting.
- Only app registrations owned by the tenant are listed, not enterprise apps from other publishers.
- A deleted app stays 30 days in the recycle bin (`Restore-MgDirectoryDeletedItem`).

## License

MIT
