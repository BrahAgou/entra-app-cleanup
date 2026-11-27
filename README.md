# Entra App Cleanup

**Sign in with your M365 admin account and clean up Entra ID app registrations: last sign-in, owners, active secrets, sensitive permissions, user restriction, Conditional Access for workload identities.**

One HTML page, 100% client-side, read-only. No backend, no telemetry: the data goes from Microsoft Graph straight to your browser.

**Try it with sample data:** `https://BrahAgou.github.io/entra-app-cleanup/` → *Try with sample data*

## What you get

- **Overview**: app registrations, sensitive apps, unused apps that still hold a valid secret, candidates to delete, apps without an active owner, apps any user can sign in to. Distribution of sensitivity, last sign-in, owners and credentials, and a priority list.
- **App registrations**: every app with its criticality, verdict, sensitivity, top permissions, last sign-in, owners (disabled owners flagged), secrets and certificates. Search (also by permission), filters, sort.
- **Sensitive apps**: apps ranked by what their **granted** permissions allow, app-only or delegated, with the **user restriction** state ("Assignment required" on the enterprise app) and the main recommendation.
- **Recommendations per app, ranked by criticality**: remove secrets on unused apps, delete, reduce application permissions (with the least-privilege alternative: Sites.Selected, RBAC for Applications in Exchange Online…), replace secrets with certificates or managed identities, require user assignment, review admin consents, user Conditional Access for delegated apps, owners, expiring or useless credentials.
- **Cross-check in the sign-in logs before deleting**: for each delete candidate (one by one or all at once), reads 30 days of interactive, non-interactive, service principal and managed identity sign-ins where the app is the client or the resource. An app with successful sign-ins in the logs is no longer proposed for deletion.
- **Conditional Access for workload identities**: for app-only apps, reads the service principal sign-in logs of the last 30 days, lists the source IPs, groups them into ranges, and generates a named location plus a blocking policy in **report-only** mode (JSON and PowerShell). Warnings come first for every app: licence, scope (no multi-tenant apps or managed identities), 30-day window, dynamic IPs, failed-only IPs left out, widened ranges.
- **KPI tracking**: save a snapshot (JSON) of the current state before cleaning up, reload it later as a baseline and see the change for each KPI (sensitive apps, unused apps with a valid secret, candidates to delete, apps without owner or user restriction, valid and expired secrets…), plus the apps removed, new, or with lower criticality. Every scan is also kept in the browser history (last 8 per tenant). KPIs export to CSV and the deck gets a *Progress since the baseline* slide.
- **Cleanup plan**: four lists in order, plus a PowerShell script where every delete line is **commented out**.
- **Exports**: full CSV (permissions, consents, owners, credentials, IPs, recommendations) and a **PowerPoint deck in English**, with the tenant logo from Entra company branding, ranked by criticality: summary, progress since the baseline, action plan, apps by criticality, one slide per critical app, warnings before Conditional Access, workload identity proposals, method.
- English and French, light and dark themes.

## How it is computed

| | Rule |
| --- | --- |
| Verdict: Delete | No sign-in of any kind for more than *N* days (default 90) and the app is older than *N* days |
| Verdict: Review | Used, but something to fix (owner, credentials, user restriction on a sensitive app…) |
| Sensitivity | Highest level among granted permissions. Critical: control of the directory, roles, apps, all mail or all files. Delegated permissions count one level lower |
| Criticality | Most severe recommendation for the app |
| User restriction | "Assignment required" off on the enterprise app = any user of the tenant can sign in |

## Setup (once per tenant)

1. Entra admin center → **App registrations** → **New registration**, single tenant.
2. **Authentication** → **Add a platform** → **Single-page application** → redirect URI: the exact URL where you open the page, for example `https://brahagou.github.io/entra-app-cleanup/` (the page shows the value to copy).
3. **API permissions** → Microsoft Graph → **Delegated**: `Application.Read.All`, `AuditLog.Read.All`, `Directory.Read.All` → **Grant admin consent**.
4. Open the page, enter the tenant ID and the client ID, sign in.

No client secret: the page uses the authorization code flow with PKCE.

### Or with Microsoft Graph PowerShell

```powershell
# Sign in with the right to create an app registration and grant consent
Connect-MgGraph -Scopes 'Application.ReadWrite.All','DelegatedPermissionGrant.ReadWrite.All'

# Find Microsoft Graph and the ids of the three delegated permissions by name
$graph  = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$scopes = 'Application.Read.All','AuditLog.Read.All','Directory.Read.All'
$access = $graph.Oauth2PermissionScopes | Where-Object Value -in $scopes |
          ForEach-Object { @{ Id = $_.Id; Type = 'Scope' } }

# Create the app registration with an SPA redirect URI (change it if you host the page elsewhere)
$app = New-MgApplication -DisplayName 'Entra App Cleanup' -SignInAudience AzureADMyOrg `
  -Spa @{ RedirectUris = @('https://brahagou.github.io/entra-app-cleanup/') } `
  -RequiredResourceAccess @{ ResourceAppId = $graph.AppId; ResourceAccess = $access }

# Create its service principal, then grant tenant-wide admin consent for the three scopes
$sp = New-MgServicePrincipal -AppId $app.AppId
New-MgOauth2PermissionGrant -ClientId $sp.Id -ConsentType AllPrincipals -ResourceId $graph.Id `
  -Scope ($scopes -join ' ') | Out-Null

# Values to paste into the page
"Client ID: $($app.AppId)"; "Tenant ID: $((Get-MgContext).TenantId)"
```

## Requirements

- A read role: **Global Reader**, or **Security Reader** plus **Reports Reader**.
- **Entra ID P1 or P2** for sign-in activity and sign-in logs. Without it, permissions, owners and credentials still work and the verdict shows *Unknown*.
- **Workload Identities Premium** to create the proposed Conditional Access policies (the analysis itself does not need it).

## Run locally

```bash
cd docs && python -m http.server 8080
# then open http://localhost:8080/ and register http://localhost:8080/ as SPA redirect URI
```

## Host on GitHub Pages

Settings → Pages → *Deploy from a branch* → `main` / `docs`.

## Limits

- Sign-in activity is kept for a limited time: an app used once a year can look inactive. Confirm with the owner before deleting.
- Application permissions are read from the APIs the apps request plus Microsoft Graph, Exchange Online and SharePoint Online.
- An app running on dynamic cloud IPs needs a fixed egress (NAT gateway, firewall) before an IP-based policy makes sense.
- Only app registrations owned by the tenant are listed, not enterprise apps from other publishers.
- A deleted app stays 30 days in the recycle bin (`Restore-MgDirectoryDeletedItem`).

## License

MIT
