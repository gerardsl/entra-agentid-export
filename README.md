# Export Microsoft Entra Agent ID objects and connections

`Export-EntraAgentIdentities.ps1` downloads the Microsoft Entra Agent ID inventory of a tenant using **Microsoft Graph REST API calls** (no Graph SDK modules required) and documents how everything is connected.

## What is exported

| Object | Graph endpoint |
|---|---|
| Agent identity blueprints | `GET /applications/microsoft.graph.agentIdentityBlueprint` |
| Blueprint principals | `GET /servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal` |
| Agent identities | `GET /servicePrincipals/microsoft.graph.agentIdentity` |
| Agent users | `GET /users/microsoft.graph.agentUser` |

Connections (fetched with JSON `$batch`, 20 requests per call, with paging and 429/5xx retry):

- Hierarchy: blueprint → blueprint principal → agent identity → agent user
- Owners, sponsors, and agent user manager
- Application permissions (`appRoleAssignments` / `appRoleAssignedTo`), resolved to permission names
- Delegated permissions (`oauth2PermissionGrants`)
- Blueprint inheritable permissions and `requiredResourceAccess`
- Blueprint credentials: client secrets, certificates, and federated identity credentials (**metadata only**, no secret values or hints)
- Group, directory role, and administrative unit memberships (`memberOf`)

The script is read-only. It only sends GET requests, plus `POST /$batch` requests that contain GET requests.

## Requirements

- PowerShell 7.2 or later
- One of these authentication options:

### Option A: App-only (recommended for a complete export)

Create an app registration and grant these **application** permissions on Microsoft Graph. Then grant admin consent.

| Permission | Used for |
|---|---|
| `AgentIdentityBlueprint.Read.All` | Blueprints, blueprint owners, and inheritable permissions |
| `AgentIdentityBlueprintPrincipal.Read.All` | Blueprint principals, their owners, and memberships |
| `AgentIdentity.Read.All` | Agent identities, their owners, and memberships |
| `Application.Read.All` | Sponsors, app role assignments, and resource lookups |
| `Directory.Read.All` | Delegated permission grants, agent user memberships, and app roles |
| `User.Read.All` | Agent users, their sponsors, and managers |

> According to the Graph docs, **sponsors of agent identities** can only be read with the `AgentIdentity.ReadWrite.All` application permission. In testing, delegated sign-in with the read scopes below returned them successfully. If the calls fail in your tenant, they're logged in `csv\errors.csv`, and the export continues.

Use a client secret or a certificate. The script uses a PS256 client assertion with `x5t#S256`.

### Option B: Delegated sign-in (browser or device code)

- The signed-in user needs an Entra role that can read Agent ID objects, such as **Agent ID Administrator**. Owners can read their own objects.
- The delegated scopes are the same as Option A and require admin consent.
- By default, the script uses the *Microsoft Graph Command Line Tools* public client. To use your own app registration, pass `-ClientId` and add the `http://localhost` redirect URI (Mobile and desktop applications). For `-DeviceCode`, enable **Allow public client flows**.
- If your tenant's Conditional Access blocks the device code flow ("You don't have access to this"), use the default browser sign-in instead.
- Avoid tokens that contain `Directory.AccessAsUser.All`, such as Azure CLI tokens. Agent ID APIs can reject them.

## Usage

```powershell
# Delegated, browser sign-in (default)
.\Export-EntraAgentIdentities.ps1 -TenantId contoso.onmicrosoft.com

# Delegated, device code
.\Export-EntraAgentIdentities.ps1 -TenantId contoso.onmicrosoft.com -DeviceCode

# App-only, client secret
$secret = Read-Host -AsSecureString 'Client secret'
.\Export-EntraAgentIdentities.ps1 -TenantId <tenant-id> -ClientId <app-id> -ClientSecret $secret

# App-only, certificate from Cert:\CurrentUser\My or Cert:\LocalMachine\My
.\Export-EntraAgentIdentities.ps1 -TenantId <tenant-id> -ClientId <app-id> -CertificateThumbprint <thumbprint>

# App-only, certificate from a PFX file
.\Export-EntraAgentIdentities.ps1 -TenantId <tenant-id> -ClientId <app-id> -CertificatePath .\app.pfx -CertificatePassword (Read-Host -AsSecureString)

# Pre-acquired Graph access token (for example, from Graph Explorer)
.\Export-EntraAgentIdentities.ps1 -AccessToken $token
```

Optional parameters:

| Parameter | Default | Description |
|---|---|---|
| `-OutputPath` | `.\AgentIdExport-<tenantId>-<timestamp>` | Output folder |
| `-ApiVersion` | `v1.0` | `v1.0` or `beta` |
| `-DiagramDetail` | `Full` | `Hierarchy` draws only blueprint → principal → agent identity → agent user. Use it for large tenants. |
| `-CsvDelimiter` | `,` | Use `;` for Excel with European regional settings |
| `-MaxRetries` | `6` | Retries for throttled (429) and transient (5xx) responses |

## Output

| File | Content |
|---|---|
| `agentid-export.json` | Full export: `exportInfo`, nested `hierarchy`, each object with `properties` and `relationships`, `edges`, and `errors` |
| `csv\blueprints.csv`, `blueprintPrincipals.csv`, `agentIdentities.csv`, `agentUsers.csv` | One row per object with summarized connections |
| `csv\ownersAndSponsors.csv` | Owners, sponsors, and managers |
| `csv\appPermissions.csv`, `delegatedPermissions.csv` | Permission grants with resolved permission names |
| `csv\inheritablePermissions.csv`, `requiredResourceAccess.csv` | Blueprint permission configuration |
| `csv\credentials.csv` | Secret, certificate, and federated credential metadata with expiry status |
| `csv\memberships.csv` | Group, role, and administrative unit memberships |
| `csv\edges.csv` | Every connection as `Source → Relationship → Target` |
| `csv\errors.csv` | Requests that failed, such as those with missing permissions |
| `agentid-graph.mmd` | Mermaid diagram source |
| `agentid-graph.html` | Diagram viewer with zoom. It loads Mermaid from `cdn.jsdelivr.net`. If the default layout fails on large graphs, it falls back to the ELK layout, then to a layout without blueprint groups. If the CDN is blocked, open the `.mmd` file in VS Code or at mermaid.live (use the ELK layout for large tenants). |

Blueprints that aren't registered in the tenant appear as *(blueprint not in this tenant)*. For example, multitenant or Microsoft-managed blueprints have only a blueprint principal locally. The JSON field `blueprintHomeTenantId` shows the blueprint's home tenant.
