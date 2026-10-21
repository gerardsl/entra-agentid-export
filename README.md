# Entra Agent ID: blueprints and their agent identities

`Export-EntraAgentIdentities.ps1` shows each Microsoft Entra Agent ID **blueprint** with its **agent identities** below it. If an agent identity has **agent users**, they appear below that agent identity:

```
Blueprint
 ├─ Agent identity
 │   └─ Agent user
 └─ Agent identity
```

The script uses read-only **Microsoft Graph REST API calls** and doesn't need any Graph SDK modules.

## What it reads

| Object | Graph endpoint |
|---|---|
| Agent identity blueprints | `GET /applications/microsoft.graph.agentIdentityBlueprint` |
| Blueprint principals | `GET /servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal` (used only to get the names of blueprints that are registered in another tenant, such as the Microsoft Copilot Studio blueprint) |
| Agent identities | `GET /servicePrincipals/microsoft.graph.agentIdentity` |
| Agent users | `GET /users/microsoft.graph.agentUser` |

An agent identity goes under the blueprint whose `appId` matches the agent identity's `agentIdentityBlueprintId`. An agent user goes under the agent identity whose ID matches the agent user's `identityParentId`.

## Requirements

- PowerShell 7.2 or later
- The following Microsoft Graph permissions, with admin consent. Use application permissions for app-only auth or delegated scopes for sign-in.

| Permission | Used for |
|---|---|
| `AgentIdentityBlueprint.Read.All` | Blueprints |
| `AgentIdentityBlueprintPrincipal.Read.All` | Names of blueprints registered in another tenant |
| `AgentIdentity.Read.All` | Agent identities |
| `User.Read.All` | Agent users |

When you sign in as a user:

- You need an Entra role that can read Agent ID objects, such as **Agent ID Administrator**.
- By default, the script uses the *Microsoft Graph Command Line Tools* public client. To use your own app registration, pass `-ClientId` and add the `http://localhost` redirect URI (Mobile and desktop applications). For `-DeviceCode`, also enable **Allow public client flows**.
- If Conditional Access blocks the device code flow and you see "You don't have access to this", use the default browser sign-in.
- Avoid tokens that contain `Directory.AccessAsUser.All`, such as Azure CLI tokens. Agent ID APIs can reject them.

## Usage

```powershell
# Browser sign-in (default)
.\Export-EntraAgentIdentities.ps1 -TenantId contoso.onmicrosoft.com

# Device code sign-in
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
| `-CsvDelimiter` | `,` | Use `;` for Excel with European regional settings |
| `-MaxRetries` | `6` | Retries for throttled (429) and transient (5xx) responses |

## Output

| File | Content |
|---|---|
| `agentid-hierarchy.html` | Each blueprint, with its agent identities below it and their agent users below those. Blueprints with the most agent identities come first. The page has a filter box and an option to hide blueprints that have no agent identities. It doesn't need internet access to open. |
| `agentid-hierarchy.json` | The same hierarchy as nested JSON |
| `agentid-hierarchy.csv` | One row per blueprint, agent identity, and agent user combination |

Some objects need special handling:

- **Blueprints registered in another tenant** have only a blueprint principal in your tenant. They're labeled *Registered in another tenant*. The Microsoft Copilot Studio blueprint is an example.
- **Agent identities whose blueprint isn't visible** are grouped under *Blueprint &lt;appId&gt;*.
- **Agent users whose agent identity can't be found** are listed at the end of the page.
