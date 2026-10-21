#Requires -Version 7.2
<#
.SYNOPSIS
    Exports Microsoft Entra Agent ID objects (agent identity blueprints, blueprint principals,
    agent identities, agent users) and their connections using Microsoft Graph REST API calls.

.DESCRIPTION
    Core objects:
      * Agent identity blueprints           GET /applications/microsoft.graph.agentIdentityBlueprint
      * Agent identity blueprint principals GET /servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal
      * Agent identities                    GET /servicePrincipals/microsoft.graph.agentIdentity
      * Agent users                         GET /users/microsoft.graph.agentUser

    Connections (fetched with JSON $batch requests):
      * Hierarchy: blueprint -> blueprint principal -> agent identity -> agent user
      * Owners, sponsors and (agent user) manager
      * Application permissions (appRoleAssignments / appRoleAssignedTo)
      * Delegated permissions (oauth2PermissionGrants)
      * Inheritable permissions and requiredResourceAccess on blueprints
      * Credentials on blueprints (client secret / certificate / federated identity credential METADATA only)
      * Group, directory role and administrative unit memberships (memberOf)

    Output (in -OutputPath):
      agentid-export.json   Full export: objects + relationships, nested hierarchy, edges, errors
      csv\*.csv             Flattened tables (one per object type / relationship type)
      agentid-graph.mmd     Mermaid relationship diagram source
      agentid-graph.html    Diagram viewer (loads Mermaid from cdn.jsdelivr.net)

    The script is read-only: it only issues GET requests (plus POST /$batch containing GETs).

.PARAMETER TenantId
    Tenant ID or verified domain. Required for app-only auth. Defaults to 'organizations' for delegated auth.

.PARAMETER ClientId
    Application (client) ID. For delegated auth defaults to the "Microsoft Graph Command Line Tools"
    public client (14d82eec-204b-4c2f-b7e8-296a70dab67e).

.PARAMETER ClientSecret
    Client secret (SecureString recommended; plain string accepted) for app-only auth.

.PARAMETER CertificateThumbprint
    Thumbprint of a certificate (with private key) in Cert:\CurrentUser\My or Cert:\LocalMachine\My.

.PARAMETER CertificatePath
    Path to a .pfx file for app-only auth. Use -CertificatePassword when the file is protected.

.PARAMETER Interactive
    Delegated sign-in in the browser (authorization code + PKCE, loopback redirect http://localhost). Default mode.

.PARAMETER DeviceCode
    Delegated sign-in using the device code flow.

.PARAMETER AccessToken
    A pre-acquired Microsoft Graph access token (SecureString or string). Cannot be refreshed.

.PARAMETER OutputPath
    Output folder. Default: .\AgentIdExport-<tenantId>-<timestamp>

.PARAMETER ApiVersion
    Microsoft Graph version: v1.0 (default) or beta.

.PARAMETER DiagramDetail
    Full (default): hierarchy + owners/sponsors + permissions + credentials + memberships.
    Hierarchy: only blueprint -> principal -> agent identity -> agent user.

.PARAMETER CsvDelimiter
    CSV delimiter. Default ','. Use ';' for Excel with a European regional configuration.

.PARAMETER MaxRetries
    Maximum retries for throttled (429) or transient (5xx) responses.

.EXAMPLE
    .\Export-EntraAgentIdentities.ps1 -TenantId contoso.onmicrosoft.com
    Delegated, interactive browser sign-in.

.EXAMPLE
    .\Export-EntraAgentIdentities.ps1 -TenantId contoso.onmicrosoft.com -DeviceCode

.EXAMPLE
    $secret = Read-Host -AsSecureString 'Client secret'
    .\Export-EntraAgentIdentities.ps1 -TenantId <tenantId> -ClientId <appId> -ClientSecret $secret

.EXAMPLE
    .\Export-EntraAgentIdentities.ps1 -TenantId <tenantId> -ClientId <appId> -CertificateThumbprint <thumbprint>

.EXAMPLE
    .\Export-EntraAgentIdentities.ps1 -TenantId <tenantId> -ClientId <appId> -CertificatePath .\app.pfx -CertificatePassword (Read-Host -AsSecureString)

.EXAMPLE
    .\Export-EntraAgentIdentities.ps1 -AccessToken $token -DiagramDetail Hierarchy -CsvDelimiter ';'
#>
[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(ParameterSetName = 'Interactive')]
    [Parameter(ParameterSetName = 'DeviceCode')]
    [Parameter(ParameterSetName = 'ClientSecret', Mandatory)]
    [Parameter(ParameterSetName = 'CertificateThumbprint', Mandatory)]
    [Parameter(ParameterSetName = 'CertificateFile', Mandatory)]
    [string]$TenantId = 'organizations',

    [Parameter(ParameterSetName = 'Interactive')]
    [Parameter(ParameterSetName = 'DeviceCode')]
    [Parameter(ParameterSetName = 'ClientSecret', Mandatory)]
    [Parameter(ParameterSetName = 'CertificateThumbprint', Mandatory)]
    [Parameter(ParameterSetName = 'CertificateFile', Mandatory)]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'ClientSecret', Mandatory)]
    [object]$ClientSecret,

    [Parameter(ParameterSetName = 'CertificateThumbprint', Mandatory)]
    [string]$CertificateThumbprint,

    [Parameter(ParameterSetName = 'CertificateFile', Mandatory)]
    [string]$CertificatePath,

    [Parameter(ParameterSetName = 'CertificateFile')]
    [object]$CertificatePassword,

    [Parameter(ParameterSetName = 'Interactive')]
    [switch]$Interactive,

    [Parameter(ParameterSetName = 'DeviceCode', Mandatory)]
    [switch]$DeviceCode,

    [Parameter(ParameterSetName = 'AccessToken', Mandatory)]
    [object]$AccessToken,

    [string]$OutputPath,

    [ValidateSet('v1.0', 'beta')]
    [string]$ApiVersion = 'v1.0',

    [ValidateSet('Full', 'Hierarchy')]
    [string]$DiagramDetail = 'Full',

    [ValidateLength(1, 1)]
    [string]$CsvDelimiter = ',',

    [ValidateRange(0, 10)]
    [int]$MaxRetries = 6
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0.0'
$LoginHost = 'https://login.microsoftonline.com'
$GraphHost = 'https://graph.microsoft.com'
$GraphBase = "$GraphHost/$ApiVersion"
$AuthMode = $PSCmdlet.ParameterSetName
if ($AuthMode -in 'Interactive', 'DeviceCode' -and -not $ClientId) {
    $ClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools
}

# Read permissions used by this script (app-only: application permissions; delegated: scopes).
$RequiredPermissions = @(
    'AgentIdentityBlueprint.Read.All'
    'AgentIdentityBlueprintPrincipal.Read.All'
    'AgentIdentity.Read.All'
    'Application.Read.All'
    'Directory.Read.All'
    'User.Read.All'
)
$DelegatedScopes = @($RequiredPermissions | ForEach-Object { "$GraphHost/$_" }) + 'offline_access', 'openid', 'profile'

class GraphApiException : System.Exception {
    [int]$StatusCode
    [string]$ErrorCode
    [string]$Uri
    GraphApiException([string]$message, [int]$statusCode, [string]$errorCode, [string]$uri) : base($message) {
        $this.StatusCode = $statusCode
        $this.ErrorCode = $errorCode
        $this.Uri = $uri
    }
}

#region Generic helpers
function Write-Step([string]$Message) { Write-Host "==> $Message" -ForegroundColor Cyan }

function ConvertTo-PlainText([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [securestring]) { return [System.Net.NetworkCredential]::new('', $Value).Password }
    return [string]$Value
}

function ConvertTo-Base64Url([byte[]]$Bytes) {
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function ConvertFrom-Base64Url([string]$Text) {
    $t = $Text.Replace('-', '+').Replace('_', '/')
    switch ($t.Length % 4) { 2 { $t += '==' } 3 { $t += '=' } }
    return , [Convert]::FromBase64String($t)
}

function Get-Sha256([byte[]]$Bytes) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return , $sha.ComputeHash($Bytes) } finally { $sha.Dispose() }
}

function Get-JwtClaims([string]$Jwt) {
    $parts = $Jwt.Split('.')
    if ($parts.Count -lt 2) { return $null }
    try { return [Text.Encoding]::UTF8.GetString((ConvertFrom-Base64Url $parts[1])) | ConvertFrom-Json } catch { return $null }
}

$script:JsonDateParams = @{}
if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $script:JsonDateParams.DateKind = 'String' }

function ConvertFrom-GraphJson([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $Text | ConvertFrom-Json -Depth 100 @script:JsonDateParams
}

function Format-Value([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') }
    if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ') }
    return $Value
}

function Join-Values([object[]]$Values, [string]$Separator = '; ') {
    (@($Values) | Where-Object { $null -ne $_ -and "$_" -ne '' } | ForEach-Object { "$(Format-Value $_)" }) -join $Separator
}

function Get-ObjectTypeName([object]$DirectoryObject) {
    $t = "$($DirectoryObject.'@odata.type')" -replace '^#?microsoft\.graph\.', ''
    if (-not $t) { return 'directoryObject' }
    return $t
}

function Get-PrincipalLabel([object]$DirectoryObject) {
    if ($null -eq $DirectoryObject) { return $null }
    $name = $DirectoryObject.displayName
    if ($DirectoryObject.userPrincipalName) { $name = "$name <$($DirectoryObject.userPrincipalName)>" }
    if (-not $name) { $name = $DirectoryObject.id }
    return $name
}
#endregion

#region Authentication
$script:Token = @{ Value = $null; ExpiresOn = [datetimeoffset]::MinValue; Refresh = $null; Tenant = $TenantId }
$script:SigningCertificate = $null

function Invoke-TokenEndpoint([hashtable]$Form, [string]$Tenant = $script:Token.Tenant) {
    $resp = Invoke-WebRequest -Method Post -Uri "$LoginHost/$Tenant/oauth2/v2.0/token" -Body $Form `
        -ContentType 'application/x-www-form-urlencoded' -SkipHttpErrorCheck -UseBasicParsing
    [pscustomobject]@{ StatusCode = [int]$resp.StatusCode; Json = ($resp.Content | ConvertFrom-Json) }
}

function Set-TokenFromResponse([object]$Json) {
    if (-not $Json.access_token) { throw "Token endpoint returned no access_token: $($Json.error) $($Json.error_description)" }
    $script:Token.Value = $Json.access_token
    $script:Token.ExpiresOn = [datetimeoffset]::UtcNow.AddSeconds([int]$Json.expires_in)
    if ($Json.refresh_token) { $script:Token.Refresh = $Json.refresh_token }
    $claims = Get-JwtClaims $Json.access_token
    if ($claims.tid) { $script:Token.Tenant = $claims.tid }
}

function Assert-TokenResponse([object]$Response, [string]$What) {
    if ($Response.StatusCode -ne 200) {
        throw "$What failed ($($Response.StatusCode)): $($Response.Json.error) - $($Response.Json.error_description)"
    }
    Set-TokenFromResponse $Response.Json
}

function Get-SigningCertificate {
    if ($script:SigningCertificate) { return $script:SigningCertificate }
    $cert = $null
    if ($CertificateThumbprint) {
        $tp = ($CertificateThumbprint -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
        foreach ($store in 'Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') {
            $cert = Get-ChildItem -Path "$store\$tp" -ErrorAction SilentlyContinue
            if ($cert) { break }
        }
        if (-not $cert) { throw "Certificate with thumbprint $tp not found in CurrentUser\My or LocalMachine\My." }
    }
    else {
        $path = (Resolve-Path -LiteralPath $CertificatePath).Path
        $password = ConvertTo-PlainText $CertificatePassword
        $flags = if ($IsMacOS) { [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet }
        else { [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet }
        $loader = 'System.Security.Cryptography.X509Certificates.X509CertificateLoader' -as [type]
        $cert = if ($loader) { $loader::LoadPkcs12FromFile($path, $password, $flags) }
        else { [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($path, $password, $flags) }
    }
    if (-not $cert.HasPrivateKey) { throw "Certificate $($cert.Thumbprint) has no accessible private key." }
    $script:SigningCertificate = $cert
    return $cert
}

function New-ClientAssertion {
    $cert = Get-SigningCertificate
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
    if (-not $rsa) { throw 'Only RSA certificates are supported for the client assertion.' }
    $now = [datetimeoffset]::UtcNow.ToUnixTimeSeconds()
    $payload = [ordered]@{
        aud = "$LoginHost/$TenantId/oauth2/v2.0/token"; iss = $ClientId; sub = $ClientId
        jti = [guid]::NewGuid().ToString(); nbf = $now; iat = $now; exp = $now + 600
    } | ConvertTo-Json -Compress
    $encodedPayload = ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($payload))
    $x5t = ConvertTo-Base64Url (Get-Sha256 $cert.RawData)

    # PS256 is recommended; fall back to RS256 for keys in legacy CryptoAPI providers that can't do PSS.
    foreach ($alg in 'PS256', 'RS256') {
        $header = [ordered]@{ alg = $alg; typ = 'JWT'; 'x5t#S256' = $x5t } | ConvertTo-Json -Compress
        $signingInput = (ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($header))) + '.' + $encodedPayload
        $padding = if ($alg -eq 'PS256') { [System.Security.Cryptography.RSASignaturePadding]::Pss } else { [System.Security.Cryptography.RSASignaturePadding]::Pkcs1 }
        try {
            $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($signingInput), [System.Security.Cryptography.HashAlgorithmName]::SHA256, $padding)
            return "$signingInput.$(ConvertTo-Base64Url $signature)"
        }
        catch { if ($alg -eq 'RS256') { throw } }
    }
}

function Get-TokenInteractive {
    $verifierBytes = [byte[]]::new(32)
    [System.Security.Cryptography.RandomNumberGenerator]::Fill($verifierBytes)
    $verifier = ConvertTo-Base64Url $verifierBytes
    $challenge = ConvertTo-Base64Url (Get-Sha256 ([Text.Encoding]::ASCII.GetBytes($verifier)))
    $state = [guid]::NewGuid().ToString('N')

    $probe = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $probe.Start(); $port = $probe.LocalEndpoint.Port; $probe.Stop()
    $redirectUri = "http://localhost:$port"

    $query = @{
        client_id = $ClientId; response_type = 'code'; redirect_uri = $redirectUri; response_mode = 'query'
        scope = ($DelegatedScopes -join ' '); state = $state; code_challenge = $challenge
        code_challenge_method = 'S256'; prompt = 'select_account'
    }
    $qs = ($query.GetEnumerator() | ForEach-Object { "$($_.Key)=$([uri]::EscapeDataString($_.Value))" }) -join '&'
    $authorizeUrl = "$LoginHost/$TenantId/oauth2/v2.0/authorize?$qs"

    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("$redirectUri/")
    $listener.Start()
    try {
        Write-Host 'Opening the browser for sign-in. If it does not open, browse to:' -ForegroundColor Yellow
        Write-Host $authorizeUrl
        try { Start-Process $authorizeUrl } catch { Write-Verbose "Could not start browser: $_" }

        $contextTask = $listener.GetContextAsync()
        $deadline = [datetime]::UtcNow.AddMinutes(5)
        while (-not $contextTask.AsyncWaitHandle.WaitOne(500)) {
            if ([datetime]::UtcNow -gt $deadline) { throw 'Timed out waiting for the browser sign-in (5 minutes).' }
        }
        $context = $contextTask.GetAwaiter().GetResult()
        $params = [System.Web.HttpUtility]::ParseQueryString($context.Request.Url.Query)
        $ok = -not $params['error'] -and $params['code'] -and $params['state'] -eq $state
        $html = if ($ok) { '<html><body style="font-family:sans-serif"><h3>Sign-in complete. You can close this window.</h3></body></html>' }
        else { '<html><body style="font-family:sans-serif"><h3>Sign-in failed. Check the PowerShell window.</h3></body></html>' }
        $buffer = [Text.Encoding]::UTF8.GetBytes($html)
        $context.Response.ContentType = 'text/html; charset=utf-8'
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()
    }
    finally { $listener.Stop(); $listener.Close() }

    if ($params['error']) { throw "Sign-in failed: $($params['error']) - $($params['error_description'])" }
    if ($params['state'] -ne $state) { throw 'Sign-in failed: state mismatch.' }

    $resp = Invoke-TokenEndpoint -Tenant $TenantId -Form @{
        client_id = $ClientId; grant_type = 'authorization_code'; code = $params['code']
        redirect_uri = $redirectUri; code_verifier = $verifier; scope = ($DelegatedScopes -join ' ')
    }
    Assert-TokenResponse $resp 'Authorization code redemption'
}

function Get-TokenDeviceCode {
    $resp = Invoke-WebRequest -Method Post -Uri "$LoginHost/$TenantId/oauth2/v2.0/devicecode" -SkipHttpErrorCheck -UseBasicParsing `
        -ContentType 'application/x-www-form-urlencoded' -Body @{ client_id = $ClientId; scope = ($DelegatedScopes -join ' ') }
    $dc = $resp.Content | ConvertFrom-Json
    if ([int]$resp.StatusCode -ne 200) { throw "Device code request failed: $($dc.error) - $($dc.error_description)" }
    Write-Host $dc.message -ForegroundColor Yellow
    $interval = [Math]::Max(1, [int]$dc.interval)
    $deadline = [datetime]::UtcNow.AddSeconds([int]$dc.expires_in)
    while ([datetime]::UtcNow -lt $deadline) {
        Start-Sleep -Seconds $interval
        $r = Invoke-TokenEndpoint -Tenant $TenantId -Form @{
            client_id = $ClientId; grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; device_code = $dc.device_code
        }
        if ($r.StatusCode -eq 200) { Set-TokenFromResponse $r.Json; return }
        switch ($r.Json.error) {
            'authorization_pending' { continue }
            'slow_down' { $interval += 5; continue }
            default { throw "Device code sign-in failed: $($r.Json.error) - $($r.Json.error_description)" }
        }
    }
    throw 'Device code expired before sign-in completed.'
}

function Get-GraphAccessToken([switch]$ForceRefresh) {
    if (-not $ForceRefresh -and $script:Token.Value -and $script:Token.ExpiresOn -gt [datetimeoffset]::UtcNow.AddMinutes(5)) {
        return $script:Token.Value
    }
    switch ($AuthMode) {
        'ClientSecret' {
            $r = Invoke-TokenEndpoint -Tenant $TenantId -Form @{
                client_id = $ClientId; grant_type = 'client_credentials'; scope = "$GraphHost/.default"
                client_secret = (ConvertTo-PlainText $ClientSecret)
            }
            Assert-TokenResponse $r 'Client credentials (secret) token request'
        }
        { $_ -in 'CertificateThumbprint', 'CertificateFile' } {
            $r = Invoke-TokenEndpoint -Tenant $TenantId -Form @{
                client_id = $ClientId; grant_type = 'client_credentials'; scope = "$GraphHost/.default"
                client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
                client_assertion = (New-ClientAssertion)
            }
            Assert-TokenResponse $r 'Client credentials (certificate) token request'
        }
        { $_ -in 'Interactive', 'DeviceCode' } {
            $refreshed = $false
            if ($script:Token.Refresh) {
                $r = Invoke-TokenEndpoint -Form @{
                    client_id = $ClientId; grant_type = 'refresh_token'; refresh_token = $script:Token.Refresh
                    scope = ($DelegatedScopes -join ' ')
                }
                if ($r.StatusCode -eq 200) { Set-TokenFromResponse $r.Json; $refreshed = $true }
            }
            if (-not $refreshed) {
                if ($AuthMode -eq 'DeviceCode') { Get-TokenDeviceCode } else { Get-TokenInteractive }
            }
        }
        'AccessToken' {
            if (-not $script:Token.Value) {
                $script:Token.Value = (ConvertTo-PlainText $AccessToken) -replace '^\s*Bearer\s+', ''
                $claims = Get-JwtClaims $script:Token.Value
                $script:Token.ExpiresOn = if ($claims.exp) { [datetimeoffset]::FromUnixTimeSeconds([long]$claims.exp) } else { [datetimeoffset]::MaxValue }
                if ($claims.tid) { $script:Token.Tenant = $claims.tid }
            }
            if ($script:Token.ExpiresOn -le [datetimeoffset]::UtcNow) { throw 'The supplied access token has expired. Acquire a new token and run again.' }
        }
    }
    return $script:Token.Value
}
#endregion

#region Microsoft Graph requests
function Get-RetryDelaySeconds([object]$RetryAfter, [int]$Attempt) {
    $seconds = 0
    $value = @($RetryAfter)[0]
    if ($null -ne $value -and [int]::TryParse("$value", [ref]$seconds) -and $seconds -gt 0) { return [Math]::Min($seconds, 120) }
    return [Math]::Min(60, [Math]::Pow(2, $Attempt))
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        [object]$Body
    )
    if ($Uri -notmatch '^https://') { $Uri = "$GraphBase$Uri" }
    $ProgressPreference = 'SilentlyContinue'
    $attempt = 0
    while ($true) {
        $attempt++
        $request = @{
            Method = $Method; Uri = $Uri; SkipHttpErrorCheck = $true; UseBasicParsing = $true
            Headers = @{ Authorization = "Bearer $(Get-GraphAccessToken)"; 'OData-Version' = '4.0' }
        }
        if ($null -ne $Body) {
            $request.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 20 -Compress))
            $request.ContentType = 'application/json; charset=utf-8'
        }
        try { $resp = Invoke-WebRequest @request }
        catch {
            if ($attempt -le $MaxRetries) {
                Write-Verbose "Network error on $Method $Uri ($($_.Exception.Message)); retry $attempt"
                Start-Sleep -Seconds (Get-RetryDelaySeconds $null $attempt); continue
            }
            throw
        }
        $status = [int]$resp.StatusCode
        if ($status -ge 200 -and $status -lt 300) { return ConvertFrom-GraphJson $resp.Content }
        if ($status -eq 401 -and $attempt -eq 1 -and $AuthMode -ne 'AccessToken') {
            $null = Get-GraphAccessToken -ForceRefresh; continue
        }
        if (($status -eq 429 -or $status -ge 500) -and $attempt -le $MaxRetries) {
            $delay = Get-RetryDelaySeconds $resp.Headers['Retry-After'] $attempt
            Write-Verbose "HTTP $status on $Method $Uri; waiting $delay s (retry $attempt)"
            Start-Sleep -Seconds $delay; continue
        }
        $graphError = $null
        try { $graphError = (ConvertFrom-GraphJson $resp.Content).error } catch { }
        throw [GraphApiException]::new("$Method $Uri -> HTTP $status $($graphError.code): $($graphError.message)", $status, $graphError.code, $Uri)
    }
}

function Get-GraphCollection([string]$Uri) {
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $page = Invoke-Graph -Uri $next
        foreach ($item in @($page.value)) { if ($null -ne $item) { $items.Add($item) } }
        $next = $page.'@odata.nextLink'
    }
    return , $items.ToArray()
}

# Tries each URI in order; falls back to the next one on 400/404 (e.g. unsupported $select or type cast casing).
function Get-GraphCollectionWithFallback([string[]]$Uris) {
    $lastError = $null
    foreach ($uri in $Uris) {
        try { return , (Get-GraphCollection $uri) }
        catch {
            $lastError = $_
            $status = if ($_.Exception -is [GraphApiException]) { $_.Exception.StatusCode } else { 0 }
            if ($status -notin 400, 404) { break }
            Write-Verbose "Falling back after HTTP $status on $uri"
        }
    }
    throw $lastError
}

function Get-BatchResponseItems([object]$Body) {
    if ($null -eq $Body) { return , @() }
    if ($Body.PSObject.Properties['value'] -and ($null -eq $Body.value -or $Body.value -is [array])) {
        $items = [System.Collections.Generic.List[object]]::new()
        foreach ($item in @($Body.value)) { if ($null -ne $item) { $items.Add($item) } }
        if ($Body.'@odata.nextLink') { foreach ($item in (Get-GraphCollection $Body.'@odata.nextLink')) { $items.Add($item) } }
        return , $items.ToArray()
    }
    return , @($Body)
}

<#
  Runs GET requests through JSON batching (20 per $batch call), following paging and retrying throttled items.
  Each request: [pscustomobject]@{ Key; Url (relative, e.g. /servicePrincipals/{id}/owners); AllowNotFound }
  Returns a hashtable Key -> [pscustomobject]@{ Ok; Status; Items; ErrorCode; ErrorMessage }
#>
function Invoke-GraphBatchGet([object[]]$Requests, [string]$Activity = 'Fetching relationships') {
    $results = @{}
    if (-not $Requests -or $Requests.Count -eq 0) { return $results }
    $pending = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $Requests) { $pending.Add($r) }
    $total = $Requests.Count; $completed = 0; $round = 0
    while ($pending.Count -gt 0) {
        $round++
        $retry = [System.Collections.Generic.List[object]]::new()
        $maxDelay = 0
        for ($offset = 0; $offset -lt $pending.Count; $offset += 20) {
            $chunk = $pending.GetRange($offset, [Math]::Min(20, $pending.Count - $offset))
            $batchRequests = for ($i = 0; $i -lt $chunk.Count; $i++) { @{ id = "$i"; method = 'GET'; url = $chunk[$i].Url } }
            $response = Invoke-Graph -Method POST -Uri "$GraphBase/`$batch" -Body @{ requests = @($batchRequests) }
            foreach ($r in @($response.responses)) {
                $request = $chunk[[int]$r.id]
                $status = [int]$r.status
                if ($status -ge 200 -and $status -lt 300) {
                    $results[$request.Key] = [pscustomobject]@{ Ok = $true; Status = $status; Items = (Get-BatchResponseItems $r.body); ErrorCode = $null; ErrorMessage = $null }
                    $completed++
                }
                elseif (($status -eq 429 -or $status -ge 500) -and $round -le $MaxRetries) {
                    $retry.Add($request)
                    $maxDelay = [Math]::Max($maxDelay, (Get-RetryDelaySeconds $r.headers.'Retry-After' $round))
                }
                elseif ($status -eq 404 -and $request.AllowNotFound) {
                    $results[$request.Key] = [pscustomobject]@{ Ok = $true; Status = $status; Items = @(); ErrorCode = $null; ErrorMessage = $null }
                    $completed++
                }
                else {
                    $results[$request.Key] = [pscustomobject]@{ Ok = $false; Status = $status; Items = @(); ErrorCode = $r.body.error.code; ErrorMessage = $r.body.error.message }
                    $completed++
                }
            }
            Write-Progress -Activity $Activity -Status "$completed / $total" -PercentComplete ([Math]::Min(100, [int](100 * $completed / $total)))
        }
        $pending = $retry
        if ($pending.Count -gt 0) {
            Write-Verbose "$($pending.Count) batched requests throttled; waiting $maxDelay s"
            Start-Sleep -Seconds $maxDelay
        }
    }
    Write-Progress -Activity $Activity -Completed
    return $results
}
#endregion

#region Main: authenticate
Write-Step "Authenticating ($AuthMode)"
if ($AuthMode -eq 'Interactive') { Add-Type -AssemblyName System.Web }
$null = Get-GraphAccessToken
$claims = Get-JwtClaims $script:Token.Value
$tenantGuid = if ($claims.tid) { $claims.tid } else { $TenantId }
$grantedPermissions = @(if ($claims.roles) { $claims.roles } elseif ($claims.scp) { $claims.scp -split ' ' })
$caller = if ($claims.upn) { $claims.upn } elseif ($claims.unique_name) { $claims.unique_name } elseif ($claims.app_displayname) { "$($claims.app_displayname) ($($claims.appid))" } else { $claims.appid }
$tokenType = if ($claims.scp) { 'Delegated' } else { 'Application' }
Write-Host "    Tenant: $tenantGuid | Caller: $caller | Token type: $tokenType"
Write-Host "    Permissions in token: $(if ($grantedPermissions) { $grantedPermissions -join ', ' } else { '(none)' })"

$missing = foreach ($p in $RequiredPermissions) {
    $alternatives = @($p, ($p -replace '\.Read\.All$', '.ReadWrite.All'))
    if ($p -eq 'User.Read.All') { $alternatives += 'Directory.Read.All', 'Directory.ReadWrite.All' }
    if (-not ($grantedPermissions | Where-Object { $_ -in $alternatives })) { $p }
}
if ($missing) { Write-Warning "Token lacks: $($missing -join ', '). Related data will be reported in the errors section." }
if ($grantedPermissions -contains 'Directory.AccessAsUser.All') {
    Write-Warning 'Token contains Directory.AccessAsUser.All, which Agent ID APIs may reject (403). If that happens, use your own app registration (-ClientId) or app-only auth.'
}
if ($tokenType -eq 'Delegated') {
    Write-Verbose 'Graph docs list sponsors of agent identities and blueprint principals as application-permission only; if those calls fail they are logged in errors.csv.'
}
#endregion

#region Main: output folder
if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location) ('AgentIdExport-{0}-{1}' -f $tenantGuid, (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
$csvPath = Join-Path $OutputPath 'csv'
$null = New-Item -ItemType Directory -Path $csvPath -Force
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path
$csvPath = Join-Path $OutputPath 'csv'
#endregion

#region Main: core objects
$exportErrors = [System.Collections.Generic.List[object]]::new()
function Add-ExportError([string]$ObjectType, [string]$ObjectId, [string]$Relation, [string]$Uri, [int]$Status, [string]$Code, [string]$Message) {
    $exportErrors.Add([pscustomobject][ordered]@{
            ObjectType = $ObjectType; ObjectId = $ObjectId; Relation = $Relation; Uri = $Uri
            Status = $Status; ErrorCode = $Code; Message = $Message
        })
}

function Get-CoreCollection([string]$Name, [string[]]$Uris) {
    Write-Step "Listing $Name"
    try {
        $items = Get-GraphCollectionWithFallback $Uris
        Write-Host "    $($items.Count) found"
        return , $items
    }
    catch {
        $ex = $_.Exception
        $status = if ($ex -is [GraphApiException]) { $ex.StatusCode } else { 0 }
        $code = if ($ex -is [GraphApiException]) { $ex.ErrorCode } else { $null }
        Write-Warning "Could not list $Name : $($ex.Message)"
        Add-ExportError $Name '' 'list' $Uris[0] $status $code $ex.Message
        return , @()
    }
}

$aiSelect = 'id,displayName,agentIdentityBlueprintId,accountEnabled,createdDateTime,createdByAppId,disabledByMicrosoftStatus,servicePrincipalType,tags,managerApplications'
$auSelect = 'id,displayName,userPrincipalName,mail,mailNickname,accountEnabled,createdDateTime,identityParentId,userType,jobTitle,department,companyName,usageLocation'

$blueprints = Get-CoreCollection 'agent identity blueprints' @('/applications/microsoft.graph.agentIdentityBlueprint')
$principals = Get-CoreCollection 'agent identity blueprint principals' @('/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal')
$agents = Get-CoreCollection 'agent identities' @(
    "/servicePrincipals/microsoft.graph.agentIdentity?`$select=$aiSelect"
    '/servicePrincipals/microsoft.graph.agentIdentity')
$agentUsers = Get-CoreCollection 'agent users' @(
    "/users/microsoft.graph.agentUser?`$select=$auSelect"
    '/users/microsoft.graph.agentUser'
    '/users/microsoft.graph.AgentUser')

# Never export secret material: drop the secret hint (first characters of client secrets).
foreach ($b in $blueprints) { foreach ($pc in @($b.passwordCredentials)) { if ($pc) { $pc.PSObject.Properties.Remove('hint'); $pc.PSObject.Properties.Remove('secretText') } } }

$organizationName = $null
try { $organizationName = (Invoke-Graph -Uri '/organization?$select=id,displayName').value[0].displayName } catch { Write-Verbose "Organization name unavailable: $($_.Exception.Message)" }
#endregion

#region Main: relationships
$relRequests = [System.Collections.Generic.List[object]]::new()
function Add-RelRequest([string]$ObjectType, [string]$ObjectId, [string]$Relation, [string]$Url, [switch]$AllowNotFound) {
    $relRequests.Add([pscustomobject]@{
            Key = "$ObjectType|$ObjectId|$Relation"; ObjectType = $ObjectType; ObjectId = $ObjectId
            Relation = $Relation; Url = $Url; AllowNotFound = [bool]$AllowNotFound
        })
}

foreach ($b in $blueprints) {
    $typed = "/applications/$($b.id)/microsoft.graph.agentIdentityBlueprint"
    Add-RelRequest Blueprint $b.id owners "$typed/owners"
    Add-RelRequest Blueprint $b.id sponsors "$typed/sponsors"
    Add-RelRequest Blueprint $b.id inheritablePermissions "$typed/inheritablePermissions"
    Add-RelRequest Blueprint $b.id federatedIdentityCredentials "/applications/$($b.id)/federatedIdentityCredentials"
}
foreach ($p in $principals) {
    $typed = "/servicePrincipals/$($p.id)/microsoft.graph.agentIdentityBlueprintPrincipal"
    Add-RelRequest BlueprintPrincipal $p.id owners "$typed/owners"
    Add-RelRequest BlueprintPrincipal $p.id sponsors "$typed/sponsors"
    Add-RelRequest BlueprintPrincipal $p.id memberOf "$typed/memberOf"
    Add-RelRequest BlueprintPrincipal $p.id appRoleAssignments "/servicePrincipals/$($p.id)/appRoleAssignments"
    Add-RelRequest BlueprintPrincipal $p.id appRoleAssignedTo "/servicePrincipals/$($p.id)/appRoleAssignedTo"
    Add-RelRequest BlueprintPrincipal $p.id oauth2PermissionGrants "/servicePrincipals/$($p.id)/oauth2PermissionGrants"
}
foreach ($a in $agents) {
    $typed = "/servicePrincipals/$($a.id)/microsoft.graph.agentIdentity"
    Add-RelRequest AgentIdentity $a.id owners "$typed/owners"
    Add-RelRequest AgentIdentity $a.id sponsors "$typed/sponsors"
    Add-RelRequest AgentIdentity $a.id memberOf "$typed/memberOf"
    Add-RelRequest AgentIdentity $a.id appRoleAssignments "/servicePrincipals/$($a.id)/appRoleAssignments"
    Add-RelRequest AgentIdentity $a.id oauth2PermissionGrants "/servicePrincipals/$($a.id)/oauth2PermissionGrants"
}
foreach ($u in $agentUsers) {
    Add-RelRequest AgentUser $u.id sponsors "/users/$($u.id)/sponsors"
    Add-RelRequest AgentUser $u.id manager "/users/$($u.id)/manager" -AllowNotFound
    Add-RelRequest AgentUser $u.id memberOf "/users/$($u.id)/memberOf"
    Add-RelRequest AgentUser $u.id appRoleAssignments "/users/$($u.id)/appRoleAssignments"
}

Write-Step "Fetching $($relRequests.Count) relationship collections (JSON batching)"
$relResults = Invoke-GraphBatchGet $relRequests.ToArray() 'Fetching relationships'
foreach ($req in $relRequests) {
    $r = $relResults[$req.Key]
    if ($r -and -not $r.Ok) { Add-ExportError $req.ObjectType $req.ObjectId $req.Relation $req.Url $r.Status $r.ErrorCode $r.ErrorMessage }
}

function Get-Rel([string]$ObjectType, [string]$ObjectId, [string]$Relation) {
    $r = $relResults["$ObjectType|$ObjectId|$Relation"]
    if ($r -and $r.Ok) { return , @($r.Items) }
    return , @()
}
#endregion

#region Main: resolve permission resources (service principals) and grant principals
$resourceById = @{}
$resourceByAppId = @{}
foreach ($p in $principals) { $resourceById[$p.id] = $p; if ($p.appId) { $resourceByAppId[$p.appId] = $p } }

$neededResourceIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$neededResourceAppIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$neededPrincipalIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($req in $relRequests) {
    if ($req.Relation -notin 'appRoleAssignments', 'oauth2PermissionGrants') { continue }
    foreach ($item in (Get-Rel $req.ObjectType $req.ObjectId $req.Relation)) {
        if ($item.resourceId) { $null = $neededResourceIds.Add($item.resourceId) }
        if ($item.consentType -eq 'Principal' -and $item.principalId) { $null = $neededPrincipalIds.Add($item.principalId) }
    }
}
foreach ($b in $blueprints) {
    foreach ($ip in (Get-Rel Blueprint $b.id inheritablePermissions)) { if ($ip.resourceAppId) { $null = $neededResourceAppIds.Add($ip.resourceAppId) } }
    foreach ($rra in @($b.requiredResourceAccess)) { if ($rra.resourceAppId) { $null = $neededResourceAppIds.Add($rra.resourceAppId) } }
}

$spSelect = '$select=id,appId,displayName,appRoles,oauth2PermissionScopes,servicePrincipalType'
$lookupRequests = @(
    foreach ($id in $neededResourceIds) { if (-not $resourceById.ContainsKey($id)) { [pscustomobject]@{ Key = "res|$id"; Url = "/servicePrincipals/$id`?$spSelect"; AllowNotFound = $true } } }
    foreach ($appId in $neededResourceAppIds) { if (-not $resourceByAppId.ContainsKey($appId)) { [pscustomobject]@{ Key = "app|$appId"; Url = "/servicePrincipals(appId='$appId')?$spSelect"; AllowNotFound = $true } } }
    foreach ($id in $neededPrincipalIds) { [pscustomobject]@{ Key = "obj|$id"; Url = "/directoryObjects/$id"; AllowNotFound = $true } }
)
$principalNameById = @{}
if ($lookupRequests.Count -gt 0) {
    Write-Step "Resolving $($lookupRequests.Count) resource applications / principals"
    $lookup = Invoke-GraphBatchGet $lookupRequests 'Resolving resources'
    foreach ($key in $lookup.Keys) {
        $r = $lookup[$key]
        if (-not $r.Ok) { Write-Verbose "Lookup $key failed: $($r.Status) $($r.ErrorCode)"; continue }
        $obj = @($r.Items)[0]
        if (-not $obj) { continue }
        if ($key.StartsWith('obj|')) { $principalNameById[$obj.id] = Get-PrincipalLabel $obj; continue }
        $resourceById[$obj.id] = $obj
        if ($obj.appId) { $resourceByAppId[$obj.appId] = $obj }
    }
}

function Resolve-AppRoleValue([object]$Resource, [string]$AppRoleId) {
    if ($AppRoleId -eq '00000000-0000-0000-0000-000000000000') { return '(default access)' }
    $role = @($Resource.appRoles) | Where-Object { $_.id -eq $AppRoleId } | Select-Object -First 1
    if ($role) { if ($role.value) { return $role.value } return $role.displayName }
    return $AppRoleId
}

function Resolve-ScopeValue([object]$Resource, [string]$ScopeId) {
    $scopes = @($Resource.oauth2PermissionScopes) + @($Resource.publishedPermissionScopes)
    $scope = $scopes | Where-Object { $_ -and $_.id -eq $ScopeId } | Select-Object -First 1
    if ($scope) { return $scope.value }
    return $ScopeId
}
#endregion

#region Main: lookups and tables
function New-Index([object[]]$Items, [string]$Property) {
    $index = @{}
    foreach ($item in $Items) {
        $key = "$($item.$Property)"
        if (-not $index.ContainsKey($key)) { $index[$key] = [System.Collections.Generic.List[object]]::new() }
        $index[$key].Add($item)
    }
    return $index
}

$blueprintByAppId = @{}; foreach ($b in $blueprints) { if ($b.appId) { $blueprintByAppId[$b.appId] = $b } }
$principalByAppId = @{}; foreach ($p in $principals) { if ($p.appId) { $principalByAppId[$p.appId] = $p } }
$agentById = @{}; foreach ($a in $agents) { $agentById[$a.id] = $a }
$agentsByBlueprint = New-Index $agents 'agentIdentityBlueprintId'
$usersByAgent = New-Index $agentUsers 'identityParentId'

$knownObjects = @{}
foreach ($b in $blueprints) { $knownObjects[$b.id] = @{ Type = 'Blueprint'; Name = $b.displayName } }
foreach ($p in $principals) { $knownObjects[$p.id] = @{ Type = 'BlueprintPrincipal'; Name = $p.displayName } }
foreach ($a in $agents) { $knownObjects[$a.id] = @{ Type = 'AgentIdentity'; Name = $a.displayName } }
foreach ($u in $agentUsers) { $knownObjects[$u.id] = @{ Type = 'AgentUser'; Name = $u.displayName } }

function Get-NormalizedType([object]$DirectoryObject) {
    if ($DirectoryObject.id -and $knownObjects.ContainsKey($DirectoryObject.id)) { return $knownObjects[$DirectoryObject.id].Type }
    $t = Get-ObjectTypeName $DirectoryObject
    switch ($t) {
        'agentIdentityBlueprintPrincipal' { return 'BlueprintPrincipal' }
        'agentIdentityBlueprint' { return 'Blueprint' }
        default { return $t.Substring(0, 1).ToUpperInvariant() + $t.Substring(1) }
    }
}

$tables = [ordered]@{}
function New-Table([string]$Name, [string[]]$Columns) {
    $tables[$Name] = [pscustomobject]@{ Columns = $Columns; Rows = [System.Collections.Generic.List[object]]::new() }
}
function Add-Row([string]$Name, [hashtable]$Values) {
    $table = $tables[$Name]
    foreach ($k in $Values.Keys) { if ($k -notin $table.Columns) { throw "Internal error: unknown column '$k' for table '$Name'." } }
    $row = [ordered]@{}
    foreach ($c in $table.Columns) { $row[$c] = Format-Value $Values[$c] }
    $table.Rows.Add([pscustomobject]$row)
}

New-Table 'blueprints' @('BlueprintObjectId', 'AppId', 'DisplayName', 'Description', 'CreatedDateTime', 'CreatedByAppId', 'SignInAudience', 'PublisherDomain', 'IdentifierUris', 'Tags', 'BlueprintPrincipalId', 'AgentIdentityCount', 'AgentUserCount', 'Owners', 'Sponsors', 'InheritablePermissions', 'ClientSecretCount', 'CertificateCount', 'FederatedCredentialCount', 'NextCredentialExpiry')
New-Table 'blueprintPrincipals' @('BlueprintPrincipalId', 'AppId', 'DisplayName', 'BlueprintInThisTenant', 'BlueprintObjectId', 'AppOwnerOrganizationId', 'PublisherName', 'AccountEnabled', 'ServicePrincipalType', 'Tags', 'AgentIdentityCount', 'Owners', 'Sponsors', 'AppPermissions', 'DelegatedPermissions', 'MemberOf')
New-Table 'agentIdentities' @('AgentIdentityId', 'DisplayName', 'BlueprintAppId', 'BlueprintName', 'BlueprintPrincipalId', 'BlueprintInThisTenant', 'AccountEnabled', 'CreatedDateTime', 'CreatedByAppId', 'DisabledByMicrosoftStatus', 'Tags', 'AgentUserCount', 'AgentUsers', 'Owners', 'Sponsors', 'AppPermissions', 'DelegatedPermissions', 'MemberOf')
New-Table 'agentUsers' @('AgentUserId', 'DisplayName', 'UserPrincipalName', 'Mail', 'AccountEnabled', 'CreatedDateTime', 'UserType', 'AgentIdentityId', 'AgentIdentityName', 'BlueprintAppId', 'BlueprintName', 'Manager', 'Sponsors', 'AppPermissions', 'MemberOf')
New-Table 'ownersAndSponsors' @('ObjectType', 'ObjectId', 'ObjectName', 'Relationship', 'PrincipalType', 'PrincipalId', 'PrincipalName', 'PrincipalUserPrincipalName')
New-Table 'appPermissions' @('AssignmentId', 'PrincipalType', 'PrincipalId', 'PrincipalName', 'ResourceId', 'ResourceAppId', 'ResourceName', 'AppRoleId', 'AppRoleValue', 'CreatedDateTime', 'Source')
New-Table 'delegatedPermissions' @('GrantId', 'ClientType', 'ClientId', 'ClientName', 'ResourceId', 'ResourceAppId', 'ResourceName', 'ConsentType', 'PrincipalId', 'PrincipalName', 'Scope')
New-Table 'inheritablePermissions' @('BlueprintObjectId', 'BlueprintAppId', 'BlueprintName', 'ResourceAppId', 'ResourceName', 'Kind', 'Scopes')
New-Table 'requiredResourceAccess' @('BlueprintObjectId', 'BlueprintAppId', 'BlueprintName', 'ResourceAppId', 'ResourceName', 'PermissionType', 'PermissionId', 'PermissionValue')
New-Table 'credentials' @('BlueprintObjectId', 'BlueprintAppId', 'BlueprintName', 'CredentialType', 'Name', 'CredentialId', 'StartDateTime', 'EndDateTime', 'Status', 'KeyType', 'KeyUsage', 'CertificateThumbprint', 'Issuer', 'Subject', 'Audiences', 'Description')
New-Table 'memberships' @('MemberType', 'MemberId', 'MemberName', 'TargetType', 'TargetId', 'TargetName', 'SecurityEnabled', 'MailEnabled', 'GroupTypes', 'RoleTemplateId')
New-Table 'edges' @('SourceType', 'SourceId', 'SourceName', 'Relationship', 'TargetType', 'TargetId', 'TargetName', 'Detail')
New-Table 'errors' @('ObjectType', 'ObjectId', 'Relation', 'Uri', 'Status', 'ErrorCode', 'Message')

function Add-Edge([string]$SourceType, [string]$SourceId, [string]$SourceName, [string]$Relationship, [string]$TargetType, [string]$TargetId, [string]$TargetName, [string]$Detail) {
    Add-Row 'edges' @{
        SourceType = $SourceType; SourceId = $SourceId; SourceName = $SourceName; Relationship = $Relationship
        TargetType = $TargetType; TargetId = $TargetId; TargetName = $TargetName; Detail = $Detail
    }
}

function Get-CredentialStatus([object]$EndDateTime) {
    $end = [datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse("$(Format-Value $EndDateTime)", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$end)) { return 'Unknown' }
    $now = [datetimeoffset]::UtcNow
    if ($end -lt $now) { return 'Expired' }
    if ($end -lt $now.AddDays(30)) { return 'ExpiringWithin30Days' }
    return 'Valid'
}
#endregion

#region Main: relationship processing
function Add-PrincipalLinks([string]$ObjectType, [object]$Object, [string]$Relationship, [object[]]$Linked) {
    foreach ($pr in $Linked) {
        if (-not $pr) { continue }
        $prType = Get-NormalizedType $pr
        Add-Row 'ownersAndSponsors' @{
            ObjectType = $ObjectType; ObjectId = $Object.id; ObjectName = $Object.displayName; Relationship = $Relationship
            PrincipalType = $prType; PrincipalId = $pr.id; PrincipalName = $pr.displayName; PrincipalUserPrincipalName = $pr.userPrincipalName
        }
        Add-Edge $prType $pr.id (Get-PrincipalLabel $pr) "${Relationship}Of" $ObjectType $Object.id $Object.displayName ''
    }
}

function Add-Memberships([string]$ObjectType, [object]$Object, [object[]]$Groups) {
    foreach ($g in $Groups) {
        if (-not $g) { continue }
        $targetType = Get-NormalizedType $g
        Add-Row 'memberships' @{
            MemberType = $ObjectType; MemberId = $Object.id; MemberName = $Object.displayName
            TargetType = $targetType; TargetId = $g.id; TargetName = $g.displayName; SecurityEnabled = $g.securityEnabled
            MailEnabled = $g.mailEnabled; GroupTypes = (Join-Values @($g.groupTypes) ','); RoleTemplateId = $g.roleTemplateId
        }
        Add-Edge $ObjectType $Object.id $Object.displayName 'MemberOf' $targetType $g.id $g.displayName ''
    }
}

$seenAssignments = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
function Add-AppRoleAssignment([object]$Assignment, [string]$Source) {
    $resource = $resourceById["$($Assignment.resourceId)"]
    $value = Resolve-AppRoleValue $resource $Assignment.appRoleId
    $Assignment | Add-Member -NotePropertyName appRoleValue -NotePropertyValue $value -Force
    if (-not $seenAssignments.Add("$($Assignment.id)")) { return }
    $principalType = if ($knownObjects.ContainsKey("$($Assignment.principalId)")) { $knownObjects[$Assignment.principalId].Type } else { "$($Assignment.principalType)" }
    $principalName = if ($knownObjects.ContainsKey("$($Assignment.principalId)")) { $knownObjects[$Assignment.principalId].Name } else { $Assignment.principalDisplayName }
    $resourceType = if ($knownObjects.ContainsKey("$($Assignment.resourceId)")) { $knownObjects[$Assignment.resourceId].Type } else { 'Resource' }
    $resourceName = if ($resource.displayName) { $resource.displayName } else { $Assignment.resourceDisplayName }
    Add-Row 'appPermissions' @{
        AssignmentId = $Assignment.id; PrincipalType = $principalType; PrincipalId = $Assignment.principalId; PrincipalName = $principalName
        ResourceId = $Assignment.resourceId; ResourceAppId = $resource.appId; ResourceName = $resourceName
        AppRoleId = $Assignment.appRoleId; AppRoleValue = $value; CreatedDateTime = $Assignment.createdDateTime; Source = $Source
    }
    Add-Edge $principalType $Assignment.principalId $principalName 'AppPermission' $resourceType $Assignment.resourceId $resourceName $value
}

function Add-DelegatedGrant([string]$ObjectType, [object]$Object, [object]$Grant) {
    $resource = $resourceById["$($Grant.resourceId)"]
    $resourceName = if ($resource.displayName) { $resource.displayName } else { $Grant.resourceId }
    $resourceType = if ($knownObjects.ContainsKey("$($Grant.resourceId)")) { $knownObjects[$Grant.resourceId].Type } else { 'Resource' }
    $principalName = if ($Grant.consentType -eq 'Principal') { if ($principalNameById["$($Grant.principalId)"]) { $principalNameById["$($Grant.principalId)"] } else { $Grant.principalId } } else { 'All users' }
    $scope = "$($Grant.scope)".Trim()
    $Grant | Add-Member -NotePropertyName resourceDisplayName -NotePropertyValue $resourceName -Force
    $Grant | Add-Member -NotePropertyName principalDisplayName -NotePropertyValue $principalName -Force
    Add-Row 'delegatedPermissions' @{
        GrantId = $Grant.id; ClientType = $ObjectType; ClientId = $Object.id; ClientName = $Object.displayName
        ResourceId = $Grant.resourceId; ResourceAppId = $resource.appId; ResourceName = $resourceName
        ConsentType = $Grant.consentType; PrincipalId = $Grant.principalId; PrincipalName = $principalName; Scope = $scope
    }
    Add-Edge $ObjectType $Object.id $Object.displayName 'DelegatedPermission' $resourceType $Grant.resourceId $resourceName "$scope [$principalName]"
}

Write-Step 'Building relationship tables'
foreach ($b in $blueprints) {
    $principal = $principalByAppId[$b.appId]
    if ($principal) { Add-Edge 'Blueprint' $b.id $b.displayName 'HasBlueprintPrincipal' 'BlueprintPrincipal' $principal.id $principal.displayName '' }
    Add-PrincipalLinks 'Blueprint' $b 'Owner' (Get-Rel Blueprint $b.id owners)
    Add-PrincipalLinks 'Blueprint' $b 'Sponsor' (Get-Rel Blueprint $b.id sponsors)

    foreach ($ip in (Get-Rel Blueprint $b.id inheritablePermissions)) {
        $resource = $resourceByAppId["$($ip.resourceAppId)"]
        $resourceName = if ($resource.displayName) { $resource.displayName } else { $ip.resourceAppId }
        $scopesObj = $ip.inheritableScopes
        $kind = if ($scopesObj.kind) { $scopesObj.kind } else { (Get-ObjectTypeName $scopesObj) }
        $scopeList = Join-Values @($scopesObj.scopes) ' '
        Add-Row 'inheritablePermissions' @{
            BlueprintObjectId = $b.id; BlueprintAppId = $b.appId; BlueprintName = $b.displayName
            ResourceAppId = $ip.resourceAppId; ResourceName = $resourceName; Kind = $kind; Scopes = $scopeList
        }
        $resourceId = if ($resource.id) { $resource.id } else { $ip.resourceAppId }
        Add-Edge 'Blueprint' $b.id $b.displayName 'InheritablePermission' 'Resource' $resourceId $resourceName (("$kind $scopeList").Trim())
    }

    foreach ($rra in @($b.requiredResourceAccess)) {
        if (-not $rra) { continue }
        $resource = $resourceByAppId["$($rra.resourceAppId)"]
        $resourceName = if ($resource.displayName) { $resource.displayName } else { $rra.resourceAppId }
        $resourceId = if ($resource.id) { $resource.id } else { $rra.resourceAppId }
        foreach ($ra in @($rra.resourceAccess)) {
            if (-not $ra) { continue }
            $isRole = $ra.type -eq 'Role'
            $value = if ($isRole) { Resolve-AppRoleValue $resource $ra.id } else { Resolve-ScopeValue $resource $ra.id }
            $permissionType = if ($isRole) { 'Application' } else { 'Delegated' }
            Add-Row 'requiredResourceAccess' @{
                BlueprintObjectId = $b.id; BlueprintAppId = $b.appId; BlueprintName = $b.displayName; ResourceAppId = $rra.resourceAppId
                ResourceName = $resourceName; PermissionType = $permissionType; PermissionId = $ra.id; PermissionValue = $value
            }
            Add-Edge 'Blueprint' $b.id $b.displayName 'RequiredResourceAccess' 'Resource' $resourceId $resourceName "${permissionType}: $value"
        }
    }

    foreach ($pc in @($b.passwordCredentials)) {
        if (-not $pc) { continue }
        Add-Row 'credentials' @{
            BlueprintObjectId = $b.id; BlueprintAppId = $b.appId; BlueprintName = $b.displayName; CredentialType = 'ClientSecret'
            Name = $pc.displayName; CredentialId = $pc.keyId; StartDateTime = $pc.startDateTime; EndDateTime = $pc.endDateTime
            Status = (Get-CredentialStatus $pc.endDateTime)
        }
    }
    foreach ($kc in @($b.keyCredentials)) {
        if (-not $kc) { continue }
        Add-Row 'credentials' @{
            BlueprintObjectId = $b.id; BlueprintAppId = $b.appId; BlueprintName = $b.displayName; CredentialType = 'Certificate'
            Name = $kc.displayName; CredentialId = $kc.keyId; StartDateTime = $kc.startDateTime; EndDateTime = $kc.endDateTime
            Status = (Get-CredentialStatus $kc.endDateTime); KeyType = $kc.type; KeyUsage = $kc.usage; CertificateThumbprint = $kc.customKeyIdentifier
        }
    }
    foreach ($fic in (Get-Rel Blueprint $b.id federatedIdentityCredentials)) {
        Add-Row 'credentials' @{
            BlueprintObjectId = $b.id; BlueprintAppId = $b.appId; BlueprintName = $b.displayName; CredentialType = 'FederatedIdentityCredential'
            Name = $fic.name; CredentialId = $fic.id; Status = 'Active'; Issuer = $fic.issuer; Subject = $fic.subject
            Audiences = (Join-Values @($fic.audiences)); Description = $fic.description
        }
        Add-Edge 'Blueprint' $b.id $b.displayName 'FederatedCredential' 'FederatedIdentity' $fic.id "$($fic.name) ($($fic.issuer))" "$($fic.issuer) | $($fic.subject)"
    }
}

foreach ($p in $principals) {
    Add-PrincipalLinks 'BlueprintPrincipal' $p 'Owner' (Get-Rel BlueprintPrincipal $p.id owners)
    Add-PrincipalLinks 'BlueprintPrincipal' $p 'Sponsor' (Get-Rel BlueprintPrincipal $p.id sponsors)
    Add-Memberships 'BlueprintPrincipal' $p (Get-Rel BlueprintPrincipal $p.id memberOf)
    foreach ($ara in (Get-Rel BlueprintPrincipal $p.id appRoleAssignments)) { Add-AppRoleAssignment $ara 'appRoleAssignments' }
    foreach ($ara in (Get-Rel BlueprintPrincipal $p.id appRoleAssignedTo)) { Add-AppRoleAssignment $ara 'appRoleAssignedTo' }
    foreach ($g in (Get-Rel BlueprintPrincipal $p.id oauth2PermissionGrants)) { Add-DelegatedGrant 'BlueprintPrincipal' $p $g }
}

foreach ($a in $agents) {
    $parentAppId = "$($a.agentIdentityBlueprintId)"
    $principal = $principalByAppId[$parentAppId]
    $blueprint = $blueprintByAppId[$parentAppId]
    if ($principal) { Add-Edge 'BlueprintPrincipal' $principal.id $principal.displayName 'HasAgentIdentity' 'AgentIdentity' $a.id $a.displayName '' }
    elseif ($blueprint) { Add-Edge 'Blueprint' $blueprint.id $blueprint.displayName 'HasAgentIdentity' 'AgentIdentity' $a.id $a.displayName '' }
    elseif ($parentAppId) { Add-Edge 'Blueprint' $parentAppId "Blueprint $parentAppId (not readable in this tenant)" 'HasAgentIdentity' 'AgentIdentity' $a.id $a.displayName '' }
    Add-PrincipalLinks 'AgentIdentity' $a 'Owner' (Get-Rel AgentIdentity $a.id owners)
    Add-PrincipalLinks 'AgentIdentity' $a 'Sponsor' (Get-Rel AgentIdentity $a.id sponsors)
    Add-Memberships 'AgentIdentity' $a (Get-Rel AgentIdentity $a.id memberOf)
    foreach ($ara in (Get-Rel AgentIdentity $a.id appRoleAssignments)) { Add-AppRoleAssignment $ara 'appRoleAssignments' }
    foreach ($g in (Get-Rel AgentIdentity $a.id oauth2PermissionGrants)) { Add-DelegatedGrant 'AgentIdentity' $a $g }
}

foreach ($u in $agentUsers) {
    $parent = $agentById["$($u.identityParentId)"]
    if ($parent) { Add-Edge 'AgentIdentity' $parent.id $parent.displayName 'HasAgentUser' 'AgentUser' $u.id $u.displayName '' }
    Add-PrincipalLinks 'AgentUser' $u 'Sponsor' (Get-Rel AgentUser $u.id sponsors)
    Add-PrincipalLinks 'AgentUser' $u 'Manager' (Get-Rel AgentUser $u.id manager)
    Add-Memberships 'AgentUser' $u (Get-Rel AgentUser $u.id memberOf)
    foreach ($ara in (Get-Rel AgentUser $u.id appRoleAssignments)) { Add-AppRoleAssignment $ara 'appRoleAssignments' }
}

foreach ($e in $exportErrors) {
    Add-Row 'errors' @{ ObjectType = $e.ObjectType; ObjectId = $e.ObjectId; Relation = $e.Relation; Uri = $e.Uri; Status = $e.Status; ErrorCode = $e.ErrorCode; Message = $e.Message }
}
#endregion

#region Main: core object tables
function Get-Indexed([hashtable]$Index, [object]$Key) {
    $k = "$Key"
    if ($Index.ContainsKey($k)) { return $Index[$k].ToArray() }
}

$linksByObject = New-Index $tables.ownersAndSponsors.Rows.ToArray() 'ObjectId'
$appPermsByPrincipal = New-Index $tables.appPermissions.Rows.ToArray() 'PrincipalId'
$delegatedByClient = New-Index $tables.delegatedPermissions.Rows.ToArray() 'ClientId'
$membershipsByMember = New-Index $tables.memberships.Rows.ToArray() 'MemberId'
$credsByBlueprint = New-Index $tables.credentials.Rows.ToArray() 'BlueprintObjectId'
$inheritableByBlueprint = New-Index $tables.inheritablePermissions.Rows.ToArray() 'BlueprintObjectId'

function Get-LinkNames([string]$ObjectId, [string]$Relationship) {
    Join-Values @(Get-Indexed $linksByObject $ObjectId | Where-Object { $_.Relationship -eq $Relationship } |
            ForEach-Object { if ($_.PrincipalUserPrincipalName) { $_.PrincipalUserPrincipalName } else { $_.PrincipalName } })
}
function Get-AppPermSummary([string]$Id) { Join-Values @(Get-Indexed $appPermsByPrincipal $Id | ForEach-Object { "$($_.ResourceName): $($_.AppRoleValue)" }) }
function Get-DelegatedSummary([string]$Id) { Join-Values @(Get-Indexed $delegatedByClient $Id | ForEach-Object { "$($_.ResourceName): $($_.Scope) [$($_.PrincipalName)]" }) }
function Get-MemberOfSummary([string]$Id) { Join-Values @(Get-Indexed $membershipsByMember $Id | ForEach-Object { "$($_.TargetName) ($($_.TargetType))" }) }
function Get-BlueprintName([string]$AppId) {
    if ($blueprintByAppId[$AppId]) { return $blueprintByAppId[$AppId].displayName }
    if ($principalByAppId[$AppId]) { return $principalByAppId[$AppId].displayName }
    return $null
}

foreach ($b in $blueprints) {
    $principal = $principalByAppId["$($b.appId)"]
    $bAgents = @(Get-Indexed $agentsByBlueprint $b.appId)
    $userCount = 0; foreach ($a in $bAgents) { $userCount += @(Get-Indexed $usersByAgent $a.id).Count }
    $creds = @(Get-Indexed $credsByBlueprint $b.id)
    $nextExpiry = $creds | Where-Object { $_.EndDateTime -and $_.Status -ne 'Expired' } | Sort-Object EndDateTime | Select-Object -First 1 -ExpandProperty EndDateTime
    Add-Row 'blueprints' @{
        BlueprintObjectId = $b.id; AppId = $b.appId; DisplayName = $b.displayName; Description = $b.description
        CreatedDateTime = $b.createdDateTime; CreatedByAppId = $b.createdByAppId; SignInAudience = $b.signInAudience
        PublisherDomain = $b.publisherDomain; IdentifierUris = (Join-Values @($b.identifierUris)); Tags = (Join-Values @($b.tags))
        BlueprintPrincipalId = $principal.id; AgentIdentityCount = $bAgents.Count; AgentUserCount = $userCount
        Owners = (Get-LinkNames $b.id 'Owner'); Sponsors = (Get-LinkNames $b.id 'Sponsor')
        InheritablePermissions = (Join-Values @(Get-Indexed $inheritableByBlueprint $b.id | ForEach-Object { "$($_.ResourceName): $($_.Kind) $($_.Scopes)".Trim() }))
        ClientSecretCount = @($creds | Where-Object { $_.CredentialType -eq 'ClientSecret' }).Count
        CertificateCount = @($creds | Where-Object { $_.CredentialType -eq 'Certificate' }).Count
        FederatedCredentialCount = @($creds | Where-Object { $_.CredentialType -eq 'FederatedIdentityCredential' }).Count
        NextCredentialExpiry = $nextExpiry
    }
}

foreach ($p in $principals) {
    $b = $blueprintByAppId["$($p.appId)"]
    Add-Row 'blueprintPrincipals' @{
        BlueprintPrincipalId = $p.id; AppId = $p.appId; DisplayName = $p.displayName; BlueprintInThisTenant = [bool]$b
        BlueprintObjectId = $b.id; AppOwnerOrganizationId = $p.appOwnerOrganizationId; PublisherName = $p.publisherName
        AccountEnabled = $p.accountEnabled; ServicePrincipalType = $p.servicePrincipalType; Tags = (Join-Values @($p.tags))
        AgentIdentityCount = @(Get-Indexed $agentsByBlueprint $p.appId).Count
        Owners = (Get-LinkNames $p.id 'Owner'); Sponsors = (Get-LinkNames $p.id 'Sponsor')
        AppPermissions = (Get-AppPermSummary $p.id); DelegatedPermissions = (Get-DelegatedSummary $p.id); MemberOf = (Get-MemberOfSummary $p.id)
    }
}

foreach ($a in $agents) {
    $appId = "$($a.agentIdentityBlueprintId)"
    $aUsers = @(Get-Indexed $usersByAgent $a.id)
    Add-Row 'agentIdentities' @{
        AgentIdentityId = $a.id; DisplayName = $a.displayName; BlueprintAppId = $appId; BlueprintName = (Get-BlueprintName $appId)
        BlueprintPrincipalId = $principalByAppId[$appId].id; BlueprintInThisTenant = [bool]$blueprintByAppId[$appId]
        AccountEnabled = $a.accountEnabled; CreatedDateTime = $a.createdDateTime; CreatedByAppId = $a.createdByAppId
        DisabledByMicrosoftStatus = $a.disabledByMicrosoftStatus; Tags = (Join-Values @($a.tags))
        AgentUserCount = $aUsers.Count; AgentUsers = (Join-Values @($aUsers | ForEach-Object { if ($_.userPrincipalName) { $_.userPrincipalName } else { $_.displayName } }))
        Owners = (Get-LinkNames $a.id 'Owner'); Sponsors = (Get-LinkNames $a.id 'Sponsor')
        AppPermissions = (Get-AppPermSummary $a.id); DelegatedPermissions = (Get-DelegatedSummary $a.id); MemberOf = (Get-MemberOfSummary $a.id)
    }
}

foreach ($u in $agentUsers) {
    $parent = $agentById["$($u.identityParentId)"]
    $appId = "$($parent.agentIdentityBlueprintId)"
    Add-Row 'agentUsers' @{
        AgentUserId = $u.id; DisplayName = $u.displayName; UserPrincipalName = $u.userPrincipalName; Mail = $u.mail
        AccountEnabled = $u.accountEnabled; CreatedDateTime = $u.createdDateTime; UserType = $u.userType
        AgentIdentityId = $u.identityParentId; AgentIdentityName = $parent.displayName
        BlueprintAppId = $(if ($parent) { $appId }); BlueprintName = $(if ($parent) { Get-BlueprintName $appId })
        Manager = (Get-LinkNames $u.id 'Manager'); Sponsors = (Get-LinkNames $u.id 'Sponsor')
        AppPermissions = (Get-AppPermSummary $u.id); MemberOf = (Get-MemberOfSummary $u.id)
    }
}
#endregion

#region Main: JSON export
Write-Step 'Writing JSON and CSV files'
function Get-RelationshipsJson([string]$ObjectType, [string]$ObjectId, [string[]]$Relations) {
    $rels = [ordered]@{}
    foreach ($r in $Relations) { $rels[$r] = Get-Rel $ObjectType $ObjectId $r }
    return $rels
}

$appIdOrder = [System.Collections.Generic.List[string]]::new()
$seenAppIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($id in @($blueprints | ForEach-Object { $_.appId }) + @($principals | ForEach-Object { $_.appId }) + @($agents | ForEach-Object { $_.agentIdentityBlueprintId })) {
    if ($id -and $seenAppIds.Add($id)) { $appIdOrder.Add($id) }
}

$hierarchy = @(foreach ($appId in $appIdOrder) {
        $b = $blueprintByAppId[$appId]; $p = $principalByAppId[$appId]
        [ordered]@{
            blueprintAppId        = $appId
            blueprintDisplayName  = Get-BlueprintName $appId
            blueprintObjectId     = $b.id
            blueprintInThisTenant = [bool]$b
            blueprintHomeTenantId = $p.appOwnerOrganizationId
            blueprintPrincipalId  = $p.id
            agentIdentities       = @(foreach ($a in (Get-Indexed $agentsByBlueprint $appId)) {
                    [ordered]@{
                        id = $a.id; displayName = $a.displayName; accountEnabled = $a.accountEnabled
                        agentUsers = @(foreach ($u in (Get-Indexed $usersByAgent $a.id)) { [ordered]@{ id = $u.id; displayName = $u.displayName; userPrincipalName = $u.userPrincipalName } })
                    }
                })
        }
    })

$export = [ordered]@{
    exportInfo          = [ordered]@{
        tool = 'Export-EntraAgentIdentities.ps1'; version = $ScriptVersion; exportedAtUtc = [datetime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
        tenantId = $tenantGuid; organizationName = $organizationName; graphApiVersion = $ApiVersion
        authMode = $AuthMode; tokenType = $tokenType; caller = $caller; permissionsInToken = $grantedPermissions
        counts = [ordered]@{
            blueprints = $blueprints.Count; blueprintPrincipals = $principals.Count; agentIdentities = $agents.Count; agentUsers = $agentUsers.Count
            edges = $tables.edges.Rows.Count; errors = $exportErrors.Count
        }
    }
    hierarchy           = $hierarchy
    blueprints          = @(foreach ($b in $blueprints) {
            $rels = Get-RelationshipsJson Blueprint $b.id @('owners', 'sponsors', 'inheritablePermissions', 'federatedIdentityCredentials')
            $rels.blueprintPrincipalId = $principalByAppId["$($b.appId)"].id
            $rels.agentIdentityIds = @(Get-Indexed $agentsByBlueprint $b.appId | ForEach-Object { $_.id })
            [ordered]@{ properties = $b; relationships = $rels }
        })
    blueprintPrincipals = @(foreach ($p in $principals) {
            $rels = Get-RelationshipsJson BlueprintPrincipal $p.id @('owners', 'sponsors', 'memberOf', 'appRoleAssignments', 'appRoleAssignedTo', 'oauth2PermissionGrants')
            $rels.blueprintObjectId = $blueprintByAppId["$($p.appId)"].id
            $rels.agentIdentityIds = @(Get-Indexed $agentsByBlueprint $p.appId | ForEach-Object { $_.id })
            [ordered]@{ properties = $p; relationships = $rels }
        })
    agentIdentities     = @(foreach ($a in $agents) {
            $rels = Get-RelationshipsJson AgentIdentity $a.id @('owners', 'sponsors', 'memberOf', 'appRoleAssignments', 'oauth2PermissionGrants')
            $rels.blueprintPrincipalId = $principalByAppId["$($a.agentIdentityBlueprintId)"].id
            $rels.agentUserIds = @(Get-Indexed $usersByAgent $a.id | ForEach-Object { $_.id })
            [ordered]@{ properties = $a; relationships = $rels }
        })
    agentUsers          = @(foreach ($u in $agentUsers) {
            $rels = Get-RelationshipsJson AgentUser $u.id @('sponsors', 'manager', 'memberOf', 'appRoleAssignments')
            $rels.agentIdentityId = $u.identityParentId
            [ordered]@{ properties = $u; relationships = $rels }
        })
    orphanAgentUserIds  = @($agentUsers | Where-Object { -not $agentById.ContainsKey("$($_.identityParentId)") } | ForEach-Object { $_.id })
    edges               = $tables.edges.Rows
    errors              = $exportErrors
}
$jsonFile = Join-Path $OutputPath 'agentid-export.json'
$export | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $jsonFile -Encoding utf8
#endregion

#region Main: CSV export
foreach ($name in $tables.Keys) {
    $table = $tables[$name]
    $file = Join-Path $csvPath "$name.csv"
    if ($table.Rows.Count -gt 0) {
        $table.Rows | Export-Csv -LiteralPath $file -NoTypeInformation -Encoding utf8BOM -Delimiter $CsvDelimiter
    }
    else {
        (($table.Columns | ForEach-Object { '"' + $_ + '"' }) -join $CsvDelimiter) | Set-Content -LiteralPath $file -Encoding utf8BOM
    }
}
#endregion

#region Main: relationship diagram (Mermaid + HTML)
function ConvertTo-MermaidText([string]$Text, [int]$MaxLength = 80) {
    if (-not $Text) { return '' }
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt $MaxLength) { $t = $t.Substring(0, $MaxLength - 1) + [char]0x2026 }
    $sb = [Text.StringBuilder]::new()
    foreach ($ch in $t.ToCharArray()) {
        $null = switch ($ch) {
            '#' { $sb.Append('#35;') } '"' { $sb.Append('#34;') } '<' { $sb.Append('#60;') } '>' { $sb.Append('#62;') }
            '&' { $sb.Append('#38;') } ';' { $sb.Append('#59;') } '`' { $sb.Append("'") } default { $sb.Append($ch) }
        }
    }
    return $sb.ToString()
}

$nodeStyles = @{
    Blueprint          = @{ Open = '[['; Close = ']]'; Class = 'blueprint'; Label = 'Agent identity blueprint' }
    BlueprintPrincipal = @{ Open = '(['; Close = '])'; Class = 'principal'; Label = 'Blueprint principal' }
    AgentIdentity      = @{ Open = '('; Close = ')'; Class = 'agent'; Label = 'Agent identity' }
    AgentUser          = @{ Open = '>'; Close = ']'; Class = 'agentuser'; Label = 'Agent user' }
    User               = @{ Open = '['; Close = ']'; Class = 'user'; Label = 'User' }
    Group              = @{ Open = '{{'; Close = '}}'; Class = 'group'; Label = 'Group' }
    DirectoryRole      = @{ Open = '{{'; Close = '}}'; Class = 'role'; Label = 'Directory role' }
    AdministrativeUnit = @{ Open = '{{'; Close = '}}'; Class = 'group'; Label = 'Administrative unit' }
    Resource           = @{ Open = '[('; Close = ')]'; Class = 'resource'; Label = 'API / resource' }
    FederatedIdentity  = @{ Open = '[/'; Close = '/]'; Class = 'fic'; Label = 'Federated credential' }
}
$hierarchyRelationships = @('HasBlueprintPrincipal', 'HasAgentIdentity', 'HasAgentUser')
$edgeText = @{
    HasBlueprintPrincipal = 'principal'; HasAgentIdentity = 'agent identity'; HasAgentUser = 'agent user'
    OwnerOf = 'owner'; SponsorOf = 'sponsor'; ManagerOf = 'manager'; MemberOf = 'member of'; FederatedCredential = 'federated credential'
    AppPermission = 'app'; DelegatedPermission = 'delegated'; InheritablePermission = 'inheritable'
}
$edgeArrow = @{ OwnerOf = '-.->'; SponsorOf = '-.->'; ManagerOf = '-.->'; AppPermission = '==>'; DelegatedPermission = '==>'; InheritablePermission = '==>' }

# Core objects are drawn inside one subgraph per blueprint (keyed by blueprint appId).
$nodeGroup = @{}
foreach ($b in $blueprints) { $nodeGroup["$($b.id)"] = $b.appId }
foreach ($p in $principals) { $nodeGroup["$($p.id)"] = $p.appId }
foreach ($a in $agents) { if ($a.agentIdentityBlueprintId) { $nodeGroup["$($a.id)"] = $a.agentIdentityBlueprintId; if (-not $blueprintByAppId[$a.agentIdentityBlueprintId]) { $nodeGroup[$a.agentIdentityBlueprintId] = $a.agentIdentityBlueprintId } } }
foreach ($u in $agentUsers) { $parent = $agentById["$($u.identityParentId)"]; if ($parent.agentIdentityBlueprintId) { $nodeGroup["$($u.id)"] = $parent.agentIdentityBlueprintId } }

$diagramNodes = [ordered]@{}
function Add-DiagramNode([string]$Type, [string]$Id, [string]$Name) {
    $key = "$Id".ToLowerInvariant()
    if (-not $diagramNodes.Contains($key)) {
        $style = if ($nodeStyles.ContainsKey($Type)) { $nodeStyles[$Type] } else { @{ Open = '['; Close = ']'; Class = 'sp'; Label = $Type } }
        $label = "$(ConvertTo-MermaidText $(if ($Name) { $Name } else { $Id }))<br/><i>$(ConvertTo-MermaidText $style.Label)</i>"
        $diagramNodes[$key] = [pscustomobject]@{ NodeId = "n$($diagramNodes.Count + 1)"; Definition = "$($style.Open)""$label""$($style.Close):::$($style.Class)"; Group = $nodeGroup[$Id] }
    }
    return $diagramNodes[$key].NodeId
}

# Make sure every core object appears even if it has no edges.
foreach ($b in $blueprints) { $null = Add-DiagramNode 'Blueprint' $b.id $b.displayName }
foreach ($p in $principals) { $null = Add-DiagramNode 'BlueprintPrincipal' $p.id $p.displayName }
foreach ($a in $agents) { $null = Add-DiagramNode 'AgentIdentity' $a.id $a.displayName }
foreach ($u in $agentUsers) { $null = Add-DiagramNode 'AgentUser' $u.id $u.displayName }

$diagramEdges = [ordered]@{}
foreach ($e in $tables.edges.Rows) {
    if ($e.Relationship -eq 'RequiredResourceAccess') { continue }
    if ($DiagramDetail -eq 'Hierarchy' -and $e.Relationship -notin $hierarchyRelationships) { continue }
    $from = Add-DiagramNode $e.SourceType $e.SourceId $e.SourceName
    $to = Add-DiagramNode $e.TargetType $e.TargetId $e.TargetName
    $key = "$from|$to|$($e.Relationship)"
    if (-not $diagramEdges.Contains($key)) {
        $diagramEdges[$key] = [pscustomobject]@{ From = $from; To = $to; Relationship = $e.Relationship; Details = [System.Collections.Generic.List[string]]::new() }
    }
    if ($e.Detail -and $e.Relationship -in 'AppPermission', 'DelegatedPermission', 'InheritablePermission' -and -not $diagramEdges[$key].Details.Contains($e.Detail)) {
        $diagramEdges[$key].Details.Add($e.Detail)
    }
}

$mmd = [System.Collections.Generic.List[string]]::new()
$mmd.Add('flowchart LR')
$mmd.Add('  classDef blueprint fill:#dbeafe,stroke:#1d4ed8,stroke-width:2px,color:#0b1f4d')
$mmd.Add('  classDef principal fill:#e0e7ff,stroke:#4338ca,color:#1e1b4b')
$mmd.Add('  classDef agent fill:#dcfce7,stroke:#15803d,stroke-width:2px,color:#052e16')
$mmd.Add('  classDef agentuser fill:#fef9c3,stroke:#a16207,color:#422006')
$mmd.Add('  classDef user fill:#fce7f3,stroke:#be185d,color:#500724')
$mmd.Add('  classDef group fill:#f3e8ff,stroke:#7e22ce,color:#3b0764')
$mmd.Add('  classDef role fill:#fee2e2,stroke:#b91c1c,color:#450a0a')
$mmd.Add('  classDef sp fill:#f1f5f9,stroke:#475569,color:#0f172a')
$mmd.Add('  classDef resource fill:#ffedd5,stroke:#c2410c,color:#431407')
$mmd.Add('  classDef fic fill:#ecfeff,stroke:#0e7490,color:#083344')

$groupIndex = 0
foreach ($appId in $appIdOrder) {
    $members = @($diagramNodes.Values | Where-Object { $_.Group -eq $appId })
    if (-not $members) { continue }
    $groupIndex++
    $title = Get-BlueprintName $appId
    if (-not $title) { $title = $appId }
    if (-not $blueprintByAppId[$appId]) { $title = "$title (blueprint not in this tenant)" }
    $mmd.Add("  subgraph sg$groupIndex[""$(ConvertTo-MermaidText "Blueprint: $title" 100)""]")
    foreach ($n in $members) { $mmd.Add("    $($n.NodeId)$($n.Definition)") }
    $mmd.Add('  end')
}
foreach ($n in $diagramNodes.Values) {
    if (-not $n.Group -or $n.Group -notin $appIdOrder) { $mmd.Add("  $($n.NodeId)$($n.Definition)") }
}
foreach ($edge in $diagramEdges.Values) {
    $text = $edgeText[$edge.Relationship]
    if (-not $text) { $text = $edge.Relationship }
    if ($edge.Details.Count -gt 0) {
        $shown = @($edge.Details | Select-Object -First 4 | ForEach-Object { if ($_.Length -gt 60) { $_.Substring(0, 59) + [char]0x2026 } else { $_ } })
        $more = if ($edge.Details.Count -gt 4) { " +$($edge.Details.Count - 4) more" } else { '' }
        $text = "${text}: $($shown -join ', ')$more"
    }
    $arrow = if ($edgeArrow.ContainsKey($edge.Relationship)) { $edgeArrow[$edge.Relationship] } else { '-->' }
    $mmd.Add("  $($edge.From) $arrow|""$(ConvertTo-MermaidText $text 300)""| $($edge.To)")
}
$mermaidSource = $mmd -join "`n"
$mmdFile = Join-Path $OutputPath 'agentid-graph.mmd'
Set-Content -LiteralPath $mmdFile -Value $mermaidSource -Encoding utf8
if ($diagramNodes.Count -gt 1500) {
    Write-Warning "The diagram has $($diagramNodes.Count) nodes and may render slowly. Re-run with -DiagramDetail Hierarchy for a lighter view."
}

$enc = { param($s) [System.Net.WebUtility]::HtmlEncode("$s") }
$srcJson = $mermaidSource | ConvertTo-Json -EscapeHandling EscapeHtml
$tenantLabel = if ($organizationName) { "$organizationName ($tenantGuid)" } else { $tenantGuid }
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Entra Agent ID relationships - $(& $enc $tenantLabel)</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; background: #f8fafc; color: #0f172a; }
  header { padding: 16px 24px; background: #0f172a; color: #fff; }
  header h1 { margin: 0; font-size: 20px; }
  header p { margin: 4px 0 0; font-size: 13px; color: #cbd5e1; }
  .stats { display: flex; gap: 12px; flex-wrap: wrap; padding: 12px 24px; }
  .stat { background: #fff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 8px 14px; font-size: 13px; }
  .stat b { display: block; font-size: 20px; }
  .toolbar { padding: 0 24px 8px; display: flex; gap: 8px; align-items: center; font-size: 13px; flex-wrap: wrap; }
  .toolbar button { padding: 4px 10px; border: 1px solid #94a3b8; background: #fff; border-radius: 6px; cursor: pointer; }
  .legend span { display: inline-block; padding: 2px 8px; border-radius: 4px; margin-right: 4px; border: 1px solid #94a3b8; }
  #graph { margin: 0 24px 16px; background: #fff; border: 1px solid #e2e8f0; border-radius: 8px; overflow: auto; height: calc(100vh - 240px); padding: 12px; }
  details { margin: 0 24px 24px; font-size: 13px; }
  pre { white-space: pre-wrap; font-size: 12px; }
</style>
</head>
<body>
<header>
  <h1>Microsoft Entra Agent ID &ndash; relationships</h1>
  <p>Tenant: $(& $enc $tenantLabel) &middot; Exported (UTC): $(& $enc $export.exportInfo.exportedAtUtc) &middot; Graph $(& $enc $ApiVersion) &middot; Detail: $(& $enc $DiagramDetail) &middot; Caller: $(& $enc $caller)</p>
</header>
<div class="stats">
  <div class="stat"><b>$($blueprints.Count)</b>Blueprints</div>
  <div class="stat"><b>$($principals.Count)</b>Blueprint principals</div>
  <div class="stat"><b>$($agents.Count)</b>Agent identities</div>
  <div class="stat"><b>$($agentUsers.Count)</b>Agent users</div>
  <div class="stat"><b>$($tables.edges.Rows.Count)</b>Connections</div>
  <div class="stat"><b>$($exportErrors.Count)</b>Errors (see csv/errors.csv)</div>
</div>
<div class="toolbar">
  <button id="zin">Zoom +</button><button id="zout">Zoom &minus;</button><button id="zfit">Fit width</button><button id="z100">100%</button>
  <span class="legend">
    <span style="background:#dbeafe">Blueprint</span><span style="background:#e0e7ff">Blueprint principal</span>
    <span style="background:#dcfce7">Agent identity</span><span style="background:#fef9c3">Agent user</span>
    <span style="background:#fce7f3">User</span><span style="background:#f3e8ff">Group</span><span style="background:#fee2e2">Directory role</span>
    <span style="background:#ffedd5">API / resource</span><span style="background:#ecfeff">Federated credential</span>
    &middot; dotted = owner/sponsor/manager &middot; thick = permissions
  </span>
  <span id="layoutNote" style="color:#64748b"></span>
</div>
<div id="graph">Rendering diagram&hellip;</div>
<details><summary>Mermaid source</summary><pre id="src"></pre></details>
<script type="module">
  const src = $srcJson;
  document.getElementById('src').textContent = src;
  const host = document.getElementById('graph');
  let scale = 1, svgEl = null, baseW = 0, baseH = 0;
  function apply() { if (!svgEl) return; svgEl.style.width = (baseW * scale) + 'px'; svgEl.style.height = (baseH * scale) + 'px'; }
  function fit() { if (!svgEl) return; scale = Math.min(1, (host.clientWidth - 30) / baseW); apply(); }
  document.getElementById('zin').onclick = () => { scale *= 1.25; apply(); };
  document.getElementById('zout').onclick = () => { scale /= 1.25; apply(); };
  document.getElementById('zfit').onclick = fit;
  document.getElementById('z100').onclick = () => { scale = 1; apply(); };
  const flat = src.split('\n').filter(l => !/^\s*subgraph\s/.test(l) && l.trim() !== 'end').join('\n');
  // Mermaid's default (dagre) layout can crash on large graphs with subgraphs; fall back to ELK, then to no subgraphs.
  const attempts = [
    { label: 'dagre layout', layout: 'dagre', text: src },
    { label: 'ELK layout', layout: 'elk', text: src, elk: true },
    { label: 'dagre layout without blueprint groups', layout: 'dagre', text: flat }
  ];
  let lastError = null;
  try {
    const { default: mermaid } = await import('https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs');
    for (let i = 0; i < attempts.length && !svgEl; i++) {
      const a = attempts[i];
      try {
        if (a.elk) {
          const { default: elkLayouts } = await import('https://cdn.jsdelivr.net/npm/@mermaid-js/layout-elk@0/dist/mermaid-layout-elk.esm.min.mjs');
          mermaid.registerLayoutLoaders(elkLayouts);
        }
        mermaid.initialize({ startOnLoad: false, securityLevel: 'strict', maxTextSize: 50000000, maxEdges: 200000,
                             layout: a.layout, flowchart: { useMaxWidth: false, htmlLabels: true } });
        const { svg } = await mermaid.render('agentIdGraph' + i, a.text);
        host.innerHTML = svg;
        svgEl = host.querySelector('svg');
        document.getElementById('layoutNote').textContent = 'Rendered with ' + a.label + (i > 0 ? ' (fallback)' : '');
      } catch (e) {
        lastError = e;
        console.warn('Render with ' + a.label + ' failed:', e);
        const leftover = document.getElementById('dagentIdGraph' + i);
        if (leftover) leftover.remove();
      }
    }
    if (!svgEl) throw lastError;
    const vb = svgEl.viewBox && svgEl.viewBox.baseVal;
    baseW = (vb && vb.width) ? vb.width : svgEl.getBBox().width;
    baseH = (vb && vb.height) ? vb.height : svgEl.getBBox().height;
    svgEl.style.maxWidth = 'none';
    fit();
  } catch (e) {
    host.textContent = 'Could not render the diagram (' + e + '). Open agentid-graph.mmd in https://mermaid.live (Layout: ELK) or VS Code to view it.';
  }
</script>
</body>
</html>
"@
$htmlFile = Join-Path $OutputPath 'agentid-graph.html'
Set-Content -LiteralPath $htmlFile -Value $html -Encoding utf8
#endregion

#region Main: summary
Write-Step 'Export complete'
Write-Host "    Blueprints:            $($blueprints.Count)"
Write-Host "    Blueprint principals:  $($principals.Count)"
Write-Host "    Agent identities:      $($agents.Count)"
Write-Host "    Agent users:           $($agentUsers.Count)"
Write-Host "    Connections (edges):   $($tables.edges.Rows.Count)"
Write-Host "    Output folder:         $OutputPath"
if ($exportErrors.Count -gt 0) {
    Write-Warning "$($exportErrors.Count) requests failed (details in csv\errors.csv):"
    $exportErrors | Group-Object ObjectType, Relation, Status, ErrorCode | Sort-Object Count -Descending | ForEach-Object {
        $first = $_.Group[0]
        Write-Host ("    {0,5} x {1}.{2} -> HTTP {3} {4}" -f $_.Count, $first.ObjectType, $first.Relation, $first.Status, $first.ErrorCode) -ForegroundColor Yellow
    }
}

[pscustomobject]@{
    OutputPath          = $OutputPath
    Json                = $jsonFile
    Html                = $htmlFile
    Mermaid             = $mmdFile
    Blueprints          = $blueprints.Count
    BlueprintPrincipals = $principals.Count
    AgentIdentities     = $agents.Count
    AgentUsers          = $agentUsers.Count
    Connections         = $tables.edges.Rows.Count
    Errors              = $exportErrors.Count
}
#endregion
