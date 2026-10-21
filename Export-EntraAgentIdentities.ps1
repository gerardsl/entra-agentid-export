#Requires -Version 7.2
<#
.SYNOPSIS
    Shows each Microsoft Entra Agent ID blueprint with its agent identities below it and, when they exist,
    the agent users of each agent identity. Uses Microsoft Graph REST API calls (read-only).

.DESCRIPTION
    Reads:
      * Agent identity blueprints           GET /applications/microsoft.graph.agentIdentityBlueprint
      * Agent identity blueprint principals GET /servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal
                                            (only used to name blueprints that are registered in another tenant)
      * Agent identities                    GET /servicePrincipals/microsoft.graph.agentIdentity
      * Agent users                         GET /users/microsoft.graph.agentUser

    Agent identities are placed under their blueprint (agentIdentityBlueprintId = blueprint appId) and
    agent users under their agent identity (identityParentId = agent identity object ID).

    Output (in -OutputPath):
      agentid-hierarchy.html   Each blueprint with its agent identities below and their agent users below those
      agentid-hierarchy.json   The same hierarchy as nested JSON
      agentid-hierarchy.csv    One row per blueprint / agent identity / agent user

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
    .\Export-EntraAgentIdentities.ps1 -AccessToken $token -CsvDelimiter ';'
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

    [ValidateLength(1, 1)]
    [string]$CsvDelimiter = ',',

    [ValidateRange(0, 10)]
    [int]$MaxRetries = 6
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '2.0.0'
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
    'User.Read.All'
)
# User.Read lets delegated sign-ins read the organization name shown in the page header.
$DelegatedScopes = @($RequiredPermissions + 'User.Read' | ForEach-Object { "$GraphHost/$_" }) + 'offline_access', 'openid', 'profile'

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
Write-Verbose "Permissions in token: $($grantedPermissions -join ', ')"

$missing = foreach ($p in $RequiredPermissions) {
    $alternatives = @($p, ($p -replace '\.Read\.All$', '.ReadWrite.All'))
    if ($p -eq 'User.Read.All') { $alternatives += 'Directory.Read.All', 'Directory.ReadWrite.All' }
    if (-not ($grantedPermissions | Where-Object { $_ -in $alternatives })) { $p }
}
if ($missing) { Write-Warning "Token lacks: $($missing -join ', '). The related objects may be missing from the output." }
if ($grantedPermissions -contains 'Directory.AccessAsUser.All') {
    Write-Warning 'Token contains Directory.AccessAsUser.All, which Agent ID APIs may reject (403). If that happens, use your own app registration (-ClientId) or app-only auth.'
}
#endregion

#region Main: output folder
if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location) ('AgentIdExport-{0}-{1}' -f $tenantGuid, (Get-Date -Format 'yyyyMMdd-HHmmss'))
}
$null = New-Item -ItemType Directory -Path $OutputPath -Force
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path
#endregion

#region Main: read blueprints, agent identities and agent users
$warnings = [System.Collections.Generic.List[string]]::new()

function Get-CoreCollection([string]$Name, [string[]]$Uris) {
    Write-Step "Listing $Name"
    try {
        $items = Get-GraphCollectionWithFallback $Uris
        Write-Host "    $($items.Count) found"
        return , $items
    }
    catch {
        $message = "Could not list $Name : $($_.Exception.Message)"
        Write-Warning $message
        $warnings.Add($message)
        return , @()
    }
}

$blueprints = Get-CoreCollection 'agent identity blueprints' @(
    '/applications/microsoft.graph.agentIdentityBlueprint?$select=id,appId,displayName'
    '/applications/microsoft.graph.agentIdentityBlueprint')
$principals = Get-CoreCollection 'agent identity blueprint principals' @(
    '/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal?$select=id,appId,displayName,appOwnerOrganizationId'
    '/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal')
$agents = Get-CoreCollection 'agent identities' @(
    '/servicePrincipals/microsoft.graph.agentIdentity?$select=id,displayName,agentIdentityBlueprintId,accountEnabled'
    '/servicePrincipals/microsoft.graph.agentIdentity')
$agentUsers = Get-CoreCollection 'agent users' @(
    '/users/microsoft.graph.agentUser?$select=id,displayName,userPrincipalName,identityParentId,accountEnabled'
    '/users/microsoft.graph.agentUser'
    '/users/microsoft.graph.AgentUser')

$organizationName = $null
try { $organizationName = (Invoke-Graph -Uri '/organization?$select=id,displayName').value[0].displayName }
catch { Write-Verbose "Organization name unavailable: $($_.Exception.Message)" }
#endregion

#region Main: build the hierarchy (blueprint -> agent identities -> agent users)
Write-Step 'Placing agent identities under their blueprints'
function Get-SortKey([object]$Item) { "$($Item.displayName)".ToLowerInvariant() }

function New-GroupIndex([object[]]$Items, [string]$Property) {
    $index = @{}
    foreach ($item in $Items) {
        $key = "$($item.$Property)"
        if (-not $index.ContainsKey($key)) { $index[$key] = [System.Collections.Generic.List[object]]::new() }
        $index[$key].Add($item)
    }
    return $index
}

function New-AgentUserNode([object]$User) {
    [ordered]@{ displayName = $User.displayName; userPrincipalName = $User.userPrincipalName; id = $User.id; accountEnabled = $User.accountEnabled }
}

$blueprintByAppId = @{}
foreach ($b in $blueprints) { if ($b.appId) { $blueprintByAppId["$($b.appId)"] = $b } }
$principalByAppId = @{}
foreach ($p in $principals) { if ($p.appId) { $principalByAppId["$($p.appId)"] = $p } }
$agentsByBlueprint = New-GroupIndex $agents 'agentIdentityBlueprintId'
$usersByAgent = New-GroupIndex $agentUsers 'identityParentId'
$agentIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($a in $agents) { $null = $agentIds.Add("$($a.id)") }

# Every blueprint visible in the tenant: registered here, present only as a blueprint principal
# (blueprint registered in another tenant), or only referenced by agent identities.
$appIds = [System.Collections.Generic.List[string]]::new()
$seenAppIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($id in @($blueprints | ForEach-Object { $_.appId }) + @($principals | ForEach-Object { $_.appId }) + @($agents | ForEach-Object { $_.agentIdentityBlueprintId })) {
    if ($id -and $seenAppIds.Add("$id")) { $appIds.Add("$id") }
}
if ($agentsByBlueprint.ContainsKey('')) { $appIds.Add('') }

$hierarchyNodes = foreach ($appId in $appIds) {
    $b = $blueprintByAppId[$appId]
    $p = $principalByAppId[$appId]
    $name = if ($b.displayName) { $b.displayName } elseif ($p.displayName) { $p.displayName } elseif ($appId) { "Blueprint $appId" } else { '(No blueprint reference)' }
    $agentNodes = @(
        foreach ($a in (@($agentsByBlueprint[$appId]) | Where-Object { $_ } | Sort-Object { Get-SortKey $_ })) {
            [ordered]@{
                displayName    = $a.displayName
                id             = $a.id
                accountEnabled = $a.accountEnabled
                agentUsers     = @(@($usersByAgent["$($a.id)"]) | Where-Object { $_ } | Sort-Object { Get-SortKey $_ } | ForEach-Object { New-AgentUserNode $_ })
            }
        })
    [ordered]@{
        displayName            = $name
        appId                  = $(if ($appId) { $appId } else { $null })
        objectId               = $b.id
        registeredInThisTenant = [bool]$b
        homeTenantId           = $p.appOwnerOrganizationId
        agentIdentities        = $agentNodes
    }
}
# Blueprints with the most agent identities first, then by name.
$hierarchy = @($hierarchyNodes | Sort-Object @{ Expression = { $_.agentIdentities.Count }; Descending = $true }, @{ Expression = { "$($_.displayName)".ToLowerInvariant() } })
$unlinkedUsers = @($agentUsers | Where-Object { -not $agentIds.Contains("$($_.identityParentId)") } | Sort-Object { Get-SortKey $_ } | ForEach-Object { New-AgentUserNode $_ })
$blueprintCount = @($hierarchy | Where-Object { $_.appId }).Count
$registeredCount = @($hierarchy | Where-Object { $_.registeredInThisTenant }).Count
#endregion

#region Main: JSON and CSV
Write-Step 'Writing output files'
$exportedAtUtc = [datetime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ')
$export = [ordered]@{
    exportInfo         = [ordered]@{
        tool = 'Export-EntraAgentIdentities.ps1'; version = $ScriptVersion; exportedAtUtc = $exportedAtUtc
        tenantId = $tenantGuid; organizationName = $organizationName; graphApiVersion = $ApiVersion; caller = $caller
        counts = [ordered]@{
            blueprints = $blueprintCount; blueprintsRegisteredInThisTenant = $registeredCount
            agentIdentities = $agents.Count; agentUsers = $agentUsers.Count
        }
        warnings = @($warnings)
    }
    blueprints         = $hierarchy
    unlinkedAgentUsers = $unlinkedUsers
}
$jsonFile = Join-Path $OutputPath 'agentid-hierarchy.json'
$export | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $jsonFile -Encoding utf8

$csvRows = [System.Collections.Generic.List[object]]::new()
function Add-CsvRow([object]$Blueprint, [object]$Agent, [object]$User) {
    $csvRows.Add([pscustomobject][ordered]@{
            BlueprintName                   = $Blueprint.displayName
            BlueprintAppId                  = $Blueprint.appId
            BlueprintRegisteredInThisTenant = $Blueprint.registeredInThisTenant
            AgentIdentityName               = $Agent.displayName
            AgentIdentityId                 = $Agent.id
            AgentIdentityEnabled            = $Agent.accountEnabled
            AgentUserName                   = $User.displayName
            AgentUserPrincipalName          = $User.userPrincipalName
            AgentUserId                     = $User.id
        })
}
foreach ($bp in $hierarchy) {
    if ($bp.agentIdentities.Count -eq 0) { Add-CsvRow $bp $null $null; continue }
    foreach ($ai in $bp.agentIdentities) {
        if ($ai.agentUsers.Count -eq 0) { Add-CsvRow $bp $ai $null; continue }
        foreach ($au in $ai.agentUsers) { Add-CsvRow $bp $ai $au }
    }
}
foreach ($au in $unlinkedUsers) { Add-CsvRow $null $null $au }
$csvFile = Join-Path $OutputPath 'agentid-hierarchy.csv'
if ($csvRows.Count -gt 0) {
    $csvRows | Export-Csv -LiteralPath $csvFile -NoTypeInformation -Encoding utf8BOM -Delimiter $CsvDelimiter
}
else {
    $header = 'BlueprintName', 'BlueprintAppId', 'BlueprintRegisteredInThisTenant', 'AgentIdentityName', 'AgentIdentityId', 'AgentIdentityEnabled', 'AgentUserName', 'AgentUserPrincipalName', 'AgentUserId'
    (($header | ForEach-Object { '"' + $_ + '"' }) -join $CsvDelimiter) | Set-Content -LiteralPath $csvFile -Encoding utf8BOM
}
#endregion

#region Main: HTML view
function ConvertTo-HtmlText([object]$Value) { [System.Net.WebUtility]::HtmlEncode("$Value") }

function ConvertTo-SearchText([object[]]$Values) {
    $text = (@($Values) | Where-Object { $_ }) -join ' '
    ConvertTo-HtmlText $text.ToLowerInvariant()
}

function Get-DisabledBadge([object]$Node) {
    if ($Node.accountEnabled -eq $false) { return ' <span class="badge off">Disabled</span>' }
    return ''
}

function Get-AgentUserHtml([object]$User) {
    $upn = if ($User.userPrincipalName) { "<div class=""ids"">$(ConvertTo-HtmlText $User.userPrincipalName)</div>" } else { '' }
    "<li><div class=""user-name""><span class=""kind"">Agent user</span> $(ConvertTo-HtmlText $User.displayName)$(Get-DisabledBadge $User)</div>$upn</li>"
}

$body = [System.Text.StringBuilder]::new()
foreach ($bp in $hierarchy) {
    $n = $bp.agentIdentities.Count
    $countText = if ($n -eq 1) { '1 agent identity' } else { "$n agent identities" }
    $sectionClass = if ($n -eq 0) { 'bp empty' } else { 'bp' }
    $external = ''
    if ($bp.appId -and -not $bp.registeredInThisTenant) {
        $title = if ($bp.homeTenantId) { " title=""Blueprint registered in tenant $(ConvertTo-HtmlText $bp.homeTenantId)""" } else { '' }
        $external = " <span class=""badge ext""$title>Registered in another tenant</span>"
    }
    $appIdText = if ($bp.appId) { "App ID $(ConvertTo-HtmlText $bp.appId)" } else { '' }
    [void]$body.Append("<section class=""$sectionClass"" data-bp=""$(ConvertTo-SearchText @($bp.displayName, $bp.appId))"">")
    [void]$body.Append("<div class=""bp-head""><span class=""kind"">Blueprint</span><span class=""bp-name"">$(ConvertTo-HtmlText $bp.displayName)</span><span class=""count"">$countText</span>$external<div class=""ids"">$appIdText</div></div>")
    if ($n -eq 0) {
        [void]$body.Append('<div class="none">No agent identities</div>')
    }
    else {
        [void]$body.Append('<div class="agents">')
        foreach ($ai in $bp.agentIdentities) {
            $search = ConvertTo-SearchText (@($ai.displayName, $ai.id) + @($ai.agentUsers | ForEach-Object { $_.displayName; $_.userPrincipalName }))
            [void]$body.Append("<div class=""agent"" data-search=""$search""><div class=""agent-name""><span class=""kind"">Agent identity</span> $(ConvertTo-HtmlText $ai.displayName)$(Get-DisabledBadge $ai)</div><div class=""ids"">$(ConvertTo-HtmlText $ai.id)</div>")
            if ($ai.agentUsers.Count -gt 0) {
                [void]$body.Append('<ul class="users">')
                foreach ($au in $ai.agentUsers) { [void]$body.Append((Get-AgentUserHtml $au)) }
                [void]$body.Append('</ul>')
            }
            [void]$body.Append('</div>')
        }
        [void]$body.Append('</div>')
    }
    [void]$body.Append('</section>')
}

$unlinkedHtml = ''
if ($unlinkedUsers.Count -gt 0) {
    $items = ($unlinkedUsers | ForEach-Object { Get-AgentUserHtml $_ }) -join ''
    $unlinkedHtml = "<section class=""unlinked""><div class=""bp-head""><span class=""bp-name"">Agent users without a matching agent identity</span><span class=""count"">$($unlinkedUsers.Count)</span></div><ul class=""users loose"">$items</ul></section>"
}
$warningHtml = ''
if ($warnings.Count -gt 0) {
    $warningHtml = "<div class=""warn""><b>Some objects could not be read, so this view may be incomplete:</b><br>$(($warnings | ForEach-Object { ConvertTo-HtmlText $_ }) -join '<br>')</div>"
}
$tenantLabel = if ($organizationName) { "$organizationName ($tenantGuid)" } else { $tenantGuid }

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Entra Agent ID blueprints - $(ConvertTo-HtmlText $tenantLabel)</title>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; background: #f8fafc; color: #0f172a; }
  header { padding: 16px 24px; background: #0f172a; color: #fff; }
  header h1 { margin: 0; font-size: 20px; }
  header p { margin: 4px 0 0; font-size: 13px; color: #cbd5e1; }
  .stats { display: flex; gap: 12px; flex-wrap: wrap; padding: 12px 24px 0; }
  .stat { background: #fff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 8px 14px; font-size: 13px; }
  .stat b { display: block; font-size: 20px; }
  .stat.s-bp { border-left: 4px solid #2563eb; } .stat.s-ai { border-left: 4px solid #16a34a; } .stat.s-au { border-left: 4px solid #ca8a04; }
  .toolbar { display: flex; gap: 16px; align-items: center; flex-wrap: wrap; padding: 12px 24px; font-size: 13px; }
  .toolbar input[type=search] { width: 380px; max-width: 100%; padding: 6px 10px; border: 1px solid #94a3b8; border-radius: 6px; font-size: 13px; }
  #shown { color: #64748b; }
  .warn { margin: 0 24px 12px; padding: 8px 12px; background: #fef3c7; border: 1px solid #f59e0b; border-radius: 6px; font-size: 13px; }
  main { padding: 0 24px 24px; display: flex; flex-direction: column; gap: 12px; }
  .bp, .unlinked { background: #fff; border: 1px solid #cbd5e1; border-radius: 10px; overflow: hidden; }
  .bp-head { background: #dbeafe; border-bottom: 1px solid #bfdbfe; padding: 10px 14px; display: flex; flex-wrap: wrap; gap: 4px 10px; align-items: baseline; }
  .bp.empty .bp-head { background: #f1f5f9; border-bottom: none; }
  .unlinked .bp-head { background: #fef9c3; border-bottom-color: #fde68a; }
  .kind { font-size: 10px; text-transform: uppercase; letter-spacing: .05em; font-weight: 600; color: #64748b; }
  .bp-head .kind { color: #1d4ed8; }
  .bp-name { font-size: 16px; font-weight: 600; }
  .count { font-size: 12px; color: #334155; background: #fff; border: 1px solid #bfdbfe; border-radius: 999px; padding: 1px 8px; }
  .badge { font-size: 11px; border-radius: 4px; padding: 1px 6px; font-weight: 400; }
  .badge.ext { background: #ede9fe; color: #5b21b6; }
  .badge.off { background: #fee2e2; color: #991b1b; }
  .ids { font-size: 11px; color: #64748b; font-family: Consolas, 'Courier New', monospace; width: 100%; overflow-wrap: anywhere; }
  .agents { position: relative; display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 10px; padding: 12px 14px 14px 32px; }
  .agents::before { content: ''; position: absolute; left: 16px; top: 0; bottom: 18px; border-left: 2px solid #93c5fd; }
  .agent { background: #f0fdf4; border: 1px solid #86efac; border-left: 4px solid #16a34a; border-radius: 8px; padding: 8px 10px; }
  .agent .kind { color: #15803d; }
  .agent-name { font-weight: 600; font-size: 14px; }
  .users { list-style: none; margin: 8px 0 0; padding: 0 0 0 12px; border-left: 2px solid #fcd34d; }
  .users.loose { margin: 12px 14px 14px 32px; }
  .users li { background: #fefce8; border: 1px solid #fde68a; border-radius: 6px; padding: 4px 8px; margin-top: 4px; font-size: 13px; }
  .users .kind { color: #a16207; }
  .none { padding: 8px 14px 10px 32px; font-size: 12px; color: #94a3b8; font-style: italic; }
  [hidden] { display: none !important; }
</style>
</head>
<body>
<header>
  <h1>Microsoft Entra Agent ID &ndash; blueprints and their agent identities</h1>
  <p>Tenant: $(ConvertTo-HtmlText $tenantLabel) &middot; Exported (UTC): $exportedAtUtc</p>
</header>
<div class="stats">
  <div class="stat s-bp"><b>$blueprintCount</b>Blueprints ($registeredCount registered in this tenant)</div>
  <div class="stat s-ai"><b>$($agents.Count)</b>Agent identities</div>
  <div class="stat s-au"><b>$($agentUsers.Count)</b>Agent users</div>
</div>
<div class="toolbar">
  <input type="search" id="q" placeholder="Filter by blueprint, agent identity or agent user" aria-label="Filter">
  <label><input type="checkbox" id="hideEmpty"> Hide blueprints without agent identities</label>
  <span id="shown"></span>
</div>
$warningHtml
<main>
$($body.ToString())
$unlinkedHtml
</main>
<script>
  const input = document.getElementById('q');
  const hideEmpty = document.getElementById('hideEmpty');
  const shown = document.getElementById('shown');
  const sections = Array.from(document.querySelectorAll('section.bp'));
  function applyFilter() {
    const q = input.value.trim().toLowerCase();
    let visible = 0;
    for (const s of sections) {
      const blueprintMatch = !q || s.dataset.bp.includes(q);
      let agentMatch = false;
      for (const a of s.querySelectorAll('.agent')) {
        const match = blueprintMatch || a.dataset.search.includes(q);
        a.hidden = !match;
        if (match) agentMatch = true;
      }
      const show = (blueprintMatch || agentMatch) && !(hideEmpty.checked && s.classList.contains('empty'));
      s.hidden = !show;
      if (show) visible++;
    }
    shown.textContent = visible + ' of ' + sections.length + ' blueprints shown';
  }
  input.addEventListener('input', applyFilter);
  hideEmpty.addEventListener('change', applyFilter);
  applyFilter();
</script>
</body>
</html>
"@
$htmlFile = Join-Path $OutputPath 'agentid-hierarchy.html'
Set-Content -LiteralPath $htmlFile -Value $html -Encoding utf8
#endregion

#region Main: summary
Write-Step 'Done'
Write-Host "    Blueprints:        $blueprintCount ($registeredCount registered in this tenant)"
Write-Host "    Agent identities:  $($agents.Count)"
Write-Host "    Agent users:       $($agentUsers.Count)$(if ($unlinkedUsers.Count) { " ($($unlinkedUsers.Count) without a matching agent identity)" })"
Write-Host "    Output folder:     $OutputPath"
if ($warnings.Count -gt 0) { Write-Warning "$($warnings.Count) list request(s) failed; the output may be incomplete (see the warning at the top of the HTML page)." }

[pscustomobject]@{
    OutputPath      = $OutputPath
    Html            = $htmlFile
    Json            = $jsonFile
    Csv             = $csvFile
    Blueprints      = $blueprintCount
    AgentIdentities = $agents.Count
    AgentUsers      = $agentUsers.Count
}
#endregion
