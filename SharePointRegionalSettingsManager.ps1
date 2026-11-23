#requires -Version 7.4
<#
.SYNOPSIS
    Discovers, exports, updates and restores SharePoint Online site collection locales.
.DESCRIPTION
    Run on Windows with PowerShell 7.4 or later:
      pwsh -NoProfile -STA -File .\SharePointRegionalSettingsManager.ps1

    Enter the commercial SharePoint tenant administration URL and the application
    (client) ID of your organisation's Entra public client app. The client ID can
    also be supplied with -ClientId or ENTRAID_APP_ID / ENTRAID_CLIENT_ID.
    Configure the app for interactive sign-in with a Mobile and desktop applications
    redirect URI of http://localhost and consented SharePoint delegated permissions
    suitable for tenant administration (normally AllSites.FullControl). The signed-in
    user needs SharePoint admin permissions for discovery and rights to manage each
    site's regional settings. The tool does not grant site access.

    Connect, Discover Sites, then Export Settings to save the observed baseline.
    Apply NZ Locale targets the discovered site collection root webs. Standalone
    subsites are outside the inventory. Rollback reads a settings or recovery CSV.
    Use Grid State CSV for an operational snapshot of the entire inventory.

    Changes are verified from SharePoint. Each attempt that needs a write records
    its original locale in a durable local recovery CSV first. Its path is shown
    in the live log. Recovery CSVs are stored in the app's LocalAppData folder.
    If a write fails ambiguously, inspect the site's current value with Refresh.

    Default discovery streams detailed CSOM pages using the same tenant API as
    Get-PnPTenantSite -Detailed. Use -UseTenantSiteCmdlet for direct cmdlet discovery.
    The cmdlet itself buffers all pages for each geo before emitting results.
    For Multi-Geo, supply each additional geography's admin URL.
.PARAMETER TenantAdminUrl
    Prefills the primary tenant admin URL.
.PARAMETER ClientId
    Prefills the Entra application client ID.
.PARAMETER AdditionalAdminUrls
    Additional admin endpoints for the same tenant, one per configured geography.
.PARAMETER Concurrency
    Maximum parallel site workers. Defaults to 5.
.PARAMETER UseTenantSiteCmdlet
    Uses Get-PnPTenantSite -Detailed for enumeration.
.EXAMPLE
    pwsh -NoProfile -STA -File .\SharePointRegionalSettingsManager.ps1 -Concurrency 5
#>
[CmdletBinding()]
param(
    [string] $TenantAdminUrl,
    [string] $ClientId = ($env:ENTRAID_APP_ID ?? $env:ENTRAID_CLIENT_ID),
    [string[]] $AdditionalAdminUrls = @(),
    [ValidateRange(1, 20)] [int] $Concurrency = 5,
    [switch] $UseTenantSiteCmdlet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Filters apply to discovery and are checked against live site metadata on rollback.
$config = @{
    Target = 5129
    IncludeOneDrive = $false
    ExcludedTemplatePatterns = @('REDIRECTSITE#*', 'APPCATALOG#*', 'SRCHCEN#*',
        'SRCHCENTERLITE#*', 'SRCHCENTERFAST#*', 'SEARCHCENTER#*', 'ENTERPRISESEARCHCENTER#*', 'TENANTADMIN#*')
    ExcludedUrlPatterns = @()
    LogPath = Join-Path ([IO.Path]::GetTempPath()) 'SPRegionalSettings.log'
    LogMaxBytes = 5MB
    LogArchives = 4
    RequestTimeoutMs = 60000
    UseTenantSiteCmdlet = [bool] $UseTenantSiteCmdlet
    RecoveryDirectory = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'SPRegionalSettingsManager'
}

# These helpers are loaded into background runspaces as source. They have no UI dependencies.
$commonCode = {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    function Get-LocaleName([int] $Id) {
        switch ($Id) {
            1033 { 'English (United States)' }
            2057 { 'English (United Kingdom)' }
            3081 { 'English (Australia)' }
            5129 { 'English (New Zealand)' }
            default {
                try { [Globalization.CultureInfo]::GetCultureInfo($Id).EnglishName }
                catch [Globalization.CultureNotFoundException] { "Unknown ($Id)" }
            }
        }
    }

    function ConvertTo-AdminUrl([string] $Value) {
        $uri = $null
        if (-not [Uri]::TryCreate($Value.Trim(), 'Absolute', [ref] $uri) -or
            $uri.Scheme -ne 'https' -or $uri.Host -notmatch '^[a-z0-9-]+-admin\.sharepoint\.com$' -or
            -not $uri.IsDefaultPort -or $uri.UserInfo -or $uri.Query -or $uri.Fragment -or
            $uri.AbsolutePath -ne '/') {
            throw 'Enter an HTTPS SharePoint admin URL in the form https://tenant-admin.sharepoint.com.'
        }
        $uri.GetLeftPart([UriPartial]::Authority)
    }

    function ConvertTo-SiteUrl([string] $Value, $Hosts) {
        $uri = $null
        if (-not [Uri]::TryCreate($Value.Trim(), 'Absolute', [ref] $uri) -or
            $uri.Scheme -ne 'https' -or -not $uri.IsDefaultPort -or $uri.UserInfo -or
            $uri.Query -or $uri.Fragment -or -not $Hosts.Contains($uri.IdnHost)) {
            throw "Site URL is invalid or outside the configured tenant hosts: $Value"
        }
        $uri.AbsoluteUri.TrimEnd('/')
    }

    function Test-ExcludedSite([string] $Url, [string] $Template, $Config) {
        $uri = [Uri] $Url
        if (-not $Config.IncludeOneDrive -and ($Template -like 'SPSPERS#*' -or
            $uri.Host -like '*-my.sharepoint.com' -or $uri.AbsolutePath -like '/personal/*')) {
            return $true
        }
        foreach ($pattern in $Config.ExcludedTemplatePatterns) {
            if ($Template -like $pattern) { return $true }
        }
        foreach ($pattern in $Config.ExcludedUrlPatterns) {
            if ($Url -like $pattern) { return $true }
        }
        return $false
    }

    function Send-Log([string] $Level, [string] $Url, [string] $Message) {
        $format = '{0:yyyy-MM-dd HH:mm:ss.fff zzz} | {1} | {2} | {3}'
        $line = $format -f [DateTimeOffset]::Now, $Level, $Url, ($Message -replace '[\r\n]+', ' ')
        $shared.Events.Enqueue(@{ Kind = 'Log'; Line = $line })
        $shared.LogQueue.Add($line)
    }

    function Get-ErrorDisposition($Failure) {
        $transient = $false
        $delay = 0
        $auth = $false
        $permanent = $false
        for ($exception = $Failure.Exception; $null -ne $exception; $exception = $exception.InnerException) {
            $responseProperty = $exception.PSObject.Properties['Response']
            $response = if ($responseProperty) { $responseProperty.Value }
            $statusProperty = $exception.PSObject.Properties['StatusCode']
            $status = if ($statusProperty -and $null -ne $statusProperty.Value) { [int] $statusProperty.Value }
            if ($response -is [Net.HttpWebResponse]) {
                $status = [int] $response.StatusCode
                $retryAfter = $response.Headers['Retry-After']
                $seconds = 0
                $until = [DateTimeOffset]::MinValue
                if ([int]::TryParse($retryAfter, [ref] $seconds)) { $delay = [Math]::Max($delay, $seconds) }
                elseif ([DateTimeOffset]::TryParse($retryAfter, [ref] $until)) {
                    $delay = [Math]::Max($delay, [Math]::Ceiling(($until - [DateTimeOffset]::UtcNow).TotalSeconds))
                }
            }
            $message = $exception.Message
            if ($status -eq 401 -or $message -match 'AADSTS|MsalUiRequired|interaction_required|invalid_grant|\b401\b|unauthenticated|authentication needs attention') {
                $auth = $true
            }
            if ($status -in 400, 403, 404 -or $message -match '\b(403|404)\b|access (is )?denied|server relative url.*does not exist') {
                $permanent = $true
            }
            if ($status -in 408, 429, 500, 502, 503, 504 -or
                $exception -is [TimeoutException] -or $exception -is [Net.Sockets.SocketException] -or
                ($exception -is [Net.Http.HttpRequestException] -and -not $status) -or
                ($exception -is [Net.WebException] -and $exception.Status -ne 'ProtocolError') -or
                $message -match '\b(408|429|500|502|503|504)\b|throttl|maximum retry attempts|timed? ?out|temporarily unavailable|connection (was |has been )?(closed|reset|aborted)|name resolution|could not be resolved|Server did not retain') {
                $transient = $true
            }
        }
        @{ Retry = ($transient -and -not $auth -and -not $permanent); Delay = $delay; Authentication = $auth }
    }

    function Invoke-WithRetry([string] $Url, [string] $Description, [scriptblock] $Action, [switch] $Stoppable) {
        for ($attempt = 0; ; $attempt++) {
            if ($Stoppable) { $operation.Cancel.Token.ThrowIfCancellationRequested() }
            try { return (& $Action) }
            catch {
                $disposition = Get-ErrorDisposition $_
                if ($attempt -ge 3 -or -not $disposition.Retry) { throw }
                $delay = [Math]::Max([int] [Math]::Pow(2, $attempt + 1), $disposition.Delay)
                Send-Log 'WARN' $Url "$Description failed. Retry $($attempt + 1) of 3 in ${delay}s. $($_.Exception.Message)"
                if ($Stoppable) {
                    [void] $operation.Cancel.Token.WaitHandle.WaitOne([TimeSpan]::FromSeconds($delay))
                    $operation.Cancel.Token.ThrowIfCancellationRequested()
                } else {
                    [Threading.Thread]::Sleep([TimeSpan]::FromSeconds($delay))
                }
            }
        }
    }

    function New-SiteConnection([string] $Url) {
        # Serialize access to the interactive token source. Each site receives a
        # separate CSOM context, so queued CSOM actions cannot cross worker threads.
        $resource = ([Uri] $Url).GetLeftPart([UriPartial]::Authority)
        $session.TokenGate.Wait()
        try {
            if ($session.ResourceErrors.ContainsKey($resource)) {
                throw "Authentication needs attention for $resource. Reconnect to renew your sign-in."
            }
            $access = $null
            if (-not $session.Tokens.TryGetValue($resource, [ref] $access) -or
                $access.Expires -le [DateTimeOffset]::UtcNow.AddMinutes(5)) {
                try {
                    $token = Get-PnPAccessToken -ResourceUrl "$resource/.default" -Connection $session.Admin
                    if ($token -isnot [string] -or [string]::IsNullOrWhiteSpace($token)) {
                        throw 'The SharePoint access token is unavailable.'
                    }
                } catch {
                    $disposition = Get-ErrorDisposition $_
                    if (-not $disposition.Retry) { $session.ResourceErrors[$resource] = $true }
                    throw
                }
                # Read expiry for cache scheduling only. SharePoint validates
                # the signed token. Parsing here avoids an assembly dependency.
                $expires = [DateTimeOffset]::UtcNow.AddMinutes(6)
                try {
                    $segments = $token.Split('.')
                    if ($segments.Length -eq 3) {
                        $encoded = $segments[1].Replace('-', '+').Replace('_', '/')
                        $encoded = $encoded.PadRight($encoded.Length + ((4 - $encoded.Length % 4) % 4), '=')
                        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encoded))
                        $claims = ConvertFrom-Json -InputObject $json -AsHashtable
                        if ($claims.Contains('exp')) { $expires = [DateTimeOffset]::FromUnixTimeSeconds([long] $claims.exp) }
                    }
                } catch {
                    Write-Verbose 'Could not read access token expiry. Using a short cache lifetime.'
                }
                $access = @{ Token = $token; Expires = $expires }
                $session.Tokens[$resource] = $access
            }
        } finally {
            [void] $session.TokenGate.Release()
        }
        $connection = Connect-PnPOnline -Url $Url -AccessToken $access.Token -ReturnConnection
        $connection.Context.RequestTimeout = $config.RequestTimeoutMs
        # PnPClientContext overrides the cmdlet retry argument. Let the app
        # schedule the three retries with a fresh context and a live recheck.
        if ($connection.Context.PSObject.Properties['RetryCount']) { $connection.Context.RetryCount = 1 }
        $connection
    }

    function Write-AtomicCsv([string] $Path, $Records) {
        $temp = Join-Path ([IO.Path]::GetDirectoryName($Path)) ('.' + [Guid]::NewGuid() + '.tmp')
        try {
            $Records | Export-Csv -LiteralPath $temp -NoTypeInformation -Encoding utf8BOM
            [IO.File]::Move($temp, $Path, $true)
        } finally {
            if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
        }
    }

    function Read-RollbackCsv([string] $Path) {
        Add-Type -AssemblyName Microsoft.VisualBasic.Core
        $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
        $jobs = [Collections.Generic.List[object]]::new()
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $skipped = 0
        try {
            $parser.SetDelimiters(',')
            $parser.HasFieldsEnclosedInQuotes = $true
            $parser.TrimWhiteSpace = $false
            $header = $parser.ReadFields()
            if ($null -eq $header) { throw 'The backup CSV is empty.' }
            $columns = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::OrdinalIgnoreCase)
            for ($index = 0; $index -lt $header.Length; $index++) { $columns.Add($header[$index], $index) }
            foreach ($name in 'SiteUrl', 'SiteTitle', 'LocaleId', 'LocaleName', 'ExportDate') {
                if (-not $columns.ContainsKey($name)) { throw "The backup CSV is missing the $name column." }
            }
            while (-not $parser.EndOfData) {
                $line = $parser.LineNumber
                $fields = $parser.ReadFields()
                if ($fields.Length -ne $header.Length) { throw "Unexpected field count at CSV line $line." }
                $url = ConvertTo-SiteUrl $fields[$columns['SiteUrl']] $session.Hosts
                if (-not $seen.Add($url)) { throw "The backup CSV contains a duplicate site URL: $url" }
                if (Test-ExcludedSite $url '' $config) { throw "The backup contains an excluded site URL: $url" }
                $raw = $fields[$columns['LocaleId']]
                if ([string]::IsNullOrWhiteSpace($raw)) { $skipped++; continue }
                $id = 0
                if (-not [int]::TryParse($raw, [Globalization.NumberStyles]::None,
                    [Globalization.CultureInfo]::InvariantCulture, [ref] $id) -or $id -lt 1 -or $id -gt 65535) {
                    throw "Invalid LocaleId on CSV line ${line}: $raw"
                }
                $jobs.Add(@{ Url = $url; Title = $fields[$columns['SiteTitle']]; Target = $id })
            }
        } finally { $parser.Dispose() }
        if ($jobs.Count -eq 0) { throw 'The backup CSV contains no restorable locale values.' }
        @{ Jobs = $jobs.ToArray(); Skipped = $skipped }
    }
}

$loggerCode = {
    param($shared, $config)
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'
    $mutex = [Threading.Mutex]::new($false, 'Local\SPRegionalSettingsLog')
    $reported = $false
    $utf8 = [Text.UTF8Encoding]::new($false)
    try {
        foreach ($line in $shared.LogQueue.GetConsumingEnumerable()) {
            $owned = $false
            try {
                try { $owned = $mutex.WaitOne(2000) }
                catch [Threading.AbandonedMutexException] { $owned = $true }
                if (-not $owned) { throw 'The rolling log is in use.' }
                $path = $config.LogPath
                if ([IO.File]::Exists($path) -and [IO.FileInfo]::new($path).Length -ge $config.LogMaxBytes) {
                    for ($index = $config.LogArchives - 1; $index -ge 1; $index--) {
                        if ([IO.File]::Exists("$path.$index")) {
                            [IO.File]::Move("$path.$index", "$path.$($index + 1)", $true)
                        }
                    }
                    [IO.File]::Move($path, "$path.1", $true)
                }
                [IO.File]::AppendAllText($path, $line + [Environment]::NewLine, $utf8)
                $reported = $false
            } catch {
                if (-not $reported) {
                    $shared.Events.Enqueue(@{ Kind = 'Log'; Line = "WARN | File logging unavailable. $($_.Exception.Message)" })
                    $reported = $true
                }
            } finally {
                if ($owned) { $mutex.ReleaseMutex() }
            }
        }
    } finally { $mutex.Dispose() }
}

$producerCode = {
    param($shared, $config, $session, $operation, $jobs, [string] $commonSource)
    . ([scriptblock]::Create($commonSource))

    function Publish-Site($Site, $Seen) {
        $operation.Cancel.Token.ThrowIfCancellationRequested()
        if ([string]::IsNullOrWhiteSpace($Site.Url)) { return }
        if (Test-ExcludedSite $Site.Url $Site.Template $config) { return }
        try { $url = ConvertTo-SiteUrl $Site.Url $session.Hosts }
        catch {
            $operation.Notes.Enqueue($_.Exception.Message)
            Send-Log 'WARN' $Site.Url $_.Exception.Message
            return
        }
        if ($Seen.Add($url)) {
            $item = @{ Url = $url; Title = [string] $Site.Title; Target = $config.Target }
            $shared.Events.Enqueue(@{ Kind = 'Discovered'; Item = $item })
            $operation.Queue.Add($item, $operation.Cancel.Token)
        }
    }

    try {
        if ($operation.Mode -in 'Apply', 'Rollback') {
            $operation.Cancel.Token.ThrowIfCancellationRequested()
            # Open the recovery file on the producer thread before any work can
            # reach the site workers. Workers flush each original value before writing.
            [void] [IO.Directory]::CreateDirectory($config.RecoveryDirectory)
            $stamp = [DateTime]::Now.ToString('yyyyMMdd_HHmmss')
            $path = Join-Path $config.RecoveryDirectory "RegionalSettingsRecovery_$($operation.Mode)_${stamp}_$([Guid]::NewGuid().ToString('N')).csv"
            $stream = [IO.FileStream]::new($path, 'CreateNew', 'Write', 'Read')
            $operation.Recovery = @{ Stream = $stream; SyncRoot = [object]::new(); Failed = $false }
            $bytes = [Text.Encoding]::UTF8.GetBytes("SiteUrl,SiteTitle,LocaleId,LocaleName,ExportDate`r`n")
            $bom = [Text.UTF8Encoding]::new($true).GetPreamble()
            $stream.Write($bom, 0, $bom.Length)
            $stream.Write($bytes, 0, $bytes.Length)
            $stream.Flush($true)
            Send-Log 'INFO' '' "Recovery CSV: $path"
        }
        if ($operation.Mode -ne 'Discover') {
            foreach ($item in $jobs) {
                $operation.Cancel.Token.ThrowIfCancellationRequested()
                $shared.Events.Enqueue(@{ Kind = 'Pending'; Item = $item })
                $operation.Queue.Add($item, $operation.Cancel.Token)
            }
            return
        }
        $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($adminUrl in $session.AdminUrls) {
            $operation.Cancel.Token.ThrowIfCancellationRequested()
            Send-Log 'INFO' $adminUrl 'Enumerating site collections.'
            $pageState = @{ Connection = $null }
            try {
                $pageState.Connection = Invoke-WithRetry $adminUrl 'Admin connection' { New-SiteConnection $adminUrl } -Stoppable
                if ($config.UseTenantSiteCmdlet) {
                    # Retry collection before publishing any results, avoiding duplicates.
                    $sites = @(Invoke-WithRetry $adminUrl 'Tenant discovery' {
                        Get-PnPTenantSite -Detailed -Connection $pageState.Connection -IncludeOneDriveSites:$config.IncludeOneDrive
                    } -Stoppable)
                    foreach ($site in $sites) { Publish-Site $site $seen }
                    continue
                }
                # Use the same detailed CSOM endpoint as Get-PnPTenantSite and emit
                # each page immediately. Rebuild the context after a page retry.
                $cursor = $null
                $cursors = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
                do {
                    $page = Invoke-WithRetry $adminUrl 'Tenant discovery page' {
                        if ($null -eq $pageState.Connection) { $pageState.Connection = New-SiteConnection $adminUrl }
                        $geo = $pageState.Connection
                        try {
                            $tenant = [Microsoft.Online.SharePoint.TenantAdministration.Tenant]::new($geo.Context)
                            $filter = [Microsoft.Online.SharePoint.TenantAdministration.SPOSitePropertiesEnumerableFilter]::new()
                            $filter.IncludeDetail = $true
                            if ($config.IncludeOneDrive) { $filter.IncludePersonalSite = 'Include' }
                            else { $filter.IncludePersonalSite = 'UseServerDefault' }
                            if ($cursor) { $filter.StartIndex = $cursor }
                            $result = $tenant.GetSitePropertiesFromSharePointByFilters($filter)
                            $geo.Context.Load($result)
                            $geo.Context.ExecuteQuery()
                            @{ Items = @($result | ForEach-Object { $_ }); Next = $result.NextStartIndexFromSharePoint }
                        } catch {
                            # A failed request can leave its CSOM batch in an uncertain
                            # state. Dispose it and use a fresh context for the retry.
                            $geo.Context.Dispose()
                            $pageState.Connection = $null
                            throw
                        }
                    } -Stoppable
                    foreach ($site in $page.Items) { Publish-Site $site $seen }
                    $cursor = $page.Next
                    if ($cursor -and -not $cursors.Add($cursor)) { throw 'SharePoint returned a repeated discovery cursor.' }
                } while (-not [string]::IsNullOrWhiteSpace($cursor))
                Send-Log 'INFO' $adminUrl "Enumeration returned $($seen.Count) unique included sites so far."
            } catch [OperationCanceledException] { throw }
            catch {
                $operation.Notes.Enqueue("$adminUrl | $($_.Exception.Message)")
                Send-Log 'ERROR' $adminUrl "Discovery failed. $($_.Exception.Message)"
            } finally {
                if ($null -ne $pageState.Connection) { $pageState.Connection.Context.Dispose() }
            }
        }
    } catch [OperationCanceledException] {
        Send-Log 'INFO' '' 'Site queuing stopped.'
    } catch {
        $operation.Notes.Enqueue($_.Exception.Message)
        Send-Log 'ERROR' '' "Could not queue site operations. $($_.Exception.Message)"
    } finally {
        $operation.Queue.CompleteAdding()
    }
}

$workerCode = {
    param($shared, $config, $session, $operation, [string] $commonSource)
    . ([scriptblock]::Create($commonSource))

    function Save-BeforeValue([string] $Url, [string] $Title, [int] $Locale) {
        $record = [pscustomobject] [ordered] @{
            SiteUrl = $Url
            SiteTitle = $Title
            LocaleId = $Locale
            LocaleName = Get-LocaleName $Locale
            ExportDate = [DateTimeOffset]::UtcNow.ToString('o')
        }
        $csv = @($record | ConvertTo-Csv -NoTypeInformation)
        $bytes = [Text.Encoding]::UTF8.GetBytes($csv[1] + [Environment]::NewLine)
        [Threading.Monitor]::Enter($operation.Recovery.SyncRoot)
        try {
            if ($operation.Recovery.Failed) { throw 'The recovery CSV is unavailable. No further writes can be submitted.' }
            $stream = $operation.Recovery.Stream
            $start = $stream.Position
            try {
                $stream.Write($bytes, 0, $bytes.Length)
                $stream.Flush($true)
            } catch {
                $operation.Recovery.Failed = $true
                try { $stream.SetLength($start); $stream.Flush($true) }
                catch { Send-Log 'ERROR' $Url "Recovery CSV could not be truncated to its last complete record. $($_.Exception.Message)" }
                throw
            }
        } finally { [Threading.Monitor]::Exit($operation.Recovery.SyncRoot) }
    }

    foreach ($item in $operation.Queue.GetConsumingEnumerable($operation.Cancel.Token)) {
        if ($operation.Cancel.IsCancellationRequested) { break }
        $mode = $operation.Mode
        $isWrite = $mode -in 'Apply', 'Rollback'
        $state = @{
            Before = $null; Current = $null; WriteStarted = $false
            Journalled = $false; Title = $item.Title; After = $null
        }
        $shared.Events.Enqueue(@{ Kind = 'Active'; Item = $item })
        $status = 'Failed'
        $errorText = ''
        try {
            $status = Invoke-WithRetry $item.Url 'Site operation' {
                $connection = New-SiteConnection $item.Url
                try {
                    $web = Get-PnPWeb -Includes RegionalSettings, Title, Url, WebTemplate, Configuration -Connection $connection
                    if ($null -eq $web -or $null -eq $web.RegionalSettings) { throw 'Regional settings were unavailable.' }
                    if ((ConvertTo-SiteUrl $web.Url $session.Hosts) -ine $item.Url) {
                        throw "SharePoint returned a different site URL: $($web.Url)"
                    }
                    $template = "$($web.WebTemplate)#$($web.Configuration)"
                    if (Test-ExcludedSite $item.Url $template $config) { throw "Excluded site template: $template" }

                    if ($mode -eq 'Rollback') {
                        $site = Get-PnPSite -Includes Url -Connection $connection
                        if ((ConvertTo-SiteUrl $site.Url $session.Hosts) -ine $item.Url) {
                            throw 'The backup URL must reference a site collection root.'
                        }
                    }

                    $state.Title = [string] $web.Title
                    $regional = $web.RegionalSettings
                    if (-not $regional.IsPropertyAvailable('LocaleId')) {
                        $connection.Context.Load($regional)
                        Invoke-PnPQuery -RetryCount 1 -Connection $connection
                    }
                    $observed = [int] $regional.LocaleId
                    if ($observed -lt 1) { throw 'SharePoint did not return a valid LocaleId.' }
                    $state.Current = $observed
                    if (-not $isWrite) {
                        if ($observed -eq $config.Target) { return 'Already NZ' }
                        if ($mode -eq 'Discover') { return 'Discovered' }
                        return 'Refreshed'
                    }
                    if (-not $state.Journalled) { $state.Before = $observed }
                    $shared.Events.Enqueue(@{
                        Kind = 'Observed'; Item = $item; Before = $state.Before
                        Current = $observed; Title = $state.Title
                    })

                    # A retry starts with a fresh server read. If the earlier request
                    # committed, this branch confirms it without submitting again.
                    if ($observed -eq $item.Target) {
                        $state.After = $observed
                        if ($mode -eq 'Rollback') { return 'Restored' }
                        if ($state.WriteStarted) { return 'Updated' }
                        return 'Already NZ'
                    }
                    if (-not $state.Journalled) {
                        Save-BeforeValue $item.Url $state.Title $state.Before
                        $state.Journalled = $true
                    } elseif ($observed -ne $state.Before) {
                        throw "The site locale changed concurrently to $observed. Review the site before retrying."
                    }

                    $regional.LocaleId = [uint32] $item.Target
                    $regional.Update()
                    $state.WriteStarted = $true
                    $state.Current = $null
                    Invoke-PnPQuery -RetryCount 1 -Connection $connection
                    # Get-PnPWeb uses EnsureProperties and may return cached values
                    # on this context. Explicitly reload the object from the server.
                    $connection.Context.Load($regional)
                    Invoke-PnPQuery -RetryCount 1 -Connection $connection
                    $state.After = [int] $regional.LocaleId
                    $state.Current = $state.After
                    if ($state.After -ne $item.Target) {
                        throw "Server did not retain the requested locale. Expected $($item.Target), read $($state.After)."
                    }
                    if ($mode -eq 'Rollback') { return 'Restored' }
                    return 'Updated'
                } finally { $connection.Context.Dispose() }
            }

            $message = if ($isWrite) { "$status. Locale $($state.Before) to $($state.After)." }
                else { "${status}. Locale $($state.Current)." }
            Send-Log 'INFO' $item.Url $message
        } catch {
            $errorText = $_.Exception.Message
            $disposition = Get-ErrorDisposition $_
            if ($disposition.Authentication) { $errorText = "Authentication failed. Reconnect if prompted. $errorText" }
            if ($state.WriteStarted -and $null -eq $state.Current) {
                $errorText = "The write outcome could not be verified. Use Refresh to check the live value. $errorText"
            }
            if ($disposition.Authentication) {
                $resource = ([Uri] $item.Url).GetLeftPart([UriPartial]::Authority)
                $discard = $null
                [void] $session.Tokens.TryRemove($resource, [ref] $discard)
            }
            Send-Log 'ERROR' $item.Url "${mode}: $errorText"
        } finally {
            $shared.Events.Enqueue(@{
                Kind = 'Result'; Item = $item; Status = $status; Title = $state.Title
                Current = $state.Current; Before = $state.Before; After = $state.After
                Error = $errorText; At = [DateTimeOffset]::Now; IsWrite = $isWrite
            })
        }
    }
}

$taskCode = {
    param($shared, $config, $session, [string] $task, $data, [string] $commonSource)
    . ([scriptblock]::Create($commonSource))
    try {
        switch ($task) {
            'Connect' {
                Send-Log 'INFO' $session.AdminUrls[0] 'Starting interactive sign-in.'
                $admin = $null
                try {
                    $admin = Connect-PnPOnline -Url $session.AdminUrls[0] -ClientId $session.ClientId -Interactive -ValidateConnection -ReturnConnection
                    $admin.Context.RequestTimeout = $config.RequestTimeoutMs
                    $session.Admin = $admin
                    $shared.Events.Enqueue(@{ Kind = 'Connected' })
                    Send-Log 'INFO' $session.AdminUrls[0] 'Interactive sign-in completed.'
                } catch {
                    if ($null -ne $admin) { $admin.Context.Dispose() }
                    $session.Admin = $null
                    throw
                }
            }
            'Read Backup' {
                $payload = Read-RollbackCsv $data
                $shared.Events.Enqueue(@{ Kind = 'BackupReady'; Payload = $payload; Path = $data })
                Send-Log 'INFO' '' "Read rollback CSV with $($payload.Jobs.Count) restorable rows from $data."
            }
            'Export Settings' {
                $at = [DateTimeOffset]::UtcNow.ToString('o')
                $records = foreach ($row in $data.Table.Rows) {
                    [pscustomobject] [ordered] @{
                        SiteUrl = $row.SiteUrl
                        SiteTitle = $row.SiteTitle
                        LocaleId = if ($row.IsNull('LocaleId')) { $null } else { $row.LocaleId }
                        LocaleName = $row.LocaleName
                        ExportDate = $at
                    }
                }
                Write-AtomicCsv $data.Path $records
                Send-Log 'INFO' '' "Exported $($data.Table.Rows.Count) settings rows to $($data.Path)."
            }
            'Export Grid' {
                $records = foreach ($row in $data.Table.Rows) {
                    [pscustomobject] [ordered] @{
                        SiteUrl = $row.SiteUrl
                        SiteTitle = $row.SiteTitle
                        LocaleId = $row.LocaleId
                        LocaleName = $row.LocaleName
                        TargetLocaleId = $row.TargetLocaleId
                        BeforeLocaleId = $row.BeforeLocaleId
                        AfterLocaleId = $row.AfterLocaleId
                        Status = $row.Status
                        LastUpdated = $row.LastUpdated
                        Error = $row.Error
                    }
                }
                Write-AtomicCsv $data.Path $records
                Send-Log 'INFO' '' "Exported $($data.Table.Rows.Count) grid rows to $($data.Path)."
            }
        }
    } catch {
        Send-Log 'ERROR' '' "${task}: $($_.Exception.Message)"
        $shared.Events.Enqueue(@{ Kind = 'TaskError'; Message = $_.Exception.Message })
    } finally {
        if ($task -in 'Export Settings', 'Export Grid') { $data.Table.Dispose() }
    }
}

# The UI, binding and dialogs run on the main STA thread. Every background
# pipeline passes plain messages to the main thread through the event queue.
if (-not $IsWindows) { throw 'SharePoint Regional Settings Manager requires Windows.' }
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    throw 'Run with pwsh -NoProfile -STA -File and the path to this script.'
}
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
. $commonCode
$commonSource = $commonCode.ToString()
$appTitle = 'SharePoint Regional Settings Manager'

function Show-Message([string] $Message, [string] $Icon = 'Information') {
    [void] [Windows.Forms.MessageBox]::Show($form, $Message, $appTitle, 'OK', $Icon)
}

function Confirm-Action([string] $Message) {
    [Windows.Forms.MessageBox]::Show($form, $Message, $appTitle, 'YesNo', 'Warning', 'Button2') -eq 'Yes'
}

function Add-Pipeline($Pool, [scriptblock] $Code, [object[]] $Arguments, [string] $Role) {
    $pipeline = [PowerShell]::Create()
    try {
        $pipeline.RunspacePool = $Pool
        if ([object]::ReferenceEquals($Pool, $ui.Pool)) {
            # Load PnP on the same background pipeline that uses it.
            $wrapper = {
                param([string] $ModulePath, [string] $Source, [object[]] $WorkerArguments)
                Import-Module -Name $ModulePath -ErrorAction Stop
                & ([scriptblock]::Create($Source)) @WorkerArguments
            }
            [void] $pipeline.AddScript($wrapper.ToString()).AddArgument($ui.ModulePath).
                AddArgument($Code.ToString()).AddArgument($Arguments)
        } else {
            [void] $pipeline.AddScript($Code.ToString())
            foreach ($argument in $Arguments) { [void] $pipeline.AddArgument($argument) }
        }
        $handle = $pipeline.BeginInvoke()
        $ui.Pipelines.Add(@{ Instance = $pipeline; Handle = $handle; Role = $Role })
    } catch {
        $pipeline.Dispose()
        throw
    }
}

$form = [Windows.Forms.Form]::new()
$form.Text = $appTitle
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = [Drawing.Size]::new(1400, 800)
$form.Size = [Drawing.Size]::new(1510, 880)
$form.AutoScaleMode = 'Dpi'
$form.BackColor = [Drawing.Color]::FromArgb(241, 245, 249)
$form.ForeColor = [Drawing.Color]::FromArgb(26, 43, 58)
$form.Font = [Drawing.Font]::new('Segoe UI', 10)

$shared = @{
    Events = [Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
    LogQueue = [Collections.Concurrent.BlockingCollection[string]]::new()
}
$ui = @{
    Pipelines = [Collections.Generic.List[hashtable]]::new()
    Operation = $null; Task = ''; TaskError = ''; Backup = $null
    CloseAfterWork = $false; DiscoveryComplete = $false
    Pools = [Collections.Generic.List[object]]::new()
    LastSummary = 'Ready to connect'
}
$session = @{
    Admin = $null; AdminUrls = @(); ClientId = ''; Hosts = $null
    Tokens = [Collections.Concurrent.ConcurrentDictionary[string, hashtable]]::new([StringComparer]::OrdinalIgnoreCase)
    ResourceErrors = @{}
    TokenGate = [Threading.SemaphoreSlim]::new(1, 1)
}

$root = [Windows.Forms.TableLayoutPanel]::new()
$root.Dock = 'Fill'
$root.Padding = [Windows.Forms.Padding]::new(18, 12, 18, 0)
$root.ColumnCount = 1
$root.RowCount = 6
foreach ($height in 44, 47, 44, 39) { [void] $root.RowStyles.Add([Windows.Forms.RowStyle]::new('Absolute', $height)) }
[void] $root.RowStyles.Add([Windows.Forms.RowStyle]::new('Percent', 100))
[void] $root.RowStyles.Add([Windows.Forms.RowStyle]::new('Absolute', 208))
$form.Controls.Add($root)

$heading = [Windows.Forms.Label]::new()
$heading.Dock = 'Fill'
$heading.Text = $appTitle
$heading.Font = [Drawing.Font]::new('Segoe UI', 19, [Drawing.FontStyle]::Bold)
$root.Controls.Add($heading, 0, 0)

function New-FlowPanel {
    $panel = [Windows.Forms.FlowLayoutPanel]::new()
    $panel.Dock = 'Fill'
    $panel.WrapContents = $false
    $panel.AutoScroll = $true
    $panel
}

function Add-FieldLabel($Panel, [string] $Text) {
    $label = [Windows.Forms.Label]::new()
    $label.AutoSize = $true
    $label.Margin = [Windows.Forms.Padding]::new(3, 9, 5, 0)
    $label.Text = $Text
    $Panel.Controls.Add($label)
}

function Add-TextField($Panel, [int] $Width, [string] $Value, [string] $Name) {
    $box = [Windows.Forms.TextBox]::new()
    $box.Width = $Width
    $box.Margin = [Windows.Forms.Padding]::new(0, 5, 13, 0)
    $box.Text = $Value
    $box.AccessibleName = $Name
    $Panel.Controls.Add($box)
    $box
}

function Add-Button($Panel, [string] $Text, [switch] $Primary) {
    $button = [Windows.Forms.Button]::new()
    $button.Text = $Text
    $button.AutoSize = $true
    $button.MinimumSize = [Drawing.Size]::new(88, 33)
    $button.Padding = [Windows.Forms.Padding]::new(8, 0, 8, 0)
    $button.FlatStyle = 'Flat'
    $button.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(184, 201, 215)
    $button.BackColor = [Drawing.Color]::White
    if ($Primary) {
        $button.BackColor = [Drawing.Color]::FromArgb(23, 83, 127)
        $button.ForeColor = [Drawing.Color]::White
    }
    $Panel.Controls.Add($button)
    $button
}

$connectionPanel = New-FlowPanel
$root.Controls.Add($connectionPanel, 0, 1)
Add-FieldLabel $connectionPanel 'Tenant Admin URL'
$urlBox = Add-TextField $connectionPanel 395 $TenantAdminUrl 'Tenant Admin URL'
Add-FieldLabel $connectionPanel 'Entra Client ID'
$clientBox = Add-TextField $connectionPanel 334 $ClientId 'Entra Client ID'
$connectButton = Add-Button $connectionPanel 'Connect' -Primary
$connectionLabel = [Windows.Forms.Label]::new()
$connectionLabel.AutoSize = $true
$connectionLabel.Margin = [Windows.Forms.Padding]::new(12, 9, 0, 0)
$connectionPanel.Controls.Add($connectionLabel)

$toolbar = New-FlowPanel
$root.Controls.Add($toolbar, 0, 2)
$discoverButton = Add-Button $toolbar 'Discover Sites'
$exportButton = Add-Button $toolbar 'Export Settings'
$applyButton = Add-Button $toolbar 'Apply NZ Locale' -Primary
$rollbackButton = Add-Button $toolbar 'Rollback'
$refreshButton = Add-Button $toolbar 'Refresh'
$stopButton = Add-Button $toolbar 'Stop'
$stateButton = Add-Button $toolbar 'Grid State CSV'

$filters = New-FlowPanel
$root.Controls.Add($filters, 0, 3)
Add-FieldLabel $filters 'Search'
$searchBox = Add-TextField $filters 315 '' 'Search sites'
Add-FieldLabel $filters 'Search field'
$searchField = [Windows.Forms.ComboBox]::new()
$searchField.AccessibleName = 'Search field'
$searchField.Width = 175
$searchField.DropDownStyle = 'DropDownList'
$searchField.Margin = [Windows.Forms.Padding]::new(2, 5, 10, 0)
[void] $searchField.Items.AddRange(@('All fields', 'Site URL', 'Title', 'Locale', 'Status'))
$searchField.SelectedIndex = 0
$filters.Controls.Add($searchField)
$visibleLabel = [Windows.Forms.Label]::new()
$visibleLabel.AutoSize = $true
$visibleLabel.Margin = [Windows.Forms.Padding]::new(6, 9, 0, 0)
$filters.Controls.Add($visibleLabel)

# A typed table keeps LCID sorting numeric. Binding owns row indexing and
# filtering without duplicating every site as a separate UI object.
$table = [Data.DataTable]::new('Sites')
$table.Locale = [Globalization.CultureInfo]::InvariantCulture
$columnSpec = @(
    @('SiteUrl', 'Site URL', [string], 325), @('SiteTitle', 'Title', [string], 175),
    @('LocaleId', 'Current LocaleId', [int], 120), @('LocaleName', 'Current Locale', [string], 165),
    @('TargetLocaleId', 'Target LocaleId', [int], 115), @('BeforeLocaleId', 'Before LocaleId', [int], 115),
    @('AfterLocaleId', 'After LocaleId', [int], 110), @('Status', 'Status', [string], 120),
    @('LastUpdated', 'Last Updated', [string], 163), @('Error', 'Error', [string], 390)
)
foreach ($spec in $columnSpec) { [void] $table.Columns.Add($spec[0], $spec[2]) }
$table.PrimaryKey = @($table.Columns['SiteUrl'])
$view = $table.DefaultView
$source = [Windows.Forms.BindingSource]::new()
$source.DataSource = $view

$grid = [Windows.Forms.DataGridView]::new()
$grid.Dock = 'Fill'
$grid.AutoGenerateColumns = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToOrderColumns = $true
$grid.AllowUserToResizeRows = $false
$grid.ReadOnly = $true
$grid.MultiSelect = $true
$grid.SelectionMode = 'FullRowSelect'
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = 'None'
$grid.BackgroundColor = [Drawing.Color]::White
$grid.GridColor = [Drawing.Color]::FromArgb(222, 232, 240)
$grid.RowTemplate.Height = 29
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(32, 62, 83)
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $grid.ColumnHeadersDefaultCellStyle.BackColor
$grid.ColumnHeadersHeight = 35
$grid.DefaultCellStyle.SelectionBackColor = [Drawing.Color]::FromArgb(21, 70, 111)
foreach ($spec in $columnSpec) {
    $column = [Windows.Forms.DataGridViewTextBoxColumn]::new()
    $column.Name = $spec[0]
    $column.HeaderText = $spec[1]
    $column.DataPropertyName = $spec[0]
    $column.Width = $spec[3]
    $column.SortMode = 'Automatic'
    [void] $grid.Columns.Add($column)
}
$grid.DataSource = $source
$root.Controls.Add($grid, 0, 4)

$footer = [Windows.Forms.SplitContainer]::new()
$footer.Dock = 'Fill'
$footer.Panel1MinSize = 0
$footer.FixedPanel = 'Panel1'
$footer.Margin = [Windows.Forms.Padding]::new(3, 8, 3, 4)
$root.Controls.Add($footer, 0, 5)

$statsLayout = [Windows.Forms.TableLayoutPanel]::new()
$statsLayout.Dock = 'Fill'
$statsLayout.Padding = [Windows.Forms.Padding]::new(8, 3, 12, 7)
$statsLayout.RowCount = 6
$statsLayout.ColumnCount = 2
foreach ($percent in 50, 50) { [void] $statsLayout.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Percent', $percent)) }
foreach ($height in 30, 29, 29, 29, 29, 22) { [void] $statsLayout.RowStyles.Add([Windows.Forms.RowStyle]::new('Absolute', $height)) }
$statsHeading = [Windows.Forms.Label]::new()
$statsHeading.Text = 'Operation progress'
$statsHeading.Font = [Drawing.Font]::new('Segoe UI Semibold', 11)
$statsHeading.Dock = 'Fill'
$statsLayout.Controls.Add($statsHeading, 0, 0)
$statsLayout.SetColumnSpan($statsHeading, 2)
$statsLabels = @{}
$statKeys = @('Total Sites', 'Processed', 'Already NZ', 'Updated', 'Restored', 'Failed', 'Pending', 'Remaining')
for ($i = 0; $i -lt $statKeys.Length; $i++) {
    $label = [Windows.Forms.Label]::new()
    $label.Dock = 'Fill'
    $label.TextAlign = 'MiddleLeft'
    $statsLabels[$statKeys[$i]] = $label
    $statsLayout.Controls.Add($label, ($i % 2), (1 + [int] [Math]::Floor($i / 2)))
}
$progress = [Windows.Forms.ProgressBar]::new()
$progress.Dock = 'Fill'
$progress.AccessibleName = 'Site operation progress'
$statsLayout.Controls.Add($progress, 0, 5)
$statsLayout.SetColumnSpan($progress, 2)
$footer.Panel1.Controls.Add($statsLayout)

$logGroup = [Windows.Forms.GroupBox]::new()
$logGroup.Text = 'Live log'
$logGroup.Dock = 'Fill'
$logGroup.Padding = [Windows.Forms.Padding]::new(7)
$liveLog = [Windows.Forms.RichTextBox]::new()
$liveLog.Dock = 'Fill'
$liveLog.ReadOnly = $true
$liveLog.DetectUrls = $false
$liveLog.WordWrap = $false
$liveLog.Font = [Drawing.Font]::new('Consolas', 9)
$liveLog.BackColor = [Drawing.Color]::White
$logGroup.Controls.Add($liveLog)
$footer.Panel2.Controls.Add($logGroup)

$statusBar = [Windows.Forms.StatusStrip]::new()
$statusBar.ShowItemToolTips = $true
$statusLabel = [Windows.Forms.ToolStripStatusLabel]::new()
$statusLabel.Spring = $true
$statusLabel.TextAlign = 'MiddleLeft'
$logPathLabel = [Windows.Forms.ToolStripStatusLabel]::new()
$logPathLabel.Text = $config.LogPath
[void] $statusBar.Items.Add($statusLabel)
[void] $statusBar.Items.Add($logPathLabel)
$form.Controls.Add($statusBar)

$colours = @{
    'Discovered' = [Drawing.Color]::White
    'Pending' = [Drawing.Color]::LightYellow
    'Reading' = [Drawing.Color]::LightYellow
    'Applying' = [Drawing.Color]::LightYellow
    'Rolling back' = [Drawing.Color]::LightYellow
    'Already NZ' = [Drawing.Color]::LightGreen
    'Updated' = [Drawing.Color]::LightBlue
    'Restored' = [Drawing.Color]::LightCyan
    'Failed' = [Drawing.Color]::LightCoral
}

$grid.Add_CellFormatting({
    param($sender, $eventArgs)
    if ($eventArgs.RowIndex -ge 0 -and $eventArgs.RowIndex -lt $view.Count) {
        $row = $view[$eventArgs.RowIndex]
        $colour = $colours[[string] $row['Status']]
        if ($null -ne $colour) { $eventArgs.CellStyle.BackColor = $colour }
    }
})

$filterTimer = [Windows.Forms.Timer]::new()
$filterTimer.Interval = 180
$tickTimer = [Windows.Forms.Timer]::new()
$tickTimer.Interval = 70

function Set-RowFilter {
    $literal = [regex]::Replace($searchBox.Text.Trim(), "['%*\[\]]", {
        param($match)
        if ($match.Value -eq "'") { return "''" }
        '[' + $match.Value + ']'
    })
    $columns = switch ($searchField.SelectedIndex) {
        1 { 'SiteUrl' }
        2 { 'SiteTitle' }
        3 { 'LocaleName'; 'LocaleId' }
        4 { 'Status' }
        default { 'SiteUrl'; 'SiteTitle'; 'LocaleName'; 'LocaleId'; 'Status' }
    }
    if (-not $literal) { $view.RowFilter = '' }
    else {
        $parts = foreach ($column in $columns) {
            if ($column -eq 'LocaleId') { "Convert(LocaleId, 'System.String') LIKE '%$literal%'" }
            else { "$column LIKE '%$literal%'" }
        }
        $view.RowFilter = $parts -join ' OR '
    }
    $grid.ClearSelection()
    $visibleLabel.Text = "$($view.Count.ToString('N0')) visible"
}

function New-RunStats {
    @{ Total = 0; Processed = 0; Active = 0; 'Already NZ' = 0; Updated = 0; Restored = 0; Failed = 0 }
}
$ui.Stats = New-RunStats

function Update-Chrome {
    $running = $null -ne $ui.Operation
    $busy = $running -or [bool] $ui.Task
    $connected = $null -ne $session.Admin
    $matchesInputs = $connected -and $urlBox.Text.Trim().TrimEnd('/') -ieq $session.AdminUrls[0] -and
        $clientBox.Text.Trim() -ieq $session.ClientId
    $inventory = $table.Rows.Count
    $connectButton.Enabled = -not $busy
    $urlBox.Enabled = -not $busy
    $clientBox.Enabled = -not $busy
    $discoverButton.Enabled = $matchesInputs -and -not $busy
    $exportButton.Enabled = $inventory -gt 0 -and -not $busy
    $stateButton.Enabled = $exportButton.Enabled
    $applyButton.Enabled = $matchesInputs -and $ui.Inventory.Count -gt 0 -and -not $busy
    $refreshButton.Enabled = $matchesInputs -and $inventory -gt 0 -and -not $busy
    $rollbackButton.Enabled = $matchesInputs -and -not $busy
    $stopButton.Enabled = $running -and -not $ui.Operation.Cancel.IsCancellationRequested
    $connectionLabel.Text = if ($ui.Task -eq 'Connect') { 'Signing in…' }
        elseif ($connected -and -not $matchesInputs) { 'Reconnect to use new settings' }
        elseif ($connected) { 'Connected' } else { 'Disconnected' }
    $connectButton.Text = if ($connected) { 'Reconnect' } else { 'Connect' }
    $stats = $ui.Stats
    $remaining = [Math]::Max(0, $stats.Total - $stats.Processed)
    $pending = [Math]::Max(0, $remaining - $stats.Active)
    $statsHeading.Text = "Operation progress  |  Sites Discovered  $($ui.Inventory.Count.ToString('N0'))"
    $statsLabels['Total Sites'].Text = "Total Sites  $($inventory.ToString('N0'))"
    $statsLabels['Processed'].Text = "Processed  $($stats.Processed.ToString('N0'))"
    foreach ($key in 'Already NZ', 'Updated', 'Restored', 'Failed') {
        $statsLabels[$key].Text = "$key  $($stats[$key].ToString('N0'))"
    }
    $statsLabels['Pending'].Text = "Pending  $($pending.ToString('N0'))"
    $statsLabels['Remaining'].Text = "Remaining  $($remaining.ToString('N0'))"
    $indeterminate = $running -and $ui.Operation.Mode -eq 'Discover' -and -not $ui.Operation.Queue.IsAddingCompleted
    $progress.Style = if ($indeterminate) { 'Marquee' } else { 'Continuous' }
    $progress.Maximum = [Math]::Max(1, $stats.Total)
    $progress.Value = [Math]::Min($progress.Maximum, $stats.Processed)
    $visibleLabel.Text = "$($view.Count.ToString('N0')) visible"
    $statusLabel.Text = if ($running) {
        if ($ui.Operation.Cancel.IsCancellationRequested) { 'Stopping. Waiting for active site operations to finish.' }
        else { "$($ui.Operation.Mode)  |  $($stats.Active) active  |  $Concurrency parallel workers" }
    } elseif ($ui.Task) { "$($ui.Task) in progress" } else { $ui.LastSummary }
    $statusLabel.ToolTipText = $statusLabel.Text
}

$ui.Inventory = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$ui.Pool = $null
$ui.ServicePool = $null

function New-AppPool([int] $MaxThreads) {
    $initial = [Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $initial.ThrowOnRunspaceOpenError = $true
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $MaxThreads, $initial, $Host)
    try {
        $pool.ApartmentState = 'MTA'
        $pool.ThreadOptions = 'ReuseThread'
        $pool.Open()
        $ui.Pools.Add($pool)
        $pool
    } catch {
        $pool.Dispose()
        throw
    }
}

function Find-PnPModule {
    Get-Module -ListAvailable -Name PnP.PowerShell |
        Where-Object {
            $_.Version -ge [version] '2.12.0' -and
            ($null -eq $_.PowerShellVersion -or $_.PowerShellVersion -le $PSVersionTable.PSVersion)
        } |
        Sort-Object Version -Descending | Select-Object -First 1
}

function Get-PnPPool {
    if ($null -eq $ui.Pool) { throw 'The PnP runspace pool is unavailable.' }
    $ui.Pool
}

function Set-RowFields($Row, [hashtable] $Fields) {
    $Row.BeginEdit()
    try {
        foreach ($name in $Fields.Keys) {
            $value = $Fields[$name]
            $Row[$name] = if ($null -eq $value) { [DBNull]::Value } else { $value }
        }
        $Row.EndEdit()
    } catch {
        $Row.CancelEdit()
        throw
    }
}

function Get-OrAddRow($Item) {
    $row = $table.Rows.Find($Item.Url)
    if ($null -eq $row) {
        $row = $table.NewRow()
        Set-RowFields $row @{ SiteUrl = $Item.Url; SiteTitle = $Item.Title
            TargetLocaleId = $Item.Target; Status = 'Discovered' }
        $table.Rows.Add($row)
    }
    $row
}

function Start-UiTask([string] $Name, $Data) {
    $pool = if ($Name -eq 'Connect') { Get-PnPPool } else { $ui.ServicePool }
    Add-Pipeline $pool $taskCode @($shared, $config, $session, $Name, $Data, $commonSource) 'Task'
    $ui.Task = $Name
    $ui.TaskError = ''
    Update-Chrome
}

function Start-SiteOperation([string] $Mode, [object[]] $Jobs = @()) {
    $pool = Get-PnPPool
    $operation = @{
        Mode = $Mode
        Cancel = [Threading.CancellationTokenSource]::new()
        Queue = [Collections.Concurrent.BlockingCollection[hashtable]]::new(256)
        Notes = [Collections.Concurrent.ConcurrentQueue[string]]::new()
        Recovery = $null
    }
    $ui.Operation = $operation
    $ui.Stats = New-RunStats
    $ui.Stats.Total = $Jobs.Count
    if ($Mode -eq 'Discover') {
        $ui.DiscoveryComplete = $false
        $ui.Inventory.Clear()
        $table.Clear()
    }
    try {
        # The producer completes the queue in its finally block. Workers exit
        # once the queue drains or Stop signals cancellation.
        Add-Pipeline $pool $producerCode @($shared, $config, $session, $operation, $Jobs, $commonSource) 'Producer'
        for ($worker = 0; $worker -lt $Concurrency; $worker++) {
            Add-Pipeline $pool $workerCode @($shared, $config, $session, $operation, $commonSource) 'Worker'
        }
    } catch {
        $operation.Notes.Enqueue("Could not start background processing. $($_.Exception.Message)")
        $operation.Cancel.Cancel()
    }
    Send-Log 'INFO' '' "Started $Mode."
    Update-Chrome
}

function Stop-SiteOperation {
    if ($null -ne $ui.Operation -and -not $ui.Operation.Cancel.IsCancellationRequested) {
        $ui.Operation.Cancel.Cancel()
        Send-Log 'INFO' '' "Stop requested during $($ui.Operation.Mode)."
        Update-Chrome
    }
}

function Complete-SiteOperation {
    $operation = $ui.Operation
    $stats = $ui.Stats
    $remaining = [Math]::Max(0, $stats.Total - $stats.Processed)
    $cancelled = $operation.Cancel.IsCancellationRequested
    if ($null -ne $operation.Recovery) { $operation.Recovery.Stream.Dispose() }
    if ($operation.Mode -eq 'Discover') {
        $ui.DiscoveryComplete = -not $cancelled -and $operation.Notes.IsEmpty -and $remaining -eq 0
    }
    $end = if ($cancelled) { 'stopped' } elseif (-not $operation.Notes.IsEmpty -or $remaining) { 'incomplete' } else { 'finished' }
    $ui.LastSummary = "$($operation.Mode) $end. Processed $($stats.Processed) of $($stats.Total). Updated $($stats.Updated). Restored $($stats.Restored). Already NZ $($stats['Already NZ']). Failed $($stats.Failed). Remaining $remaining."
    if ($cancelled -and $operation.Mode -eq 'Discover') { $ui.LastSummary += ' Tenant enumeration may be partial.' }
    foreach ($note in $operation.Notes.ToArray()) { Send-Log 'ERROR' '' $note }
    $level = if ($end -ne 'finished' -or $stats.Failed) { 'WARN' } else { 'INFO' }
    Send-Log $level '' $ui.LastSummary
    if ($remaining -gt 0) {
        foreach ($row in $table.Select("Status = 'Pending'")) {
            Set-RowFields $row @{ Status = 'Not processed' }
        }
    }
    $operation.Queue.Dispose()
    $operation.Cancel.Dispose()
    $ui.Operation = $null
}

function Read-AppEvents {
    $logBatch = [Text.StringBuilder]::new()
    $budget = [Diagnostics.Stopwatch]::StartNew()
    $entry = $null
    for ($count = 0; $count -lt 300 -and ($count -eq 0 -or $budget.ElapsedMilliseconds -lt 24); $count++) {
        if (-not $shared.Events.TryDequeue([ref] $entry)) { break }
        switch ($entry.Kind) {
            'Log' { [void] $logBatch.AppendLine($entry.Line) }
            'Connected' { $ui.LastSummary = "Connected to $($session.AdminUrls[0])" }
            'BackupReady' { $ui.Backup = $entry }
            'TaskError' { $ui.TaskError = $entry.Message }
            'Discovered' {
                if ($ui.Inventory.Add($entry.Item.Url)) { $ui.Stats.Total++ }
                $null = Get-OrAddRow $entry.Item
            }
            'Pending' {
                $row = Get-OrAddRow $entry.Item
                Set-RowFields $row @{ Status = 'Pending'; TargetLocaleId = $entry.Item.Target }
            }
            'Active' {
                $row = Get-OrAddRow $entry.Item
                $operationState = switch ($ui.Operation.Mode) {
                    'Apply' { 'Applying' }
                    'Rollback' { 'Rolling back' }
                    default { 'Reading' }
                }
                $fields = @{ Status = $operationState; Error = '' }
                if ($ui.Operation.Mode -in 'Apply', 'Rollback') {
                    $fields.TargetLocaleId = $entry.Item.Target
                    $fields.BeforeLocaleId = $null
                    $fields.AfterLocaleId = $null
                }
                Set-RowFields $row $fields
                $ui.Stats.Active++
            }
            'Observed' {
                $row = Get-OrAddRow $entry.Item
                Set-RowFields $row @{
                    SiteTitle = $entry.Title; LocaleId = $entry.Current
                    LocaleName = Get-LocaleName $entry.Current; BeforeLocaleId = $entry.Before
                }
            }
            'Result' {
                $row = Get-OrAddRow $entry.Item
                $fields = @{
                    SiteTitle = $entry.Title; LocaleId = $entry.Current
                    LocaleName = if ($null -ne $entry.Current) { Get-LocaleName $entry.Current } else { '' }
                    Status = $entry.Status; LastUpdated = $entry.At.ToString('yyyy-MM-dd HH:mm:ss')
                    Error = $entry.Error
                }
                if ($entry.IsWrite) {
                    $fields.BeforeLocaleId = $entry.Before
                    $fields.AfterLocaleId = $entry.After
                }
                Set-RowFields $row $fields
                $ui.Stats.Active--
                $ui.Stats.Processed++
                if ($ui.Stats.ContainsKey($entry.Status)) { $ui.Stats[$entry.Status]++ }
            }
        }
    }
    if ($logBatch.Length -gt 0) {
        if ($liveLog.TextLength -gt 220000) {
            $liveLog.Select(0, 80000)
            $liveLog.SelectedText = ''
        }
        $liveLog.AppendText($logBatch.ToString())
        $liveLog.ScrollToCaret()
    }
}

function Start-ApprovedRollback {
    $backup = $ui.Backup
    $ui.Backup = $null
    $jobs = $backup.Payload.Jobs
    $message = "Restore the LocaleId values for $($jobs.Count) sites using this CSV?`r`n`r`n$($backup.Path)"
    if ($backup.Payload.Skipped) {
        $message += "`r`n`r`n$($backup.Payload.Skipped) rows have no saved LocaleId and will be skipped."
    }
    if (Confirm-Action $message) {
        Send-Log 'INFO' '' "Rollback confirmed. Backup: $($backup.Path)"
        Start-SiteOperation 'Rollback' $jobs
    } else { $ui.LastSummary = 'Rollback cancelled before changes' }
}

function Complete-UiTask {
    $name = $ui.Task
    $ui.Task = ''
    if ($ui.TaskError) {
        $message = $ui.TaskError
        $ui.LastSummary = "$name failed. See the live log."
        if (-not $ui.CloseAfterWork) { Show-Message $message 'Error' }
        return
    }
    $ui.LastSummary = "$name completed"
    if ($name -eq 'Connect') { $ui.LastSummary = "Connected to $($session.AdminUrls[0])" }
    if ($name -eq 'Read Backup' -and $null -ne $ui.Backup -and -not $ui.CloseAfterWork) {
        Start-ApprovedRollback
    }
}

function Invoke-UiAction([scriptblock] $Action) {
    try { & $Action }
    catch {
        $message = $_.Exception.Message
        $ui.LastSummary = $message
        Send-Log 'ERROR' '' $message
        Show-Message $message 'Error'
        Update-Chrome
    }
}

function Receive-Pipelines {
    for ($index = $ui.Pipelines.Count - 1; $index -ge 0; $index--) {
        $job = $ui.Pipelines[$index]
        if (-not $job.Handle.IsCompleted) { continue }
        try {
            $null = $job.Instance.EndInvoke($job.Handle)
            if ($job.Instance.HadErrors) {
                $failure = $job.Instance.Streams.Error | Select-Object -First 1
                if ($null -ne $failure) { throw $failure }
            }
        } catch {
            $isCancelled = $null -ne $ui.Operation -and $ui.Operation.Cancel.IsCancellationRequested
            $exception = $_.Exception
            while ($null -ne $exception.InnerException) { $exception = $exception.InnerException }
            if (-not ($isCancelled -and $exception -is [OperationCanceledException])) {
                $message = $exception.Message
                Send-Log 'ERROR' '' "$($job.Role): $message"
                if ($job.Role -eq 'Task') { $ui.TaskError = $message }
                if ($null -ne $ui.Operation) {
                    $ui.Operation.Notes.Enqueue($message)
                    # A worker bootstrap failure could leave the producer waiting
                    # on a full queue. Cancelling releases that wait immediately.
                    $ui.Operation.Cancel.Cancel()
                }
            }
        } finally {
            $job.Instance.Dispose()
            $ui.Pipelines.RemoveAt($index)
        }
    }
}

function Save-Grid([switch] $State) {
    $dialog = [Windows.Forms.SaveFileDialog]::new()
    try {
        $dialog.Filter = 'CSV files (*.csv)|*.csv'
        $dialog.DefaultExt = 'csv'
        $prefix = if ($State) { 'RegionalSettingsState' } else { 'RegionalSettingsBackup' }
        $dialog.FileName = "$($prefix)_$([DateTime]::Now.ToString('yyyyMMdd_HHmmss')).csv"
        if ($dialog.ShowDialog($form) -ne 'OK') { return }
        $path = $dialog.FileName
    } finally { $dialog.Dispose() }
    if (-not $State) {
        $missing = $table.Select('LocaleId IS NULL').Length
        if ($missing -gt 0 -and -not (Confirm-Action "The current locale is unavailable for $missing sites. Their CSV rows will contain a blank LocaleId and cannot be used for rollback. Export anyway?")) {
            return
        }
    }
    $task = if ($State) { 'Export Grid' } else { 'Export Settings' }
    $snapshot = $table.Copy()
    try { Start-UiTask $task @{ Path = $path; Table = $snapshot } }
    catch { $snapshot.Dispose(); throw }
}

$filterTimer.Add_Tick({
    $filterTimer.Stop()
    Invoke-UiAction { Set-RowFilter }
})
$searchBox.Add_TextChanged({ $filterTimer.Stop(); $filterTimer.Start() })
$searchField.Add_SelectedIndexChanged({ $filterTimer.Stop(); $filterTimer.Start() })
$form.Add_Shown({
    $footer.SplitterDistance = 403
    $footer.Panel1MinSize = 365
    $footer.Panel2MinSize = 400
    $grid.ClearSelection()
})
$grid.Add_DataBindingComplete({
    param($sender, $eventArgs)
    if ($eventArgs.ListChangedType -eq 'Reset') { $grid.ClearSelection() }
})

$urlBox.Add_TextChanged({ Update-Chrome })
$clientBox.Add_TextChanged({ Update-Chrome })
$connectButton.Add_Click({
    Invoke-UiAction {
        $candidate = ConvertTo-AdminUrl $urlBox.Text
        $clientGuid = [Guid]::Empty
        if (-not [Guid]::TryParse($clientBox.Text.Trim(), [ref] $clientGuid) -or $clientGuid -eq [Guid]::Empty) {
            throw 'Enter the Entra application client ID as a GUID.'
        }
        $admins = [Collections.Generic.List[string]]::new()
        $hosts = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($admin in @($candidate) + $AdditionalAdminUrls) {
            $admin = ConvertTo-AdminUrl $admin
            if ($admins.Contains($admin)) { continue }
            $admins.Add($admin)
            $siteHost = ([Uri] $admin).Host -replace '-admin\.sharepoint\.com$', '.sharepoint.com'
            [void] $hosts.Add($siteHost)
            if ($config.IncludeOneDrive) { [void] $hosts.Add($siteHost -replace '\.sharepoint\.com$', '-my.sharepoint.com') }
        }
        if ($session.AdminUrls.Count -gt 0 -and $session.AdminUrls[0] -ne $candidate) {
            if ($table.Rows.Count -gt 0 -and -not (Confirm-Action 'Switching tenants will clear the current inventory from the grid. Continue?')) { return }
            $table.Clear()
            $ui.Inventory.Clear()
            $ui.DiscoveryComplete = $false
            $ui.Stats = New-RunStats
        }
        if ($null -ne $session.Admin) { $session.Admin.Context.Dispose(); $session.Admin = $null }
        $session.Tokens.Clear()
        $session.ResourceErrors.Clear()
        $session.AdminUrls = $admins.ToArray()
        $session.Hosts = $hosts
        $session.ClientId = $clientGuid.ToString()
        $urlBox.Text = $candidate
        $clientBox.Text = $session.ClientId
        Start-UiTask 'Connect' $null
    }
})
$discoverButton.Add_Click({ Invoke-UiAction { Start-SiteOperation 'Discover' } })
$exportButton.Add_Click({ Invoke-UiAction { Save-Grid } })
$stateButton.Add_Click({ Invoke-UiAction { Save-Grid -State } })
$stopButton.Add_Click({ Invoke-UiAction { Stop-SiteOperation } })
$applyButton.Add_Click({
    Invoke-UiAction {
        $total = $ui.Inventory.Count
        if ($total -eq 0) { throw 'Run Discover Sites before applying the NZ locale.' }
        $message = "You are about to modify regional settings on up to $total sites. Sites already set to 5129 will be skipped after a live check.`r`n`r`nContinue?"
        if (-not $ui.DiscoveryComplete) {
            $message += "`r`n`r`nDiscovery was incomplete. Only discovered sites will be checked."
        }
        if (Confirm-Action $message) {
            $jobs = foreach ($url in $ui.Inventory) {
                $row = $table.Rows.Find($url)
                @{ Url = $url; Title = $row.SiteTitle; Target = $config.Target }
            }
            Start-SiteOperation 'Apply' @($jobs)
        }
    }
})
$rollbackButton.Add_Click({
    Invoke-UiAction {
        $dialog = [Windows.Forms.OpenFileDialog]::new()
        try {
            $dialog.Filter = 'CSV files (*.csv)|*.csv'
            $dialog.CheckFileExists = $true
            if ($dialog.ShowDialog($form) -eq 'OK') { Start-UiTask 'Read Backup' $dialog.FileName }
        } finally { $dialog.Dispose() }
    }
})
$refreshButton.Add_Click({
    Invoke-UiAction {
        $selected = @($grid.SelectedRows | ForEach-Object {
            if ($null -ne $_.DataBoundItem) { $_.DataBoundItem.Row }
        })
        $rows = $table.Rows
        if ($selected.Count -gt 0 -and $selected.Count -lt $table.Rows.Count) {
            $choice = [Windows.Forms.MessageBox]::Show($form,
                "Refresh $($selected.Count) selected sites? Choose No to refresh all $($table.Rows.Count) sites.",
                $appTitle, 'YesNoCancel', 'Question')
            if ($choice -eq 'Cancel') { return }
            if ($choice -eq 'Yes') { $rows = $selected }
        }
        $jobs = foreach ($row in $rows) {
            @{ Url = $row.SiteUrl; Title = $row.SiteTitle; Target = $config.Target }
        }
        Start-SiteOperation 'Refresh' @($jobs)
    }
})

$form.Add_FormClosing({
    param($sender, $closeArgs)
    if ($null -ne $ui.Operation -or $ui.Task) {
        $closeArgs.Cancel = $true
        $ui.CloseAfterWork = $true
        if ($null -ne $ui.Operation) { Stop-SiteOperation }
        else { $statusLabel.Text = "Closing after $($ui.Task) finishes." }
    }
})

$tickTimer.Add_Tick({
    $tickTimer.Stop()
    try {
        Receive-Pipelines
        Read-AppEvents
        if ($ui.Pipelines.Count -eq 0 -and $shared.Events.IsEmpty) {
            if ($null -ne $ui.Operation) { Complete-SiteOperation }
            if ($ui.Task) { Complete-UiTask }
            if ($ui.CloseAfterWork) { $form.Close(); return }
        }
        Update-Chrome
    } catch {
        Send-Log 'ERROR' '' "UI event processing failed. $($_.Exception.Message)"
        $ui.LastSummary = $_.Exception.Message
        if ($null -ne $ui.Operation) { Stop-SiteOperation }
    } finally {
        if (-not $form.IsDisposed -and -not $form.Disposing) { $tickTimer.Start() }
    }
})

$logger = $null
$loggerHandle = $null
try {
    $module = Find-PnPModule
    if ($null -eq $module) {
        if (-not (Confirm-Action "PnP.PowerShell 2.12 or later is required. Install PnP.PowerShell for your Windows account from PowerShell Gallery?")) { return }
        $installForm = [Windows.Forms.Form]::new()
        $installForm.Text = $appTitle
        $installForm.StartPosition = 'CenterScreen'
        $installForm.ClientSize = [Drawing.Size]::new(445, 87)
        $installForm.ControlBox = $false
        $installLabel = [Windows.Forms.Label]::new()
        $installLabel.Dock = 'Fill'
        $installLabel.TextAlign = 'MiddleCenter'
        $installLabel.Text = 'Installing PnP.PowerShell for the current user…'
        $installForm.Controls.Add($installLabel)
        $installTimer = [Windows.Forms.Timer]::new()
        $installer = [PowerShell]::Create()
        try {
            [void] $installer.AddScript({
                $ErrorActionPreference = 'Stop'
                Install-Module -Name PnP.PowerShell -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -Confirm:$false
            }.ToString())
            $installHandle = $installer.BeginInvoke()
            $installTimer.Interval = 200
            $installTimer.Add_Tick({ if ($installHandle.IsCompleted) { $installForm.Close() } })
            $installTimer.Start()
            [void] $installForm.ShowDialog()
            $null = $installer.EndInvoke($installHandle)
            $module = Find-PnPModule
            if ($null -eq $module) { throw 'No compatible PnP.PowerShell module was found after installation.' }
        } catch {
            Show-Message "PnP.PowerShell is unavailable. $($_.Exception.Message)" 'Error'
            return
        } finally {
            $installTimer.Dispose()
            $installer.Dispose()
            $installForm.Dispose()
        }
    }

    $ui.ModulePath = $module.Path
    $ui.ServicePool = New-AppPool 1
    # One extra runspace lets enumeration and site workers run concurrently.
    $ui.Pool = New-AppPool ($Concurrency + 1)
    $logger = [PowerShell]::Create()
    [void] $logger.AddScript($loggerCode.ToString()).AddArgument($shared).AddArgument($config)
    $loggerHandle = $logger.BeginInvoke()
    Send-Log 'INFO' '' "Application started. PnP.PowerShell $($module.Version). Concurrency $Concurrency."
    Update-Chrome
    $tickTimer.Start()
    [void] $form.ShowDialog()
} finally {
    $tickTimer.Stop()
    $filterTimer.Stop()
    if ($null -ne $ui.Operation) { $ui.Operation.Cancel.Cancel() }
    foreach ($job in $ui.Pipelines) {
        try { $null = $job.Instance.EndInvoke($job.Handle) }
        catch { Write-Verbose $_.Exception.Message }
        finally { $job.Instance.Dispose() }
    }
    if ($null -ne $ui.Operation) {
        if ($null -ne $ui.Operation.Recovery) { $ui.Operation.Recovery.Stream.Dispose() }
        $ui.Operation.Queue.Dispose()
        $ui.Operation.Cancel.Dispose()
    }
    foreach ($pool in $ui.Pools) { $pool.Dispose() }
    if ($null -ne $session.Admin) { $session.Admin.Context.Dispose() }
    $session.Tokens.Clear()
    $session.TokenGate.Dispose()
    $shared.LogQueue.CompleteAdding()
    if ($null -ne $logger) {
        try { if ($null -ne $loggerHandle) { $null = $logger.EndInvoke($loggerHandle) } }
        catch { Write-Verbose $_.Exception.Message }
        finally { $logger.Dispose() }
    }
    $shared.LogQueue.Dispose()
    $filterTimer.Dispose()
    $tickTimer.Dispose()
    $form.Dispose()
    $source.Dispose()
    $view.Dispose()
    $table.Dispose()
}
