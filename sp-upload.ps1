<#
.SYNOPSIS
    Bulk-upload a local folder tree to a SharePoint document library.

.DESCRIPTION
    Uses Azure CLI for auth + Microsoft Graph API (no app registration required).
    PowerShell 5.1+ port of sp-upload.sh with identical functionality.

.EXAMPLE
    .\sp-upload.ps1 --source .\my-folder --site-url https://contoso.sharepoint.com/sites/team-x --library "Shared Documents"

.EXAMPLE
    .\sp-upload.ps1 --source .\reports --site-url https://contoso.sharepoint.com/sites/finance --library Documents --remote-path "2026/Q1" --dry-run
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

###############################################################################
# Constants
###############################################################################
$GRAPH_BASE       = 'https://graph.microsoft.com/v1.0'
$SMALL_FILE_LIMIT = 4 * 1024 * 1024          # 4 MB threshold
$TOKEN_TTL        = 2700                      # Refresh token every 45 min
$MIN_CHUNK_SIZE   = 320 * 1024                # 320 KiB - Graph API minimum
$MAX_CHUNK_SIZE   = 60 * 1024 * 1024          # 60 MiB  - Graph API maximum
$MAX_RETRIES      = 5                         # Max retry attempts on throttling
$INITIAL_BACKOFF  = 5                         # Initial retry backoff (seconds)
$MAX_BACKOFF      = 120                       # Max retry backoff (seconds)

###############################################################################
# Global state
###############################################################################
$script:Source          = ''
$script:SiteUrl         = ''
$script:Library         = ''
$script:RemotePath      = ''
$script:ChunkSize       = 10 * 1024 * 1024    # default 10 MiB per chunk
$script:DryRun          = $false
$script:LogFile         = ''
$script:LedgerFile      = ''
$script:AccessToken     = ''
$script:TokenAcquiredAt = 0
$script:SiteId          = ''
$script:DriveId         = ''

###############################################################################
# Helpers
###############################################################################
function Get-Timestamp {
    return (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
}

function Write-Log {
    param([string]$Message)
    $msg = "[$(Get-Timestamp)] $Message"
    Write-Host $msg -ForegroundColor Cyan
    if ($script:LogFile) { $msg | Out-File -FilePath $script:LogFile -Append -Encoding utf8 }
}

function Write-Ok {
    param([string]$Message)
    $msg = "[$(Get-Timestamp)] $([char]0x2713) $Message"
    Write-Host $msg -ForegroundColor Green
    if ($script:LogFile) { $msg | Out-File -FilePath $script:LogFile -Append -Encoding utf8 }
}

function Write-Warn {
    param([string]$Message)
    $msg = "[$(Get-Timestamp)] $([char]0x26A0) $Message"
    Write-Host $msg -ForegroundColor Yellow
    if ($script:LogFile) { $msg | Out-File -FilePath $script:LogFile -Append -Encoding utf8 }
}

function Write-Err {
    param([string]$Message)
    $msg = "[$(Get-Timestamp)] $([char]0x2717) $Message"
    Write-Host $msg -ForegroundColor Red
    if ($script:LogFile) { $msg | Out-File -FilePath $script:LogFile -Append -Encoding utf8 }
}

function Stop-WithError {
    param([string]$Message)
    Write-Err $Message
    exit 1
}

function Show-Usage {
    $scriptName = Split-Path -Leaf $MyInvocation.ScriptName
    if (-not $scriptName) { $scriptName = 'sp-upload.ps1' }
    Write-Host @"
Usage: .\$scriptName [options]

Upload a local folder tree to a SharePoint document library.
Uses Azure CLI for authentication (no app registration required).

Required:
  --source <path>       Local folder to upload
  --site-url <url>      SharePoint site URL
  --library <name>      Target document library (e.g. "Shared Documents")

Optional:
  --remote-path <path>  Sub-path inside the library (default: library root)
  --chunk-size <bytes>  Upload chunk size in bytes (default: 10485760 / 10 MiB)
                        Min: 327680 (320 KiB), Max: 62914560 (60 MiB)
                        Must be a multiple of 327680 (320 KiB)
  --dry-run             Preview operations without executing
  --log <path>          Log file (default: .\sp-upload.log)
  --ledger <path>       Ledger file for resume tracking
                        (default: <source>\.sp-upload-ledger)
  -h, --help            Show this help message

Prerequisites:
  * Azure CLI installed (https://aka.ms/installazurecli)
  * Logged in via: az login
"@
    exit 0
}

function Get-HumanSize {
    param([long]$Bytes)
    if ($Bytes -ge 1073741824) {
        return "$([math]::Floor($Bytes / 1073741824)) GB"
    } elseif ($Bytes -ge 1048576) {
        return "$([math]::Floor($Bytes / 1048576)) MB"
    } elseif ($Bytes -ge 1024) {
        return "$([math]::Floor($Bytes / 1024)) KB"
    } else {
        return "$Bytes B"
    }
}

function Get-HumanTime {
    param([int]$Seconds)
    if ($Seconds -le 0) { return '< 1s' }
    $h = [math]::Floor($Seconds / 3600)
    $m = [math]::Floor(($Seconds % 3600) / 60)
    $s = $Seconds % 60
    $out = ''
    if ($h -gt 0) { $out += "${h}h " }
    if ($m -gt 0) { $out += "${m}m " }
    $out += "${s}s"
    return $out
}

function Get-UnixTime {
    return [long][DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
}

function Get-GraphErrorDetail {
    param([string]$Body)
    try {
        $parsed = $Body | ConvertFrom-Json -ErrorAction SilentlyContinue
        $code = $parsed.error.code
        $msg  = $parsed.error.message
        if ($code -or $msg) {
            $codeStr = if ($code) { $code } else { 'unknown' }
            $msgStr  = if ($msg)  { $msg }  else { '<none>' }
            return "code=$codeStr, message=$msgStr"
        }
    } catch {}
    if ($Body.Length -gt 500) { return $Body.Substring(0, 500) }
    return $Body
}

###############################################################################
# Argument parsing
###############################################################################
function Read-Arguments {
    param([string[]]$RawArgs)

    $i = 0
    while ($i -lt $RawArgs.Count) {
        switch ($RawArgs[$i]) {
            '--source'      { $script:Source     = $RawArgs[++$i]; break }
            '--site-url'    { $script:SiteUrl    = $RawArgs[++$i]; break }
            '--library'     { $script:Library    = $RawArgs[++$i]; break }
            '--remote-path' { $script:RemotePath = $RawArgs[++$i]; break }
            '--chunk-size'  { $script:ChunkSize  = $RawArgs[++$i]; break }
            '--dry-run'     { $script:DryRun     = $true;          break }
            '--log'         { $script:LogFile    = $RawArgs[++$i]; break }
            '--ledger'      { $script:LedgerFile = $RawArgs[++$i]; break }
            { $_ -eq '-h' -or $_ -eq '--help' } { Show-Usage }
            default         { Stop-WithError "Unknown option: $($RawArgs[$i])" }
        }
        $i++
    }

    if (-not $script:Source)  { Stop-WithError 'Missing required option: --source' }
    if (-not $script:SiteUrl) { Stop-WithError 'Missing required option: --site-url' }
    if (-not $script:Library) { Stop-WithError 'Missing required option: --library' }

    if (-not (Test-Path -LiteralPath $script:Source)) {
        Stop-WithError "Source path does not exist: $($script:Source)"
    }
    $script:Source = (Resolve-Path -LiteralPath $script:Source -ErrorAction Stop).Path
    if (-not $script:LogFile)    { $script:LogFile    = '.\sp-upload.log' }
    if (-not $script:LedgerFile) { $script:LedgerFile = Join-Path $script:Source '.sp-upload-ledger' }

    # Normalize remote path: backslashes to forward slashes, trim separators
    if ($script:RemotePath) {
        $script:RemotePath = ($script:RemotePath -replace '\\', '/').Trim('/')
    }

    # Validate chunk size
    if (-not ($script:ChunkSize -match '^\d+$')) {
        Stop-WithError "--chunk-size must be a positive integer (bytes), got: $($script:ChunkSize)"
    }
    $script:ChunkSize = [long]$script:ChunkSize
    if ($script:ChunkSize -lt $MIN_CHUNK_SIZE) {
        Stop-WithError "--chunk-size must be at least $MIN_CHUNK_SIZE bytes (320 KiB), got: $($script:ChunkSize)"
    }
    if ($script:ChunkSize -gt $MAX_CHUNK_SIZE) {
        Stop-WithError "--chunk-size must be at most $MAX_CHUNK_SIZE bytes (60 MiB), got: $($script:ChunkSize)"
    }
    if ($script:ChunkSize % $MIN_CHUNK_SIZE -ne 0) {
        Stop-WithError "--chunk-size must be a multiple of $MIN_CHUNK_SIZE bytes (320 KiB), got: $($script:ChunkSize)"
    }
}

###############################################################################
# Pre-flight checks
###############################################################################
function Test-Preflight {
    if (-not $script:DryRun) {
        $azPath = Get-Command az -ErrorAction SilentlyContinue
        if (-not $azPath) { Stop-WithError 'Azure CLI not found. Install: https://aka.ms/installazurecli' }

        Write-Log 'Checking Azure CLI login status...'
        $null = az account show -o none 2>$null
        if ($LASTEXITCODE -ne 0) {
            Stop-WithError 'Not logged in. Run: az login'
        }
        Write-Ok 'Authenticated via Azure CLI'
    } else {
        Write-Warn 'Dry-run mode - skipping Azure CLI and authentication checks'
    }

    if (-not (Test-Path -LiteralPath $script:Source -PathType Container)) {
        Stop-WithError "Source directory does not exist: $($script:Source)"
    }
    Write-Ok "Source directory: $($script:Source)"
}

###############################################################################
# Token management - refreshes every 45 min (tokens valid ~1 hour)
###############################################################################
function Update-Token {
    $now = Get-UnixTime
    if (($now - $script:TokenAcquiredAt) -lt $TOKEN_TTL -and $script:AccessToken) {
        return
    }

    $script:AccessToken = az account get-access-token `
        --resource https://graph.microsoft.com `
        --query accessToken -o tsv 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $script:AccessToken) {
        Stop-WithError 'Failed to acquire access token for Microsoft Graph'
    }
    $script:TokenAcquiredAt = $now
}

###############################################################################
# URL-encode a path (preserving forward slashes)
###############################################################################
function Get-EncodedPath {
    param([string]$Path)
    $encoded = [System.Text.StringBuilder]::new($Path.Length * 2)
    foreach ($c in $Path.ToCharArray()) {
        if ($c -match '[a-zA-Z0-9.~_/\-]') {
            $null = $encoded.Append($c)
        } elseif ($c -eq ' ') {
            $null = $encoded.Append('%20')
        } else {
            $null = $encoded.Append(('%{0:X2}' -f [int]$c))
        }
    }
    return $encoded.ToString()
}

###############################################################################
# Retry-aware HTTP client - backs off on 429 / 503, then aborts
# Returns: hashtable with Body (string) and StatusCode (int)
###############################################################################
function Invoke-GraphRequest {
    param(
        [string]$Uri,
        [string]$Method = 'GET',
        [hashtable]$Headers = @{},
        [object]$Body = $null,
        [string]$ContentType = $null,
        [string]$InFile = $null
    )

    $attempt  = 0
    $backoff  = $INITIAL_BACKOFF
    $hdrs     = @{} + $Headers
    if (-not $hdrs.ContainsKey('Authorization')) {
        $hdrs['Authorization'] = "Bearer $($script:AccessToken)"
    }

    while ($true) {
        $response     = $null
        $statusCode   = 0
        $responseBody = ''
        $responseHeaders = @{}

        try {
            $params = @{
                Uri             = $Uri
                Method          = $Method
                Headers         = $hdrs
                UseBasicParsing = $true
                ErrorAction     = 'Stop'
            }
            if ($ContentType) { $params['ContentType'] = $ContentType }
            if ($Body)        { $params['Body']        = $Body }
            if ($InFile)      { $params['InFile']      = $InFile }

            $response        = Invoke-WebRequest @params
            $statusCode      = [int]$response.StatusCode
            $responseBody    = $response.Content
            $responseHeaders = $response.Headers
        } catch {
            $ex = $_.Exception
            if ($ex.PSObject.Properties['Response'] -and $ex.Response) {
                $statusCode      = [int]$ex.Response.StatusCode
                $responseHeaders = @{}
                try {
                    foreach ($key in $ex.Response.Headers.AllKeys) {
                        $responseHeaders[$key] = $ex.Response.Headers[$key]
                    }
                } catch {}
                try {
                    $stream = $ex.Response.GetResponseStream()
                    $reader = [System.IO.StreamReader]::new($stream)
                    $responseBody = $reader.ReadToEnd()
                    $reader.Close()
                    $stream.Close()
                } catch {
                    $responseBody = $ex.Message
                }
            } else {
                throw
            }
        }

        switch ($statusCode) {
            401 {
                $attempt++
                if ($attempt -gt $MAX_RETRIES) {
                    $detail = Get-GraphErrorDetail $responseBody
                    Stop-WithError "Authentication failed (HTTP 401) after $MAX_RETRIES retries - aborting. Detail: $detail"
                }
                Write-Warn "Token expired (HTTP 401), refreshing token (attempt $attempt/$MAX_RETRIES)..."
                $script:TokenAcquiredAt = 0
                Update-Token
                $hdrs['Authorization'] = "Bearer $($script:AccessToken)"
            }
            { $_ -eq 429 -or $_ -eq 503 } {
                $attempt++
                if ($attempt -gt $MAX_RETRIES) {
                    $detail = Get-GraphErrorDetail $responseBody
                    Stop-WithError "Persistent throttling (HTTP $statusCode) after $MAX_RETRIES retries - aborting. Detail: $detail. Re-run to resume."
                }

                $waitTime = $backoff
                $retryAfter = $null
                if ($responseHeaders.ContainsKey('Retry-After')) {
                    $retryAfter = $responseHeaders['Retry-After']
                }
                if ($retryAfter -match '^\d+$') {
                    $waitTime = [int]$retryAfter
                }

                Write-Warn "Throttled (HTTP $statusCode), waiting ${waitTime}s before retry $attempt/$MAX_RETRIES..."
                Start-Sleep -Seconds $waitTime

                $backoff = $backoff * 2
                if ($backoff -gt $MAX_BACKOFF) { $backoff = $MAX_BACKOFF }
            }
            default {
                return @{
                    Body       = $responseBody
                    StatusCode = $statusCode
                    Headers    = $responseHeaders
                }
            }
        }
    }
}

###############################################################################
# Resolve SharePoint site ID and drive (library) ID via Graph API
###############################################################################
function Resolve-SiteAndDrive {
    Update-Token

    # Parse site URL -> hostname + site path
    $url = $script:SiteUrl -replace '^https?://', ''
    $url = $url.TrimEnd('/')
    $slashIdx = $url.IndexOf('/')
    if ($slashIdx -gt 0) {
        $hostname  = $url.Substring(0, $slashIdx)
        $sitePath  = '/' + $url.Substring($slashIdx + 1)
    } else {
        $hostname  = $url
        $sitePath  = ''
    }

    # Resolve site
    Write-Log 'Resolving SharePoint site...'
    if ($sitePath) {
        $endpoint = "/sites/${hostname}:${sitePath}"
    } else {
        $endpoint = "/sites/${hostname}"
    }

    $result = Invoke-GraphRequest -Uri "${GRAPH_BASE}${endpoint}"
    try {
        $siteObj = $result.Body | ConvertFrom-Json
        $script:SiteId = $siteObj.id
    } catch {
        $script:SiteId = ''
    }
    if (-not $script:SiteId) {
        $detail = Get-GraphErrorDetail $result.Body
        Stop-WithError "Could not resolve site (HTTP $($result.StatusCode)): $($script:SiteUrl) - $detail"
    }
    Write-Ok "Site ID: $($script:SiteId)"

    # Resolve drive (library)
    Write-Log "Resolving document library '$($script:Library)'..."
    $result = Invoke-GraphRequest -Uri "${GRAPH_BASE}/sites/$($script:SiteId)/drives?`$select=id,name"
    $drive = $null
    try {
        $drivesObj = $result.Body | ConvertFrom-Json
        $drive = $drivesObj.value | Where-Object { $_.name -eq $script:Library } | Select-Object -First 1
    } catch {
        $drivesObj = $null
    }
    if (-not $drive) {
        if ($result.StatusCode -ne 200) {
            $detail = Get-GraphErrorDetail $result.Body
            Stop-WithError "Failed to list drives (HTTP $($result.StatusCode)): $detail"
        }
        $available = ''
        if ($drivesObj -and $drivesObj.value) {
            $available = ($drivesObj.value | ForEach-Object { $_.name }) -join ', '
        }
        if (-not $available) { $available = 'none' }
        Stop-WithError "Library '$($script:Library)' not found. Available: $available"
    }
    $script:DriveId = $drive.id
    Write-Ok "Drive ID: $($script:DriveId)"
}

###############################################################################
# Create a remote folder via Graph API
###############################################################################
function New-RemoteFolder {
    param([string]$FolderPath)

    $parentPath = Split-Path -Parent $FolderPath
    $leaf       = Split-Path -Leaf   $FolderPath
    # Normalize parent path separators to forward slashes
    if ($parentPath) { $parentPath = $parentPath -replace '\\', '/' }

    if (-not $parentPath -or $parentPath -eq '.') {
        $endpoint = "/drives/$($script:DriveId)/root/children"
    } else {
        $encoded  = Get-EncodedPath $parentPath
        $endpoint = "/drives/$($script:DriveId)/root:/${encoded}:/children"
    }

    $bodyObj = @{
        name                                = $leaf
        folder                              = @{}
        '@microsoft.graph.conflictBehavior' = 'fail'
    }
    $bodyJson = $bodyObj | ConvertTo-Json -Compress

    $result = Invoke-GraphRequest -Uri "${GRAPH_BASE}${endpoint}" `
        -Method 'POST' -ContentType 'application/json' -Body $bodyJson

    switch ($result.StatusCode) {
        { $_ -eq 200 -or $_ -eq 201 } { return $true }
        409 { return $true }  # Already exists
        default {
            $detail = Get-GraphErrorDetail $result.Body
            Write-Err "Failed to create folder (HTTP $($result.StatusCode)): $FolderPath - $detail"
            return $false
        }
    }
}

###############################################################################
# Upload - small files (< 4 MB): simple PUT
###############################################################################
function Send-SmallFile {
    param([string]$LocalPath, [string]$RemotePath)

    $encoded = Get-EncodedPath $RemotePath

    $result = Invoke-GraphRequest `
        -Uri "${GRAPH_BASE}/drives/$($script:DriveId)/root:/${encoded}:/content" `
        -Method 'PUT' `
        -ContentType 'application/octet-stream' `
        -InFile $LocalPath

    switch ($result.StatusCode) {
        { $_ -eq 200 -or $_ -eq 201 } { return $true }
        default {
            $detail = Get-GraphErrorDetail $result.Body
            Write-Err "Upload failed (HTTP $($result.StatusCode)): $RemotePath - $detail"
            return $false
        }
    }
}

###############################################################################
# Upload session helpers
###############################################################################
function Invoke-UploadSessionRequest {
    param(
        [string]$Uri,
        [string]$Method = 'PUT',
        [hashtable]$Headers = @{},
        [object]$Body = $null,
        [string]$ContentType = $null
    )

    $statusCode      = 0
    $responseBody    = ''
    $responseHeaders = @{}
    $networkError    = $false

    try {
        $params = @{
            Uri             = $Uri
            Method          = $Method
            Headers         = $Headers
            UseBasicParsing = $true
            ErrorAction     = 'Stop'
        }
        if ($ContentType) { $params['ContentType'] = $ContentType }
        if ($PSBoundParameters.ContainsKey('Body')) { $params['Body'] = $Body }

        $response        = Invoke-WebRequest @params
        $statusCode      = [int]$response.StatusCode
        $responseBody    = $response.Content
        $responseHeaders = $response.Headers
    } catch {
        $ex = $_.Exception
        if ($ex.PSObject.Properties['Response'] -and $ex.Response) {
            $statusCode = [int]$ex.Response.StatusCode
            try {
                foreach ($key in $ex.Response.Headers.AllKeys) {
                    $responseHeaders[$key] = $ex.Response.Headers[$key]
                }
            } catch {}
            try {
                $stream = $ex.Response.GetResponseStream()
                $reader = [System.IO.StreamReader]::new($stream)
                $responseBody = $reader.ReadToEnd()
                $reader.Close()
                $stream.Close()
            } catch {
                $responseBody = $ex.Message
            }
        } else {
            $networkError = $true
            $responseBody = $ex.Message
        }
    }

    return @{
        Body           = $responseBody
        StatusCode     = $statusCode
        Headers        = $responseHeaders
        IsNetworkError = $networkError
    }
}

function Get-UploadSessionNextOffset {
    param([string]$Body)

    if (-not $Body) { return $null }

    try {
        $parsed = $Body | ConvertFrom-Json -ErrorAction Stop
        $range = $parsed.nextExpectedRanges | Select-Object -First 1
        if ($range -and $range -match '^(\d+)-') {
            return [long]$Matches[1]
        }
    } catch {}

    return $null
}

function Get-UploadSessionStatus {
    param([string]$UploadUrl)

    $result = Invoke-UploadSessionRequest -Uri $UploadUrl -Method 'GET'
    $nextOffset = $null
    if ($result.StatusCode -eq 200) {
        $nextOffset = Get-UploadSessionNextOffset -Body $result.Body
    }

    return @{
        Body           = $result.Body
        StatusCode     = $result.StatusCode
        Headers        = $result.Headers
        IsNetworkError = $result.IsNetworkError
        NextOffset     = $nextOffset
    }
}

function Remove-UploadSession {
    param([string]$UploadUrl)

    if (-not $UploadUrl) { return }
    $null = Invoke-UploadSessionRequest -Uri $UploadUrl -Method 'DELETE'
}

###############################################################################
# Upload - large files (>= 4 MB): chunked upload session
###############################################################################
function Send-LargeFile {
    param([string]$LocalPath, [string]$RemotePath)

    $fileSize = (Get-Item -LiteralPath $LocalPath).Length
    $encoded  = Get-EncodedPath $RemotePath

    # Create upload session (with retries + exponential backoff)
    $uploadUrl      = ''
    $sessionAttempt = 0
    $sessionBackoff = $INITIAL_BACKOFF

    while ($true) {
        $result = Invoke-GraphRequest `
            -Uri "${GRAPH_BASE}/drives/$($script:DriveId)/root:/${encoded}:/createUploadSession" `
            -Method 'POST' `
            -ContentType 'application/json' `
            -Body '{"item":{"@microsoft.graph.conflictBehavior":"replace"}}'

        $uploadUrl = ''
        try {
            $sessionObj = $result.Body | ConvertFrom-Json
            $uploadUrl  = $sessionObj.uploadUrl
        } catch {}
        if ($uploadUrl) { break }

        $sessionAttempt++
        if ($sessionAttempt -gt $MAX_RETRIES) {
            $detail = Get-GraphErrorDetail $result.Body
            Write-Err "Failed to create upload session (HTTP $($result.StatusCode)) after $MAX_RETRIES retries: $RemotePath - $detail"
            return $false
        }

        $detail = Get-GraphErrorDetail $result.Body
        Write-Warn "Upload session creation failed (HTTP $($result.StatusCode), attempt $sessionAttempt/$MAX_RETRIES): $RemotePath - $detail"
        Write-Warn "Retrying in ${sessionBackoff}s..."
        Start-Sleep -Seconds $sessionBackoff
        $sessionBackoff = $sessionBackoff * 2
        if ($sessionBackoff -gt $MAX_BACKOFF) { $sessionBackoff = $MAX_BACKOFF }
        Update-Token
    }

    # Upload in chunks
    $offset      = [long]0
    $chunkIdx    = 0
    $totalChunks = [math]::Ceiling(([double]$fileSize) / ([double]$script:ChunkSize))
    $fs          = $null

    try {
        $fs = [System.IO.FileStream]::new($LocalPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)

        while ($offset -lt $fileSize) {
            $remaining = $fileSize - $offset
            if ($remaining -lt $script:ChunkSize) {
                $thisChunk = [long]$remaining
            } else {
                $thisChunk = [long]$script:ChunkSize
            }
            $end = $offset + $thisChunk - 1
            $chunkIdx++

            Write-Log "  Chunk ${chunkIdx}/${totalChunks}: bytes ${offset}-${end}/${fileSize}"

            # Read chunk into byte array
            $buffer = [byte[]]::new([int]$thisChunk)
            $null = $fs.Seek($offset, [System.IO.SeekOrigin]::Begin)
            $bytesRead = 0
            while ($bytesRead -lt $thisChunk) {
                $n = $fs.Read($buffer, $bytesRead, [int]($thisChunk - $bytesRead))
                if ($n -eq 0) { break }
                $bytesRead += $n
            }
            if ($bytesRead -ne $thisChunk) {
                Remove-UploadSession $uploadUrl
                Write-Err "Failed to read chunk data at offset $offset of $RemotePath"
                return $false
            }

            $chunkAttempt  = 0
            $chunkBackoff  = $INITIAL_BACKOFF
            $chunkComplete = $false

            while (-not $chunkComplete) {
                $chunkHeaders = @{
                    'Content-Length' = [string]$thisChunk
                    'Content-Range'  = "bytes ${offset}-${end}/${fileSize}"
                }
                $chunkResult = Invoke-UploadSessionRequest -Uri $uploadUrl -Method 'PUT' `
                    -Headers $chunkHeaders -ContentType 'application/octet-stream' -Body $buffer

                if ($chunkResult.IsNetworkError) {
                    $status = Get-UploadSessionStatus $uploadUrl
                    if ($null -ne $status.NextOffset -and $status.NextOffset -gt $offset) {
                        Write-Warn "Recovered upload session state after interrupted chunk; resuming from byte $($status.NextOffset)"
                        $offset = [long]$status.NextOffset
                        $chunkComplete = $true
                        continue
                    }

                    $chunkAttempt++
                    if ($chunkAttempt -gt $MAX_RETRIES) {
                        Remove-UploadSession $uploadUrl
                        Write-Err "Chunk upload failed (network error) after $MAX_RETRIES retries at offset $offset of $RemotePath"
                        return $false
                    }

                    Write-Warn "Chunk upload interrupted, waiting ${chunkBackoff}s before retry $chunkAttempt/$MAX_RETRIES..."
                    Start-Sleep -Seconds $chunkBackoff
                    $chunkBackoff = $chunkBackoff * 2
                    if ($chunkBackoff -gt $MAX_BACKOFF) { $chunkBackoff = $MAX_BACKOFF }
                    continue
                }

                switch ($chunkResult.StatusCode) {
                    { $_ -eq 200 -or $_ -eq 201 } {
                        $offset = $end + 1
                        $chunkComplete = $true
                        break
                    }
                    202 {
                        $nextOffset = Get-UploadSessionNextOffset -Body $chunkResult.Body
                        if ($null -eq $nextOffset) {
                            $offset = $end + 1
                        } elseif ($nextOffset -le $offset) {
                            $status = Get-UploadSessionStatus $uploadUrl
                            if ($null -ne $status.NextOffset -and $status.NextOffset -gt $offset) {
                                $offset = [long]$status.NextOffset
                            } else {
                                Remove-UploadSession $uploadUrl
                                Write-Err "Upload session returned an invalid nextExpectedRanges value at offset $offset of $RemotePath"
                                return $false
                            }
                        } else {
                            $offset = [long]$nextOffset
                        }
                        $chunkComplete = $true
                        break
                    }
                    416 {
                        $status = Get-UploadSessionStatus $uploadUrl
                        if ($null -ne $status.NextOffset -and $status.NextOffset -gt $offset) {
                            Write-Warn "Upload session already advanced to byte $($status.NextOffset); resynchronizing chunk upload state"
                            $offset = [long]$status.NextOffset
                            $chunkComplete = $true
                            break
                        }

                        $detail = Get-GraphErrorDetail $chunkResult.Body
                        Remove-UploadSession $uploadUrl
                        Write-Err "Chunk range rejected (HTTP 416) at offset $offset of $RemotePath - $detail"
                        return $false
                    }
                    404 {
                        Write-Err "Upload session expired or was not found while uploading $RemotePath"
                        return $false
                    }
                    { $_ -eq 429 -or $_ -eq 500 -or $_ -eq 502 -or $_ -eq 503 -or $_ -eq 504 } {
                        $status = Get-UploadSessionStatus $uploadUrl
                        if ($null -ne $status.NextOffset -and $status.NextOffset -gt $offset) {
                            Write-Warn "Upload session already advanced to byte $($status.NextOffset); resuming from reported server offset"
                            $offset = [long]$status.NextOffset
                            $chunkComplete = $true
                            break
                        }

                        $chunkAttempt++
                        if ($chunkAttempt -gt $MAX_RETRIES) {
                            $detail = Get-GraphErrorDetail $chunkResult.Body
                            Remove-UploadSession $uploadUrl
                            Stop-WithError "Chunk upload failed persistently (HTTP $($chunkResult.StatusCode)) after $MAX_RETRIES retries at offset $offset of $RemotePath - aborting. Detail: $detail"
                        }

                        $waitTime = $chunkBackoff
                        $retryAfter = $null
                        if ($chunkResult.Headers.ContainsKey('Retry-After')) {
                            $retryAfter = $chunkResult.Headers['Retry-After']
                        }
                        if ($retryAfter -match '^\d+$') {
                            $waitTime = [int]$retryAfter
                        }

                        Write-Warn "Chunk upload transient failure (HTTP $($chunkResult.StatusCode)), waiting ${waitTime}s before retry $chunkAttempt/$MAX_RETRIES..."
                        Start-Sleep -Seconds $waitTime
                        $chunkBackoff = $chunkBackoff * 2
                        if ($chunkBackoff -gt $MAX_BACKOFF) { $chunkBackoff = $MAX_BACKOFF }
                        continue
                    }
                    default {
                        $detail = Get-GraphErrorDetail $chunkResult.Body
                        Remove-UploadSession $uploadUrl
                        Write-Err "Chunk upload failed (HTTP $($chunkResult.StatusCode)) at offset $offset of $RemotePath - $detail"
                        return $false
                    }
                }
            }
        }
    } finally {
        if ($fs) { $fs.Close() }
    }

    return $true
}

###############################################################################
# Ledger helpers - path-based resume tracking
###############################################################################
function Test-LedgerEntry {
    param([string]$RelPath)
    if (-not (Test-Path -LiteralPath $script:LedgerFile)) { return $false }
    $content = Get-Content -LiteralPath $script:LedgerFile -ErrorAction SilentlyContinue
    return ($content -contains $RelPath)
}

function Add-LedgerEntry {
    param([string]$RelPath)
    $RelPath | Out-File -FilePath $script:LedgerFile -Append -Encoding utf8
}

###############################################################################
# Progress display
###############################################################################
function Write-Progress-Info {
    param(
        [int]$FilesDone, [int]$FilesTotal,
        [long]$BytesDone, [long]$BytesTotal,
        [long]$UploadBytes, [long]$UploadSecs
    )

    $filesLeft = $FilesTotal - $FilesDone
    $bytesLeft = $BytesTotal - $BytesDone

    Write-Log "  Progress : ${FilesDone}/${FilesTotal} files, $(Get-HumanSize $BytesDone)/$(Get-HumanSize $BytesTotal)"
    Write-Log "  Remaining: ${filesLeft} files, $(Get-HumanSize $bytesLeft)"

    if ($UploadSecs -gt 0 -and $UploadBytes -gt 0) {
        $throughput = [long]($UploadBytes / $UploadSecs)
        if ($throughput -gt 0 -and $bytesLeft -gt 0) {
            $etaSecs = [int]($bytesLeft / $throughput)
            Write-Log "  ETA      : ~$(Get-HumanTime $etaSecs)"
        } elseif ($bytesLeft -eq 0) {
            Write-Log '  ETA      : done'
        }
    } else {
        if ($bytesLeft -gt 0) {
            Write-Log '  ETA      : estimating...'
        } else {
            Write-Log '  ETA      : done'
        }
    }
}

###############################################################################
# Upload workflow
###############################################################################
function Start-UploadAll {
    $uploaded    = 0
    $skipped     = 0
    $failed      = 0
    $dirsCreated = 0
    $basePath    = $script:RemotePath

    # Create --remote-path directories first
    if ($basePath -and -not $script:DryRun) {
        Update-Token
        $parts       = $basePath -split '/'
        $cumulative  = ''
        foreach ($part in $parts) {
            if (-not $part) { continue }
            if (-not $cumulative) { $cumulative = $part } else { $cumulative = "$cumulative/$part" }
            Write-Log "Ensuring base folder: $cumulative"
            if (New-RemoteFolder $cumulative) { $dirsCreated++ }
        }
    } elseif ($basePath -and $script:DryRun) {
        $parts      = $basePath -split '/'
        $cumulative = ''
        foreach ($part in $parts) {
            if (-not $part) { continue }
            if (-not $cumulative) { $cumulative = $part } else { $cumulative = "$cumulative/$part" }
            Write-Log "[DRY-RUN] Would create base folder: $cumulative"
        }
    }

    # Collect subdirectories from source, sorted shallowest-first
    $allDirs = Get-ChildItem -LiteralPath $script:Source -Recurse -Directory |
        Sort-Object { $_.FullName }
    $relDirs = @()
    foreach ($d in $allDirs) {
        $rel = $d.FullName.Substring($script:Source.Length).TrimStart('\', '/')
        if ($rel) { $relDirs += $rel }
    }

    # Create remote subdirectories
    foreach ($relDir in $relDirs) {
        # Normalize to forward slashes for Graph API
        $relDirNorm = $relDir -replace '\\', '/'
        if ($basePath) {
            $remoteFolder = "$basePath/$relDirNorm"
        } else {
            $remoteFolder = $relDirNorm
        }

        if ($script:DryRun) {
            Write-Log "[DRY-RUN] Would create folder: $remoteFolder"
        } else {
            Update-Token
            Write-Log "Creating folder: $remoteFolder"
            if (New-RemoteFolder $remoteFolder) {
                Write-Ok "Folder ready: $remoteFolder"
                $dirsCreated++
            } else {
                Write-Warn "Folder may already exist: $remoteFolder"
            }
        }
    }

    # Pre-scan: compute totals
    $allFiles = Get-ChildItem -LiteralPath $script:Source -Recurse -File |
        Where-Object { $_.Name -ne '.sp-upload-ledger' } |
        Sort-Object { $_.FullName }
    $total        = $allFiles.Count
    $totalBytes   = [long]0
    $largestBytes = [long]0
    $largestName  = ''
    foreach ($f in $allFiles) {
        $totalBytes += $f.Length
        if ($f.Length -gt $largestBytes) {
            $largestBytes = $f.Length
            $largestName  = $f.FullName.Substring($script:Source.Length).TrimStart('\', '/') -replace '\\', '/'
        }
    }

    Write-Host ''
    Write-Log ([string][char]0x2500 * 39)
    Write-Log "Files to process  : $total"
    Write-Log "Total size        : $(Get-HumanSize $totalBytes)"
    if ($total -gt 0) {
        Write-Log "Largest file      : $(Get-HumanSize $largestBytes) ($largestName)"
    }
    Write-Log ([string][char]0x2500 * 39)
    Write-Host ''

    # Upload files with progress tracking
    $bytesProcessed     = [long]0
    $filesProcessed     = 0
    $uploadBytesDone    = [long]0
    $uploadTimeElapsed  = [long]0

    foreach ($file in $allFiles) {
        $fsize = $file.Length
        $rel   = $file.FullName.Substring($script:Source.Length).TrimStart('\', '/') -replace '\\', '/'

        if (Test-LedgerEntry $rel) {
            $skipped++
            Write-Log "Skipping (already uploaded): $rel"
            $filesProcessed++
            $bytesProcessed += $fsize
            Write-Progress-Info $filesProcessed $total $bytesProcessed $totalBytes $uploadBytesDone $uploadTimeElapsed
            continue
        }

        if ($basePath) {
            $remoteFilePath = "$basePath/$rel"
        } else {
            $remoteFilePath = $rel
        }

        if ($script:DryRun) {
            $method = 'simple'
            if ($fsize -ge $SMALL_FILE_LIMIT) { $method = 'chunked' }
            Write-Log "[DRY-RUN] Would upload ($method, $(Get-HumanSize $fsize)): $rel"
            $filesProcessed++
            $bytesProcessed += $fsize
            continue
        }

        Update-Token
        Write-Log "Uploading: $rel ($(Get-HumanSize $fsize))"

        $fileStart = Get-UnixTime
        $uploadOk  = $false

        if ($fsize -lt $SMALL_FILE_LIMIT) {
            $uploadOk = Send-SmallFile $file.FullName $remoteFilePath
        } else {
            $uploadOk = Send-LargeFile $file.FullName $remoteFilePath
        }

        $fileEnd = Get-UnixTime

        if ($uploadOk) {
            Write-Ok "Uploaded: $rel"
            Add-LedgerEntry $rel
            $uploaded++
            $uploadBytesDone   += $fsize
            $uploadTimeElapsed += ($fileEnd - $fileStart)
        } else {
            Write-Err "Failed: $rel"
            $failed++
        }

        $filesProcessed++
        $bytesProcessed += $fsize
        Write-Progress-Info $filesProcessed $total $bytesProcessed $totalBytes $uploadBytesDone $uploadTimeElapsed
    }

    # Summary
    Write-Host ''
    Write-Log ([string][char]0x2550 * 39)
    Write-Log 'Upload complete'
    Write-Log "  Directories created : $dirsCreated"
    Write-Log "  Files total         : $total"
    Write-Log "  Uploaded            : $uploaded"
    Write-Log "  Skipped (resuming)  : $skipped"
    Write-Log "  Failed              : $failed"
    Write-Log ([string][char]0x2550 * 39)

    if ($script:DryRun) {
        Write-Warn 'Dry-run mode - no changes were made.'
    }

    if ($failed -gt 0) { return $false }
    return $true
}

###############################################################################
# Main
###############################################################################
function Main {
    Read-Arguments $args

    Test-Preflight

    if (-not $script:DryRun) {
        Resolve-SiteAndDrive
    }

    Write-Log ([string][char]0x2550 * 39)
    Write-Log 'SharePoint Bulk Upload'
    Write-Log "  Source      : $($script:Source)"
    Write-Log "  Site URL    : $($script:SiteUrl)"
    Write-Log "  Library     : $($script:Library)"
    $rp = if ($script:RemotePath) { $script:RemotePath } else { '(root)' }
    Write-Log "  Remote path : $rp"
    Write-Log "  Ledger      : $($script:LedgerFile)"
    Write-Log "  Chunk size  : $(Get-HumanSize $script:ChunkSize)"
    Write-Log "  Dry-run     : $($script:DryRun)"
    if ($script:DriveId) {
        Write-Log "  Drive ID    : $($script:DriveId)"
    }
    Write-Log ([string][char]0x2550 * 39)
    Write-Host ''

    $success = Start-UploadAll
    if (-not $success) { exit 1 }
}

# Invoke main, passing through script arguments
Main @args