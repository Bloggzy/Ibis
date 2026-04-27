function Get-IbisConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $configPath = Join-Path $ProjectRoot 'config.json'
    if (-not (Test-Path -LiteralPath $configPath)) {
        throw "Ibis config was not found at: $configPath"
    }

    Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
}

function Save-IbisConfigPathSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [string]$ToolsRoot,

        [string]$SourceRoot,

        [string]$OutputRoot,

        [bool]$CompletionBeepEnabled
    )

    $configPath = Join-Path $ProjectRoot 'config.json'
    if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
        throw "Ibis config was not found at: $configPath"
    }

    $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    if ($PSBoundParameters.ContainsKey('ToolsRoot')) {
        $config.defaultToolsRoot = $ToolsRoot
    }
    if ($PSBoundParameters.ContainsKey('SourceRoot')) {
        $config.defaultSourceRoot = $SourceRoot
    }
    if ($PSBoundParameters.ContainsKey('OutputRoot')) {
        $config.defaultOutputRoot = $OutputRoot
    }
    if ($PSBoundParameters.ContainsKey('CompletionBeepEnabled')) {
        if ($null -eq $config.PSObject.Properties['completionBeepEnabled']) {
            $config | Add-Member -NotePropertyName 'completionBeepEnabled' -NotePropertyValue $CompletionBeepEnabled
        }
        else {
            $config.completionBeepEnabled = $CompletionBeepEnabled
        }
    }

    $config | ConvertTo-Json -Depth 20 | Out-File -LiteralPath $configPath -Encoding UTF8
    $config
}

function Get-IbisToolDefinition {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot,

        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    $toolDefinitionRoot = Join-Path $ProjectRoot $Config.toolDefinitionPath
    if (-not (Test-Path -LiteralPath $toolDefinitionRoot)) {
        return @()
    }

    $definitions = @()
    $files = Get-ChildItem -LiteralPath $toolDefinitionRoot -Filter '*.json' -File | Sort-Object Name
    foreach ($file in $files) {
        try {
            $definition = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
            $definition | Add-Member -NotePropertyName DefinitionPath -NotePropertyValue $file.FullName -Force
            $definitions += $definition
        }
        catch {
            $definitions += [pscustomobject]@{
                id = $file.BaseName
                name = $file.BaseName
                executablePath = ''
                downloadUrl = ''
                manualUrl = ''
                notes = "Definition could not be parsed: $($_.Exception.Message)"
                DefinitionPath = $file.FullName
                DefinitionError = $_.Exception.Message
            }
        }
    }

    $definitions
}

function Test-IbisToolStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object[]]$ToolDefinitions
    )

    foreach ($tool in $ToolDefinitions) {
        if ($tool.executablePath) {
            $installState = Test-IbisToolInstallState -ToolsRoot $ToolsRoot -ToolDefinition $tool
            $expectedPath = $installState.ExpectedPath
            $present = $installState.Present
            $status = $installState.Status
        }
        elseif ($tool.DefinitionError) {
            $expectedPath = $null
            $present = $false
            $status = 'Definition Error'
        }
        else {
            $expectedPath = $null
            $present = $false
            $status = 'Missing'
        }

        [pscustomobject]@{
            Id = $tool.id
            Name = $tool.name
            Status = $status
            Present = $present
            ExpectedPath = $expectedPath
            DownloadUrl = $tool.downloadUrl
            ManualUrl = $tool.manualUrl
            Notes = $tool.notes
        }
    }
}

function Get-IbisToolAcquisitionPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ToolStatuses
    )

    foreach ($status in $ToolStatuses) {
        if (-not $status.Present) {
            $source = $status.DownloadUrl
            if ([string]::IsNullOrWhiteSpace($source) -or $source -eq 'latest-release') {
                $source = $status.ManualUrl
            }

            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                ExpectedPath = $status.ExpectedPath
                AcquisitionSource = $source
                Notes = $status.Notes
            }
        }
    }
}

function Format-IbisToolAcquisitionPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$AcquisitionPlan
    )

    if ($AcquisitionPlan.Count -eq 0) {
        return 'All configured tools are present.'
    }

    $lines = @()
    $lines += 'Ibis tool guidance'
    $lines += '=================='
    $lines += ''
    foreach ($item in $AcquisitionPlan) {
        $lines += $item.Name
        $lines += ('  Expected path: {0}' -f $item.ExpectedPath)
        $lines += ('  Get from:      {0}' -f $item.AcquisitionSource)
        if (-not [string]::IsNullOrWhiteSpace($item.Notes)) {
            $lines += ('  Notes:         {0}' -f $item.Notes)
        }
        $lines += ''
    }

    $lines -join [Environment]::NewLine
}

function Write-IbisProgressEvent {
    [CmdletBinding()]
    param(
        [string]$ProgressPath,

        [string]$ToolId,

        [string]$ToolName,

        [string]$Stage,

        [string]$Message,

        [int]$Index = 0,

        [int]$Total = 0,

        [string]$Status = 'Info'
    )

    if ([string]::IsNullOrWhiteSpace($ProgressPath)) {
        return
    }

    $directory = Split-Path -Path $ProgressPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $progressEvent = [pscustomobject]@{
        Time = (Get-Date).ToString('s')
        ToolId = $ToolId
        ToolName = $ToolName
        Stage = $Stage
        Message = $Message
        Index = $Index
        Total = $Total
        Status = $Status
    }

    $line = ($progressEvent | ConvertTo-Json -Compress -Depth 4) + [Environment]::NewLine
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $stream = $null
        $writer = $null
        try {
            $stream = [System.IO.FileStream]::new(
                $ProgressPath,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite
            )
            $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
            $writer.Write($line)
            return
        }
        catch {
            if ($attempt -eq 10) {
                throw
            }
            Start-Sleep -Milliseconds (25 * $attempt)
        }
        finally {
            if ($writer) {
                $writer.Dispose()
            }
            elseif ($stream) {
                $stream.Dispose()
            }
        }
    }
}

function Get-IbisToolDefinitionById {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$Id
    )

    $ToolDefinitions | Where-Object { $_.id -eq $Id } | Select-Object -First 1
}

function Resolve-IbisToolDownloadUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition
    )

    if ([string]::IsNullOrWhiteSpace($ToolDefinition.downloadUrl)) {
        throw "No download URL is configured for $($ToolDefinition.name)."
    }

    if ($ToolDefinition.downloadUrl -ne 'latest-release') {
        return $ToolDefinition.downloadUrl
    }

    if ([string]::IsNullOrWhiteSpace($ToolDefinition.githubRepo)) {
        throw "$($ToolDefinition.name) uses latest-release but has no githubRepo metadata."
    }

    if ([string]::IsNullOrWhiteSpace($ToolDefinition.assetPattern)) {
        throw "$($ToolDefinition.name) uses latest-release but has no assetPattern metadata."
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $apiUrl = 'https://api.github.com/repos/{0}/releases/latest' -f $ToolDefinition.githubRepo
    $release = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
    $asset = $release.assets | Where-Object { $_.name -like $ToolDefinition.assetPattern } | Select-Object -First 1
    if ($null -eq $asset) {
        throw "No GitHub release asset matching '$($ToolDefinition.assetPattern)' was found for $($ToolDefinition.githubRepo)."
    }

    $asset.browser_download_url
}

function Get-IbisToolInstallDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition
    )

    $relativeDirectory = $ToolDefinition.installDirectory
    if ([string]::IsNullOrWhiteSpace($relativeDirectory)) {
        $relativeDirectory = Split-Path -Path $ToolDefinition.executablePath -Parent
    }

    if ([string]::IsNullOrWhiteSpace($relativeDirectory) -or $relativeDirectory -eq '.') {
        return $ToolsRoot
    }

    Join-Path $ToolsRoot $relativeDirectory
}

function Get-IbisToolExpectedPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition
    )

    Join-Path $ToolsRoot $ToolDefinition.executablePath
}

function Test-IbisIsAdministrator {
    [CmdletBinding()]
    param()

    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-IbisDefenderExclusionRecommendation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object[]]$ToolDefinitions
    )

    foreach ($tool in $ToolDefinitions) {
        if ($tool.defenderExclusionRecommended -eq $true) {
            [pscustomobject]@{
                Id = $tool.id
                Name = $tool.name
                Path = (Get-IbisToolInstallDirectory -ToolsRoot $ToolsRoot -ToolDefinition $tool)
                Reason = $tool.defenderExclusionReason
            }
        }
    }
}

function Get-IbisDefenderExclusionStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object[]]$ToolDefinitions
    )

    $recommendations = @(Get-IbisDefenderExclusionRecommendation -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions)
    $isAdministrator = Test-IbisIsAdministrator
    $mpPreferenceCommand = Get-Command Get-MpPreference -ErrorAction SilentlyContinue
    if ($null -eq $mpPreferenceCommand) {
        foreach ($recommendation in $recommendations) {
            [pscustomobject]@{
                Id = $recommendation.Id
                Name = $recommendation.Name
                Path = $recommendation.Path
                Present = $false
                Status = 'Unavailable'
                Message = 'Get-MpPreference is not available on this system.'
                Reason = $recommendation.Reason
                IsAdministrator = $isAdministrator
                IsAuthoritative = $false
            }
        }
        return
    }

    try {
        $preferences = Get-MpPreference -ErrorAction Stop
        $existingExclusions = @($preferences.ExclusionPath)
        foreach ($recommendation in $recommendations) {
            $resolvedRecommendation = Resolve-IbisComparablePath -Path $recommendation.Path
            $present = $false
            foreach ($existingExclusion in $existingExclusions) {
                if ((Resolve-IbisComparablePath -Path $existingExclusion) -eq $resolvedRecommendation) {
                    $present = $true
                    break
                }
            }

            [pscustomobject]@{
                Id = $recommendation.Id
                Name = $recommendation.Name
                Path = $recommendation.Path
                Present = $present
                Status = $(if ($present) { 'Present' } else { 'Missing' })
                Message = $(if ($present) { 'Defender exclusion is present.' } elseif ($isAdministrator) { 'Defender exclusion is missing.' } else { 'Defender exclusion was not visible from a standard-user session. Run as Administrator for an authoritative check.' })
                Reason = $recommendation.Reason
                IsAdministrator = $isAdministrator
                IsAuthoritative = $isAdministrator
            }
        }
    }
    catch {
        $message = $_.Exception.Message
        if (-not $isAdministrator) {
            $message = "$message Run as Administrator for an authoritative Defender exclusion check."
        }

        foreach ($recommendation in $recommendations) {
            [pscustomobject]@{
                Id = $recommendation.Id
                Name = $recommendation.Name
                Path = $recommendation.Path
                Present = $false
                Status = 'Failed'
                Message = $message
                Reason = $recommendation.Reason
                IsAdministrator = $isAdministrator
                IsAuthoritative = $false
            }
        }
    }
}

function Add-IbisDefenderExclusion {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object[]]$ToolDefinitions
    )

    $statuses = @(Get-IbisDefenderExclusionStatus -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions)
    if ($statuses.Count -eq 0) {
        return
    }

    $addPreferenceCommand = Get-Command Add-MpPreference -ErrorAction SilentlyContinue
    if ($null -eq $addPreferenceCommand) {
        foreach ($status in $statuses) {
            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                Path = $status.Path
                Status = 'Unavailable'
                Message = 'Add-MpPreference is not available on this system.'
                IsAdministrator = (Test-IbisIsAdministrator)
            }
        }
        return
    }

    foreach ($status in $statuses) {
        if ($status.Present) {
            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                Path = $status.Path
                Status = 'Present'
                Message = 'Defender exclusion already exists.'
                IsAdministrator = (Test-IbisIsAdministrator)
            }
            continue
        }

        try {
            if (-not (Test-Path -LiteralPath $status.Path)) {
                New-Item -ItemType Directory -Path $status.Path -Force | Out-Null
            }

            if ($PSCmdlet.ShouldProcess($status.Path, 'Add Windows Defender exclusion')) {
                Add-MpPreference -ExclusionPath $status.Path -ErrorAction Stop
            }

            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                Path = $status.Path
                Status = 'Added'
                Message = 'Defender exclusion added.'
                IsAdministrator = (Test-IbisIsAdministrator)
            }
        }
        catch {
            $message = $_.Exception.Message
            if (-not (Test-IbisIsAdministrator)) {
                $message = "$message Run as Administrator to add Defender exclusions."
            }

            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                Path = $status.Path
                Status = 'Failed'
                Message = $message
                IsAdministrator = (Test-IbisIsAdministrator)
            }
        }
    }
}

function Remove-IbisDefenderExclusion {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object[]]$ToolDefinitions
    )

    $statuses = @(Get-IbisDefenderExclusionStatus -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions)
    if ($statuses.Count -eq 0) {
        return
    }

    $removePreferenceCommand = Get-Command Remove-MpPreference -ErrorAction SilentlyContinue
    if ($null -eq $removePreferenceCommand) {
        foreach ($status in $statuses) {
            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                Path = $status.Path
                Status = 'Unavailable'
                Message = 'Remove-MpPreference is not available on this system.'
                IsAdministrator = (Test-IbisIsAdministrator)
            }
        }
        return
    }

    foreach ($status in $statuses) {
        if (-not $status.Present) {
            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                Path = $status.Path
                Status = 'Missing'
                Message = 'Defender exclusion is not present.'
                IsAdministrator = (Test-IbisIsAdministrator)
            }
            continue
        }

        try {
            if ($PSCmdlet.ShouldProcess($status.Path, 'Remove Windows Defender exclusion')) {
                Remove-MpPreference -ExclusionPath $status.Path -ErrorAction Stop
            }

            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                Path = $status.Path
                Status = 'Removed'
                Message = 'Defender exclusion removed.'
                IsAdministrator = (Test-IbisIsAdministrator)
            }
        }
        catch {
            $message = $_.Exception.Message
            if (-not (Test-IbisIsAdministrator)) {
                $message = "$message Run as Administrator to remove Defender exclusions."
            }

            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                Path = $status.Path
                Status = 'Failed'
                Message = $message
                IsAdministrator = (Test-IbisIsAdministrator)
            }
        }
    }
}

function Resolve-IbisComparablePath {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ''
    }

    try {
        $fullPath = [System.IO.Path]::GetFullPath($Path)
    }
    catch {
        $fullPath = $Path
    }

    $fullPath.TrimEnd([char[]]@('\', '/')).ToLowerInvariant()
}

function Invoke-IbisDownloadFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    try {
        Start-BitsTransfer -Source $Uri -Destination $DestinationPath -ErrorAction Stop
    }
    catch {
        Invoke-WebRequest -Uri $Uri -OutFile $DestinationPath -UseBasicParsing
    }
}

function Expand-IbisArchive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    try {
        Import-Module Microsoft.PowerShell.Archive -ErrorAction Stop
        Expand-Archive -LiteralPath $LiteralPath -DestinationPath $DestinationPath -Force
        return
    }
    catch {
        $archiveError = $_.Exception.Message
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
            [System.IO.Compression.ZipFile]::ExtractToDirectory($LiteralPath, $DestinationPath)
            return
        }
        catch {
            throw "Unable to extract ZIP archive. Expand-Archive failed with: $archiveError. .NET fallback failed with: $($_.Exception.Message)"
        }
    }
}

function New-IbisToolInstallWorkspace {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition
    )

    $id = [System.Guid]::NewGuid().ToString()
    $installDirectory = Get-IbisToolInstallDirectory -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition

    if ($ToolDefinition.defenderExclusionRecommended -eq $true) {
        if (-not (Test-Path -LiteralPath $installDirectory)) {
            New-Item -ItemType Directory -Path $installDirectory -Force | Out-Null
        }
        $workspaceRoot = Join-Path (Join-Path $installDirectory '_ibis-staging') $id
    }
    else {
        $workspaceRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('IbisToolInstall-' + $id)
    }

    $downloadDirectory = Join-Path $workspaceRoot 'download'
    $extractDirectory = Join-Path $workspaceRoot 'extract'
    New-Item -ItemType Directory -Path $downloadDirectory -Force | Out-Null
    New-Item -ItemType Directory -Path $extractDirectory -Force | Out-Null

    [pscustomobject]@{
        Root = $workspaceRoot
        DownloadDirectory = $downloadDirectory
        ExtractDirectory = $extractDirectory
        InstallDirectory = $installDirectory
        IsDefenderAware = ($ToolDefinition.defenderExclusionRecommended -eq $true)
    }
}

function Get-IbisToolPublishSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExtractDirectory,

        [object]$ToolDefinition
    )

    if ($ToolDefinition -and $ToolDefinition.renameExtractedDirectoryFrom) {
        return $ExtractDirectory
    }

    $items = @(Get-ChildItem -LiteralPath $ExtractDirectory -Force)
    $directories = @($items | Where-Object { $_.PSIsContainer })
    $files = @($items | Where-Object { -not $_.PSIsContainer })

    if ($directories.Count -eq 1 -and $files.Count -eq 0) {
        return $directories[0].FullName
    }

    $ExtractDirectory
}

function Backup-IbisToolInstallDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstallDirectory,

        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [object]$ToolDefinition
    )

    if (-not (Test-Path -LiteralPath $InstallDirectory)) {
        return $null
    }

    $items = @()
    if ((Resolve-IbisComparablePath -Path $InstallDirectory) -eq (Resolve-IbisComparablePath -Path $ToolsRoot)) {
        if ($ToolDefinition -and $ToolDefinition.renameExtractedDirectoryTo) {
            $candidate = Join-Path $ToolsRoot $ToolDefinition.renameExtractedDirectoryTo
            if (Test-Path -LiteralPath $candidate) {
                $items += Get-Item -LiteralPath $candidate
            }
        }
        if ($ToolDefinition -and $ToolDefinition.renameExtractedDirectoryFrom) {
            $candidate = Join-Path $ToolsRoot $ToolDefinition.renameExtractedDirectoryFrom
            if (Test-Path -LiteralPath $candidate) {
                $items += Get-Item -LiteralPath $candidate
            }
        }
    }
    else {
        $items = @(Get-ChildItem -LiteralPath $InstallDirectory -Force | Where-Object {
            $_.Name -ne '_ibis-staging' -and $_.Name -ne '_ibis-backup'
        })
    }

    if ($items.Count -eq 0) {
        return $null
    }

    $backupRoot = Join-Path $InstallDirectory '_ibis-backup'
    $backupPath = Join-Path $backupRoot (Get-Date -Format 'yyyyMMdd-HHmmss')
    New-Item -ItemType Directory -Path $backupPath -Force | Out-Null

    foreach ($item in $items) {
        Move-Item -LiteralPath $item.FullName -Destination $backupPath -Force
    }

    $backupPath
}

function Publish-IbisStagedToolInstall {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$StagedSourcePath,

        [Parameter(Mandatory = $true)]
        [string]$InstallDirectory,

        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [object]$ToolDefinition
    )

    if (-not (Test-Path -LiteralPath $InstallDirectory)) {
        New-Item -ItemType Directory -Path $InstallDirectory -Force | Out-Null
    }

    $backupPath = $null
    $items = @(Get-ChildItem -LiteralPath $StagedSourcePath -Force)
    foreach ($item in $items) {
        $destinationPath = Join-Path $InstallDirectory $item.Name
        if (Test-Path -LiteralPath $destinationPath) {
            if ($null -eq $backupPath) {
                $backupPath = New-IbisToolBackupPath -InstallDirectory $InstallDirectory
            }
            Move-Item -LiteralPath $destinationPath -Destination $backupPath -Force
        }
        Move-Item -LiteralPath $item.FullName -Destination $InstallDirectory -Force
    }

    $backupPath
}

function New-IbisToolBackupPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstallDirectory
    )

    $backupRoot = Join-Path $InstallDirectory '_ibis-backup'
    $backupPath = Join-Path $backupRoot (Get-Date -Format 'yyyyMMdd-HHmmss')
    $suffix = 0
    while (Test-Path -LiteralPath $backupPath) {
        $suffix++
        $backupPath = Join-Path $backupRoot ('{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $suffix)
    }

    New-Item -ItemType Directory -Path $backupPath -Force | Out-Null
    $backupPath
}

function Test-IbisToolInstallState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition
    )

    $expectedPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
    $installDirectory = Get-IbisToolInstallDirectory -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition

    if (Test-Path -LiteralPath $expectedPath -PathType Leaf) {
        return [pscustomobject]@{
            Id = $ToolDefinition.id
            Name = $ToolDefinition.name
            Status = 'Present'
            Present = $true
            ExpectedPath = $expectedPath
            InstallDirectory = $installDirectory
            Message = 'Expected executable is present.'
        }
    }

    if (Test-Path -LiteralPath $installDirectory -PathType Container) {
        $items = @(Get-ChildItem -LiteralPath $installDirectory -Force | Where-Object {
            $_.Name -ne '_ibis-staging' -and $_.Name -ne '_ibis-backup'
        })
        if ((Resolve-IbisComparablePath -Path $installDirectory) -eq (Resolve-IbisComparablePath -Path $ToolsRoot) -and $ToolDefinition.renameExtractedDirectoryTo) {
            $toolFolder = Join-Path $ToolsRoot $ToolDefinition.renameExtractedDirectoryTo
            $items = @()
            if (Test-Path -LiteralPath $toolFolder) {
                $items = @(Get-Item -LiteralPath $toolFolder)
            }
        }
        if ($items.Count -gt 0) {
            return [pscustomobject]@{
                Id = $ToolDefinition.id
                Name = $ToolDefinition.name
                Status = 'Partial'
                Present = $false
                ExpectedPath = $expectedPath
                InstallDirectory = $installDirectory
                Message = 'Install directory contains files, but the expected executable is missing.'
            }
        }
    }

    [pscustomobject]@{
        Id = $ToolDefinition.id
        Name = $ToolDefinition.name
        Status = 'Missing'
        Present = $false
        ExpectedPath = $expectedPath
        InstallDirectory = $installDirectory
        Message = 'Expected executable is missing.'
    }
}

function Invoke-IbisInstallTool {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition,

        [string]$ProgressPath,

        [int]$ProgressIndex = 0,

        [int]$ProgressTotal = 0
    )

    $installState = Test-IbisToolInstallState -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
    $expectedPath = $installState.ExpectedPath
    if ($installState.Present) {
        Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'Present' -Message 'Tool already present.' -Index $ProgressIndex -Total $ProgressTotal -Status 'Skipped'
        return [pscustomobject]@{
            Id = $ToolDefinition.id
            Name = $ToolDefinition.name
            Status = 'Present'
            ExpectedPath = $expectedPath
            Message = 'Tool already present.'
        }
    }

    Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'Resolve' -Message 'Resolving download URL.' -Index $ProgressIndex -Total $ProgressTotal
    $downloadUrl = Resolve-IbisToolDownloadUrl -ToolDefinition $ToolDefinition
    $installDirectory = Get-IbisToolInstallDirectory -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
    $backupPath = $null
    $packageType = $ToolDefinition.packageType
    if ([string]::IsNullOrWhiteSpace($packageType)) {
        if ($downloadUrl -like '*.zip') {
            $packageType = 'zip'
        }
        else {
            $packageType = 'file'
        }
    }

    if ($PSCmdlet.ShouldProcess($ToolDefinition.name, "Download and install to $installDirectory")) {
        if (-not (Test-Path -LiteralPath $ToolsRoot)) {
            New-Item -ItemType Directory -Path $ToolsRoot | Out-Null
        }

        Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'Workspace' -Message 'Creating install workspace.' -Index $ProgressIndex -Total $ProgressTotal
        $workspace = New-IbisToolInstallWorkspace -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
        try {
            $downloadFileName = $ToolDefinition.downloadFileName
            if ([string]::IsNullOrWhiteSpace($downloadFileName)) {
                $downloadFileName = Split-Path -Path ([uri]$downloadUrl).AbsolutePath -Leaf
            }
            if ([string]::IsNullOrWhiteSpace($downloadFileName)) {
                $downloadFileName = $ToolDefinition.id + '.download'
            }

            $downloadPath = Join-Path $workspace.DownloadDirectory $downloadFileName
            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'Download' -Message "Downloading $downloadFileName." -Index $ProgressIndex -Total $ProgressTotal
            Invoke-IbisDownloadFile -Uri $downloadUrl -DestinationPath $downloadPath
            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'Downloaded' -Message "Downloaded $downloadFileName." -Index $ProgressIndex -Total $ProgressTotal

            if ($packageType -eq 'zip') {
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'Extract' -Message 'Extracting ZIP archive.' -Index $ProgressIndex -Total $ProgressTotal
                Expand-IbisArchive -LiteralPath $downloadPath -DestinationPath $workspace.ExtractDirectory
            }
            elseif ($packageType -eq 'file') {
                $targetName = $ToolDefinition.downloadFileName
                if ([string]::IsNullOrWhiteSpace($targetName)) {
                    $targetName = Split-Path -Path $ToolDefinition.executablePath -Leaf
                }
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'Copy' -Message "Copying $targetName into staging." -Index $ProgressIndex -Total $ProgressTotal
                Copy-Item -LiteralPath $downloadPath -Destination (Join-Path $workspace.ExtractDirectory $targetName) -Force
            }
            else {
                throw "Unsupported packageType '$packageType' for $($ToolDefinition.name)."
            }

            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'Publish' -Message "Publishing files to $installDirectory." -Index $ProgressIndex -Total $ProgressTotal
            $publishSource = Get-IbisToolPublishSource -ExtractDirectory $workspace.ExtractDirectory -ToolDefinition $ToolDefinition
            $backupPath = Publish-IbisStagedToolInstall -StagedSourcePath $publishSource -InstallDirectory $installDirectory -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'PostInstall' -Message 'Running post-install checks.' -Index $ProgressIndex -Total $ProgressTotal
            Invoke-IbisToolPostInstall -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
        }
        finally {
            if ($workspace -and (Test-Path -LiteralPath $workspace.Root)) {
                try {
                    Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'Cleanup' -Message 'Removing install workspace.' -Index $ProgressIndex -Total $ProgressTotal
                    Remove-Item -LiteralPath $workspace.Root -Recurse -Force
                }
                catch {
                    Write-Warning "Unable to remove Ibis install workspace: $($workspace.Root). $($_.Exception.Message)"
                }
            }
        }
    }

    $present = Test-Path -LiteralPath $expectedPath -PathType Leaf
    $status = 'Installed'
    $message = 'Tool installed.'
    if (-not $present) {
        $status = 'Install Incomplete'
        $message = 'Download completed, but expected executable was not found. Manual review is needed.'
    }

    Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage $status -Message $message -Index $ProgressIndex -Total $ProgressTotal -Status $(if ($present) { 'Completed' } else { 'Warning' })

    [pscustomobject]@{
        Id = $ToolDefinition.id
        Name = $ToolDefinition.name
        Status = $status
        ExpectedPath = $expectedPath
        DownloadUrl = $downloadUrl
        BackupPath = $backupPath
        Message = $message
    }
}

function Invoke-IbisInstallMissingTools {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [string[]]$ToolIds,

        [string]$ProgressPath
    )

    $statuses = @(Test-IbisToolStatus -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions)
    $missing = @($statuses | Where-Object { -not $_.Present })
    if ($ToolIds -and $ToolIds.Count -gt 0) {
        $missing = @($missing | Where-Object { $ToolIds -contains $_.Id })
    }

    $total = $missing.Count
    Write-IbisProgressEvent -ProgressPath $ProgressPath -Stage 'Start' -Message "Starting install for $total missing tool(s)." -Index 0 -Total $total
    $index = 0
    foreach ($status in $missing) {
        $index++
        $definition = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id $status.Id
        if ($null -eq $definition) {
            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $status.Id -ToolName $status.Name -Stage 'Skipped' -Message 'No matching tool definition was found.' -Index $index -Total $total -Status 'Warning'
            [pscustomobject]@{
                Id = $status.Id
                Name = $status.Name
                Status = 'Skipped'
                ExpectedPath = $status.ExpectedPath
                Message = 'No matching tool definition was found.'
            }
            continue
        }

        try {
            if ($PSCmdlet.ShouldProcess($definition.name, 'Install missing Ibis tool')) {
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $definition.id -ToolName $definition.name -Stage 'ToolStart' -Message "Starting $($definition.name)." -Index $index -Total $total
                Invoke-IbisInstallTool -ToolsRoot $ToolsRoot -ToolDefinition $definition -ProgressPath $ProgressPath -ProgressIndex $index -ProgressTotal $total
            }
        }
        catch {
            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $definition.id -ToolName $definition.name -Stage 'Failed' -Message $_.Exception.Message -Index $index -Total $total -Status 'Failed'
            [pscustomobject]@{
                Id = $definition.id
                Name = $definition.name
                Status = 'Failed'
                ExpectedPath = (Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $definition)
                Message = $_.Exception.Message
            }
        }
    }

    Write-IbisProgressEvent -ProgressPath $ProgressPath -Stage 'Finished' -Message 'Install run finished.' -Index $total -Total $total -Status 'Completed'
}

function Invoke-IbisToolPostInstall {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition
    )

    if ($ToolDefinition.renameExtractedDirectoryFrom -and $ToolDefinition.renameExtractedDirectoryTo) {
        $fromPath = Join-Path $ToolsRoot $ToolDefinition.renameExtractedDirectoryFrom
        $toPath = Join-Path $ToolsRoot $ToolDefinition.renameExtractedDirectoryTo
        if (Test-Path -LiteralPath $fromPath) {
            if (Test-Path -LiteralPath $toPath) {
                $backupPath = New-IbisToolBackupPath -InstallDirectory $ToolsRoot
                Move-Item -LiteralPath $toPath -Destination $backupPath -Force
            }
            Rename-Item -LiteralPath $fromPath -NewName $ToolDefinition.renameExtractedDirectoryTo
        }
    }

    if ($ToolDefinition.renameExecutablePattern -and $ToolDefinition.renameExecutableTo) {
        $installDirectory = Get-IbisToolInstallDirectory -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
        $renameCandidates = @(Get-IbisExecutableRenameCandidate -InstallDirectory $installDirectory -Pattern $ToolDefinition.renameExecutablePattern)
        if ($renameCandidates.Count -eq 1) {
            $target = Join-Path $installDirectory $ToolDefinition.renameExecutableTo
            if ($renameCandidates[0].FullName -ne $target) {
                Move-Item -LiteralPath $renameCandidates[0].FullName -Destination $target -Force
            }
        }
    }
}

function Get-IbisExecutableRenameCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstallDirectory,

        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    if (-not (Test-Path -LiteralPath $InstallDirectory -PathType Container)) {
        return @()
    }

    Get-ChildItem -LiteralPath $InstallDirectory -Filter $Pattern -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object {
            $relativePath = $_.FullName.Substring($InstallDirectory.Length).TrimStart('\', '/')
            $pathParts = $relativePath -split '[\\/]'
            -not ($pathParts -contains '_ibis-backup') -and -not ($pathParts -contains '_ibis-staging')
        } |
        Sort-Object FullName
}

function New-IbisCommandSpec {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolId,

        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$ArgumentList = @(),

        [string]$WorkingDirectory,

        [string]$Description,

        [string]$ExpectedOutputPath
    )

    [pscustomobject]@{
        ToolId = $ToolId
        FilePath = $FilePath
        ArgumentList = @($ArgumentList)
        WorkingDirectory = $WorkingDirectory
        Description = $Description
        ExpectedOutputPath = $ExpectedOutputPath
    }
}

function ConvertTo-IbisCommandLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$CommandSpec
    )

    $parts = @()
    $parts += ConvertTo-IbisQuotedArgument -Value $CommandSpec.FilePath
    foreach ($argument in $CommandSpec.ArgumentList) {
        $parts += ConvertTo-IbisQuotedArgument -Value $argument
    }

    $parts -join ' '
}

function ConvertTo-IbisQuotedArgument {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return '""'
    }

    $escaped = $Value -replace '"', '\"'
    if ($escaped -match '\s|["]') {
        return '"' + $escaped + '"'
    }

    $escaped
}

function Resolve-IbisComparablePath {
    [CmdletBinding()]
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    try {
        ([System.IO.Path]::GetFullPath($Path)).TrimEnd('\', '/')
    }
    catch {
        $Path.TrimEnd('\', '/')
    }
}

function Test-IbisPathInsideRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $resolvedPath = Resolve-IbisComparablePath -Path $Path
    $resolvedRoot = Resolve-IbisComparablePath -Path $RootPath
    if ([string]::IsNullOrWhiteSpace($resolvedPath) -or [string]::IsNullOrWhiteSpace($resolvedRoot)) {
        return $false
    }

    if ([string]::Equals($resolvedPath, $resolvedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    $rootWithSeparator = $resolvedRoot
    if (-not $rootWithSeparator.EndsWith('\') -and -not $rootWithSeparator.EndsWith('/')) {
        $rootWithSeparator += [System.IO.Path]::DirectorySeparatorChar
    }

    $resolvedPath.StartsWith($rootWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-IbisSourceWriteBoundary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$WritablePaths
    )

    $violations = @()
    foreach ($item in $WritablePaths) {
        if ($null -eq $item) {
            continue
        }

        $name = [string]$item.Name
        $path = [string]$item.Path
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = 'Writable path'
        }
        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        $insideSource = Test-IbisPathInsideRoot -Path $path -RootPath $SourceRoot
        if ($insideSource) {
            $violations += [pscustomobject]@{
                Name = $name
                Path = $path
                SourceRoot = $SourceRoot
            }
        }
    }

    [pscustomobject]@{
        SourceRoot = $SourceRoot
        Passed = ($violations.Count -eq 0)
        Violations = $violations
        Message = if ($violations.Count -eq 0) {
            'Writable paths are outside the evidence source root.'
        }
        else {
            'One or more writable paths are inside the evidence source root.'
        }
    }
}

function Test-IbisEvidenceRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    $checks = @(
        @{ Name = 'Windows registry hives'; RelativePath = 'Windows\System32\config'; Required = $true },
        @{ Name = 'Windows Event Logs'; RelativePath = 'Windows\System32\winevt\Logs'; Required = $false },
        @{ Name = 'Prefetch'; RelativePath = 'Windows\Prefetch'; Required = $false },
        @{ Name = 'Amcache'; RelativePath = 'Windows\appcompat\Programs'; Required = $false },
        @{ Name = 'Users'; RelativePath = 'Users'; Required = $true }
    )

    $results = @()
    foreach ($check in $checks) {
        $path = Join-Path $SourceRoot $check.RelativePath
        $exists = Test-Path -LiteralPath $path
        $results += [pscustomobject]@{
            Name = $check.Name
            RelativePath = $check.RelativePath
            Path = $path
            Required = [bool]$check.Required
            Present = [bool]$exists
        }
    }

    $requiredMissing = @($results | Where-Object { $_.Required -and -not $_.Present })
    [pscustomobject]@{
        SourceRoot = $SourceRoot
        Present = Test-Path -LiteralPath $SourceRoot
        LooksLikeWindowsEvidence = ($requiredMissing.Count -eq 0)
        Checks = $results
    }
}

function Find-IbisVelociraptorResultsPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    $candidates = @()
    $current = $SourceRoot
    for ($i = 0; $i -lt 5; $i++) {
        if ([string]::IsNullOrWhiteSpace($current)) {
            break
        }

        $candidates += Join-Path $current 'Results'
        $parent = Split-Path -Path $current -Parent
        if ($parent -eq $current) {
            break
        }
        $current = $parent
    }

    foreach ($candidate in $candidates | Select-Object -Unique) {
        if (Test-Path -LiteralPath $candidate -PathType Container) {
            return $candidate
        }
    }

    $null
}

function Invoke-IbisVelociraptorResultsCopy {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $resultsPath = Find-IbisVelociraptorResultsPath -SourceRoot $SourceRoot
    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''

    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $destinationRoot = Join-Path $hostOutputRoot 'Velociraptor-Results'
    $destinationPath = $destinationRoot

    if ($null -eq $resultsPath) {
        return [pscustomobject]@{
            ModuleId = 'velociraptor-results'
            Status = 'Skipped'
            SourcePath = $null
            OutputPath = $destinationPath
            SourceItemCount = 0
            CopiedItemCount = 0
            Message = 'No Velociraptor Results folder was found near the source root.'
        }
    }

    $sourceItems = @(Get-ChildItem -LiteralPath $resultsPath -Recurse -Force -ErrorAction SilentlyContinue)
    $sourceFiles = @($sourceItems | Where-Object { -not $_.PSIsContainer })
    if ($sourceFiles.Count -eq 0) {
        return [pscustomobject]@{
            ModuleId = 'velociraptor-results'
            Status = 'Skipped'
            SourcePath = $resultsPath
            OutputPath = $destinationPath
            SourceItemCount = $sourceItems.Count
            CopiedItemCount = 0
            Message = 'Velociraptor Results folder was found, but it did not contain any files.'
        }
    }

    if ($PSCmdlet.ShouldProcess($resultsPath, "Copy Velociraptor Results to $destinationPath")) {
        if (-not (Test-Path -LiteralPath $destinationRoot)) {
            New-Item -ItemType Directory -Path $destinationRoot -Force | Out-Null
        }

        $items = Get-ChildItem -LiteralPath $resultsPath -Force
        foreach ($item in $items) {
            Copy-Item -LiteralPath $item.FullName -Destination $destinationPath -Recurse -Force
        }
    }

    $copiedItems = @()
    if (Test-Path -LiteralPath $destinationPath -PathType Container) {
        $copiedItems = @(Get-ChildItem -LiteralPath $destinationPath -Recurse -Force -ErrorAction SilentlyContinue)
    }

    $status = 'Completed'
    $message = "Velociraptor Results copied. $($sourceItems.Count) source item(s), $($copiedItems.Count) copied item(s)."
    if ($copiedItems.Count -lt $sourceItems.Count) {
        $status = 'Completed With Warnings'
        $message = "Velociraptor Results copied, but destination item count is lower than the source count. Source: $($sourceItems.Count), destination: $($copiedItems.Count)."
    }

    [pscustomobject]@{
        ModuleId = 'velociraptor-results'
        Status = $status
        SourcePath = $resultsPath
        OutputPath = $destinationPath
        SourceItemCount = $sourceItems.Count
        CopiedItemCount = $copiedItems.Count
        Message = $message
    }
}

function ConvertTo-IbisSafeFileName {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Value,

        [string]$DefaultValue = 'HOST'
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        $Value = $DefaultValue
    }

    $invalidPattern = '[{0}]' -f ([regex]::Escape((-join [System.IO.Path]::GetInvalidFileNameChars())))
    $safe = $Value -replace $invalidPattern, '_'
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return $DefaultValue
    }

    $safe
}

function Get-IbisHostFilePrefix {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Hostname
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace($safeHost)) {
        return ''
    }

    '{0}-' -f $safeHost
}

function New-IbisHostPrefixedFileName {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [string]$Suffix
    )

    (Get-IbisHostFilePrefix -Hostname $Hostname) + $Suffix
}

function Format-IbisHostPrefixedValue {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [string]$Format,

        [object[]]$ArgumentList = @()
    )

    (Get-IbisHostFilePrefix -Hostname $Hostname) + ($Format -f $ArgumentList)
}

function Get-IbisHostOutputRoot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $trimmedOutputRoot = $OutputRoot.TrimEnd([char[]]@('\', '/'))
    if ([string]::IsNullOrWhiteSpace($safeHost)) {
        return $trimmedOutputRoot
    }

    $leaf = [System.IO.Path]::GetFileName($trimmedOutputRoot)
    if ($leaf -and $leaf.Equals($safeHost, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $trimmedOutputRoot
    }

    Join-Path $trimmedOutputRoot $safeHost
}

function Get-IbisSystemHivePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('SAM', 'SECURITY', 'SOFTWARE', 'SYSTEM')]
        [string]$HiveName
    )

    [System.IO.Path]::Combine($SourceRoot, ('Windows\System32\config\{0}' -f $HiveName))
}

function Get-IbisWindowsRegistryHiveName {
    [CmdletBinding()]
    param()

    @('SAM', 'SECURITY', 'SOFTWARE', 'SYSTEM')
}

function Test-IbisRegistryHiveTransactionState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$HivePath
    )

    if (-not (Test-Path -LiteralPath $HivePath -PathType Leaf)) {
        return [pscustomobject]@{
            HivePath = $HivePath
            Status = 'Missing'
            IsDirty = $null
            ExitCode = $null
            Message = 'Hive was not found.'
            StandardOutput = ''
            StandardError = ''
        }
    }

    $regRipper = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'regripper'
    if ($null -eq $regRipper) {
        return [pscustomobject]@{
            HivePath = $HivePath
            Status = 'Unknown'
            IsDirty = $null
            ExitCode = $null
            Message = 'RegRipper is not configured, so hive transaction state could not be checked.'
            StandardOutput = ''
            StandardError = ''
        }
    }

    $ripPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $regRipper
    if (-not (Test-Path -LiteralPath $ripPath -PathType Leaf)) {
        return [pscustomobject]@{
            HivePath = $HivePath
            Status = 'Unknown'
            IsDirty = $null
            ExitCode = $null
            Message = "RegRipper is missing at: $ripPath"
            StandardOutput = ''
            StandardError = ''
        }
    }

    try {
        $result = Invoke-IbisProcessCapture `
            -FilePath $ripPath `
            -ArgumentList @('-r', $HivePath, '-d') `
            -WorkingDirectory (Split-Path -Path $ripPath -Parent)
        $combinedOutput = @($result.StandardOutput, $result.StandardError) -join [Environment]::NewLine

        if ($combinedOutput -match 'The hive \(.+\) is dirty\.|Hive is dirty|is dirty') {
            $status = 'Dirty'
            $isDirty = $true
            $message = 'Hive appears dirty; transaction logs may need to be replayed.'
        }
        elseif ($combinedOutput -match 'Hive is not dirty|is not dirty|not dirty') {
            $status = 'Clean'
            $isDirty = $false
            $message = 'Hive is not dirty.'
        }
        else {
            $status = 'Unknown'
            $isDirty = $null
            $message = 'RegRipper completed, but Ibis could not determine whether the hive is dirty.'
        }

        [pscustomobject]@{
            HivePath = $HivePath
            Status = $status
            IsDirty = $isDirty
            ExitCode = $result.ExitCode
            CommandLine = $result.CommandLine
            Message = $message
            StandardOutput = $result.StandardOutput
            StandardError = $result.StandardError
        }
    }
    catch {
        [pscustomobject]@{
            HivePath = $HivePath
            Status = 'Unknown'
            IsDirty = $null
            ExitCode = $null
            Message = "Hive transaction state check failed: $($_.Exception.Message)"
            StandardOutput = ''
            StandardError = ''
        }
    }
}

function Copy-IbisRegistryHiveToCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceHivePath,

        [Parameter(Mandatory = $true)]
        [string]$CacheDirectory
    )

    if (-not (Test-Path -LiteralPath $CacheDirectory)) {
        New-Item -ItemType Directory -Path $CacheDirectory -Force | Out-Null
    }

    $hiveName = Split-Path -Path $SourceHivePath -Leaf
    $destinationHivePath = Join-Path $CacheDirectory $hiveName
    Copy-Item -LiteralPath $SourceHivePath -Destination $destinationHivePath -Force

    $sourceDirectory = Split-Path -Path $SourceHivePath -Parent
    $transactionLogs = @(Get-ChildItem -LiteralPath $sourceDirectory -Filter "$hiveName.LOG*" -File -Force -ErrorAction SilentlyContinue)
    foreach ($log in $transactionLogs) {
        Copy-Item -LiteralPath $log.FullName -Destination (Join-Path $CacheDirectory $log.Name) -Force
    }

    [pscustomobject]@{
        HivePath = $destinationHivePath
        TransactionLogCount = $transactionLogs.Count
        TransactionLogs = @($transactionLogs | ForEach-Object { $_.Name })
    }
}

function Invoke-IbisRegistryHiveTransactionReplay {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$HivePath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory
    )

    $rla = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'zimmerman-rla'
    if ($null -eq $rla) {
        return [pscustomobject]@{
            Status = 'Skipped'
            ExitCode = $null
            Message = 'rla is not configured, so registry transaction logs could not be replayed.'
            StandardOutput = ''
            StandardError = ''
        }
    }

    $rlaPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $rla
    if (-not (Test-Path -LiteralPath $rlaPath -PathType Leaf)) {
        return [pscustomobject]@{
            Status = 'Skipped'
            ExitCode = $null
            Message = "rla is missing at: $rlaPath"
            StandardOutput = ''
            StandardError = ''
        }
    }

    try {
        $result = Invoke-IbisProcessCapture `
            -FilePath $rlaPath `
            -ArgumentList @('-f', $HivePath, '--out', $OutputDirectory, '--nop') `
            -WorkingDirectory (Split-Path -Path $rlaPath -Parent)
        $combinedOutput = @($result.StandardOutput, $result.StandardError) -join [Environment]::NewLine

        $status = 'Completed'
        $message = 'rla completed.'
        if ($combinedOutput -match 'There was an error|error occurred|exception') {
            $status = 'Failed'
            $message = 'rla reported an error while replaying transaction logs.'
        }
        elseif ($combinedOutput -match 'At least one transaction log was applied') {
            $message = 'rla applied at least one transaction log.'
        }
        elseif ($combinedOutput -match 'is not dirty') {
            $message = 'rla reported that the hive was not dirty.'
        }
        elseif ($combinedOutput -match 'is dirty, but no logs were found|no logs were found') {
            $status = 'Completed With Warnings'
            $message = 'rla reported that the hive is dirty, but no transaction logs were found.'
        }
        elseif ($result.ExitCode -ne 0) {
            $status = 'Failed'
            $message = "rla exited with code $($result.ExitCode)."
        }

        [pscustomobject]@{
            Status = $status
            ExitCode = $result.ExitCode
            CommandLine = $result.CommandLine
            Message = $message
            StandardOutput = $result.StandardOutput
            StandardError = $result.StandardError
        }
    }
    catch {
        [pscustomobject]@{
            Status = 'Failed'
            ExitCode = $null
            Message = "rla failed: $($_.Exception.Message)"
            StandardOutput = ''
            StandardError = ''
        }
    }
}

function Get-IbisCachedRegistryHivePreparation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$MetadataPath,

        [Parameter(Mandatory = $true)]
        [string]$SourceHivePath
    )

    if (-not (Test-Path -LiteralPath $MetadataPath -PathType Leaf)) {
        return $null
    }

    try {
        $cached = Get-Content -LiteralPath $MetadataPath -Raw | ConvertFrom-Json
        if (-not (Test-Path -LiteralPath $cached.PreparedHivePath -PathType Leaf)) {
            return $null
        }

        $sourceItem = Get-Item -LiteralPath $SourceHivePath -ErrorAction Stop
        if ($cached.SourceLength -ne $sourceItem.Length) {
            return $null
        }
        if ($null -ne $cached.SourceLastWriteTimeUtcTicks) {
            $sourceDeltaTicks = [math]::Abs(([int64]$cached.SourceLastWriteTimeUtcTicks) - $sourceItem.LastWriteTimeUtc.Ticks)
            if ($sourceDeltaTicks -gt [TimeSpan]::TicksPerSecond) {
                return $null
            }
        }
        else {
            try {
                $cachedLastWrite = [datetime]::Parse(
                    [string]$cached.SourceLastWriteTimeUtc,
                    [Globalization.CultureInfo]::InvariantCulture,
                    [Globalization.DateTimeStyles]::RoundtripKind
                )
            }
            catch {
                return $null
            }
            $lastWriteDelta = [math]::Abs(($cachedLastWrite.ToUniversalTime() - $sourceItem.LastWriteTimeUtc.ToUniversalTime()).TotalSeconds)
            if ($lastWriteDelta -gt 1) {
                return $null
            }
        }

        $cached | Add-Member -NotePropertyName CacheHit -NotePropertyValue $true -Force
        $cached
    }
    catch {
        $null
    }
}

function Invoke-IbisPrepareRegistryHive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [ValidateSet('SAM', 'SECURITY', 'SOFTWARE', 'SYSTEM')]
        [string]$HiveName
    )

    $sourceHivePath = Get-IbisSystemHivePath -SourceRoot $SourceRoot -HiveName $HiveName
    Invoke-IbisPrepareRegistryHiveFile `
        -ToolsRoot $ToolsRoot `
        -ToolDefinitions $ToolDefinitions `
        -SourceHivePath $sourceHivePath `
        -OutputRoot $OutputRoot `
        -Hostname $Hostname `
        -HiveName $HiveName `
        -CacheGroup 'Registry-Hives'
}

function Invoke-IbisPrepareRegistryHiveFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceHivePath,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [string]$HiveName,

        [string]$CacheGroup = 'Registry-Hives',

        [string]$CacheKey
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    if ([string]::IsNullOrWhiteSpace($CacheKey)) {
        $CacheKey = $HiveName
    }
    $safeCacheGroup = ConvertTo-IbisSafeFileName -Value $CacheGroup -DefaultValue 'Registry-Hives'
    $safeHiveName = ConvertTo-IbisSafeFileName -Value $HiveName -DefaultValue 'Hive'
    $safeCacheKey = ConvertTo-IbisSafeFileName -Value $CacheKey -DefaultValue $safeHiveName
    $cacheRoot = Join-Path $hostOutputRoot (Join-Path $safeCacheGroup '_Working\Prepared-Hives')
    $hiveCacheDirectory = Join-Path $cacheRoot $safeCacheKey
    $metadataPath = Join-Path $hiveCacheDirectory ('{0}-Ibis-Hive-Preparation.json' -f $safeCacheKey)

    if (-not (Test-Path -LiteralPath $SourceHivePath -PathType Leaf)) {
        return [pscustomobject]@{
            HiveName = $HiveName
            Status = 'Skipped'
            SourceHivePath = $SourceHivePath
            PreparedHivePath = $null
            CacheDirectory = $hiveCacheDirectory
            CacheHit = $false
            IsDirtyBefore = $null
            IsDirtyAfter = $null
            TransactionLogCount = 0
            CheckBefore = $null
            Replay = $null
            CheckAfter = $null
            Message = 'Source hive was not found.'
        }
    }

    $cached = Get-IbisCachedRegistryHivePreparation -MetadataPath $metadataPath -SourceHivePath $SourceHivePath
    if ($null -ne $cached) {
        return $cached
    }

    $sourceItem = Get-Item -LiteralPath $SourceHivePath
    $copyResult = Copy-IbisRegistryHiveToCache -SourceHivePath $SourceHivePath -CacheDirectory $hiveCacheDirectory
    $checkBefore = Test-IbisRegistryHiveTransactionState -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $copyResult.HivePath

    $replay = $null
    $checkAfter = $null
    $status = 'Prepared'
    $message = 'Hive was copied to the prepared hive cache.'

    if ($checkBefore.IsDirty -eq $true) {
        $replay = Invoke-IbisRegistryHiveTransactionReplay `
            -ToolsRoot $ToolsRoot `
            -ToolDefinitions $ToolDefinitions `
            -HivePath $copyResult.HivePath `
            -OutputDirectory $hiveCacheDirectory
        $checkAfter = Test-IbisRegistryHiveTransactionState -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $copyResult.HivePath

        if ($checkAfter.IsDirty -eq $false) {
            $status = 'Cleaned'
            $message = 'Hive was dirty and transaction replay appears to have cleaned it.'
        }
        elseif ($checkAfter.IsDirty -eq $true) {
            $status = 'Prepared With Warnings'
            $message = 'Hive still appears dirty after transaction replay. Processing will continue using the cached copy.'
        }
        else {
            $status = 'Prepared With Warnings'
            $message = 'Hive transaction replay was attempted, but the final dirty state could not be confirmed.'
        }
    }
    elseif ($checkBefore.IsDirty -eq $false) {
        $status = 'Prepared'
        $message = 'Hive was not dirty and was cached for processing.'
    }
    else {
        $status = 'Prepared With Warnings'
        $message = 'Hive was cached, but its dirty state could not be determined.'
    }

    $prepared = [pscustomobject]@{
        HiveName = $HiveName
        CacheGroup = $safeCacheGroup
        CacheKey = $safeCacheKey
        Status = $status
        SourceHivePath = $SourceHivePath
        SourceLength = $sourceItem.Length
        SourceLastWriteTimeUtc = $sourceItem.LastWriteTimeUtc.ToString('o')
        SourceLastWriteTimeUtcTicks = $sourceItem.LastWriteTimeUtc.Ticks
        PreparedHivePath = $copyResult.HivePath
        CacheDirectory = $hiveCacheDirectory
        CacheHit = $false
        IsDirtyBefore = $checkBefore.IsDirty
        IsDirtyAfter = $(if ($null -ne $checkAfter) { $checkAfter.IsDirty } else { $checkBefore.IsDirty })
        TransactionLogCount = $copyResult.TransactionLogCount
        TransactionLogs = @($copyResult.TransactionLogs)
        CheckBefore = $checkBefore
        Replay = $replay
        CheckAfter = $checkAfter
        Message = $message
    }

    $prepared | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $metadataPath -Encoding UTF8
    $prepared
}

function Invoke-IbisPrepareRegistryHives {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST',

        [string[]]$HiveNames = @()
    )

    if ($HiveNames.Count -eq 0) {
        $HiveNames = @(Get-IbisWindowsRegistryHiveName)
    }

    foreach ($hiveName in $HiveNames) {
        Invoke-IbisPrepareRegistryHive `
            -ToolsRoot $ToolsRoot `
            -ToolDefinitions $ToolDefinitions `
            -SourceRoot $SourceRoot `
            -OutputRoot $OutputRoot `
            -Hostname $Hostname `
            -HiveName $hiveName
    }
}

function Invoke-IbisProcessCapture {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$ArgumentList = @(),

        [string]$WorkingDirectory
    )

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $FilePath
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true
    try {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $startInfo.StandardOutputEncoding = $utf8NoBom
        $startInfo.StandardErrorEncoding = $utf8NoBom
    }
    catch {
    }
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
        $startInfo.WorkingDirectory = $WorkingDirectory
    }

    $quotedArguments = @()
    foreach ($argument in $ArgumentList) {
        $quotedArguments += ConvertTo-IbisQuotedArgument -Value $argument
    }
    $startInfo.Arguments = $quotedArguments -join ' '
    $commandLine = ConvertTo-IbisCommandLine -CommandSpec ([pscustomobject]@{
        FilePath = $FilePath
        ArgumentList = @($ArgumentList)
    })

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    [void]$process.Start()
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    [pscustomobject]@{
        FilePath = $FilePath
        Arguments = @($ArgumentList)
        CommandLine = $commandLine
        WorkingDirectory = $WorkingDirectory
        ExitCode = $process.ExitCode
        StandardOutput = $stdout
        StandardError = $stderr
    }
}

function Invoke-IbisRegRipperPlugin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$HivePath,

        [Parameter(Mandatory = $true)]
        [string]$Plugin,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $regRipper = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'regripper'
    if ($null -eq $regRipper) {
        throw 'RegRipper is not configured.'
    }

    $ripPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $regRipper
    if (-not (Test-Path -LiteralPath $ripPath -PathType Leaf)) {
        throw "RegRipper is missing at: $ripPath"
    }

    if (-not (Test-Path -LiteralPath $HivePath -PathType Leaf)) {
        return [pscustomobject]@{
            Plugin = $Plugin
            HivePath = $HivePath
            OutputPath = $OutputPath
            Status = 'Skipped'
            ExitCode = $null
            Message = 'Hive was not found.'
        }
    }

    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    $result = Invoke-IbisProcessCapture `
        -FilePath $ripPath `
        -ArgumentList @('-r', $HivePath, '-p', $Plugin) `
        -WorkingDirectory (Split-Path -Path $ripPath -Parent)

    $result.StandardOutput | Out-File -LiteralPath $OutputPath -Encoding UTF8
    if ($result.ExitCode -ne 0) {
        $errorPath = $OutputPath + '.stderr.txt'
        $result.StandardError | Out-File -LiteralPath $errorPath -Encoding UTF8
    }

    $status = 'Completed'
    $message = 'RegRipper plugin completed.'
    if ($result.ExitCode -ne 0) {
        $status = 'Failed'
        $message = "RegRipper plugin exited with code $($result.ExitCode)."
    }

    [pscustomobject]@{
        Plugin = $Plugin
        HivePath = $HivePath
        OutputPath = $OutputPath
        Status = $status
        ExitCode = $result.ExitCode
        CommandLine = $result.CommandLine
        Message = $message
    }
}

function Invoke-IbisRegRipperHiveMode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$HivePath,

        [Parameter(Mandatory = $true)]
        [ValidateSet('All', 'Timeline')]
        [string]$Mode,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $regRipper = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'regripper'
    if ($null -eq $regRipper) {
        return [pscustomobject]@{
            Mode = $Mode
            HivePath = $HivePath
            OutputPath = $OutputPath
            Status = 'Failed'
            ExitCode = $null
            Message = 'RegRipper is not configured.'
        }
    }

    $ripPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $regRipper
    if (-not (Test-Path -LiteralPath $ripPath -PathType Leaf)) {
        return [pscustomobject]@{
            Mode = $Mode
            HivePath = $HivePath
            OutputPath = $OutputPath
            Status = 'Failed'
            ExitCode = $null
            Message = "RegRipper is missing at: $ripPath"
        }
    }

    if (-not (Test-Path -LiteralPath $HivePath -PathType Leaf)) {
        return [pscustomobject]@{
            Mode = $Mode
            HivePath = $HivePath
            OutputPath = $OutputPath
            Status = 'Skipped'
            ExitCode = $null
            Message = 'Hive was not found.'
        }
    }

    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }

    $modeArgument = '-a'
    if ($Mode -eq 'Timeline') {
        $modeArgument = '-aT'
    }

    $result = Invoke-IbisProcessCapture `
        -FilePath $ripPath `
        -ArgumentList @('-r', $HivePath, $modeArgument) `
        -WorkingDirectory (Split-Path -Path $ripPath -Parent)

    $result.StandardOutput | Out-File -LiteralPath $OutputPath -Encoding UTF8
    if ($result.ExitCode -ne 0) {
        $result.StandardError | Out-File -LiteralPath ($OutputPath + '.stderr.txt') -Encoding UTF8
    }

    $status = 'Completed'
    $message = 'RegRipper hive mode completed.'
    if ($result.ExitCode -ne 0) {
        $status = 'Failed'
        $message = "RegRipper exited with code $($result.ExitCode)."
    }
    elseif ((Test-Path -LiteralPath $OutputPath -PathType Leaf) -and ((Get-Item -LiteralPath $OutputPath).Length -eq 0)) {
        $status = 'Completed With Warnings'
        $message = 'RegRipper output file was created, but it is empty.'
    }

    [pscustomobject]@{
        Mode = $Mode
        HivePath = $HivePath
        OutputPath = $OutputPath
        Status = $status
        ExitCode = $result.ExitCode
        CommandLine = $result.CommandLine
        Message = $message
    }
}

function Invoke-IbisWindowsRegistryHives {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'Registry-Hives'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $preparedHives = @(Invoke-IbisPrepareRegistryHives `
        -ToolsRoot $ToolsRoot `
        -ToolDefinitions $ToolDefinitions `
        -SourceRoot $SourceRoot `
        -OutputRoot $OutputRoot `
        -Hostname $safeHost)

    $ripResults = @()
    foreach ($preparedHive in $preparedHives) {
        if ([string]::IsNullOrWhiteSpace($preparedHive.PreparedHivePath)) {
            continue
        }

        $allOutputPath = Join-Path $outputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}.txt' -ArgumentList @($preparedHive.HiveName))
        $ripResults += Invoke-IbisRegRipperHiveMode `
            -ToolsRoot $ToolsRoot `
            -ToolDefinitions $ToolDefinitions `
            -HivePath $preparedHive.PreparedHivePath `
            -Mode 'All' `
            -OutputPath $allOutputPath

        $timelineOutputPath = Join-Path $outputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-TLN.txt' -ArgumentList @($preparedHive.HiveName))
        $ripResults += Invoke-IbisRegRipperHiveMode `
            -ToolsRoot $ToolsRoot `
            -ToolDefinitions $ToolDefinitions `
            -HivePath $preparedHive.PreparedHivePath `
            -Mode 'Timeline' `
            -OutputPath $timelineOutputPath
    }

    $softwareHive = $preparedHives | Where-Object { $_.HiveName -eq 'SOFTWARE' } | Select-Object -First 1
    if ($null -ne $softwareHive -and -not [string]::IsNullOrWhiteSpace($softwareHive.PreparedHivePath)) {
        $runOutputPath = Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'RR-SOFTWARE-Run-AutoStart.txt')
        try {
            $ripResults += Invoke-IbisRegRipperPlugin `
                -ToolsRoot $ToolsRoot `
                -ToolDefinitions $ToolDefinitions `
                -HivePath $softwareHive.PreparedHivePath `
                -Plugin 'run' `
                -OutputPath $runOutputPath
        }
        catch {
            $ripResults += [pscustomobject]@{
                Plugin = 'run'
                HivePath = $softwareHive.PreparedHivePath
                OutputPath = $runOutputPath
                Status = 'Failed'
                ExitCode = $null
                Message = $_.Exception.Message
            }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Registry-Hives.json')
    $payload = [pscustomobject]@{
        ModuleId = 'registry'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        PreparedHives = $preparedHives
        RegRipperResults = $ripResults
    }
    $payload | ConvertTo-Json -Depth 10 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $failed = @($ripResults | Where-Object { $_.Status -eq 'Failed' })
    $warnings = @($preparedHives | Where-Object { $_.Status -match 'Warnings' })
    $warnings += @($ripResults | Where-Object { $_.Status -match 'Warnings' })
    $processed = @($preparedHives | Where-Object { -not [string]::IsNullOrWhiteSpace($_.PreparedHivePath) })

    $status = 'Completed'
    $message = "Registry hive processing completed for $($processed.Count) hive(s)."
    if ($failed.Count -gt 0) {
        $status = 'Failed'
        $message = "$($failed.Count) RegRipper operation(s) failed. See registry summary JSON for details."
    }
    elseif ($warnings.Count -gt 0) {
        $status = 'Completed With Warnings'
        $message = "Registry hive processing completed with $($warnings.Count) warning(s). See registry summary JSON for details."
    }

    [pscustomobject]@{
        ModuleId = 'registry'
        Status = $status
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        PreparedHiveCount = $processed.Count
        Message = $message
    }
}

function Invoke-IbisHayabusaRuleUpdate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions
    )

    $hayabusa = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'hayabusa'
    if ($null -eq $hayabusa) {
        return [pscustomobject]@{
            ModuleId = 'hayabusa-rule-update'
            ToolId = 'hayabusa'
            Status = 'Failed'
            ExitCode = $null
            CommandLine = $null
            WorkingDirectory = $null
            StandardOutput = $null
            StandardError = $null
            Message = 'Hayabusa is not configured.'
        }
    }

    $hayabusaPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $hayabusa
    if (-not (Test-Path -LiteralPath $hayabusaPath -PathType Leaf)) {
        return [pscustomobject]@{
            ModuleId = 'hayabusa-rule-update'
            ToolId = $hayabusa.id
            Status = 'Failed'
            ExitCode = $null
            CommandLine = $null
            WorkingDirectory = Split-Path -Path $hayabusaPath -Parent
            StandardOutput = $null
            StandardError = $null
            Message = "Hayabusa is missing at: $hayabusaPath"
        }
    }

    $workingDirectory = Split-Path -Path $hayabusaPath -Parent
    $processResult = Invoke-IbisProcessCapture `
        -FilePath $hayabusaPath `
        -ArgumentList @('update-rules') `
        -WorkingDirectory $workingDirectory

    $status = 'Completed'
    $message = 'Hayabusa rules updated.'
    if ($processResult.ExitCode -ne 0) {
        $status = 'Failed'
        $message = "Hayabusa update-rules exited with code $($processResult.ExitCode)."
    }

    [pscustomobject]@{
        ModuleId = 'hayabusa-rule-update'
        ToolId = $hayabusa.id
        Status = $status
        ExitCode = $processResult.ExitCode
        CommandLine = $processResult.CommandLine
        WorkingDirectory = $workingDirectory
        StandardOutput = $processResult.StandardOutput
        StandardError = $processResult.StandardError
        Message = $message
    }
}

function Get-IbisAmcacheHivePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    [System.IO.Path]::Combine($SourceRoot, 'Windows\appcompat\Programs\Amcache.hve')
}

function Invoke-IbisAmcacheParser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$HivePath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputFileName
    )

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'zimmerman-amcacheparser'
    $outputPath = Join-Path $OutputDirectory $OutputFileName
    if ($null -eq $tool) {
        return [pscustomobject]@{
            ToolId = 'zimmerman-amcacheparser'
            HivePath = $HivePath
            OutputPath = $outputPath
            Status = 'Failed'
            ExitCode = $null
            Message = 'AmcacheParser is not configured.'
        }
    }

    $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        return [pscustomobject]@{
            ToolId = $tool.id
            HivePath = $HivePath
            OutputPath = $outputPath
            Status = 'Failed'
            ExitCode = $null
            Message = "AmcacheParser is missing at: $toolPath"
        }
    }

    if (-not (Test-Path -LiteralPath $HivePath -PathType Leaf)) {
        return [pscustomobject]@{
            ToolId = $tool.id
            HivePath = $HivePath
            OutputPath = $outputPath
            Status = 'Skipped'
            ExitCode = $null
            Message = 'Amcache hive was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $result = Invoke-IbisProcessCapture `
        -FilePath $toolPath `
        -ArgumentList @('-f', $HivePath, '--csv', $OutputDirectory, '--csvf', $OutputFileName) `
        -WorkingDirectory (Split-Path -Path $toolPath -Parent)

    if ($result.ExitCode -ne 0) {
        $result.StandardError | Out-File -LiteralPath ($outputPath + '.stderr.txt') -Encoding UTF8
    }

    $status = 'Completed'
    $message = 'AmcacheParser completed.'
    if ($result.ExitCode -ne 0) {
        $status = 'Failed'
        $message = "AmcacheParser exited with code $($result.ExitCode)."
    }
    elseif (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
        $status = 'Completed With Warnings'
        $message = 'AmcacheParser completed, but the expected CSV was not found.'
    }
    elseif ((Get-Item -LiteralPath $outputPath).Length -eq 0) {
        $status = 'Completed With Warnings'
        $message = 'AmcacheParser output CSV was created, but it is empty.'
    }

    [pscustomobject]@{
        ToolId = $tool.id
        HivePath = $HivePath
        OutputPath = $outputPath
        Status = $status
        ExitCode = $result.ExitCode
        CommandLine = $result.CommandLine
        Message = $message
    }
}

function Invoke-IbisAmcache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'Amcache'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $sourceHivePath = Get-IbisAmcacheHivePath -SourceRoot $SourceRoot

    $preparedHive = Invoke-IbisPrepareRegistryHiveFile `
        -ToolsRoot $ToolsRoot `
        -ToolDefinitions $ToolDefinitions `
        -SourceHivePath $sourceHivePath `
        -OutputRoot $OutputRoot `
        -Hostname $safeHost `
        -HiveName 'Amcache.hve' `
        -CacheGroup 'Amcache' `
        -CacheKey 'Amcache'

    if ([string]::IsNullOrWhiteSpace($preparedHive.PreparedHivePath)) {
        return [pscustomobject]@{
            ModuleId = 'amcache'
            Status = 'Skipped'
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            PreparedHive = $preparedHive
            Message = 'Amcache.hve was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $toolResults = @()
    $toolResults += Invoke-IbisAmcacheParser `
        -ToolsRoot $ToolsRoot `
        -ToolDefinitions $ToolDefinitions `
        -HivePath $preparedHive.PreparedHivePath `
        -OutputDirectory $outputDirectory `
        -OutputFileName (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'EZ-Amcache.csv')

    $toolResults += Invoke-IbisRegRipperHiveMode `
        -ToolsRoot $ToolsRoot `
        -ToolDefinitions $ToolDefinitions `
        -HivePath $preparedHive.PreparedHivePath `
        -Mode 'All' `
        -OutputPath (Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'RR-Amcache.txt'))

    $toolResults += Invoke-IbisRegRipperHiveMode `
        -ToolsRoot $ToolsRoot `
        -ToolDefinitions $ToolDefinitions `
        -HivePath $preparedHive.PreparedHivePath `
        -Mode 'Timeline' `
        -OutputPath (Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'RR-Amcache-TLN.txt'))

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Amcache.json')
    $payload = [pscustomobject]@{
        ModuleId = 'amcache'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        PreparedHive = $preparedHive
        ToolResults = $toolResults
    }
    $payload | ConvertTo-Json -Depth 10 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $failed = @($toolResults | Where-Object { $_.Status -eq 'Failed' })
    $warnings = @($toolResults | Where-Object { $_.Status -match 'Warnings' })
    if ($preparedHive.Status -match 'Warnings') {
        $warnings += $preparedHive
    }

    $status = 'Completed'
    $message = 'Amcache processing completed.'
    if ($failed.Count -gt 0) {
        $status = 'Failed'
        $message = "$($failed.Count) Amcache operation(s) failed. See Amcache summary JSON for details."
    }
    elseif ($warnings.Count -gt 0) {
        $status = 'Completed With Warnings'
        $message = "Amcache processing completed with $($warnings.Count) warning(s). See Amcache summary JSON for details."
    }

    [pscustomobject]@{
        ModuleId = 'amcache'
        Status = $status
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        PreparedHive = $preparedHive
        Message = $message
    }
}

function Invoke-IbisAppCompatCacheParser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$HivePath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputFileName
    )

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'zimmerman-appcompatcacheparser'
    $outputPath = Join-Path $OutputDirectory $OutputFileName
    if ($null -eq $tool) {
        return [pscustomobject]@{
            ToolId = 'zimmerman-appcompatcacheparser'
            HivePath = $HivePath
            OutputPath = $outputPath
            Status = 'Failed'
            ExitCode = $null
            Message = 'AppCompatCacheParser is not configured.'
        }
    }

    $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        return [pscustomobject]@{
            ToolId = $tool.id
            HivePath = $HivePath
            OutputPath = $outputPath
            Status = 'Failed'
            ExitCode = $null
            Message = "AppCompatCacheParser is missing at: $toolPath"
        }
    }

    if (-not (Test-Path -LiteralPath $HivePath -PathType Leaf)) {
        return [pscustomobject]@{
            ToolId = $tool.id
            HivePath = $HivePath
            OutputPath = $outputPath
            Status = 'Skipped'
            ExitCode = $null
            Message = 'SYSTEM hive was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $result = Invoke-IbisProcessCapture `
        -FilePath $toolPath `
        -ArgumentList @('-f', $HivePath, '--csv', $OutputDirectory, '--csvf', $OutputFileName) `
        -WorkingDirectory (Split-Path -Path $toolPath -Parent)

    if ($result.ExitCode -ne 0) {
        $result.StandardError | Out-File -LiteralPath ($outputPath + '.stderr.txt') -Encoding UTF8
    }

    $status = 'Completed'
    $message = 'AppCompatCacheParser completed.'
    if ($result.ExitCode -ne 0) {
        $status = 'Failed'
        $message = "AppCompatCacheParser exited with code $($result.ExitCode)."
    }
    elseif (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
        $status = 'Completed With Warnings'
        $message = 'AppCompatCacheParser completed, but the expected CSV was not found.'
    }
    elseif ((Get-Item -LiteralPath $outputPath).Length -eq 0) {
        $status = 'Completed With Warnings'
        $message = 'AppCompatCacheParser output CSV was created, but it is empty.'
    }

    [pscustomobject]@{
        ToolId = $tool.id
        HivePath = $HivePath
        OutputPath = $outputPath
        Status = $status
        ExitCode = $result.ExitCode
        CommandLine = $result.CommandLine
        Message = $message
    }
}

function Invoke-IbisAppCompatCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'AppCompatCache-ShimCache'
    $workingsDirectory = Join-Path $outputDirectory '_Working'

    $preparedHive = Invoke-IbisPrepareRegistryHive `
        -ToolsRoot $ToolsRoot `
        -ToolDefinitions $ToolDefinitions `
        -SourceRoot $SourceRoot `
        -OutputRoot $OutputRoot `
        -Hostname $safeHost `
        -HiveName 'SYSTEM'

    if ([string]::IsNullOrWhiteSpace($preparedHive.PreparedHivePath)) {
        return [pscustomobject]@{
            ModuleId = 'appcompatcache'
            Status = 'Skipped'
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            PreparedHive = $preparedHive
            Message = 'SYSTEM hive was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $toolResult = Invoke-IbisAppCompatCacheParser `
        -ToolsRoot $ToolsRoot `
        -ToolDefinitions $ToolDefinitions `
        -HivePath $preparedHive.PreparedHivePath `
        -OutputDirectory $outputDirectory `
        -OutputFileName (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'EZ-AppCompatCacheParser-Output.csv')

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'AppCompatCache-ShimCache.json')
    $payload = [pscustomobject]@{
        ModuleId = 'appcompatcache'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        PreparedHive = $preparedHive
        ToolResults = @($toolResult)
    }
    $payload | ConvertTo-Json -Depth 10 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'AppCompatCache/ShimCache processing completed.'
    if ($toolResult.Status -eq 'Failed') {
        $status = 'Failed'
        $message = 'AppCompatCacheParser failed. See AppCompatCache/ShimCache summary JSON for details.'
    }
    elseif ($toolResult.Status -match 'Warnings' -or $preparedHive.Status -match 'Warnings') {
        $status = 'Completed With Warnings'
        $message = 'AppCompatCache/ShimCache processing completed with warning(s). See summary JSON for details.'
    }

    [pscustomobject]@{
        ModuleId = 'appcompatcache'
        Status = $status
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        PreparedHive = $preparedHive
        Message = $message
    }
}

function Get-IbisPrefetchPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    [System.IO.Path]::Combine($SourceRoot, 'Windows\Prefetch')
}

function Rename-IbisPrefetchOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        return @()
    }

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $renamed = @()
    $files = @(Get-ChildItem -LiteralPath $OutputDirectory -File -Force -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        $regexMatch = [regex]::Match($file.Name, '^\d+_(PECmd_.+)$')
        if ($regexMatch.Success) {
            if ([string]::IsNullOrWhiteSpace($safeHost)) {
                $newName = $regexMatch.Groups[1].Value
            }
            else {
                $newName = '{0}_{1}' -f $safeHost, $regexMatch.Groups[1].Value
            }
            $newPath = Join-Path $file.DirectoryName $newName
            Move-Item -LiteralPath $file.FullName -Destination $newPath -Force
            $renamed += [pscustomobject]@{
                OriginalPath = $file.FullName
                NewPath = $newPath
            }
        }
    }

    $renamed
}

function Invoke-IbisPrefetch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $sourceDirectory = Get-IbisPrefetchPath -SourceRoot $SourceRoot
    $outputDirectory = Join-Path $hostOutputRoot 'Prefetch'
    $workingsDirectory = Join-Path $outputDirectory '_Working'

    if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
        return [pscustomobject]@{
            ModuleId = 'prefetch'
            Status = 'Skipped'
            SourceDirectory = $sourceDirectory
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            Message = 'Prefetch folder was not found. This can be normal on Windows Server systems.'
        }
    }

    $prefetchFiles = @(Get-ChildItem -LiteralPath $sourceDirectory -Filter '*.pf' -File -Force -ErrorAction SilentlyContinue)
    if ($prefetchFiles.Count -eq 0) {
        return [pscustomobject]@{
            ModuleId = 'prefetch'
            Status = 'Skipped'
            SourceDirectory = $sourceDirectory
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            Message = 'Prefetch folder was found, but it did not contain any .pf files.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'zimmerman-pecmd'
    $toolResult = $null
    $renamedOutputs = @()
    if ($null -eq $tool) {
        $toolResult = [pscustomobject]@{
            ToolId = 'zimmerman-pecmd'
            SourceDirectory = $sourceDirectory
            OutputDirectory = $outputDirectory
            Status = 'Failed'
            ExitCode = $null
            Message = 'PECmd is not configured.'
        }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResult = [pscustomobject]@{
                ToolId = $tool.id
                SourceDirectory = $sourceDirectory
                OutputDirectory = $outputDirectory
                Status = 'Failed'
                ExitCode = $null
                Message = "PECmd is missing at: $toolPath"
            }
        }
        else {
            $processResult = Invoke-IbisProcessCapture `
                -FilePath $toolPath `
                -ArgumentList @('-d', $sourceDirectory, '--csv', $outputDirectory) `
                -WorkingDirectory (Split-Path -Path $toolPath -Parent)

            if ($processResult.ExitCode -ne 0) {
                $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'PECmd.stderr.txt')) -Encoding UTF8
            }

            $renamedOutputs = @(Rename-IbisPrefetchOutput -OutputDirectory $outputDirectory -Hostname $safeHost)
            $outputFiles = @(Get-ChildItem -LiteralPath $outputDirectory -File -Force -ErrorAction SilentlyContinue)

            $status = 'Completed'
            $message = 'PECmd completed.'
            if ($processResult.ExitCode -ne 0) {
                $status = 'Failed'
                $message = "PECmd exited with code $($processResult.ExitCode)."
            }
            elseif ($outputFiles.Count -eq 0) {
                $status = 'Completed With Warnings'
                $message = 'PECmd completed, but no output files were found.'
            }

            $toolResult = [pscustomobject]@{
                ToolId = $tool.id
                SourceDirectory = $sourceDirectory
                OutputDirectory = $outputDirectory
                Status = $status
                ExitCode = $processResult.ExitCode
                CommandLine = $processResult.CommandLine
                Message = $message
            }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Prefetch.json')
    $payload = [pscustomobject]@{
        ModuleId = 'prefetch'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        SourceDirectory = $sourceDirectory
        SourcePrefetchFileCount = $prefetchFiles.Count
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        ToolResults = @($toolResult)
        RenamedOutputs = @($renamedOutputs)
    }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'Prefetch processing completed.'
    if ($toolResult.Status -eq 'Failed') {
        $status = 'Failed'
        $message = 'PECmd failed. See Prefetch summary JSON for details.'
    }
    elseif ($toolResult.Status -match 'Warnings') {
        $status = 'Completed With Warnings'
        $message = 'Prefetch processing completed with warning(s). See summary JSON for details.'
    }

    [pscustomobject]@{
        ModuleId = 'prefetch'
        Status = $status
        SourceDirectory = $sourceDirectory
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        SourcePrefetchFileCount = $prefetchFiles.Count
        Message = $message
    }
}

function Find-IbisNtfsArtifactPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('$MFT', '$J')]
        [string]$ArtifactName
    )

    $candidatePaths = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $addCandidate = {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path)) {
            return
        }

        $key = $Path.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $candidatePaths.Add($Path)
        }
    }

    $current = $SourceRoot
    for ($i = 0; $i -lt 7; $i++) {
        if ([string]::IsNullOrWhiteSpace($current)) {
            break
        }

        & $addCandidate ([System.IO.Path]::Combine($current, $ArtifactName))
        & $addCandidate ([System.IO.Path]::Combine($current, 'uploads\ntfs\%5C%5C.%5CC%3A', $ArtifactName))
        & $addCandidate ([System.IO.Path]::Combine($current, 'ntfs\%5C%5C.%5CC%3A', $ArtifactName))

        foreach ($ntfsRoot in @(
            [System.IO.Path]::Combine($current, 'uploads\ntfs'),
            [System.IO.Path]::Combine($current, 'ntfs')
        )) {
            if (Test-Path -LiteralPath $ntfsRoot -PathType Container) {
                foreach ($deviceDirectory in @(Get-ChildItem -LiteralPath $ntfsRoot -Directory -Force -ErrorAction SilentlyContinue)) {
                    & $addCandidate (Join-Path $deviceDirectory.FullName $ArtifactName)
                }
            }
        }

        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) {
            break
        }
        $current = $parent
    }

    foreach ($candidatePath in $candidatePaths) {
        if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
            return $candidatePath
        }
    }

    $null
}

function Test-IbisNtfsSpecialFilePath {
    [CmdletBinding()]
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        return $true
    }

    $usnStreamMatch = [regex]::Match($Path, '^(?<BasePath>.*\\\$UsnJrnl):\$J$')
    if ($usnStreamMatch.Success) {
        return (Test-Path -LiteralPath $usnStreamMatch.Groups['BasePath'].Value -PathType Leaf)
    }

    $false
}

function Find-IbisUsnJournalPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    $candidatePaths = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    $addCandidate = {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path)) {
            return
        }

        $key = $Path.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $candidatePaths.Add($Path)
        }
    }

    try {
        $sourceRootPath = [System.IO.Path]::GetFullPath($SourceRoot)
        $rootPath = [System.IO.Path]::GetPathRoot($sourceRootPath)
        if ([string]::Equals($sourceRootPath.TrimEnd('\', '/'), $rootPath.TrimEnd('\', '/'), [System.StringComparison]::OrdinalIgnoreCase)) {
            $drive = [System.IO.Path]::GetPathRoot($sourceRootPath).TrimEnd('\', '/')
            if (-not [string]::Equals($drive, 'C:', [System.StringComparison]::OrdinalIgnoreCase)) {
                & $addCandidate ([System.IO.Path]::Combine($SourceRoot, '$Extend\$UsnJrnl:$J'))
            }
        }
    }
    catch {
    }

    $current = $SourceRoot
    for ($i = 0; $i -lt 7; $i++) {
        if ([string]::IsNullOrWhiteSpace($current)) {
            break
        }

        & $addCandidate ([System.IO.Path]::Combine($current, 'uploads\ntfs\%5C%5C.%5CC%3A\$Extend\$UsnJrnl%3A$J'))
        & $addCandidate ([System.IO.Path]::Combine($current, 'ntfs\%5C%5C.%5CC%3A\$Extend\$UsnJrnl%3A$J'))

        foreach ($ntfsRoot in @(
            [System.IO.Path]::Combine($current, 'uploads\ntfs'),
            [System.IO.Path]::Combine($current, 'ntfs')
        )) {
            if (Test-Path -LiteralPath $ntfsRoot -PathType Container) {
                foreach ($deviceDirectory in @(Get-ChildItem -LiteralPath $ntfsRoot -Directory -Force -ErrorAction SilentlyContinue)) {
                    & $addCandidate ([System.IO.Path]::Combine($deviceDirectory.FullName, '$Extend\$UsnJrnl%3A$J'))
                    & $addCandidate ([System.IO.Path]::Combine($deviceDirectory.FullName, '$Extend\$UsnJrnl:$J'))
                }
            }
        }

        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) {
            break
        }
        $current = $parent
    }

    foreach ($candidatePath in $candidatePaths) {
        if (Test-IbisNtfsSpecialFilePath -Path $candidatePath) {
            return $candidatePath
        }
    }

    $null
}

function Invoke-IbisMftECmdArtifact {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$ArtifactPath,

        [Parameter(Mandatory = $true)]
        [string]$ArtifactName,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputFileName,

        [string]$MftPath,

        [Parameter(Mandatory = $true)]
        [string]$WorkingsDirectory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname
    )

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'zimmerman-mftecmd'
    if ($null -eq $tool) {
        return [pscustomobject]@{
            ToolId = 'zimmerman-mftecmd'
            ArtifactName = $ArtifactName
            SourcePath = $ArtifactPath
            MftPath = $MftPath
            OutputDirectory = $OutputDirectory
            OutputPath = Join-Path $OutputDirectory $OutputFileName
            Status = 'Failed'
            ExitCode = $null
            Message = 'MFTECmd is not configured.'
        }
    }

    $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        return [pscustomobject]@{
            ToolId = $tool.id
            ArtifactName = $ArtifactName
            SourcePath = $ArtifactPath
            MftPath = $MftPath
            OutputDirectory = $OutputDirectory
            OutputPath = Join-Path $OutputDirectory $OutputFileName
            Status = 'Failed'
            ExitCode = $null
            Message = "MFTECmd is missing at: $toolPath"
        }
    }

    $argumentList = @('-f', $ArtifactPath)
    if (-not [string]::IsNullOrWhiteSpace($MftPath)) {
        $argumentList += @('-m', $MftPath)
    }
    $argumentList += @('--csv', $OutputDirectory, '--csvf', $OutputFileName)

    $processResult = Invoke-IbisProcessCapture `
        -FilePath $toolPath `
        -ArgumentList $argumentList `
        -WorkingDirectory (Split-Path -Path $toolPath -Parent)

    if ($processResult.ExitCode -ne 0) {
        $stderrName = Format-IbisHostPrefixedValue -Hostname $Hostname -Format 'MFTECmd-{0}.stderr.txt' -ArgumentList @((ConvertTo-IbisSafeFileName -Value $ArtifactName))
        $processResult.StandardError | Out-File -LiteralPath (Join-Path $WorkingsDirectory $stderrName) -Encoding UTF8
    }

    $outputPath = Join-Path $OutputDirectory $OutputFileName
    $status = 'Completed'
    $message = 'MFTECmd completed.'
    if ($processResult.ExitCode -ne 0) {
        $status = 'Failed'
        $message = "MFTECmd exited with code $($processResult.ExitCode)."
    }
    elseif (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
        $status = 'Completed With Warnings'
        $message = 'MFTECmd completed, but the expected CSV output file was not found.'
    }

    [pscustomobject]@{
        ToolId = $tool.id
        ArtifactName = $ArtifactName
        SourcePath = $ArtifactPath
        MftPath = $MftPath
        OutputDirectory = $OutputDirectory
        OutputPath = $outputPath
        Status = $status
        ExitCode = $processResult.ExitCode
        CommandLine = $processResult.CommandLine
        Message = $message
    }
}

function Invoke-IbisNtfsMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'NTFS-Metadata'
    $workingsDirectory = Join-Path $outputDirectory '_Working'

    $mftPath = Find-IbisNtfsArtifactPath -SourceRoot $SourceRoot -ArtifactName '$MFT'
    $usnJournalPath = Find-IbisUsnJournalPath -SourceRoot $SourceRoot
    $locatedArtifacts = @(
        [pscustomobject]@{
            Name = '$MFT'
            SourcePath = $mftPath
            MftPath = $null
            OutputFileName = (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'MFTECmd-MFT-Output.csv')
            Found = -not [string]::IsNullOrWhiteSpace($mftPath)
            ReadyToProcess = -not [string]::IsNullOrWhiteSpace($mftPath)
            Message = if ([string]::IsNullOrWhiteSpace($mftPath)) { '$MFT was not found.' } else { '$MFT was found.' }
        },
        [pscustomobject]@{
            Name = '$UsnJrnl:$J'
            SourcePath = $usnJournalPath
            MftPath = $mftPath
            OutputFileName = (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'MFTECmd-UsnJrnl-J-Output.csv')
            Found = -not [string]::IsNullOrWhiteSpace($usnJournalPath)
            ReadyToProcess = (-not [string]::IsNullOrWhiteSpace($usnJournalPath) -and -not [string]::IsNullOrWhiteSpace($mftPath))
            Message = if ([string]::IsNullOrWhiteSpace($usnJournalPath)) {
                'USN Journal $J was not found.'
            }
            elseif ([string]::IsNullOrWhiteSpace($mftPath)) {
                'USN Journal $J was found, but $MFT was not found; MFTECmd requires $MFT for USN Journal processing.'
            }
            else {
                'USN Journal $J and $MFT were found.'
            }
        }
    )

    if (@($locatedArtifacts | Where-Object { $_.ReadyToProcess }).Count -eq 0) {
        return [pscustomobject]@{
            ModuleId = 'ntfs-metadata'
            Status = 'Skipped'
            SourceRoot = $SourceRoot
            LocatedArtifacts = $locatedArtifacts
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            Message = 'No processable NTFS metadata files were found. Checked the source root and nearby Velociraptor ntfs upload folders.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $toolResults = @()
    foreach ($artifact in @($locatedArtifacts | Where-Object { $_.ReadyToProcess })) {
        $toolResults += Invoke-IbisMftECmdArtifact `
            -ToolsRoot $ToolsRoot `
            -ToolDefinitions $ToolDefinitions `
            -ArtifactPath $artifact.SourcePath `
            -ArtifactName $artifact.Name `
            -OutputDirectory $outputDirectory `
            -OutputFileName $artifact.OutputFileName `
            -MftPath $artifact.MftPath `
            -WorkingsDirectory $workingsDirectory `
            -Hostname $safeHost
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'NTFS-Metadata.json')
    $payload = [pscustomobject]@{
        ModuleId = 'ntfs-metadata'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        LocatedArtifacts = $locatedArtifacts
        ToolResults = @($toolResults)
    }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'NTFS metadata processing completed.'
    if (@($toolResults | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0) {
        $status = 'Failed'
        $message = 'MFTECmd failed for one or more NTFS metadata files. See summary JSON for details.'
    }
    elseif (@($toolResults | Where-Object { $_.Status -match 'Warnings' }).Count -gt 0) {
        $status = 'Completed With Warnings'
        $message = 'NTFS metadata processing completed with warning(s). See summary JSON for details.'
    }

    [pscustomobject]@{
        ModuleId = 'ntfs-metadata'
        Status = $status
        SourceRoot = $SourceRoot
        LocatedArtifacts = $locatedArtifacts
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        Message = $message
    }
}

function Get-IbisSrumDatabasePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    [System.IO.Path]::Combine($SourceRoot, 'Windows\System32\sru\SRUDB.dat')
}

function Invoke-IbisSrum {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $srumPath = Get-IbisSrumDatabasePath -SourceRoot $SourceRoot
    $outputDirectory = Join-Path $hostOutputRoot 'SRUM'
    $workingsDirectory = Join-Path $outputDirectory '_Working'

    $preparedSoftware = Invoke-IbisPrepareRegistryHive `
        -ToolsRoot $ToolsRoot `
        -ToolDefinitions $ToolDefinitions `
        -SourceRoot $SourceRoot `
        -OutputRoot $OutputRoot `
        -Hostname $safeHost `
        -HiveName 'SOFTWARE'

    if (-not (Test-Path -LiteralPath $srumPath -PathType Leaf) -or [string]::IsNullOrWhiteSpace($preparedSoftware.PreparedHivePath)) {
        $missing = @()
        if (-not (Test-Path -LiteralPath $srumPath -PathType Leaf)) { $missing += 'SRUDB.dat' }
        if ([string]::IsNullOrWhiteSpace($preparedSoftware.PreparedHivePath)) { $missing += 'SOFTWARE hive' }
        return [pscustomobject]@{
            ModuleId = 'srum'
            Status = 'Skipped'
            SourcePath = $srumPath
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            PreparedHive = $preparedSoftware
            Message = "Unable to process SRUM because required source item(s) were missing: $($missing -join ', ')."
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'zimmerman-srumecmd'
    $renamedOutputs = @()
    if ($null -eq $tool) {
        $toolResult = [pscustomobject]@{
            ToolId = 'zimmerman-srumecmd'
            SourcePath = $srumPath
            OutputDirectory = $outputDirectory
            Status = 'Failed'
            ExitCode = $null
            Message = 'SrumECmd is not configured.'
        }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResult = [pscustomobject]@{
                ToolId = $tool.id
                SourcePath = $srumPath
                OutputDirectory = $outputDirectory
                Status = 'Failed'
                ExitCode = $null
                Message = "SrumECmd is missing at: $toolPath"
            }
        }
        else {
            $processResult = Invoke-IbisProcessCapture `
                -FilePath $toolPath `
                -ArgumentList @('-f', $srumPath, '-r', $preparedSoftware.PreparedHivePath, '--csv', $outputDirectory) `
                -WorkingDirectory (Split-Path -Path $toolPath -Parent)

            if ($processResult.ExitCode -ne 0) {
                $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'SrumECmd.stderr.txt')) -Encoding UTF8
            }

            $renamedOutputs = @(Rename-IbisSrumECmdOutput -OutputDirectory $outputDirectory -Hostname $safeHost)
            $outputFiles = @(Get-ChildItem -LiteralPath $outputDirectory -File -Force -ErrorAction SilentlyContinue)
            $status = 'Completed'
            $message = 'SrumECmd completed.'
            if ($processResult.ExitCode -ne 0) {
                $status = 'Failed'
                $message = "SrumECmd exited with code $($processResult.ExitCode)."
            }
            elseif ($outputFiles.Count -eq 0) {
                $status = 'Completed With Warnings'
                $message = 'SrumECmd completed, but no output files were found.'
            }

            $toolResult = [pscustomobject]@{
                ToolId = $tool.id
                SourcePath = $srumPath
                OutputDirectory = $outputDirectory
                Status = $status
                ExitCode = $processResult.ExitCode
                CommandLine = $processResult.CommandLine
                RenamedOutputs = @($renamedOutputs)
                Message = $message
            }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'SRUM.json')
    $payload = [pscustomobject]@{
        ModuleId = 'srum'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        SourcePath = $srumPath
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        PreparedHive = $preparedSoftware
        ToolResults = @($toolResult)
        RenamedOutputs = @($renamedOutputs)
    }
    $payload | ConvertTo-Json -Depth 10 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'SRUM processing completed.'
    if ($toolResult.Status -eq 'Failed') {
        $status = 'Failed'
        $message = 'SrumECmd failed. See SRUM summary JSON for details.'
    }
    elseif ($toolResult.Status -match 'Warnings' -or $preparedSoftware.Status -match 'Warnings') {
        $status = 'Completed With Warnings'
        $message = 'SRUM processing completed with warning(s). See summary JSON for details.'
    }

    [pscustomobject]@{
        ModuleId = 'srum'
        Status = $status
        SourcePath = $srumPath
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        PreparedHive = $preparedSoftware
        Message = $message
    }
}

function Rename-IbisSrumECmdOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        return @()
    }

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $renamed = @()
    $files = @(Get-ChildItem -LiteralPath $OutputDirectory -File -Force -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        $regexMatch = [regex]::Match($file.Name, '^\d+_SrumECmd_(.+)$')
        if ($regexMatch.Success) {
            $newName = Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'SrumECmd-{0}' -ArgumentList @($regexMatch.Groups[1].Value)
            $newPath = Join-Path $file.DirectoryName $newName
            Move-Item -LiteralPath $file.FullName -Destination $newPath -Force
            $renamed += [pscustomobject]@{
                OriginalPath = $file.FullName
                NewPath = $newPath
            }
        }
    }

    $renamed
}

function Get-IbisUserProfile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    $usersRoot = Join-Path $SourceRoot 'Users'
    if (-not (Test-Path -LiteralPath $usersRoot -PathType Container)) {
        return @()
    }

    Get-ChildItem -LiteralPath $usersRoot -Directory -Force -ErrorAction SilentlyContinue | ForEach-Object {
        [pscustomobject]@{
            UserName = $_.Name
            ProfilePath = $_.FullName
            NtUserPath = Join-Path $_.FullName 'NTUSER.dat'
            UsrClassPath = Join-Path $_.FullName 'AppData\Local\Microsoft\Windows\UsrClass.dat'
            RecentPath = Join-Path $_.FullName 'AppData\Roaming\Microsoft\Windows\Recent'
            PSReadLinePath = Join-Path $_.FullName 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine'
        }
    }
}

function Invoke-IbisUserDirectoryTool {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$ToolId,

        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList,

        [Parameter(Mandatory = $true)]
        [string]$Description,

        [string]$Hostname = 'HOST',

        [string]$UserName = 'User'
    )

    if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
        return [pscustomobject]@{
            ToolId = $ToolId
            Description = $Description
            SourceDirectory = $SourceDirectory
            OutputDirectory = $OutputDirectory
            Status = 'Skipped'
            ExitCode = $null
            Message = 'Source directory was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id $ToolId
    if ($null -eq $tool) {
        return [pscustomobject]@{
            ToolId = $ToolId
            Description = $Description
            SourceDirectory = $SourceDirectory
            OutputDirectory = $OutputDirectory
            Status = 'Failed'
            ExitCode = $null
            Message = "$Description tool is not configured."
        }
    }

    $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        return [pscustomobject]@{
            ToolId = $tool.id
            Description = $Description
            SourceDirectory = $SourceDirectory
            OutputDirectory = $OutputDirectory
            Status = 'Failed'
            ExitCode = $null
            Message = "$Description tool is missing at: $toolPath"
        }
    }

    $beforeOutputFiles = @(Get-ChildItem -LiteralPath $OutputDirectory -File -Recurse -Force -ErrorAction SilentlyContinue)
    $result = Invoke-IbisProcessCapture `
        -FilePath $toolPath `
        -ArgumentList $ArgumentList `
        -WorkingDirectory (Split-Path -Path $toolPath -Parent)
    if ($result.ExitCode -ne 0) {
        $stderrPath = Join-Path $OutputDirectory (($Description -replace '[\\/:*?"<>| ]', '_') + '.stderr.txt')
        $result.StandardError | Out-File -LiteralPath $stderrPath -Encoding UTF8
    }

    $renamedOutputs = @(Rename-IbisUserArtifactToolOutput -OutputDirectory $OutputDirectory -Hostname $Hostname -UserName $UserName -ToolName $Description)
    $afterOutputFiles = @(Get-ChildItem -LiteralPath $OutputDirectory -File -Recurse -Force -ErrorAction SilentlyContinue)
    $status = 'Completed'
    $message = "$Description completed."
    if ($result.ExitCode -ne 0) {
        $status = 'Failed'
        $message = "$Description exited with code $($result.ExitCode)."
    }
    elseif ($afterOutputFiles.Count -le $beforeOutputFiles.Count) {
        $status = 'Completed With Warnings'
        $message = "$Description completed, but no new output files were found."
    }

    [pscustomobject]@{
        ToolId = $tool.id
        Description = $Description
        SourceDirectory = $SourceDirectory
        OutputDirectory = $OutputDirectory
        Status = $status
        ExitCode = $result.ExitCode
        CommandLine = $result.CommandLine
        RenamedOutputs = $renamedOutputs
        Message = $message
    }
}

function Rename-IbisUserArtifactToolOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [string]$UserName,

        [Parameter(Mandatory = $true)]
        [string]$ToolName
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        return @()
    }

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $safeUserName = ConvertTo-IbisSafeFileName -Value $UserName -DefaultValue 'User'
    $safeToolName = ConvertTo-IbisSafeFileName -Value $ToolName -DefaultValue 'Tool'
    $renamed = @()

    $children = @(Get-ChildItem -LiteralPath $OutputDirectory -Force -ErrorAction SilentlyContinue)
    foreach ($child in $children) {
        $newName = $null
        if ($safeToolName -eq 'JLECmd' -and $child.Name -match '^\d+_(.+)$') {
            $newName = Format-IbisHostPrefixedValue -Hostname $safeHost -Format '{0}-{1}-{2}' -ArgumentList @($safeUserName, $safeToolName, $Matches[1])
        }
        elseif ($safeToolName -eq 'LECmd' -and $child.Name -match '^\d+_LECmd_(.+)$') {
            $newName = Format-IbisHostPrefixedValue -Hostname $safeHost -Format '{0}-{1}-{2}' -ArgumentList @($safeUserName, $safeToolName, $Matches[1])
        }
        elseif ($safeToolName -eq 'SBECmd' -and $child.Name -match '^(.+?)_(NTUSER|UsrClass)(.*)$') {
            $newName = Format-IbisHostPrefixedValue -Hostname $safeHost -Format '{0}-{1}-{2}{3}' -ArgumentList @($safeUserName, $safeToolName, $Matches[2], $Matches[3])
        }

        if (-not [string]::IsNullOrWhiteSpace($newName) -and $newName -ne $child.Name) {
            $newPath = Join-Path $OutputDirectory $newName
            Move-Item -LiteralPath $child.FullName -Destination $newPath -Force
            $renamed += [pscustomobject]@{
                OriginalPath = $child.FullName
                NewPath = $newPath
            }
        }
    }

    $renamed
}

function Copy-IbisPSReadLineHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [string]$Hostname = 'HOST',

        [string]$UserName = 'User'
    )

    if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
        return [pscustomobject]@{
            SourceDirectory = $SourceDirectory
            OutputDirectory = $OutputDirectory
            Status = 'Skipped'
            CopiedItemCount = 0
            Message = 'PSReadLine source directory was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $safeUserName = ConvertTo-IbisSafeFileName -Value $UserName -DefaultValue 'User'
    $copiedItems = @()
    $sourceFiles = @(Get-ChildItem -LiteralPath $SourceDirectory -File -Force -ErrorAction SilentlyContinue)
    foreach ($sourceFile in $sourceFiles) {
        $destinationName = Format-IbisHostPrefixedValue -Hostname $safeHost -Format '{0}-PSReadLine-{1}' -ArgumentList @($safeUserName, $sourceFile.Name)
        $destinationPath = Join-Path $OutputDirectory $destinationName
        Copy-Item -LiteralPath $sourceFile.FullName -Destination $destinationPath -Force
        $copiedItems += Get-Item -LiteralPath $destinationPath
    }

    $sourceDirectories = @(Get-ChildItem -LiteralPath $SourceDirectory -Directory -Force -ErrorAction SilentlyContinue)
    foreach ($sourceChildDirectory in $sourceDirectories) {
        Copy-Item -LiteralPath $sourceChildDirectory.FullName -Destination $OutputDirectory -Recurse -Force
        $copiedItems += Get-ChildItem -LiteralPath (Join-Path $OutputDirectory $sourceChildDirectory.Name) -Recurse -Force -ErrorAction SilentlyContinue
    }

    $status = 'Completed'
    $message = 'PSReadLine history copied.'
    if ($copiedItems.Count -eq 0) {
        $status = 'Completed With Warnings'
        $message = 'PSReadLine source was copied, but no files were found in the destination.'
    }

    [pscustomobject]@{
        SourceDirectory = $SourceDirectory
        OutputDirectory = $OutputDirectory
        DestinationDirectory = $OutputDirectory
        Status = $status
        CopiedItemCount = $copiedItems.Count
        Message = $message
    }
}

function Invoke-IbisUserArtifacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'Users'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $profiles = @(Get-IbisUserProfile -SourceRoot $SourceRoot)
    if ($profiles.Count -eq 0) {
        return [pscustomobject]@{
            ModuleId = 'user-artifacts'
            Status = 'Skipped'
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            UserCount = 0
            Message = 'No user profile directories were found.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $userResults = @()
    foreach ($profile in $profiles) {
        $safeUserName = ConvertTo-IbisSafeFileName -Value $profile.UserName -DefaultValue 'User'
        $userOutputDirectory = Join-Path $outputDirectory $safeUserName
        if (-not (Test-Path -LiteralPath $userOutputDirectory)) {
            New-Item -ItemType Directory -Path $userOutputDirectory -Force | Out-Null
        }

        $preparedHives = @()
        $toolResults = @()

        if (Test-Path -LiteralPath $profile.NtUserPath -PathType Leaf) {
            $ntUser = Invoke-IbisPrepareRegistryHiveFile `
                -ToolsRoot $ToolsRoot `
                -ToolDefinitions $ToolDefinitions `
                -SourceHivePath $profile.NtUserPath `
                -OutputRoot $OutputRoot `
                -Hostname $safeHost `
                -HiveName 'NTUSER.dat' `
                -CacheGroup 'Users' `
                -CacheKey ($safeUserName + '-NTUSER')
            $preparedHives += $ntUser
            if (-not [string]::IsNullOrWhiteSpace($ntUser.PreparedHivePath)) {
                $toolResults += Invoke-IbisRegRipperHiveMode -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $ntUser.PreparedHivePath -Mode 'All' -OutputPath (Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-NTUSER.txt' -ArgumentList @($safeUserName)))
                $toolResults += Invoke-IbisRegRipperHiveMode -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $ntUser.PreparedHivePath -Mode 'Timeline' -OutputPath (Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-NTUSER-TLN.txt' -ArgumentList @($safeUserName)))
                try {
                    $toolResults += Invoke-IbisRegRipperPlugin -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $ntUser.PreparedHivePath -Plugin 'run' -OutputPath (Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-NTUSER-Run-AutoStart.txt' -ArgumentList @($safeUserName)))
                    $toolResults += Invoke-IbisRegRipperPlugin -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $ntUser.PreparedHivePath -Plugin 'userassist' -OutputPath (Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-NTUSER-UserAssist.txt' -ArgumentList @($safeUserName)))
                }
                catch {
                    $toolResults += [pscustomobject]@{
                        Plugin = 'run/userassist'
                        HivePath = $ntUser.PreparedHivePath
                        OutputPath = $userOutputDirectory
                        Status = 'Failed'
                        ExitCode = $null
                        Message = $_.Exception.Message
                    }
                }
            }
        }

        if (Test-Path -LiteralPath $profile.UsrClassPath -PathType Leaf) {
            $usrClass = Invoke-IbisPrepareRegistryHiveFile `
                -ToolsRoot $ToolsRoot `
                -ToolDefinitions $ToolDefinitions `
                -SourceHivePath $profile.UsrClassPath `
                -OutputRoot $OutputRoot `
                -Hostname $safeHost `
                -HiveName 'UsrClass.dat' `
                -CacheGroup 'Users' `
                -CacheKey ($safeUserName + '-UsrClass')
            $preparedHives += $usrClass
            if (-not [string]::IsNullOrWhiteSpace($usrClass.PreparedHivePath)) {
                $toolResults += Invoke-IbisRegRipperHiveMode -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $usrClass.PreparedHivePath -Mode 'All' -OutputPath (Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-UsrClass.txt' -ArgumentList @($safeUserName)))
                $toolResults += Invoke-IbisRegRipperHiveMode -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $usrClass.PreparedHivePath -Mode 'Timeline' -OutputPath (Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-UsrClass-TLN.txt' -ArgumentList @($safeUserName)))
            }
        }

        $jumpListOutput = Join-Path $userOutputDirectory 'JumpLists'
        $recentLnkOutput = Join-Path $userOutputDirectory 'RecentLNKs'
        $shellBagOutput = Join-Path $userOutputDirectory 'ShellBags'
        $toolResults += Invoke-IbisUserDirectoryTool -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -ToolId 'zimmerman-jlecmd' -SourceDirectory $profile.RecentPath -OutputDirectory $jumpListOutput -ArgumentList @('-d', $profile.RecentPath, '--all', '--csv', $jumpListOutput, '--html', $jumpListOutput, '-q', '--fd') -Description 'JLECmd' -Hostname $safeHost -UserName $safeUserName
        $toolResults += Invoke-IbisUserDirectoryTool -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -ToolId 'zimmerman-lecmd' -SourceDirectory $profile.RecentPath -OutputDirectory $recentLnkOutput -ArgumentList @('-d', $profile.RecentPath, '--all', '--csv', $recentLnkOutput, '-q') -Description 'LECmd' -Hostname $safeHost -UserName $safeUserName
        $toolResults += Invoke-IbisUserDirectoryTool -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -ToolId 'zimmerman-sbecmd' -SourceDirectory $profile.ProfilePath -OutputDirectory $shellBagOutput -ArgumentList @('-d', $profile.ProfilePath, '--csv', $shellBagOutput) -Description 'SBECmd' -Hostname $safeHost -UserName $safeUserName

        $psReadLineResult = Copy-IbisPSReadLineHistory -SourceDirectory $profile.PSReadLinePath -OutputDirectory (Join-Path $userOutputDirectory 'PSReadLine') -Hostname $safeHost -UserName $safeUserName

        $userResults += [pscustomobject]@{
            UserName = $profile.UserName
            ProfilePath = $profile.ProfilePath
            OutputDirectory = $userOutputDirectory
            PreparedHives = $preparedHives
            ToolResults = $toolResults
            PSReadLine = $psReadLineResult
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'User-Artifacts.json')
    $payload = [pscustomobject]@{
        ModuleId = 'user-artifacts'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        Users = $userResults
    }
    $payload | ConvertTo-Json -Depth 12 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $toolResultsAll = @($userResults | ForEach-Object { $_.ToolResults })
    $preparedAll = @($userResults | ForEach-Object { $_.PreparedHives })
    $failed = @($toolResultsAll | Where-Object { $_.Status -eq 'Failed' })
    $warnings = @($toolResultsAll | Where-Object { $_.Status -match 'Warnings' })
    $warnings += @($preparedAll | Where-Object { $_.Status -match 'Warnings' })
    $status = 'Completed'
    $message = "User artefact processing completed for $($profiles.Count) user profile(s)."
    if ($failed.Count -gt 0) {
        $status = 'Failed'
        $message = "$($failed.Count) user artefact operation(s) failed. See User Artifacts summary JSON for details."
    }
    elseif ($warnings.Count -gt 0) {
        $status = 'Completed With Warnings'
        $message = "User artefact processing completed with $($warnings.Count) warning(s). See summary JSON for details."
    }

    [pscustomobject]@{
        ModuleId = 'user-artifacts'
        Status = $status
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        UserCount = $profiles.Count
        Message = $message
    }
}

function Get-IbisEventLogPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    [System.IO.Path]::Combine($SourceRoot, 'Windows\System32\winevt\Logs')
}

function Invoke-IbisEvtxECmdEventLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $sourceDirectory = Get-IbisEventLogPath -SourceRoot $SourceRoot
    $outputDirectory = Join-Path $hostOutputRoot 'EventLogs'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $outputPath = Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'EvtxECmd-EventLogs-Output.csv')

    if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
        return [pscustomobject]@{
            ModuleId = 'eventlogs'
            Status = 'Skipped'
            SourceDirectory = $sourceDirectory
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            Message = 'Windows Event Log folder was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'zimmerman-evtxecmd'
    if ($null -eq $tool) {
        $toolResult = [pscustomobject]@{
            ToolId = 'zimmerman-evtxecmd'
            SourceDirectory = $sourceDirectory
            OutputPath = $outputPath
            Status = 'Failed'
            ExitCode = $null
            Message = 'EvtxECmd is not configured.'
        }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResult = [pscustomobject]@{
                ToolId = $tool.id
                SourceDirectory = $sourceDirectory
                OutputPath = $outputPath
                Status = 'Failed'
                ExitCode = $null
                Message = "EvtxECmd is missing at: $toolPath"
            }
        }
        else {
            $processResult = Invoke-IbisProcessCapture `
                -FilePath $toolPath `
                -ArgumentList @('-d', $sourceDirectory, '--csv', $outputDirectory, '--csvf', (Split-Path -Path $outputPath -Leaf)) `
                -WorkingDirectory (Split-Path -Path $toolPath -Parent)
            if ($processResult.ExitCode -ne 0) {
                $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'EvtxECmd.stderr.txt')) -Encoding UTF8
            }

            $status = 'Completed'
            $message = 'EvtxECmd completed.'
            if ($processResult.ExitCode -ne 0) {
                $status = 'Failed'
                $message = "EvtxECmd exited with code $($processResult.ExitCode)."
            }
            elseif (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
                $status = 'Completed With Warnings'
                $message = 'EvtxECmd completed, but the expected CSV was not found.'
            }

            $toolResult = [pscustomobject]@{
                ToolId = $tool.id
                SourceDirectory = $sourceDirectory
                OutputPath = $outputPath
                Status = $status
                ExitCode = $processResult.ExitCode
                CommandLine = $processResult.CommandLine
                Message = $message
            }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'EventLogs.json')
    $payload = [pscustomobject]@{
        ModuleId = 'eventlogs'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        SourceDirectory = $sourceDirectory
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        ToolResults = @($toolResult)
    }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'Windows Event Log processing completed.'
    if ($toolResult.Status -eq 'Failed') {
        $status = 'Failed'
        $message = 'EvtxECmd failed. See EventLogs summary JSON for details.'
    }
    elseif ($toolResult.Status -match 'Warnings') {
        $status = 'Completed With Warnings'
        $message = 'Windows Event Log processing completed with warning(s). See summary JSON for details.'
    }

    [pscustomobject]@{
        ModuleId = 'eventlogs'
        Status = $status
        SourceDirectory = $sourceDirectory
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        Message = $message
    }
}

function Get-IbisEvtxECmdCsvPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    Join-Path (Join-Path $hostOutputRoot 'EventLogs') (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'EvtxECmd-EventLogs-Output.csv')
}

function Get-IbisDuckDbEventLogQueryDefinition {
    [CmdletBinding()]
    param(
        [string]$ProjectRoot
    )

    if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
        $ProjectRoot = Split-Path -Path $PSScriptRoot -Parent
    }

    $queryRoot = Join-Path $ProjectRoot 'queries\eventlogs'
    @(
        [pscustomobject]@{
            Id = 'time-span'
            Name = 'Event log time span'
            QueryPath = Join-Path $queryRoot 'time-span.sql'
            OutputFileNameFormat = 'Event-Log-Time-Span-Info.csv'
        }
        [pscustomobject]@{
            Id = 'logons'
            Name = 'Event log user logons'
            QueryPath = Join-Path $queryRoot 'logons.sql'
            OutputFileNameFormat = 'Event-Log-User-Logons.csv'
        }
        [pscustomobject]@{
            Id = 'outbound-rdp'
            Name = 'Event log outbound RDP'
            QueryPath = Join-Path $queryRoot 'outbound-rdp.sql'
            OutputFileNameFormat = 'Event-Log-Outbound-RDP.csv'
        }
    )
}

function ConvertTo-IbisDuckDbSqlLiteral {
    [CmdletBinding()]
    param(
        [string]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    $Value.Replace("'", "''")
}

function Expand-IbisDuckDbSqlTemplate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemplatePath,

        [Parameter(Mandatory = $true)]
        [string]$InputCsvPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputCsvPath
    )

    $template = Get-Content -LiteralPath $TemplatePath -Raw
    $template = $template.Replace('{{INPUT_CSV}}', (ConvertTo-IbisDuckDbSqlLiteral -Value $InputCsvPath))
    $template.Replace('{{OUTPUT_CSV}}', (ConvertTo-IbisDuckDbSqlLiteral -Value $OutputCsvPath))
}

function Invoke-IbisDuckDbEventLogSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST',

        [string]$ProjectRoot
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'EventLogs'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $evtxCsvPath = Get-IbisEvtxECmdCsvPath -OutputRoot $OutputRoot -Hostname $safeHost

    if (-not (Test-Path -LiteralPath $evtxCsvPath -PathType Leaf)) {
        return [pscustomobject]@{
            ModuleId = 'duckdb-eventlogs'
            Status = 'Skipped'
            EvtxECmdCsvPath = $evtxCsvPath
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            Message = 'DuckDB event log summaries skipped because EvtxECmd CSV output was not available. Run Windows Event Logs first.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) { New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'duckdb'
    $queryDefinitions = @(Get-IbisDuckDbEventLogQueryDefinition -ProjectRoot $ProjectRoot)
    $toolResults = @()
    if ($null -eq $tool) {
        $toolResults += [pscustomobject]@{ ToolId = 'duckdb'; QueryId = 'all'; InputPath = $evtxCsvPath; Status = 'Failed'; ExitCode = $null; Message = 'DuckDB CLI is not configured.' }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResults += [pscustomobject]@{ ToolId = $tool.id; QueryId = 'all'; InputPath = $evtxCsvPath; Status = 'Failed'; ExitCode = $null; Message = "DuckDB CLI is missing at: $toolPath" }
        }
        else {
            foreach ($queryDefinition in $queryDefinitions) {
                $outputPath = Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix $queryDefinition.OutputFileNameFormat)
                $renderedSqlPath = Join-Path $workingsDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'DuckDB-{0}.sql' -ArgumentList @($queryDefinition.Id))
                if (-not (Test-Path -LiteralPath $queryDefinition.QueryPath -PathType Leaf)) {
                    $toolResults += [pscustomobject]@{ ToolId = $tool.id; QueryId = $queryDefinition.Id; QueryPath = $queryDefinition.QueryPath; OutputPath = $outputPath; RenderedSqlPath = $null; Status = 'Failed'; ExitCode = $null; Message = "DuckDB SQL template was not found: $($queryDefinition.QueryPath)" }
                    continue
                }

                $sql = Expand-IbisDuckDbSqlTemplate -TemplatePath $queryDefinition.QueryPath -InputCsvPath $evtxCsvPath -OutputCsvPath $outputPath
                $sql | Out-File -LiteralPath $renderedSqlPath -Encoding UTF8
                $processResult = Invoke-IbisProcessCapture -FilePath $toolPath -ArgumentList @('-c', $sql) -WorkingDirectory (Split-Path -Path $toolPath -Parent)
                if ($processResult.ExitCode -ne 0) {
                    $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'DuckDB-{0}.stderr.txt' -ArgumentList @($queryDefinition.Id))) -Encoding UTF8
                }

                $status = 'Completed'
                $message = "$($queryDefinition.Name) completed."
                if ($processResult.ExitCode -ne 0) {
                    $status = 'Failed'
                    $message = "$($queryDefinition.Name) exited with code $($processResult.ExitCode)."
                }
                elseif (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
                    $status = 'Completed With Warnings'
                    $message = "$($queryDefinition.Name) completed, but the expected CSV was not found."
                }

                $toolResults += [pscustomobject]@{
                    ToolId = $tool.id
                    QueryId = $queryDefinition.Id
                    QueryPath = $queryDefinition.QueryPath
                    InputPath = $evtxCsvPath
                    OutputPath = $outputPath
                    RenderedSqlPath = $renderedSqlPath
                    Status = $status
                    ExitCode = $processResult.ExitCode
                    CommandLine = $processResult.CommandLine
                    Message = $message
                }
            }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'DuckDB-EventLogs.json')
    $payload = [pscustomobject]@{
        ModuleId = 'duckdb-eventlogs'
        Created = (Get-Date).ToString('s')
        EvtxECmdCsvPath = $evtxCsvPath
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        QueryDefinitions = $queryDefinitions
        ToolResults = $toolResults
    }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $failed = @($toolResults | Where-Object { $_.Status -eq 'Failed' })
    $warnings = @($toolResults | Where-Object { $_.Status -match 'Warnings' })
    $status = 'Completed'
    $message = 'DuckDB event log summaries completed.'
    if ($failed.Count -gt 0) { $status = 'Failed'; $message = "$($failed.Count) DuckDB event log summarisation operation(s) failed. See summary JSON for details." }
    elseif ($warnings.Count -gt 0) { $status = 'Completed With Warnings'; $message = "DuckDB event log summaries completed with $($warnings.Count) warning(s). See summary JSON for details." }

    [pscustomobject]@{
        ModuleId = 'duckdb-eventlogs'
        Status = $status
        EvtxECmdCsvPath = $evtxCsvPath
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        Message = $message
    }
}

function Move-IbisExistingDirectoryToBackup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath,

        [Parameter(Mandatory = $true)]
        [string]$BackupRoot
    )

    if (-not (Test-Path -LiteralPath $DirectoryPath -PathType Container)) {
        return $null
    }

    if (-not (Test-Path -LiteralPath $BackupRoot)) {
        New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
    }

    $backupPath = Join-Path $BackupRoot ((Split-Path -Path $DirectoryPath -Leaf) + '-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    Move-Item -LiteralPath $DirectoryPath -Destination $backupPath -Force
    $backupPath
}

function Invoke-IbisHayabusaEventLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $sourceDirectory = Get-IbisEventLogPath -SourceRoot $SourceRoot
    $outputDirectory = Join-Path $hostOutputRoot 'EventLogs'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $hayabusaCsvPath = Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Hayabusa-EventLogs-Output.csv')
    $hayabusaJsonlPath = Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Hayabusa-EventLogs-SuperVerbose.jsonl')

    if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
        return [pscustomobject]@{
            ModuleId = 'hayabusa'
            Status = 'Skipped'
            SourceDirectory = $sourceDirectory
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            Message = 'Windows Event Log folder was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $toolResults = @()
    $hayabusa = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'hayabusa'
    if ($null -eq $hayabusa) {
        $toolResults += [pscustomobject]@{ ToolId = 'hayabusa'; OutputPath = $hayabusaJsonlPath; Status = 'Failed'; ExitCode = $null; Message = 'Hayabusa is not configured.' }
    }
    else {
        $hayabusaPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $hayabusa
        if (-not (Test-Path -LiteralPath $hayabusaPath -PathType Leaf)) {
            $toolResults += [pscustomobject]@{ ToolId = $hayabusa.id; OutputPath = $hayabusaJsonlPath; Status = 'Failed'; ExitCode = $null; Message = "Hayabusa is missing at: $hayabusaPath" }
        }
        else {
            $csvResult = Invoke-IbisProcessCapture -FilePath $hayabusaPath -ArgumentList @('csv-timeline', '-d', $sourceDirectory, '--ISO-8601', '-o', $hayabusaCsvPath, '-w') -WorkingDirectory (Split-Path -Path $hayabusaPath -Parent)
            $csvStatus = 'Completed'
            $csvMessage = 'Hayabusa CSV timeline completed.'
            if ($csvResult.ExitCode -ne 0) { $csvStatus = 'Failed'; $csvMessage = "Hayabusa CSV timeline exited with code $($csvResult.ExitCode)." }
            elseif (-not (Test-Path -LiteralPath $hayabusaCsvPath -PathType Leaf)) { $csvStatus = 'Completed With Warnings'; $csvMessage = 'Hayabusa CSV timeline completed, but the expected CSV was not found.' }
            $toolResults += [pscustomobject]@{ ToolId = $hayabusa.id; Mode = 'csv-timeline'; OutputPath = $hayabusaCsvPath; Status = $csvStatus; ExitCode = $csvResult.ExitCode; CommandLine = $csvResult.CommandLine; Message = $csvMessage }

            $jsonResult = Invoke-IbisProcessCapture -FilePath $hayabusaPath -ArgumentList @('json-timeline', '-d', $sourceDirectory, '-x', '-U', '-a', '-A', '-w', '-p', 'super-verbose', '-L', '-o', $hayabusaJsonlPath) -WorkingDirectory (Split-Path -Path $hayabusaPath -Parent)
            $jsonStatus = 'Completed'
            $jsonMessage = 'Hayabusa JSONL timeline completed.'
            if ($jsonResult.ExitCode -ne 0) { $jsonStatus = 'Failed'; $jsonMessage = "Hayabusa JSONL timeline exited with code $($jsonResult.ExitCode)." }
            elseif (-not (Test-Path -LiteralPath $hayabusaJsonlPath -PathType Leaf)) { $jsonStatus = 'Completed With Warnings'; $jsonMessage = 'Hayabusa JSONL timeline completed, but the expected JSONL was not found.' }
            $toolResults += [pscustomobject]@{ ToolId = $hayabusa.id; Mode = 'json-timeline'; OutputPath = $hayabusaJsonlPath; Status = $jsonStatus; ExitCode = $jsonResult.ExitCode; CommandLine = $jsonResult.CommandLine; Message = $jsonMessage }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Hayabusa.json')
    $payload = [pscustomobject]@{
        ModuleId = 'hayabusa'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        SourceDirectory = $sourceDirectory
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        ToolResults = $toolResults
    }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $failed = @($toolResults | Where-Object { $_.Status -eq 'Failed' })
    $warnings = @($toolResults | Where-Object { $_.Status -match 'Warnings' })
    $status = 'Completed'
    $message = 'Hayabusa processing completed.'
    if ($failed.Count -gt 0) { $status = 'Failed'; $message = "$($failed.Count) Hayabusa operation(s) failed. See summary JSON for details." }
    elseif ($warnings.Count -gt 0) { $status = 'Completed With Warnings'; $message = "Hayabusa processing completed with $($warnings.Count) warning(s). See summary JSON for details." }

    [pscustomobject]@{
        ModuleId = 'hayabusa'
        Status = $status
        SourceDirectory = $sourceDirectory
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        JsonPath = $summaryPath
        Message = $message
    }
}

function Get-IbisHayabusaJsonlPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    Join-Path (Join-Path $hostOutputRoot 'EventLogs') (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Hayabusa-EventLogs-SuperVerbose.jsonl')
}

function Invoke-IbisTakajoEventLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'EventLogs'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $hayabusaJsonlPath = Get-IbisHayabusaJsonlPath -OutputRoot $OutputRoot -Hostname $safeHost
    $takajoOutputDirectory = Join-Path $outputDirectory 'Takajo'

    if (-not (Test-Path -LiteralPath $hayabusaJsonlPath -PathType Leaf)) {
        return [pscustomobject]@{
            ModuleId = 'takajo'
            Status = 'Skipped'
            HayabusaJsonlPath = $hayabusaJsonlPath
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $takajoOutputDirectory
            JsonPath = $null
            Message = 'Takajo skipped because Hayabusa JSONL output was not available. Run Hayabusa first.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) { New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null }

    $toolResults = @()
    $takajo = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'takajo'
    if ($null -eq $takajo) {
        $toolResults += [pscustomobject]@{ ToolId = 'takajo'; Mode = 'all'; OutputPath = $takajoOutputDirectory; Status = 'Failed'; ExitCode = $null; Message = 'Takajo is not configured.' }
    }
    else {
        $takajoPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $takajo
        if (-not (Test-Path -LiteralPath $takajoPath -PathType Leaf)) {
            $toolResults += [pscustomobject]@{ ToolId = $takajo.id; Mode = 'all'; OutputPath = $takajoOutputDirectory; Status = 'Failed'; ExitCode = $null; Message = "Takajo is missing at: $takajoPath" }
        }
        else {
            $takajoBackup = Move-IbisExistingDirectoryToBackup -DirectoryPath $takajoOutputDirectory -BackupRoot (Join-Path $workingsDirectory 'Takajo-Backups')
            $takajoResult = Invoke-IbisProcessCapture -FilePath $takajoPath -ArgumentList @('automagic', '-t', $hayabusaJsonlPath, '-o', $takajoOutputDirectory, '-s') -WorkingDirectory (Split-Path -Path $takajoPath -Parent)
            $takajoStatus = 'Completed'
            $takajoMessage = 'Takajo automagic completed.'
            if ($takajoResult.ExitCode -ne 0) { $takajoStatus = 'Failed'; $takajoMessage = "Takajo automagic exited with code $($takajoResult.ExitCode)." }
            elseif (-not (Test-Path -LiteralPath $takajoOutputDirectory -PathType Container)) { $takajoStatus = 'Completed With Warnings'; $takajoMessage = 'Takajo automagic completed, but the output folder was not found.' }
            $toolResults += [pscustomobject]@{ ToolId = $takajo.id; Mode = 'automagic'; OutputPath = $takajoOutputDirectory; BackupPath = $takajoBackup; Status = $takajoStatus; ExitCode = $takajoResult.ExitCode; CommandLine = $takajoResult.CommandLine; Message = $takajoMessage }

            if (-not (Test-Path -LiteralPath $takajoOutputDirectory)) { New-Item -ItemType Directory -Path $takajoOutputDirectory -Force | Out-Null }
            $stackCommands = @(
                'stack-cmdlines',
                'stack-computers',
                'stack-dns',
                'stack-ip-addresses',
                'stack-logons',
                'stack-processes',
                'stack-services',
                'stack-tasks',
                'stack-users'
            )
            foreach ($stackCommand in $stackCommands) {
                $stackOutputPath = Join-Path $takajoOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'Takajo-{0}.csv' -ArgumentList @($stackCommand))
                $stackResult = Invoke-IbisProcessCapture -FilePath $takajoPath -ArgumentList @($stackCommand, '-t', $hayabusaJsonlPath, '-o', $stackOutputPath, '-s', '-q') -WorkingDirectory (Split-Path -Path $takajoPath -Parent)
                $stackStatus = 'Completed'
                $stackMessage = "$stackCommand completed."
                if ($stackResult.ExitCode -ne 0) { $stackStatus = 'Failed'; $stackMessage = "$stackCommand exited with code $($stackResult.ExitCode)." }
                elseif (-not (Test-Path -LiteralPath $stackOutputPath -PathType Leaf)) { $stackStatus = 'Completed With Warnings'; $stackMessage = "$stackCommand completed, but the expected CSV was not found." }
                $toolResults += [pscustomobject]@{ ToolId = $takajo.id; Mode = $stackCommand; OutputPath = $stackOutputPath; Status = $stackStatus; ExitCode = $stackResult.ExitCode; CommandLine = $stackResult.CommandLine; Message = $stackMessage }
            }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Takajo.json')
    $payload = [pscustomobject]@{
        ModuleId = 'takajo'
        Created = (Get-Date).ToString('s')
        HayabusaJsonlPath = $hayabusaJsonlPath
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $takajoOutputDirectory
        WorkingsDirectory = $workingsDirectory
        ToolResults = $toolResults
    }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $failed = @($toolResults | Where-Object { $_.Status -eq 'Failed' })
    $warnings = @($toolResults | Where-Object { $_.Status -match 'Warnings' })
    $status = 'Completed'
    $message = 'Takajo processing completed.'
    if ($failed.Count -gt 0) { $status = 'Failed'; $message = "$($failed.Count) Takajo operation(s) failed. See summary JSON for details." }
    elseif ($warnings.Count -gt 0) { $status = 'Completed With Warnings'; $message = "Takajo processing completed with $($warnings.Count) warning(s). See summary JSON for details." }

    [pscustomobject]@{
        ModuleId = 'takajo'
        Status = $status
        HayabusaJsonlPath = $hayabusaJsonlPath
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $takajoOutputDirectory
        JsonPath = $summaryPath
        Message = $message
    }
}

function Rename-IbisChainsawOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$StagingDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname
    )

    if (-not (Test-Path -LiteralPath $StagingDirectory -PathType Container)) {
        return @()
    }

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $moved = @()
    $files = @(Get-ChildItem -LiteralPath $StagingDirectory -File -Force -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        $newPath = Join-Path $OutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'Chainsaw-{0}' -ArgumentList @($file.Name))
        Move-Item -LiteralPath $file.FullName -Destination $newPath -Force
        $moved += [pscustomobject]@{ OriginalPath = $file.FullName; NewPath = $newPath }
    }

    $remaining = @(Get-ChildItem -LiteralPath $StagingDirectory -Force -ErrorAction SilentlyContinue)
    if ($remaining.Count -eq 0) {
        Remove-Item -LiteralPath $StagingDirectory -Force
    }

    $moved
}

function Invoke-IbisChainsawEventLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $sourceDirectory = Get-IbisEventLogPath -SourceRoot $SourceRoot
    $outputDirectory = Join-Path $hostOutputRoot 'EventLogs'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $chainsawStagingDirectory = Join-Path $outputDirectory 'Chainsaw'

    if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
        return [pscustomobject]@{ ModuleId = 'chainsaw'; Status = 'Skipped'; SourceDirectory = $sourceDirectory; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; JsonPath = $null; Message = 'Windows Event Log folder was not found.' }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) { New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'chainsaw'
    $movedOutputs = @()
    if ($null -eq $tool) {
        $toolResult = [pscustomobject]@{ ToolId = 'chainsaw'; OutputDirectory = $chainsawStagingDirectory; Status = 'Failed'; ExitCode = $null; Message = 'Chainsaw is not configured.' }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResult = [pscustomobject]@{ ToolId = $tool.id; OutputDirectory = $chainsawStagingDirectory; Status = 'Failed'; ExitCode = $null; Message = "Chainsaw is missing at: $toolPath" }
        }
        else {
            $toolDirectory = Split-Path -Path $toolPath -Parent
            $sigmaPath = Join-Path $toolDirectory 'sigma'
            $mappingPath = Join-Path $toolDirectory 'mappings\sigma-event-logs-all.yml'
            $rulesPath = Join-Path $toolDirectory 'rules'
            $backupPath = Move-IbisExistingDirectoryToBackup -DirectoryPath $chainsawStagingDirectory -BackupRoot (Join-Path $workingsDirectory 'Chainsaw-Backups')
            $processResult = Invoke-IbisProcessCapture -FilePath $toolPath -ArgumentList @('hunt', $sourceDirectory, '-s', $sigmaPath, '--mapping', $mappingPath, '-r', $rulesPath, '--csv', '--output', $chainsawStagingDirectory, '--skip-errors') -WorkingDirectory $toolDirectory
            if ($processResult.ExitCode -ne 0) {
                $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Chainsaw.stderr.txt')) -Encoding UTF8
            }
            $movedOutputs = @(Rename-IbisChainsawOutput -StagingDirectory $chainsawStagingDirectory -OutputDirectory $outputDirectory -Hostname $safeHost)
            $status = 'Completed'
            $message = 'Chainsaw hunt completed.'
            if ($processResult.ExitCode -ne 0) { $status = 'Failed'; $message = "Chainsaw exited with code $($processResult.ExitCode)." }
            elseif ($movedOutputs.Count -eq 0) { $status = 'Completed With Warnings'; $message = 'Chainsaw completed, but no output files were found.' }
            $toolResult = [pscustomobject]@{ ToolId = $tool.id; OutputDirectory = $outputDirectory; StagingDirectory = $chainsawStagingDirectory; BackupPath = $backupPath; Status = $status; ExitCode = $processResult.ExitCode; CommandLine = $processResult.CommandLine; MovedOutputs = $movedOutputs; Message = $message }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Chainsaw.json')
    $payload = [pscustomobject]@{ ModuleId = 'chainsaw'; Created = (Get-Date).ToString('s'); SourceRoot = $SourceRoot; SourceDirectory = $sourceDirectory; ToolsRoot = $ToolsRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; WorkingsDirectory = $workingsDirectory; ToolResults = @($toolResult) }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'Chainsaw event log processing completed.'
    if ($toolResult.Status -eq 'Failed') { $status = 'Failed'; $message = 'Chainsaw failed. See Chainsaw summary JSON for details.' }
    elseif ($toolResult.Status -match 'Warnings') { $status = 'Completed With Warnings'; $message = 'Chainsaw event log processing completed with warning(s). See summary JSON for details.' }

    [pscustomobject]@{ ModuleId = 'chainsaw'; Status = $status; SourceDirectory = $sourceDirectory; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; JsonPath = $summaryPath; Message = $message }
}

function Get-IbisUserAccessLogPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    [System.IO.Path]::Combine($SourceRoot, 'Windows\System32\LogFiles\Sum')
}

function Rename-IbisSumECmdOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        return @()
    }

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $renamed = @()
    $files = @(Get-ChildItem -LiteralPath $OutputDirectory -File -Force -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        $regexMatch = [regex]::Match($file.Name, '^\d+_SumECmd_(.+)$')
        if ($regexMatch.Success) {
            $newName = Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'SumECmd-{0}' -ArgumentList @($regexMatch.Groups[1].Value)
            $newPath = Join-Path $file.DirectoryName $newName
            Move-Item -LiteralPath $file.FullName -Destination $newPath -Force
            $renamed += [pscustomobject]@{
                OriginalPath = $file.FullName
                NewPath = $newPath
            }
        }
    }

    $renamed
}

function Invoke-IbisUserAccessLogsSum {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $sourceDirectory = Get-IbisUserAccessLogPath -SourceRoot $SourceRoot
    $outputDirectory = Join-Path $hostOutputRoot 'User Access Logs (SUM)'
    $workingsDirectory = Join-Path $outputDirectory '_Working'

    if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
        return [pscustomobject]@{
            ModuleId = 'ual'
            Status = 'Skipped'
            SourceDirectory = $sourceDirectory
            HostOutputRoot = $hostOutputRoot
            OutputDirectory = $outputDirectory
            JsonPath = $null
            Message = 'User Access Logs / SUM folder was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) { New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'zimmerman-sumecmd'
    $renamedOutputs = @()
    if ($null -eq $tool) {
        $toolResult = [pscustomobject]@{ ToolId = 'zimmerman-sumecmd'; SourceDirectory = $sourceDirectory; OutputDirectory = $outputDirectory; Status = 'Failed'; ExitCode = $null; Message = 'SumECmd is not configured.' }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceDirectory = $sourceDirectory; OutputDirectory = $outputDirectory; Status = 'Failed'; ExitCode = $null; Message = "SumECmd is missing at: $toolPath" }
        }
        else {
            $processResult = Invoke-IbisProcessCapture -FilePath $toolPath -ArgumentList @('-d', $sourceDirectory, '--csv', $outputDirectory) -WorkingDirectory (Split-Path -Path $toolPath -Parent)
            if ($processResult.ExitCode -ne 0) {
                $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'SumECmd.stderr.txt')) -Encoding UTF8
            }

            $renamedOutputs = @(Rename-IbisSumECmdOutput -OutputDirectory $outputDirectory -Hostname $safeHost)
            $outputFiles = @(Get-ChildItem -LiteralPath $outputDirectory -File -Force -ErrorAction SilentlyContinue)

            $status = 'Completed'
            $message = 'SumECmd completed.'
            if ($processResult.ExitCode -ne 0) {
                $status = 'Failed'
                $message = "SumECmd exited with code $($processResult.ExitCode)."
            }
            elseif ($outputFiles.Count -eq 0) {
                $status = 'Completed With Warnings'
                $message = 'SumECmd completed, but no output files were found.'
            }

            $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceDirectory = $sourceDirectory; OutputDirectory = $outputDirectory; Status = $status; ExitCode = $processResult.ExitCode; CommandLine = $processResult.CommandLine; Message = $message }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'User-Access-Logs-SUM.json')
    $payload = [pscustomobject]@{ ModuleId = 'ual'; Created = (Get-Date).ToString('s'); SourceRoot = $SourceRoot; SourceDirectory = $sourceDirectory; ToolsRoot = $ToolsRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; WorkingsDirectory = $workingsDirectory; ToolResults = @($toolResult); RenamedOutputs = @($renamedOutputs) }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'User Access Logs / SUM processing completed.'
    if ($toolResult.Status -eq 'Failed') { $status = 'Failed'; $message = 'SumECmd failed. See User Access Logs / SUM summary JSON for details.' }
    elseif ($toolResult.Status -match 'Warnings') { $status = 'Completed With Warnings'; $message = 'User Access Logs / SUM processing completed with warning(s). See summary JSON for details.' }

    [pscustomobject]@{ ModuleId = 'ual'; Status = $status; SourceDirectory = $sourceDirectory; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; JsonPath = $summaryPath; Message = $message }
}

function Get-IbisBrowserHistoryUsersPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    [System.IO.Path]::Combine($SourceRoot, 'Users')
}

function Invoke-IbisBrowsingHistoryView {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $sourceDirectory = Get-IbisBrowserHistoryUsersPath -SourceRoot $SourceRoot
    $outputDirectory = Join-Path $hostOutputRoot 'BrowsingHistoryView'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $outputPath = Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'BrowsingHistoryView-All-Users.csv')

    if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
        return [pscustomobject]@{ ModuleId = 'browser-history'; Status = 'Skipped'; SourceDirectory = $sourceDirectory; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; OutputPath = $outputPath; JsonPath = $null; Message = 'Users folder was not found.' }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) { New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'browsinghistoryview'
    if ($null -eq $tool) {
        $toolResult = [pscustomobject]@{ ToolId = 'browsinghistoryview'; SourceDirectory = $sourceDirectory; OutputPath = $outputPath; Status = 'Failed'; ExitCode = $null; Message = 'BrowsingHistoryView is not configured.' }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceDirectory = $sourceDirectory; OutputPath = $outputPath; Status = 'Failed'; ExitCode = $null; Message = "BrowsingHistoryView is missing at: $toolPath" }
        }
        else {
            $processResult = Invoke-IbisProcessCapture -FilePath $toolPath -ArgumentList @('/HistorySource', '3', '/HistorySourceFolder', $sourceDirectory, '/scomma', $outputPath, '/VisitTimeFilterType', '1') -WorkingDirectory (Split-Path -Path $toolPath -Parent)
            if ($processResult.ExitCode -ne 0) {
                $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'BrowsingHistoryView.stderr.txt')) -Encoding UTF8
            }

            $status = 'Completed'
            $message = 'BrowsingHistoryView completed.'
            if ($processResult.ExitCode -ne 0) {
                $status = 'Failed'
                $message = "BrowsingHistoryView exited with code $($processResult.ExitCode)."
            }
            elseif (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) {
                $status = 'Completed With Warnings'
                $message = 'BrowsingHistoryView completed, but the expected CSV was not found.'
            }
            elseif ((Get-Item -LiteralPath $outputPath).Length -eq 0) {
                $status = 'Completed With Warnings'
                $message = 'BrowsingHistoryView output CSV was created, but it is empty.'
            }

            $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceDirectory = $sourceDirectory; OutputPath = $outputPath; Status = $status; ExitCode = $processResult.ExitCode; CommandLine = $processResult.CommandLine; Message = $message }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'BrowsingHistoryView.json')
    $payload = [pscustomobject]@{ ModuleId = 'browser-history'; Created = (Get-Date).ToString('s'); SourceRoot = $SourceRoot; SourceDirectory = $sourceDirectory; ToolsRoot = $ToolsRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; WorkingsDirectory = $workingsDirectory; OutputPath = $outputPath; ToolResults = @($toolResult) }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'Browser history processing completed.'
    if ($toolResult.Status -eq 'Failed') { $status = 'Failed'; $message = 'BrowsingHistoryView failed. See Browser history summary JSON for details.' }
    elseif ($toolResult.Status -match 'Warnings') { $status = 'Completed With Warnings'; $message = 'Browser history processing completed with warning(s). See summary JSON for details.' }

    [pscustomobject]@{ ModuleId = 'browser-history'; Status = $status; SourceDirectory = $sourceDirectory; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; OutputPath = $outputPath; JsonPath = $summaryPath; Message = $message }
}

function Move-IbisForensicWebHistoryOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$StagingDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname
    )

    if (-not (Test-Path -LiteralPath $StagingDirectory -PathType Container)) {
        return @()
    }

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $moved = @()
    $files = @(Get-ChildItem -LiteralPath $StagingDirectory -File -Recurse -Force -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        $relativePath = $file.FullName.Substring($StagingDirectory.TrimEnd('\', '/').Length).TrimStart('\', '/')
        $safeRelativeName = ConvertTo-IbisSafeFileName -Value ($relativePath -replace '[\\/]+', '-')
        $newPath = Join-Path $OutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'ForensicWebHistory-{0}' -ArgumentList @($safeRelativeName))
        Move-Item -LiteralPath $file.FullName -Destination $newPath -Force
        $moved += [pscustomobject]@{
            OriginalPath = $file.FullName
            NewPath = $newPath
        }
    }

    $remainingFiles = @(Get-ChildItem -LiteralPath $StagingDirectory -File -Force -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($remainingFiles.Count -eq 0) {
        Remove-Item -LiteralPath $StagingDirectory -Recurse -Force
    }

    $moved
}

function Invoke-IbisForensicWebHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'ForensicWebHistory'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $stagingDirectory = Join-Path $workingsDirectory 'ForensicWebHistory-Staging'

    if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
        return [pscustomobject]@{ ModuleId = 'forensic-webhistory'; Status = 'Skipped'; SourceRoot = $SourceRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; JsonPath = $null; Message = 'Evidence source root was not found.' }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) { New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null }
    $backupPath = Move-IbisExistingDirectoryToBackup -DirectoryPath $stagingDirectory -BackupRoot (Join-Path $workingsDirectory 'ForensicWebHistory-Staging-Backups')
    if (-not (Test-Path -LiteralPath $stagingDirectory)) { New-Item -ItemType Directory -Path $stagingDirectory -Force | Out-Null }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'forensic-webhistory'
    $movedOutputs = @()
    if ($null -eq $tool) {
        $toolResult = [pscustomobject]@{ ToolId = 'forensic-webhistory'; SourceRoot = $SourceRoot; OutputDirectory = $outputDirectory; StagingDirectory = $stagingDirectory; BackupPath = $backupPath; Status = 'Failed'; ExitCode = $null; Message = 'Forensic webhistory is not configured.' }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceRoot = $SourceRoot; OutputDirectory = $outputDirectory; StagingDirectory = $stagingDirectory; BackupPath = $backupPath; Status = 'Failed'; ExitCode = $null; Message = "Forensic webhistory is missing at: $toolPath" }
        }
        else {
            $processResult = Invoke-IbisProcessCapture -FilePath $toolPath -ArgumentList @('scan', '-d', $SourceRoot, '-o', $stagingDirectory, '--date-format', 'iso') -WorkingDirectory (Split-Path -Path $toolPath -Parent)
            if ($processResult.ExitCode -ne 0) {
                $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'ForensicWebHistory.stderr.txt')) -Encoding UTF8
            }

            $movedOutputs = @(Move-IbisForensicWebHistoryOutput -StagingDirectory $stagingDirectory -OutputDirectory $outputDirectory -Hostname $safeHost)
            $status = 'Completed'
            $message = 'Forensic webhistory completed.'
            if ($processResult.ExitCode -ne 0) {
                $status = 'Failed'
                $message = "Forensic webhistory exited with code $($processResult.ExitCode)."
            }
            elseif ($movedOutputs.Count -eq 0) {
                $status = 'Completed With Warnings'
                $message = 'Forensic webhistory completed, but no output files were found.'
            }

            $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceRoot = $SourceRoot; OutputDirectory = $outputDirectory; StagingDirectory = $stagingDirectory; BackupPath = $backupPath; Status = $status; ExitCode = $processResult.ExitCode; CommandLine = $processResult.CommandLine; MovedOutputs = $movedOutputs; Message = $message }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'ForensicWebHistory.json')
    $payload = [pscustomobject]@{ ModuleId = 'forensic-webhistory'; Created = (Get-Date).ToString('s'); SourceRoot = $SourceRoot; ToolsRoot = $ToolsRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; WorkingsDirectory = $workingsDirectory; StagingDirectory = $stagingDirectory; ToolResults = @($toolResult) }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'Forensic webhistory processing completed.'
    if ($toolResult.Status -eq 'Failed') { $status = 'Failed'; $message = 'Forensic webhistory failed. See summary JSON for details.' }
    elseif ($toolResult.Status -match 'Warnings') { $status = 'Completed With Warnings'; $message = 'Forensic webhistory processing completed with warning(s). See summary JSON for details.' }

    [pscustomobject]@{ ModuleId = 'forensic-webhistory'; Status = $status; SourceRoot = $SourceRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; JsonPath = $summaryPath; Message = $message }
}

function Invoke-IbisParseUsbArtifacts {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'USB'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $logPath = Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'ParseUSBs-Log.txt')

    if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
        return [pscustomobject]@{ ModuleId = 'usb'; Status = 'Skipped'; SourceRoot = $SourceRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; JsonPath = $null; Message = 'Evidence source root was not found.' }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) { New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'parseusbs'
    if ($null -eq $tool) {
        $toolResult = [pscustomobject]@{ ToolId = 'parseusbs'; SourceRoot = $SourceRoot; OutputDirectory = $outputDirectory; LogPath = $logPath; Status = 'Failed'; ExitCode = $null; Message = 'parseusbs is not configured.' }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceRoot = $SourceRoot; OutputDirectory = $outputDirectory; LogPath = $logPath; Status = 'Failed'; ExitCode = $null; Message = "parseusbs is missing at: $toolPath" }
        }
        else {
            $sourceRootForTool = $SourceRoot.TrimEnd('\', '/')
            $processResult = Invoke-IbisProcessCapture -FilePath $toolPath -ArgumentList @('-v', $sourceRootForTool, '-o', 'csv', '-d', $outputDirectory) -WorkingDirectory (Split-Path -Path $toolPath -Parent)
            $processResult.StandardOutput | Out-File -LiteralPath $logPath -Encoding UTF8
            if ($processResult.ExitCode -ne 0 -or -not [string]::IsNullOrWhiteSpace($processResult.StandardError)) {
                $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'ParseUSBs.stderr.txt')) -Encoding UTF8
            }

            $csvFiles = @(Get-ChildItem -LiteralPath $outputDirectory -Filter '*.csv' -File -Force -ErrorAction SilentlyContinue)
            $status = 'Completed'
            $message = 'parseusbs completed.'
            if ($processResult.ExitCode -ne 0) {
                $status = 'Failed'
                $message = "parseusbs exited with code $($processResult.ExitCode)."
            }
            elseif ($csvFiles.Count -eq 0) {
                $status = 'Completed With Warnings'
                $message = 'parseusbs completed, but no CSV output files were found.'
            }

            $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceRoot = $sourceRootForTool; OutputDirectory = $outputDirectory; LogPath = $logPath; Status = $status; ExitCode = $processResult.ExitCode; CommandLine = $processResult.CommandLine; Message = $message }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'USB.json')
    $payload = [pscustomobject]@{ ModuleId = 'usb'; Created = (Get-Date).ToString('s'); SourceRoot = $SourceRoot; ToolsRoot = $ToolsRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; WorkingsDirectory = $workingsDirectory; ToolResults = @($toolResult) }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'USB artefact processing completed.'
    if ($toolResult.Status -eq 'Failed') { $status = 'Failed'; $message = 'parseusbs failed. See USB summary JSON for details.' }
    elseif ($toolResult.Status -match 'Warnings') { $status = 'Completed With Warnings'; $message = 'USB artefact processing completed with warning(s). See summary JSON for details.' }

    [pscustomobject]@{ ModuleId = 'usb'; Status = $status; SourceRoot = $SourceRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; JsonPath = $summaryPath; Message = $message }
}

function Get-IbisRegexValue {
    [CmdletBinding()]
    param(
        [string[]]$Lines,

        [Parameter(Mandatory = $true)]
        [string]$Pattern,

        [int]$Group = 1,

        [string]$DefaultValue = 'Unknown'
    )

    foreach ($line in $Lines) {
        $match = [regex]::Match($line, $Pattern)
        if ($match.Success -and $match.Groups.Count -gt $Group) {
            $value = $match.Groups[$Group].Value.Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }

    $DefaultValue
}

function ConvertFrom-IbisSystemSummaryRegRipperOutput {
    [CmdletBinding()]
    param(
        [string[]]$CompNameOutput = @(),

        [string[]]$WinVerOutput = @(),

        [string[]]$IpsOutput = @(),

        [string[]]$ShutdownOutput = @(),

        [string[]]$TimeZoneOutput = @()
    )

    $hostName = Get-IbisRegexValue -Lines $CompNameOutput -Pattern '^ComputerName\s+=\s+(.+)$'
    $productName = Get-IbisRegexValue -Lines $WinVerOutput -Pattern '^ProductName\s+(.+)$'
    $buildNumber = Get-IbisRegexValue -Lines $WinVerOutput -Pattern '^BuildLab\s+.*?(\d{5})' -DefaultValue ''
    $installDate = Get-IbisRegexValue -Lines $WinVerOutput -Pattern '^InstallDate\s+(.+)$'
    $shutdownDate = Get-IbisRegexValue -Lines $ShutdownOutput -Pattern '^LastWrite time:\s*(.+)$'
    $timeZone = Get-IbisRegexValue -Lines $TimeZoneOutput -Pattern 'TimeZoneKeyName->\s*(.+)$'

    $windowsVersion = $productName
    if (-not [string]::IsNullOrWhiteSpace($buildNumber)) {
        $windowsVersion = '{0} (Build: {1})' -f $windowsVersion, $buildNumber
        $buildInt = 0
        if ([int]::TryParse($buildNumber, [ref]$buildInt) -and $buildInt -gt 20000 -and $buildInt -lt 30000) {
            $windowsVersion += ' (Potentially Windows 11 or Server - double check the build number).'
        }
    }

    $ipLines = @()
    $captureIpLines = $false
    foreach ($line in $IpsOutput) {
        if ($line -match '^IPAddress\s+Domain') {
            $captureIpLines = $true
            continue
        }
        if ($captureIpLines -and -not [string]::IsNullOrWhiteSpace($line)) {
            $ipLines += $line.Trim()
        }
    }
    if ($ipLines.Count -eq 0) {
        $ipLines += 'Unknown'
    }

    [pscustomobject]@{
        HostName = $hostName
        OperatingSystem = $windowsVersion
        BuildNumber = $buildNumber
        TimeZone = $timeZone
        InstallDate = $installDate
        LastShutdown = $shutdownDate
        IpAddressSummary = @($ipLines)
    }
}

function Format-IbisSystemSummaryText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Summary
    )

    $lines = @()
    $lines += '========================================'
    $lines += 'System Information'
    $lines += '========================================'
    $lines += ('Host name: {0}' -f $Summary.HostName)
    $lines += ('Operating system: {0}' -f $Summary.OperatingSystem)
    $lines += ('Time zone: {0}' -f $Summary.TimeZone)
    $lines += ('Install date: {0}' -f $Summary.InstallDate)
    $lines += ('Last shutdown: {0}' -f $Summary.LastShutdown)
    $lines += 'IP Address(es) / Domain(s):'
    foreach ($ipLine in $Summary.IpAddressSummary) {
        $lines += $ipLine
    }
    $lines += '========================================'
    $lines -join [Environment]::NewLine
}

function Invoke-IbisExtractHostName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $systemHive = Get-IbisSystemHivePath -SourceRoot $SourceRoot -HiveName 'SYSTEM'
    $extractedHostName = 'Unknown'
    $status = 'Completed'
    $message = 'Hostname extraction completed.'
    $exitCode = $null
    $standardError = ''
    $commandLine = $null

    try {
        $regRipper = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'regripper'
        if ($null -eq $regRipper) {
            throw 'RegRipper is not configured.'
        }

        $ripPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $regRipper
        if (-not (Test-Path -LiteralPath $ripPath -PathType Leaf)) {
            throw "RegRipper is missing at: $ripPath"
        }

        if (-not (Test-Path -LiteralPath $systemHive -PathType Leaf)) {
            $status = 'Skipped'
            $message = 'SYSTEM hive was not found.'
        }
        else {
            $processResult = Invoke-IbisProcessCapture `
                -FilePath $ripPath `
                -ArgumentList @('-r', $systemHive, '-p', 'compname') `
                -WorkingDirectory (Split-Path -Path $ripPath -Parent)

            $exitCode = $processResult.ExitCode
            $standardError = $processResult.StandardError
            $commandLine = $processResult.CommandLine
            $lines = @($processResult.StandardOutput -split "\r?\n")
            $extractedHostName = Get-IbisRegexValue -Lines $lines -Pattern '^ComputerName\s+=\s+(.+)$'

            if ($processResult.ExitCode -ne 0) {
                $status = 'Failed'
                $message = "RegRipper compname exited with code $($processResult.ExitCode)."
            }
            elseif ($extractedHostName -eq 'Unknown') {
                $status = 'Completed With Warnings'
                $message = 'RegRipper ran, but no hostname was found in the compname output.'
            }
            else {
                $message = "Extracted hostname: $extractedHostName"
            }
        }
    }
    catch {
        $status = 'Failed'
        $message = $_.Exception.Message
    }

    $finalHost = $Hostname
    if ($extractedHostName -and $extractedHostName -ne 'Unknown') {
        $finalHost = $extractedHostName
    }
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $finalHost

    [pscustomobject]@{
        ModuleId = 'extract-hostname'
        Status = $status
        HostName = $extractedHostName
        HostOutputRoot = $hostOutputRoot
        SystemHive = $systemHive
        OutputPath = $null
        ExitCode = $exitCode
        CommandLine = $commandLine
        StandardError = $standardError
        Message = $message
    }
}

function Invoke-IbisSystemSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ToolDefinitions,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [string]$Hostname = 'HOST'
    )

    $systemHive = Get-IbisSystemHivePath -SourceRoot $SourceRoot -HiveName 'SYSTEM'
    $softwareHive = Get-IbisSystemHivePath -SourceRoot $SourceRoot -HiveName 'SOFTWARE'

    $initialSafeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $workingDirectory = Join-Path $OutputRoot '_ibis-system-summary'
    if (-not (Test-Path -LiteralPath $workingDirectory)) {
        New-Item -ItemType Directory -Path $workingDirectory -Force | Out-Null
    }

    $pluginResults = @()
    $pluginOutput = @{}

    $compNameOutputPath = Join-Path $workingDirectory (New-IbisHostPrefixedFileName -Hostname $initialSafeHost -Suffix 'RR-compname.txt')
    try {
        $compNameResult = Invoke-IbisRegRipperPlugin `
            -ToolsRoot $ToolsRoot `
            -ToolDefinitions $ToolDefinitions `
            -HivePath $systemHive `
            -Plugin 'compname' `
            -OutputPath $compNameOutputPath
    }
    catch {
        $compNameResult = [pscustomobject]@{
            Plugin = 'compname'
            HivePath = $systemHive
            OutputPath = $compNameOutputPath
            Status = 'Failed'
            ExitCode = $null
            Message = $_.Exception.Message
        }
    }
    $pluginResults += $compNameResult

    if (Test-Path -LiteralPath $compNameOutputPath -PathType Leaf) {
        $pluginOutput['compname'] = @(Get-Content -LiteralPath $compNameOutputPath)
    }
    else {
        $pluginOutput['compname'] = @()
    }

    $preserveBlankHostname = [string]::IsNullOrWhiteSpace($Hostname)
    $summaryHost = Get-IbisRegexValue -Lines $pluginOutput['compname'] -Pattern '^ComputerName\s+=\s+(.+)$' -DefaultValue $initialSafeHost
    $pathHost = if ($preserveBlankHostname) { '' } else { $summaryHost }
    $safeHost = ConvertTo-IbisSafeFileName -Value $pathHost -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'System-Summary'
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    }
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    if (-not (Test-Path -LiteralPath $workingsDirectory)) {
        New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null
    }

    $finalCompNameOutputPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'RR-compname.txt')
    if (Test-Path -LiteralPath $compNameOutputPath -PathType Leaf) {
        Move-Item -LiteralPath $compNameOutputPath -Destination $finalCompNameOutputPath -Force
    }
    $compNameResult.OutputPath = $finalCompNameOutputPath
    if (Test-Path -LiteralPath $finalCompNameOutputPath -PathType Leaf) {
        $pluginOutput['compname'] = @(Get-Content -LiteralPath $finalCompNameOutputPath)
    }

    if (Test-Path -LiteralPath $workingDirectory -PathType Container) {
        $remaining = @(Get-ChildItem -LiteralPath $workingDirectory -Force)
        if ($remaining.Count -eq 0) {
            Remove-Item -LiteralPath $workingDirectory -Force
        }
    }

    $pluginSpecs = @(
        @{ Name = 'ips'; HivePath = $systemHive },
        @{ Name = 'shutdown'; HivePath = $systemHive },
        @{ Name = 'timezone'; HivePath = $systemHive },
        @{ Name = 'winver'; HivePath = $softwareHive }
    )

    foreach ($spec in $pluginSpecs) {
        $outputPath = Join-Path $workingsDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}.txt' -ArgumentList @($spec.Name))
        try {
            $result = Invoke-IbisRegRipperPlugin `
                -ToolsRoot $ToolsRoot `
                -ToolDefinitions $ToolDefinitions `
                -HivePath $spec.HivePath `
                -Plugin $spec.Name `
                -OutputPath $outputPath
        }
        catch {
            $result = [pscustomobject]@{
                Plugin = $spec.Name
                HivePath = $spec.HivePath
                OutputPath = $outputPath
                Status = 'Failed'
                ExitCode = $null
                Message = $_.Exception.Message
            }
        }
        $pluginResults += $result

        if (Test-Path -LiteralPath $outputPath -PathType Leaf) {
            $pluginOutput[$spec.Name] = @(Get-Content -LiteralPath $outputPath)
        }
        else {
            $pluginOutput[$spec.Name] = @()
        }
    }

    $summary = ConvertFrom-IbisSystemSummaryRegRipperOutput `
        -CompNameOutput $pluginOutput['compname'] `
        -WinVerOutput $pluginOutput['winver'] `
        -IpsOutput $pluginOutput['ips'] `
        -ShutdownOutput $pluginOutput['shutdown'] `
        -TimeZoneOutput $pluginOutput['timezone']

    $textPath = Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'RR-System-Summary.txt')
    $jsonPath = Join-Path $outputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'System-Summary.json')
    Format-IbisSystemSummaryText -Summary $summary | Out-File -LiteralPath $textPath -Encoding UTF8

    $payload = [pscustomobject]@{
        ModuleId = 'system-summary'
        Created = (Get-Date).ToString('s')
        SourceRoot = $SourceRoot
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $outputDirectory
        WorkingsDirectory = $workingsDirectory
        SystemHive = $systemHive
        SoftwareHive = $softwareHive
        Summary = $summary
        RegRipperPlugins = $pluginResults
    }
    $payload | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $jsonPath -Encoding UTF8

    $failed = @($pluginResults | Where-Object { $_.Status -eq 'Failed' })
    $skipped = @($pluginResults | Where-Object { $_.Status -eq 'Skipped' })
    $status = 'Completed'
    $message = 'System summary completed with RegRipper.'
    if ($failed.Count -gt 0) {
        $status = 'Failed'
        $message = "$($failed.Count) RegRipper plugin(s) failed."
    }
    elseif ($skipped.Count -gt 0) {
        $status = 'Completed With Warnings'
        $message = "$($skipped.Count) RegRipper plugin(s) skipped because source hives were missing."
    }

    [pscustomobject]@{
        ModuleId = 'system-summary'
        Status = $status
        HostName = $summary.HostName
        HostOutputRoot = $hostOutputRoot
        OutputPath = $textPath
        JsonPath = $jsonPath
        Message = $message
    }
}

function New-IbisRunSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [object[]]$SelectedModules
    )

    [pscustomobject]@{
        Created = (Get-Date).ToString('s')
        ToolsRoot = $ToolsRoot
        SourceRoot = $SourceRoot
        OutputRoot = $OutputRoot
        Hostname = $Hostname
        SelectedModules = @($SelectedModules | ForEach-Object { $_.id })
        Note = 'Initial test summary only. No external DFIR tools were executed.'
    }
}

function Save-IbisRunSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Summary
    )

    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $Summary.OutputRoot -Hostname $Summary.Hostname
    if (-not (Test-Path -LiteralPath $hostOutputRoot)) {
        New-Item -ItemType Directory -Path $hostOutputRoot | Out-Null
    }

    $safeHost = ConvertTo-IbisSafeFileName -Value $Summary.Hostname -DefaultValue ''
    $fileName = New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Ibis-Initial-Test-Summary.json'
    $path = Join-Path $hostOutputRoot $fileName
    $Summary | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $path -Encoding UTF8
    $path
}

Export-ModuleMember -Function Get-IbisConfig
Export-ModuleMember -Function Save-IbisConfigPathSetting
Export-ModuleMember -Function Get-IbisToolDefinition
Export-ModuleMember -Function Test-IbisToolStatus
Export-ModuleMember -Function Get-IbisToolAcquisitionPlan
Export-ModuleMember -Function Format-IbisToolAcquisitionPlan
Export-ModuleMember -Function Write-IbisProgressEvent
Export-ModuleMember -Function Get-IbisToolDefinitionById
Export-ModuleMember -Function Resolve-IbisToolDownloadUrl
Export-ModuleMember -Function Get-IbisToolInstallDirectory
Export-ModuleMember -Function Get-IbisToolExpectedPath
Export-ModuleMember -Function Test-IbisIsAdministrator
Export-ModuleMember -Function Get-IbisDefenderExclusionRecommendation
Export-ModuleMember -Function Get-IbisDefenderExclusionStatus
Export-ModuleMember -Function Add-IbisDefenderExclusion
Export-ModuleMember -Function Remove-IbisDefenderExclusion
Export-ModuleMember -Function New-IbisToolInstallWorkspace
Export-ModuleMember -Function Get-IbisToolPublishSource
Export-ModuleMember -Function Backup-IbisToolInstallDirectory
Export-ModuleMember -Function Publish-IbisStagedToolInstall
Export-ModuleMember -Function New-IbisToolBackupPath
Export-ModuleMember -Function Test-IbisToolInstallState
Export-ModuleMember -Function Get-IbisExecutableRenameCandidate
Export-ModuleMember -Function Invoke-IbisToolPostInstall
Export-ModuleMember -Function Invoke-IbisInstallTool
Export-ModuleMember -Function Invoke-IbisInstallMissingTools
Export-ModuleMember -Function New-IbisCommandSpec
Export-ModuleMember -Function ConvertTo-IbisCommandLine
Export-ModuleMember -Function Resolve-IbisComparablePath
Export-ModuleMember -Function Test-IbisPathInsideRoot
Export-ModuleMember -Function Test-IbisSourceWriteBoundary
Export-ModuleMember -Function Test-IbisEvidenceRoot
Export-ModuleMember -Function Find-IbisVelociraptorResultsPath
Export-ModuleMember -Function Invoke-IbisVelociraptorResultsCopy
Export-ModuleMember -Function ConvertTo-IbisSafeFileName
Export-ModuleMember -Function Get-IbisHostFilePrefix
Export-ModuleMember -Function New-IbisHostPrefixedFileName
Export-ModuleMember -Function Format-IbisHostPrefixedValue
Export-ModuleMember -Function Get-IbisHostOutputRoot
Export-ModuleMember -Function Get-IbisSystemHivePath
Export-ModuleMember -Function Get-IbisWindowsRegistryHiveName
Export-ModuleMember -Function Test-IbisRegistryHiveTransactionState
Export-ModuleMember -Function Copy-IbisRegistryHiveToCache
Export-ModuleMember -Function Invoke-IbisRegistryHiveTransactionReplay
Export-ModuleMember -Function Get-IbisCachedRegistryHivePreparation
Export-ModuleMember -Function Invoke-IbisPrepareRegistryHiveFile
Export-ModuleMember -Function Invoke-IbisPrepareRegistryHive
Export-ModuleMember -Function Invoke-IbisPrepareRegistryHives
Export-ModuleMember -Function Invoke-IbisProcessCapture
Export-ModuleMember -Function Invoke-IbisRegRipperPlugin
Export-ModuleMember -Function Invoke-IbisRegRipperHiveMode
Export-ModuleMember -Function Invoke-IbisHayabusaRuleUpdate
Export-ModuleMember -Function Invoke-IbisWindowsRegistryHives
Export-ModuleMember -Function Get-IbisAmcacheHivePath
Export-ModuleMember -Function Invoke-IbisAmcacheParser
Export-ModuleMember -Function Invoke-IbisAmcache
Export-ModuleMember -Function Invoke-IbisAppCompatCacheParser
Export-ModuleMember -Function Invoke-IbisAppCompatCache
Export-ModuleMember -Function Get-IbisPrefetchPath
Export-ModuleMember -Function Rename-IbisPrefetchOutput
Export-ModuleMember -Function Invoke-IbisPrefetch
Export-ModuleMember -Function Find-IbisNtfsArtifactPath
Export-ModuleMember -Function Test-IbisNtfsSpecialFilePath
Export-ModuleMember -Function Find-IbisUsnJournalPath
Export-ModuleMember -Function Invoke-IbisMftECmdArtifact
Export-ModuleMember -Function Invoke-IbisNtfsMetadata
Export-ModuleMember -Function Get-IbisSrumDatabasePath
Export-ModuleMember -Function Rename-IbisSrumECmdOutput
Export-ModuleMember -Function Invoke-IbisSrum
Export-ModuleMember -Function Get-IbisUserProfile
Export-ModuleMember -Function Invoke-IbisUserDirectoryTool
Export-ModuleMember -Function Rename-IbisUserArtifactToolOutput
Export-ModuleMember -Function Copy-IbisPSReadLineHistory
Export-ModuleMember -Function Invoke-IbisUserArtifacts
Export-ModuleMember -Function Get-IbisEventLogPath
Export-ModuleMember -Function Invoke-IbisEvtxECmdEventLogs
Export-ModuleMember -Function Get-IbisEvtxECmdCsvPath
Export-ModuleMember -Function Get-IbisDuckDbEventLogQueryDefinition
Export-ModuleMember -Function ConvertTo-IbisDuckDbSqlLiteral
Export-ModuleMember -Function Expand-IbisDuckDbSqlTemplate
Export-ModuleMember -Function Invoke-IbisDuckDbEventLogSummary
Export-ModuleMember -Function Move-IbisExistingDirectoryToBackup
Export-ModuleMember -Function Invoke-IbisHayabusaEventLogs
Export-ModuleMember -Function Get-IbisHayabusaJsonlPath
Export-ModuleMember -Function Invoke-IbisTakajoEventLogs
Export-ModuleMember -Function Rename-IbisChainsawOutput
Export-ModuleMember -Function Invoke-IbisChainsawEventLogs
Export-ModuleMember -Function Get-IbisUserAccessLogPath
Export-ModuleMember -Function Rename-IbisSumECmdOutput
Export-ModuleMember -Function Invoke-IbisUserAccessLogsSum
Export-ModuleMember -Function Get-IbisBrowserHistoryUsersPath
Export-ModuleMember -Function Invoke-IbisBrowsingHistoryView
Export-ModuleMember -Function Move-IbisForensicWebHistoryOutput
Export-ModuleMember -Function Invoke-IbisForensicWebHistory
Export-ModuleMember -Function Invoke-IbisParseUsbArtifacts
Export-ModuleMember -Function Get-IbisRegexValue
Export-ModuleMember -Function ConvertFrom-IbisSystemSummaryRegRipperOutput
Export-ModuleMember -Function Format-IbisSystemSummaryText
Export-ModuleMember -Function Invoke-IbisExtractHostName
Export-ModuleMember -Function Invoke-IbisSystemSummary
Export-ModuleMember -Function New-IbisRunSummary
Export-ModuleMember -Function Save-IbisRunSummary


