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

    $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    $localSettingsPath = Join-Path $ProjectRoot 'config.local.json'
    if (Test-Path -LiteralPath $localSettingsPath -PathType Leaf) {
        try {
            $localSettings = Get-Content -LiteralPath $localSettingsPath -Raw | ConvertFrom-Json
            foreach ($propertyName in @('defaultToolsRoot', 'defaultSourceRoot', 'defaultOutputRoot', 'completionBeepEnabled')) {
                $property = $localSettings.PSObject.Properties[$propertyName]
                if ($null -ne $property) {
                    $config.$propertyName = $property.Value
                }
            }
        }
        catch {
            throw "Ibis local settings could not be read at ${localSettingsPath}: $($_.Exception.Message)"
        }
    }

    $config
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

    $localSettingsPath = Join-Path $ProjectRoot 'config.local.json'
    $settings = [ordered]@{}
    if (Test-Path -LiteralPath $localSettingsPath -PathType Leaf) {
        try {
            $existingSettings = Get-Content -LiteralPath $localSettingsPath -Raw | ConvertFrom-Json
            foreach ($property in $existingSettings.PSObject.Properties) {
                $settings[$property.Name] = $property.Value
            }
        }
        catch {
            throw "Ibis local settings could not be read at ${localSettingsPath}: $($_.Exception.Message)"
        }
    }

    if ($PSBoundParameters.ContainsKey('ToolsRoot')) {
        $settings.defaultToolsRoot = $ToolsRoot
    }
    if ($PSBoundParameters.ContainsKey('SourceRoot')) {
        $settings.defaultSourceRoot = $SourceRoot
    }
    if ($PSBoundParameters.ContainsKey('OutputRoot')) {
        $settings.defaultOutputRoot = $OutputRoot
    }
    if ($PSBoundParameters.ContainsKey('CompletionBeepEnabled')) {
        $settings.completionBeepEnabled = $CompletionBeepEnabled
    }

    [pscustomobject]$settings | ConvertTo-Json -Depth 20 | Out-File -LiteralPath $localSettingsPath -Encoding UTF8
    Get-IbisConfig -ProjectRoot $ProjectRoot
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
    $files = Get-ChildItem -LiteralPath $toolDefinitionRoot -Filter '*.json' -File -Force | Sort-Object Name
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
        $installState = $null
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
            Assessment = $installState
            Message = if ($installState) { $installState.Message } else { $null }
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
                Status = $status.Status
                ExpectedPath = $status.ExpectedPath
                AcquisitionSource = $source
                Notes = $status.Notes
                Message = $status.Message
                ExistingFiles = if ($status.Assessment) { @($status.Assessment.VersionedExecutables) } else { @() }
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
        if (-not [string]::IsNullOrWhiteSpace([string]$item.Status)) {
            $lines += ('  Assessment:    {0}' -f $item.Status)
        }
        $lines += ('  Expected path: {0}' -f $item.ExpectedPath)
        $lines += ('  Get from:      {0}' -f $item.AcquisitionSource)
        if (-not [string]::IsNullOrWhiteSpace($item.Notes)) {
            $lines += ('  Notes:         {0}' -f $item.Notes)
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$item.Message)) {
            $lines += ('  Detail:        {0}' -f $item.Message)
        }
        if ($item.ExistingFiles -and $item.ExistingFiles.Count -gt 0) {
            $lines += ('  Existing:      {0}' -f (($item.ExistingFiles | ForEach-Object { $_.Name }) -join ', '))
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

function Get-IbisSevenZipPath {
    [CmdletBinding()]
    param()

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace(${env:ProgramFiles})) {
        $candidates += (Join-Path ${env:ProgramFiles} '7-Zip\7z.exe')
    }
    if (-not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
        $candidates += (Join-Path ${env:ProgramFiles(x86)} '7-Zip\7z.exe')
    }
    foreach ($commandName in @('7z.exe', '7za.exe')) {
        $command = Get-Command $commandName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command) {
            $candidates += $command.Source
        }
    }

    foreach ($candidate in @($candidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    $null
}

function Expand-IbisArchiveWithSevenZip {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $sevenZipPath = Get-IbisSevenZipPath
    if ([string]::IsNullOrWhiteSpace($sevenZipPath)) {
        throw '7-Zip fallback was not available. Install 7-Zip or enable Windows long paths, then retry.'
    }

    $result = Invoke-IbisProcessCapture `
        -FilePath $sevenZipPath `
        -ArgumentList @('x', '-y', "-o$DestinationPath", $LiteralPath) `
        -WorkingDirectory (Split-Path -Path $sevenZipPath -Parent)

    if ($result.ExitCode -ne 0) {
        throw "7-Zip fallback failed with exit code $($result.ExitCode). $($result.StandardError)"
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
            $dotNetError = $_.Exception.Message
            try {
                Expand-IbisArchiveWithSevenZip -LiteralPath $LiteralPath -DestinationPath $DestinationPath
                return
            }
            catch {
                throw "Unable to extract ZIP archive. Expand-Archive failed with: $archiveError. .NET fallback failed with: $dotNetError. 7-Zip fallback failed with: $($_.Exception.Message)"
            }
        }
    }
}

function Get-IbisLongPathsEnabled {
    [CmdletBinding()]
    param()

    try {
        $value = Get-ItemPropertyValue -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -ErrorAction Stop
        [bool]([int]$value -eq 1)
    }
    catch {
        $false
    }
}

function Set-IbisLongPathsEnabled {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enabled
    )

    if ($PSCmdlet.ShouldProcess('HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem\LongPathsEnabled', "Set to $([int]$Enabled)")) {
        New-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -Name 'LongPathsEnabled' -PropertyType DWord -Value ([int]$Enabled) -Force -ErrorAction Stop | Out-Null
    }

    [pscustomobject]@{
        Enabled = (Get-IbisLongPathsEnabled)
        Requested = $Enabled
        IsAdministrator = (Test-IbisIsAdministrator)
        Message = if ($Enabled) { 'Windows long path support enabled. A restart may be required before all processes use the setting.' } else { 'Windows long path support disabled. A restart may be required before all processes use the setting.' }
    }
}

function Get-IbisVisualCppRedistributableStatus {
    [CmdletBinding()]
    param(
        [ValidateSet('x64', 'x86')]
        [string]$Architecture = 'x64'
    )

    $runtimeKey = "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\$Architecture"
    $installed = $false
    $version = $null
    $source = $null
    $displayNames = @()

    try {
        if (Test-Path -LiteralPath $runtimeKey -PathType Container) {
            $runtime = Get-ItemProperty -LiteralPath $runtimeKey -ErrorAction Stop
            if ([int]$runtime.Installed -eq 1) {
                $installed = $true
                $version = [string]$runtime.Version
                $source = $runtimeKey
            }
        }
    }
    catch {
    }

    foreach ($uninstallRoot in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )) {
        try {
            if (-not (Test-Path -LiteralPath $uninstallRoot -PathType Container)) {
                continue
            }

            foreach ($item in @(Get-ChildItem -LiteralPath $uninstallRoot -Force -ErrorAction SilentlyContinue)) {
                try {
                    $properties = Get-ItemProperty -LiteralPath $item.PSPath -ErrorAction Stop
                    $displayName = [string]$properties.DisplayName
                    if ($displayName -match 'Microsoft Visual C\+\+ (2015|2017|2019|2022|2015-2022).*Redistributable' -and $displayName -match "\($Architecture\)") {
                        $displayNames += $displayName
                        if (-not $installed) {
                            $installed = $true
                            $version = [string]$properties.DisplayVersion
                            $source = $item.PSPath
                        }
                    }
                }
                catch {
                }
            }
        }
        catch {
        }
    }

    $message = if ($installed) {
        "Microsoft Visual C++ Redistributable 2015+ ($Architecture) appears to be installed."
    }
    else {
        "Microsoft Visual C++ Redistributable 2015+ ($Architecture) was not detected. Some DFIR tools may fail to start until it is installed."
    }

    [pscustomobject]@{
        Present = $installed
        Architecture = $Architecture
        Version = $version
        Source = $source
        DisplayNames = @($displayNames | Select-Object -Unique)
        MicrosoftUrl = 'https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist'
        Message = $message
    }
}

function Get-IbisProjectScriptFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    @(
        (Join-Path $ProjectRoot 'Run-Ibis.ps1'),
        (Join-Path $ProjectRoot 'modules\Ibis.Core.psm1'),
        (Join-Path $ProjectRoot 'modules\Ibis.Gui.psm1')
    )
}

function Get-IbisPowerShellReadiness {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    $projectFiles = @(Get-IbisProjectScriptFile -ProjectRoot $ProjectRoot)
    $executionPolicies = @()
    $policyError = $null
    try {
        $executionPolicies = @(Get-ExecutionPolicy -List -ErrorAction Stop | ForEach-Object {
            [pscustomobject]@{
                Scope = [string]$_.Scope
                ExecutionPolicy = [string]$_.ExecutionPolicy
            }
        })
    }
    catch {
        $policyError = $_.Exception.Message
    }

    $effectivePolicy = $null
    try {
        $effectivePolicy = [string](Get-ExecutionPolicy -ErrorAction Stop)
    }
    catch {
        if ([string]::IsNullOrWhiteSpace($policyError)) {
            $policyError = $_.Exception.Message
        }
    }

    $enforcedPolicy = @($executionPolicies | Where-Object {
        $_.Scope -in @('MachinePolicy', 'UserPolicy') -and $_.ExecutionPolicy -ne 'Undefined'
    })

    $fileStatuses = @()
    foreach ($projectFile in $projectFiles) {
        $exists = Test-Path -LiteralPath $projectFile -PathType Leaf
        $zoneIdentifier = $null
        if ($exists) {
            try {
                $zoneIdentifier = Get-Content -LiteralPath ($projectFile + ':Zone.Identifier') -Raw -ErrorAction Stop
            }
            catch {
            }
        }

        $fileStatuses += [pscustomobject]@{
            Path = $projectFile
            Exists = $exists
            HasZoneIdentifier = -not [string]::IsNullOrWhiteSpace($zoneIdentifier)
            ZoneIdentifier = $zoneIdentifier
        }
    }

    $backgroundImport = [pscustomobject]@{
        Passed = $false
        ErrorMessage = $null
        ErrorRecords = @()
    }
    $coreModulePath = Join-Path $ProjectRoot 'modules\Ibis.Core.psm1'
    if (Test-Path -LiteralPath $coreModulePath -PathType Leaf) {
        $backgroundPowerShell = [PowerShell]::Create()
        try {
            $importScript = {
                param([string]$ModulePath)
                $ErrorActionPreference = 'Stop'
                Import-Module -Name $ModulePath -Force -ErrorAction Stop
                'Ibis core module import passed.'
            }
            [void]$backgroundPowerShell.AddScript($importScript)
            [void]$backgroundPowerShell.AddArgument($coreModulePath)
            $handle = $backgroundPowerShell.BeginInvoke()
            try {
                [void]$backgroundPowerShell.EndInvoke($handle)
                $backgroundImport.Passed = -not $backgroundPowerShell.HadErrors
            }
            catch {
                $backgroundImport.ErrorMessage = $_.Exception.Message
            }

            $backgroundImport.ErrorRecords = @($backgroundPowerShell.Streams.Error | ForEach-Object { $_.ToString() })
            if (-not $backgroundImport.Passed -and $backgroundImport.ErrorRecords.Count -gt 0) {
                $backgroundImport.ErrorMessage = $backgroundImport.ErrorRecords[0]
            }
        }
        catch {
            $backgroundImport.ErrorMessage = $_.Exception.Message
        }
        finally {
            $backgroundPowerShell.Dispose()
        }
    }
    else {
        $backgroundImport.ErrorMessage = "Ibis core module was not found at: $coreModulePath"
    }

    $guidance = @()
    if ($enforcedPolicy.Count -gt 0) {
        $policyNames = ($enforcedPolicy | ForEach-Object { "$($_.Scope)=$($_.ExecutionPolicy)" }) -join ', '
        $guidance += "PowerShell execution policy is enforced by policy: $policyNames. Ibis cannot override Group Policy; ask IT to allow or sign Ibis."
    }
    elseif ($effectivePolicy -in @('AllSigned', 'Restricted')) {
        $guidance += "Effective PowerShell execution policy is $effectivePolicy. This can block Ibis modules or background work. Relaunch with an approved process-only execution policy, or ask IT for guidance."
    }
    elseif (-not [string]::IsNullOrWhiteSpace($policyError)) {
        $guidance += "Unable to determine the effective PowerShell execution policy: $policyError"
    }

    $markedFiles = @($fileStatuses | Where-Object { $_.HasZoneIdentifier })
    if ($markedFiles.Count -gt 0) {
        $guidance += "$($markedFiles.Count) Ibis PowerShell file(s) are marked as downloaded from the internet. After verifying this trusted checkout, use Unblock Ibis Files."
    }
    if (-not $backgroundImport.Passed) {
        $guidance += "Background PowerShell import failed: $($backgroundImport.ErrorMessage)"
    }
    if ($guidance.Count -eq 0) {
        $guidance += 'PowerShell readiness passed. Ibis can start its background tool download runspace.'
    }

    [pscustomobject]@{
        Status = if ($backgroundImport.Passed -and $guidance.Count -eq 1 -and $guidance[0] -like 'PowerShell readiness passed*') { 'Ready' } elseif ($backgroundImport.Passed) { 'Warning' } else { 'Blocked' }
        CanStartBackgroundWork = $backgroundImport.Passed
        EffectiveExecutionPolicy = $effectivePolicy
        ExecutionPolicies = $executionPolicies
        EnforcedPolicies = $enforcedPolicy
        PolicyError = $policyError
        Files = $fileStatuses
        MarkedFiles = $markedFiles
        BackgroundImport = $backgroundImport
        Guidance = $guidance
    }
}

function Unblock-IbisProjectScriptFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectRoot
    )

    foreach ($projectFile in @(Get-IbisProjectScriptFile -ProjectRoot $ProjectRoot)) {
        $hasZoneIdentifier = $false
        if (Test-Path -LiteralPath $projectFile -PathType Leaf) {
            try {
                $hasZoneIdentifier = -not [string]::IsNullOrWhiteSpace((Get-Content -LiteralPath ($projectFile + ':Zone.Identifier') -Raw -ErrorAction Stop))
            }
            catch {
            }
        }

        $status = 'Skipped'
        $message = 'File is not marked as downloaded from the internet.'
        if (-not (Test-Path -LiteralPath $projectFile -PathType Leaf)) {
            $message = 'File was not found.'
        }
        elseif ($hasZoneIdentifier -and $PSCmdlet.ShouldProcess($projectFile, 'Remove internet-origin mark from trusted Ibis file')) {
            try {
                Unblock-File -LiteralPath $projectFile -ErrorAction Stop
                $status = 'Unblocked'
                $message = 'Removed the internet-origin mark.'
            }
            catch {
                $status = 'Failed'
                $message = $_.Exception.Message
            }
        }

        [pscustomobject]@{
            Path = $projectFile
            Status = $status
            Message = $message
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
    $shortId = $id.Substring(0, 8)
    $installDirectory = Get-IbisToolInstallDirectory -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition

    if ($ToolDefinition.defenderExclusionRecommended -eq $true) {
        if (-not (Test-Path -LiteralPath $installDirectory)) {
            New-Item -ItemType Directory -Path $installDirectory -Force | Out-Null
        }
        $workspaceRoot = Join-Path (Join-Path $installDirectory '_s') $shortId
    }
    else {
        $workspaceRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('IbisToolInstall-' + $id)
    }

    $downloadDirectory = Join-Path $workspaceRoot 'd'
    $extractDirectory = Join-Path $workspaceRoot 'x'
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
                $items += Get-Item -LiteralPath $candidate -Force
            }
        }
        if ($ToolDefinition -and $ToolDefinition.renameExtractedDirectoryFrom) {
            $candidate = Join-Path $ToolsRoot $ToolDefinition.renameExtractedDirectoryFrom
            if (Test-Path -LiteralPath $candidate) {
                $items += Get-Item -LiteralPath $candidate -Force
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

function Clear-IbisToolPreviousInstall {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition
    )

    $installDirectory = Get-IbisToolInstallDirectory -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
    if (-not (Test-Path -LiteralPath $installDirectory -PathType Container)) {
        return $null
    }

    $items = @()
    if ((Resolve-IbisComparablePath -Path $installDirectory) -eq (Resolve-IbisComparablePath -Path $ToolsRoot)) {
        foreach ($folderName in @($ToolDefinition.renameExtractedDirectoryTo, $ToolDefinition.renameExtractedDirectoryFrom)) {
            if (-not [string]::IsNullOrWhiteSpace($folderName)) {
                $path = Join-Path $ToolsRoot $folderName
                if (Test-Path -LiteralPath $path) {
                    $items += Get-Item -LiteralPath $path -Force
                }
            }
        }
    }
    elseif ($ToolDefinition.installDirectory -like 'EZTools\net9*') {
        $expectedPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
        if (Test-Path -LiteralPath $expectedPath) {
            $items += Get-Item -LiteralPath $expectedPath -Force
        }
    }
    else {
        $items = @(Get-ChildItem -LiteralPath $installDirectory -Force | Where-Object {
            $_.Name -ne '_ibis-backup' -and $_.Name -ne '_s' -and $_.Name -ne '_ibis-staging'
        })
    }

    if ($items.Count -eq 0) {
        return $null
    }

    $backupPath = New-IbisToolBackupPath -InstallDirectory $installDirectory
    foreach ($item in $items) {
        if ($PSCmdlet.ShouldProcess($item.FullName, "Archive previous $($ToolDefinition.name) install")) {
            Move-Item -LiteralPath $item.FullName -Destination $backupPath -Force
        }
    }
    $backupPath
}

function Get-IbisToolFileVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $version = $null
    try {
        $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($File.FullName)
        if (-not [string]::IsNullOrWhiteSpace($versionInfo.ProductVersion)) {
            $version = $versionInfo.ProductVersion
        }
        elseif (-not [string]::IsNullOrWhiteSpace($versionInfo.FileVersion)) {
            $version = $versionInfo.FileVersion
        }
    }
    catch {
    }

    [pscustomobject]@{
        Name = $File.Name
        Path = $File.FullName
        Version = $version
        Length = $File.Length
        LastWriteTime = $File.LastWriteTime
    }
}

function Remove-IbisEmptyToolWorkspaceDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstallDirectory
    )

    foreach ($workspaceName in @('_s', '_ibis-staging')) {
        $workspacePath = Join-Path $InstallDirectory $workspaceName
        if (-not (Test-Path -LiteralPath $workspacePath -PathType Container)) {
            continue
        }

        $items = @(Get-ChildItem -LiteralPath $workspacePath -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($items.Count -eq 0) {
            Remove-Item -LiteralPath $workspacePath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-IbisToolInstallAssessment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition
    )

    $expectedPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
    $installDirectory = Get-IbisToolInstallDirectory -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
    $canonicalPresent = Test-Path -LiteralPath $expectedPath -PathType Leaf
    $versionedExecutables = @()
    if ($ToolDefinition.renameExecutablePattern) {
        $versionedExecutables = @(Get-IbisExecutableRenameCandidate -InstallDirectory $installDirectory -Pattern $ToolDefinition.renameExecutablePattern |
            Where-Object { (Resolve-IbisComparablePath -Path $_.FullName) -ne (Resolve-IbisComparablePath -Path $expectedPath) } |
            ForEach-Object { Get-IbisToolFileVersion -File $_ })
    }

    $status = 'Missing'
    $message = 'Expected executable is missing.'
    if ($canonicalPresent) {
        $status = if ($versionedExecutables.Count -gt 0) { 'Present with legacy files' } else { 'Present' }
        $message = if ($versionedExecutables.Count -gt 0) {
            "Expected executable is present; $($versionedExecutables.Count) versioned executable(s) also remain."
        }
        else {
            'Expected executable is present.'
        }
    }
    elseif ($versionedExecutables.Count -eq 1) {
        $status = 'Needs Normalization'
        $message = "Found versioned executable '$($versionedExecutables[0].Name)' but not the configured canonical filename. Download Missing Tools can install the latest release, or use Reinstall Selected to archive it first."
    }
    elseif ($versionedExecutables.Count -gt 1) {
        $status = 'Ambiguous Existing Versions'
        $message = "Found $($versionedExecutables.Count) versioned executables but not the configured canonical filename. Reinstall Selected will archive the active tool folder and install the latest release."
    }
    elseif (Test-Path -LiteralPath $installDirectory -PathType Container) {
        $items = @(Get-ChildItem -LiteralPath $installDirectory -Force | Where-Object {
            $_.Name -ne '_ibis-staging' -and $_.Name -ne '_s' -and $_.Name -ne '_ibis-backup'
        })
        if ((Resolve-IbisComparablePath -Path $installDirectory) -eq (Resolve-IbisComparablePath -Path $ToolsRoot) -and $ToolDefinition.renameExtractedDirectoryTo) {
            $toolFolder = Join-Path $ToolsRoot $ToolDefinition.renameExtractedDirectoryTo
            $items = if (Test-Path -LiteralPath $toolFolder) { @(Get-Item -LiteralPath $toolFolder -Force) } else { @() }
        }
        if ($items.Count -gt 0) {
            $status = 'Partial'
            $message = 'Install directory contains files, but the expected executable is missing.'
        }
    }

    [pscustomobject]@{
        Id = $ToolDefinition.id
        Name = $ToolDefinition.name
        Status = $status
        Present = $canonicalPresent
        ExpectedPath = $expectedPath
        InstallDirectory = $installDirectory
        VersionedExecutables = @($versionedExecutables)
        Message = $message
    }
}

function Test-IbisToolInstallState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ToolsRoot,

        [Parameter(Mandatory = $true)]
        [object]$ToolDefinition
    )

    Get-IbisToolInstallAssessment -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
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

        [int]$ProgressTotal = 0,

        [switch]$ForceReinstall
    )

    $installState = Test-IbisToolInstallState -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
    $expectedPath = $installState.ExpectedPath
    if ($installState.Present -and -not $ForceReinstall) {
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
            $publishedExecutableRelativePath = $null
            if ($ToolDefinition.renameExecutablePattern) {
                $stagedCandidates = @(Get-IbisExecutableRenameCandidate -InstallDirectory $publishSource -Pattern $ToolDefinition.renameExecutablePattern)
                if ($stagedCandidates.Count -eq 1) {
                    $publishedExecutableRelativePath = $stagedCandidates[0].FullName.Substring($publishSource.Length).TrimStart('\', '/')
                }
            }
            if ($ForceReinstall) {
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'ArchivePrevious' -Message 'Archiving previous tool files before reinstall.' -Index $ProgressIndex -Total $ProgressTotal
                $backupPath = Clear-IbisToolPreviousInstall -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
            }
            $publishBackupPath = Publish-IbisStagedToolInstall -StagedSourcePath $publishSource -InstallDirectory $installDirectory -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition
            if ($publishBackupPath) {
                $backupPath = $publishBackupPath
            }
            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId $ToolDefinition.id -ToolName $ToolDefinition.name -Stage 'PostInstall' -Message 'Running post-install checks.' -Index $ProgressIndex -Total $ProgressTotal
            Invoke-IbisToolPostInstall -ToolsRoot $ToolsRoot -ToolDefinition $ToolDefinition -PublishedExecutableRelativePath $publishedExecutableRelativePath
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
            Remove-IbisEmptyToolWorkspaceDirectory -InstallDirectory $installDirectory
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

        [string]$ProgressPath,

        [switch]$ForceReinstall
    )

    $statuses = @(Test-IbisToolStatus -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions)
    $missing = @($statuses | Where-Object { -not $_.Present })
    if ($ForceReinstall -and $ToolIds -and $ToolIds.Count -gt 0) {
        $missing = @($statuses | Where-Object { $ToolIds -contains $_.Id })
    }
    elseif ($ToolIds -and $ToolIds.Count -gt 0) {
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
                Invoke-IbisInstallTool -ToolsRoot $ToolsRoot -ToolDefinition $definition -ProgressPath $ProgressPath -ProgressIndex $index -ProgressTotal $total -ForceReinstall:$ForceReinstall
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
        [object]$ToolDefinition,

        [string]$PublishedExecutableRelativePath
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
        $target = Join-Path $installDirectory $ToolDefinition.renameExecutableTo
        $publishedCandidate = $null
        if (-not [string]::IsNullOrWhiteSpace($PublishedExecutableRelativePath)) {
            $candidatePath = Join-Path $installDirectory $PublishedExecutableRelativePath
            if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
                $publishedCandidate = Get-Item -LiteralPath $candidatePath -Force
            }
        }
        if ($publishedCandidate) {
            if ((Resolve-IbisComparablePath -Path $publishedCandidate.FullName) -ne (Resolve-IbisComparablePath -Path $target)) {
                Move-Item -LiteralPath $publishedCandidate.FullName -Destination $target -Force
            }
        }
        else {
            $renameCandidates = @(Get-IbisExecutableRenameCandidate -InstallDirectory $installDirectory -Pattern $ToolDefinition.renameExecutablePattern)
            if ($renameCandidates.Count -eq 1 -and (Resolve-IbisComparablePath -Path $renameCandidates[0].FullName) -ne (Resolve-IbisComparablePath -Path $target)) {
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

    Get-ChildItem -LiteralPath $InstallDirectory -Filter $Pattern -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $relativePath = $_.FullName.Substring($InstallDirectory.Length).TrimStart('\', '/')
            $pathParts = $relativePath -split '[\\/]'
            -not ($pathParts -contains '_ibis-backup') -and -not ($pathParts -contains '_ibis-staging') -and -not ($pathParts -contains '_s')
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

function Copy-IbisEvidenceFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force -ErrorAction Stop

    # Evidence is normally mounted or held read-only.  Copy-Item carries the
    # read-only attribute onto the working copy, which would stop tools such as
    # rla and ParseUSBs from writing a replayed hive into our own cache. Clearing
    # it on the copy keeps the source untouched and the working copy usable.
    $copiedFile = Get-Item -LiteralPath $DestinationPath -Force
    if ($copiedFile.IsReadOnly) {
        $copiedFile.IsReadOnly = $false
    }

    $copiedFile
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
    Copy-IbisEvidenceFile -SourcePath $SourceHivePath -DestinationPath $destinationHivePath | Out-Null

    $sourceDirectory = Split-Path -Path $SourceHivePath -Parent
    $transactionLogs = @(Get-ChildItem -LiteralPath $sourceDirectory -Filter "$hiveName.LOG*" -File -Force -ErrorAction SilentlyContinue)
    foreach ($log in $transactionLogs) {
        Copy-IbisEvidenceFile -SourcePath $log.FullName -DestinationPath (Join-Path $CacheDirectory $log.Name) | Out-Null
    }

    [pscustomobject]@{
        HivePath = $destinationHivePath
        TransactionLogCount = $transactionLogs.Count
        TransactionLogs = @($transactionLogs | ForEach-Object { $_.Name })
    }
}

function Find-IbisUsbEventLogCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$EventLogDirectory,

        [Parameter(Mandatory = $true)]
        [string]$CanonicalName
    )

    # Velociraptor percent-encodes the '%' in a channel name, so the on-disk name
    # for Microsoft-Windows-Partition%4Diagnostic.evtx becomes ...Partition%254Diagnostic.evtx.
    # Prefer the canonical name, then the encoded collection form.
    $candidates = @(
        [pscustomobject]@{ Name = $CanonicalName; DiscoveryMethod = 'Canonical name' },
        [pscustomobject]@{ Name = $CanonicalName.Replace('%', '%25'); DiscoveryMethod = 'Velociraptor encoded name' }
    )

    foreach ($candidate in $candidates) {
        $candidatePath = Join-Path $EventLogDirectory $candidate.Name
        if (Test-Path -LiteralPath $candidatePath -PathType Leaf) {
            return [pscustomobject]@{
                CanonicalName = $CanonicalName
                SourceName = $candidate.Name
                SourcePath = $candidatePath
                DiscoveryMethod = $candidate.DiscoveryMethod
            }
        }
    }

    $null
}

function Copy-IbisParseUsbEvidenceToStaging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot,

        [Parameter(Mandatory = $true)]
        [string]$StagingDirectory
    )

    $sourceConfigDirectory = Join-Path $SourceRoot 'Windows\System32\config'
    $stagingConfigDirectory = Join-Path $StagingDirectory 'Windows\System32\config'
    if (-not (Test-Path -LiteralPath $sourceConfigDirectory -PathType Container)) {
        return [pscustomobject]@{
            Status = 'Skipped'
            SourceConfigDirectory = $sourceConfigDirectory
            StagingDirectory = $StagingDirectory
            StagingConfigDirectory = $stagingConfigDirectory
            FileCount = 0
            Files = @()
            Message = 'Windows registry configuration directory was not found.'
        }
    }

    if (-not (Test-Path -LiteralPath $stagingConfigDirectory)) {
        New-Item -ItemType Directory -Path $stagingConfigDirectory -Force | Out-Null
    }

    # Copy all top-level configuration files so each hive retains its transaction logs
    # and any auxiliary registry files ParseUSBs may need, while never exposing source
    # evidence to a tool that can replay transactions in place.
    $sourceFiles = @(Get-ChildItem -LiteralPath $sourceConfigDirectory -File -Force -ErrorAction Stop)
    foreach ($sourceFile in $sourceFiles) {
        Copy-IbisEvidenceFile -SourcePath $sourceFile.FullName -DestinationPath (Join-Path $stagingConfigDirectory $sourceFile.Name) | Out-Null
    }

    # ParseUSBs volume mode expects a Windows-volume layout.  Preserve only the
    # artefacts it uses for additional USB context rather than copying whole profiles.
    $stagedUserHives = @()
    $stagedLnkFiles = @()
    $profiles = @(Get-IbisUserProfile -SourceRoot $SourceRoot)
    foreach ($profile in $profiles) {
        $stagedProfileDirectory = Join-Path (Join-Path $StagingDirectory 'Users') $profile.UserName
        if (Test-Path -LiteralPath $profile.NtUserPath -PathType Leaf) {
            $cachedHive = Copy-IbisRegistryHiveToCache -SourceHivePath $profile.NtUserPath -CacheDirectory $stagedProfileDirectory
            $stagedUserHives += [pscustomobject]@{
                UserName = $profile.UserName
                SourceHivePath = $profile.NtUserPath
                HivePath = $cachedHive.HivePath
                TransactionLogCount = $cachedHive.TransactionLogCount
                TransactionLogs = $cachedHive.TransactionLogs
                IsDefaultProfile = ($profile.UserName -like 'Default*')
            }
        }

        # LNK files can exist outside Recent.  Copy the LNK files only, retaining
        # their relative locations so ParseUSBs can report the original profile path.
        $sourceLnkFiles = @(Get-ChildItem -LiteralPath $profile.ProfilePath -Filter '*.lnk' -File -Recurse -Force -ErrorAction SilentlyContinue)
        foreach ($sourceLnkFile in $sourceLnkFiles) {
            $relativePath = $sourceLnkFile.FullName.Substring($profile.ProfilePath.Length).TrimStart('\', '/')
            $destinationLnkPath = Join-Path $stagedProfileDirectory $relativePath
            $destinationLnkDirectory = Split-Path -Path $destinationLnkPath -Parent
            if (-not (Test-Path -LiteralPath $destinationLnkDirectory)) {
                New-Item -ItemType Directory -Path $destinationLnkDirectory -Force | Out-Null
            }
            Copy-IbisEvidenceFile -SourcePath $sourceLnkFile.FullName -DestinationPath $destinationLnkPath | Out-Null
            $stagedLnkFiles += [pscustomobject]@{ UserName = $profile.UserName; SourcePath = $sourceLnkFile.FullName; StagingPath = $destinationLnkPath }
        }
    }

    $stagedEventLogs = @()
    $eventLogDirectory = Join-Path $SourceRoot 'Windows\System32\winevt\Logs'
    $stagingEventLogDirectory = Join-Path $StagingDirectory 'Windows\System32\winevt\Logs'
    $usbEventLogNames = @('Microsoft-Windows-Partition%4Diagnostic.evtx', 'Microsoft-Windows-Storsvc%4Diagnostic.evtx')
    foreach ($eventLogName in $usbEventLogNames) {
        $candidate = Find-IbisUsbEventLogCandidate -EventLogDirectory $eventLogDirectory -CanonicalName $eventLogName
        if ($null -ne $candidate) {
            if (-not (Test-Path -LiteralPath $stagingEventLogDirectory)) {
                New-Item -ItemType Directory -Path $stagingEventLogDirectory -Force | Out-Null
            }

            # Stage under the canonical Windows name whatever the collection called it,
            # because ParseUSBs looks for that exact name during volume processing.
            $stagingEventLogPath = Join-Path $stagingEventLogDirectory $eventLogName
            Copy-IbisEvidenceFile -SourcePath $candidate.SourcePath -DestinationPath $stagingEventLogPath | Out-Null
            $stagedEventLogs += [pscustomobject]@{
                Name = $eventLogName
                SourceName = $candidate.SourceName
                DiscoveryMethod = $candidate.DiscoveryMethod
                SourcePath = $candidate.SourcePath
                StagingPath = $stagingEventLogPath
            }
        }
    }

    [pscustomobject]@{
        Status = 'Staged'
        SourceConfigDirectory = $sourceConfigDirectory
        StagingDirectory = $StagingDirectory
        StagingConfigDirectory = $stagingConfigDirectory
        FileCount = $sourceFiles.Count
        Files = @($sourceFiles | ForEach-Object { $_.Name })
        SystemHivePath = Join-Path $stagingConfigDirectory 'SYSTEM'
        SoftwareHivePath = Join-Path $stagingConfigDirectory 'SOFTWARE'
        UserHives = $stagedUserHives
        UserHiveCount = $stagedUserHives.Count
        LnkFiles = $stagedLnkFiles
        LnkFileCount = $stagedLnkFiles.Count
        EventLogs = $stagedEventLogs
        EventLogCount = $stagedEventLogs.Count
        Message = 'Copied ParseUSBs registry, user-hive, LNK, and supported event-log evidence to the staging area.'
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

        $sourceItem = Get-Item -LiteralPath $SourceHivePath -Force -ErrorAction Stop
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

    $sourceItem = Get-Item -LiteralPath $SourceHivePath -Force
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
    try {
        [void]$process.Start()
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        $stdoutTask.Wait()
        $stderrTask.Wait()
        $stdout = $stdoutTask.Result
        $stderr = $stderrTask.Result
        $exitCode = $process.ExitCode
    }
    finally {
        if ($null -ne $process) {
            $process.Dispose()
        }
    }

    [pscustomobject]@{
        FilePath = $FilePath
        Arguments = @($ArgumentList)
        CommandLine = $commandLine
        WorkingDirectory = $WorkingDirectory
        ExitCode = $exitCode
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
    elseif ((Test-Path -LiteralPath $OutputPath -PathType Leaf) -and ((Get-Item -LiteralPath $OutputPath -Force).Length -eq 0)) {
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
    elseif ((Get-Item -LiteralPath $outputPath -Force).Length -eq 0) {
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
    elseif ((Get-Item -LiteralPath $outputPath -Force).Length -eq 0) {
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

function Find-IbisUsnJournalCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    $candidatePaths = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $addCandidate = {
        param(
            [string]$Path,
            [string]$DiscoveryMethod,
            [string]$DiscoveryConfidence = 'High',
            [bool]$AllowMountedStreamBase = $false
        )

        if ([string]::IsNullOrWhiteSpace($Path)) {
            return
        }

        $key = $Path.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $candidatePaths.Add([pscustomobject]@{
                Path = $Path
                DiscoveryMethod = $DiscoveryMethod
                DiscoveryConfidence = $DiscoveryConfidence
                AllowMountedStreamBase = $AllowMountedStreamBase
            })
        }
    }

    $addExtractedCandidates = {
        param([string]$ExtendDirectory)

        if (-not (Test-Path -LiteralPath $ExtendDirectory -PathType Container)) {
            return
        }

        foreach ($candidateFile in @(Get-ChildItem -LiteralPath $ExtendDirectory -File -Force -ErrorAction SilentlyContinue | Sort-Object Name)) {
            if ($candidateFile.Name -match '^\$UsnJrnl(?:[_\-.])\$J(?:\.[^\\/]+)?$') {
                & $addCandidate $candidateFile.FullName 'RenamedUsnJournalStream' 'Medium'
            }
            elseif ($candidateFile.Name -match '^\$J(?:\.[^\\/]+)?$') {
                & $addCandidate $candidateFile.FullName 'StandaloneJStream' 'Medium'
            }
            elseif ($candidateFile.Name -ieq '$UsnJrnl') {
                & $addCandidate $candidateFile.FullName 'FlattenedUsnJournalStream' 'Low'
            }
        }
    }

    try {
        $sourceRootPath = [System.IO.Path]::GetFullPath($SourceRoot)
        $rootPath = [System.IO.Path]::GetPathRoot($sourceRootPath)
        if ([string]::Equals($sourceRootPath.TrimEnd('\', '/'), $rootPath.TrimEnd('\', '/'), [System.StringComparison]::OrdinalIgnoreCase)) {
            $drive = [System.IO.Path]::GetPathRoot($sourceRootPath).TrimEnd('\', '/')
            if (-not [string]::Equals($drive, 'C:', [System.StringComparison]::OrdinalIgnoreCase)) {
                & $addCandidate ([System.IO.Path]::Combine($SourceRoot, '$Extend\$UsnJrnl:$J')) 'MountedNtfsJournalStream' 'High' $true
                & $addExtractedCandidates ([System.IO.Path]::Combine($SourceRoot, '$Extend'))
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

        & $addCandidate ([System.IO.Path]::Combine($current, 'uploads\ntfs\%5C%5C.%5CC%3A\$Extend\$UsnJrnl%3A$J')) 'VelociraptorEncodedJournalStream'
        & $addCandidate ([System.IO.Path]::Combine($current, 'ntfs\%5C%5C.%5CC%3A\$Extend\$UsnJrnl%3A$J')) 'NtfsUploadEncodedJournalStream'
        & $addExtractedCandidates ([System.IO.Path]::Combine($current, '$Extend'))

        foreach ($ntfsRoot in @(
            [System.IO.Path]::Combine($current, 'uploads\ntfs'),
            [System.IO.Path]::Combine($current, 'ntfs')
        )) {
            if (Test-Path -LiteralPath $ntfsRoot -PathType Container) {
                foreach ($deviceDirectory in @(Get-ChildItem -LiteralPath $ntfsRoot -Directory -Force -ErrorAction SilentlyContinue)) {
                    $extendDirectory = [System.IO.Path]::Combine($deviceDirectory.FullName, '$Extend')
                    & $addCandidate ([System.IO.Path]::Combine($extendDirectory, '$UsnJrnl%3A$J')) 'EncodedJournalStream'
                    & $addCandidate ([System.IO.Path]::Combine($extendDirectory, '$UsnJrnl:$J')) 'LiteralJournalStream'
                    & $addExtractedCandidates $extendDirectory
                }
            }
        }

        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) {
            break
        }
        $current = $parent
    }

    foreach ($candidate in $candidatePaths) {
        $found = if ($candidate.AllowMountedStreamBase) {
            Test-IbisNtfsSpecialFilePath -Path $candidate.Path
        }
        else {
            Test-Path -LiteralPath $candidate.Path -PathType Leaf
        }

        if ($found) {
            return $candidate
        }
    }

    $null
}

function Find-IbisUsnJournalPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceRoot
    )

    $candidate = Find-IbisUsnJournalCandidate -SourceRoot $SourceRoot
    if ($null -ne $candidate) {
        return $candidate.Path
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
    $usnJournalCandidate = Find-IbisUsnJournalCandidate -SourceRoot $SourceRoot
    $usnJournalPath = if ($null -ne $usnJournalCandidate) { $usnJournalCandidate.Path } else { $null }
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
            DiscoveryMethod = if ($null -ne $usnJournalCandidate) { $usnJournalCandidate.DiscoveryMethod } else { $null }
            DiscoveryConfidence = if ($null -ne $usnJournalCandidate) { $usnJournalCandidate.DiscoveryConfidence } else { $null }
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
                "USN Journal `$J and `$MFT were found via $($usnJournalCandidate.DiscoveryMethod) discovery."
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
            UserName = $UserName
            Artifact = $Description
            Operation = 'Process source directory'
            SourceDirectory = $SourceDirectory
            OutputDirectory = $OutputDirectory
            Status = 'Skipped'
            ExitCode = $null
            Message = "$Description source directory was not found for user $UserName."
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
            UserName = $UserName
            Artifact = $Description
            Operation = 'Process source directory'
            SourceDirectory = $SourceDirectory
            OutputDirectory = $OutputDirectory
            Status = 'Failed'
            ExitCode = $null
            Message = "$Description tool is not configured for user $UserName."
        }
    }

    $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        return [pscustomobject]@{
            ToolId = $tool.id
            Description = $Description
            UserName = $UserName
            Artifact = $Description
            Operation = 'Process source directory'
            SourceDirectory = $SourceDirectory
            OutputDirectory = $OutputDirectory
            Status = 'Failed'
            ExitCode = $null
            Message = "$Description tool is missing at: $toolPath for user $UserName."
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
        UserName = $UserName
        Artifact = $Description
        Operation = 'Process source directory'
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
            ToolId = 'psreadline'
            Description = 'PSReadLine'
            UserName = $UserName
            Artifact = 'PSReadLine'
            Operation = 'Copy history'
            SourceDirectory = $SourceDirectory
            OutputDirectory = $OutputDirectory
            Status = 'Skipped'
            CopiedItemCount = 0
            Message = "PSReadLine source directory was not found for user $UserName."
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
        $copiedItems += Get-Item -LiteralPath $destinationPath -Force
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
        ToolId = 'psreadline'
        Description = 'PSReadLine'
        UserName = $UserName
        Artifact = 'PSReadLine'
        Operation = 'Copy history'
        SourceDirectory = $SourceDirectory
        OutputDirectory = $OutputDirectory
        DestinationDirectory = $OutputDirectory
        Status = $status
        CopiedItemCount = $copiedItems.Count
        Message = $message
    }
}

function New-IbisUserArtifactIssue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserName,

        [Parameter(Mandatory = $true)]
        [string]$Artifact,

        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Failed', 'Skipped')]
        [string]$Status,

        [Parameter(Mandatory = $true)]
        [string]$Message,

        [string]$ToolId,

        [string]$Description,

        [string]$SourcePath,

        [string]$SourceDirectory,

        [string]$OutputPath,

        [string]$OutputDirectory
    )

    if ([string]::IsNullOrWhiteSpace($ToolId)) {
        $ToolId = (@($UserName, $Artifact, $Operation) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join '/'
    }
    if ([string]::IsNullOrWhiteSpace($Description)) {
        $Description = $Operation
    }

    [pscustomobject]@{
        ToolId = $ToolId
        Description = $Description
        UserName = $UserName
        Artifact = $Artifact
        Operation = $Operation
        SourcePath = $SourcePath
        SourceDirectory = $SourceDirectory
        OutputPath = $OutputPath
        OutputDirectory = $OutputDirectory
        Status = $Status
        ExitCode = $null
        Message = $Message
    }
}

function Add-IbisUserArtifactResultContext {
    [CmdletBinding()]
    param(
        [object]$Result,

        [string]$ToolId,

        [string]$Description,

        [Parameter(Mandatory = $true)]
        [string]$UserName,

        [Parameter(Mandatory = $true)]
        [string]$Artifact,

        [Parameter(Mandatory = $true)]
        [string]$Operation
    )

    foreach ($item in @($Result)) {
        if ($null -eq $item -or $item -is [string]) {
            continue
        }

        $propertyNames = @($item.PSObject.Properties.Name)
        if (-not [string]::IsNullOrWhiteSpace($ToolId) -and -not ($propertyNames -contains 'ToolId')) {
            $item | Add-Member -NotePropertyName ToolId -NotePropertyValue $ToolId
        }
        if (-not [string]::IsNullOrWhiteSpace($Description) -and -not ($propertyNames -contains 'Description')) {
            $item | Add-Member -NotePropertyName Description -NotePropertyValue $Description
        }
        if (-not ($propertyNames -contains 'UserName')) {
            $item | Add-Member -NotePropertyName UserName -NotePropertyValue $UserName
        }
        if (-not ($propertyNames -contains 'Artifact')) {
            $item | Add-Member -NotePropertyName Artifact -NotePropertyValue $Artifact
        }
        if (-not ($propertyNames -contains 'Operation')) {
            $item | Add-Member -NotePropertyName Operation -NotePropertyValue $Operation
        }
    }

    $Result
}

function Write-IbisUserArtifactProgressResult {
    [CmdletBinding()]
    param(
        [object]$Result,

        [string]$ProgressPath,

        [int]$ProgressIndex = 0,

        [int]$ProgressTotal = 0
    )

    if ([string]::IsNullOrWhiteSpace($ProgressPath) -or $null -eq $Result) {
        return
    }

    foreach ($item in @($Result)) {
        if ($null -eq $item -or $item -is [string]) {
            continue
        }

        $userName = [string]$item.UserName
        $artifact = [string]$item.Artifact
        $operation = [string]$item.Operation
        $status = [string]$item.Status
        $message = [string]$item.Message
        if ([string]::IsNullOrWhiteSpace($status)) {
            $status = 'Info'
        }
        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = 'No message was provided.'
        }

        $context = @($userName, $artifact, $operation) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }
        $stage = 'User artefact'
        if ($context.Count -gt 0) {
            $stage = $context -join ' / '
        }

        Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'user-artifacts' -ToolName 'User artefacts' -Stage $stage -Message "${stage}: ${status} - ${message}" -Index $ProgressIndex -Total $ProgressTotal -Status $status

        if (-not [string]::IsNullOrWhiteSpace([string]$item.CommandLine)) {
            Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'user-artifacts' -ToolName 'User artefacts' -Stage 'Command line hint' -Message "${stage}: $($item.CommandLine)" -Index $ProgressIndex -Total $ProgressTotal -Status 'Info'
        }

        foreach ($renamedOutput in @($item.RenamedOutputs)) {
            if ($null -eq $renamedOutput) {
                continue
            }
            $originalPath = [string]$renamedOutput.OriginalPath
            $newPath = [string]$renamedOutput.NewPath
            if (-not [string]::IsNullOrWhiteSpace($originalPath) -and -not [string]::IsNullOrWhiteSpace($newPath)) {
                Write-IbisProgressEvent -ProgressPath $ProgressPath -ToolId 'user-artifacts' -ToolName 'User artefacts' -Stage 'File operation' -Message "${stage}: File moved/renamed: $originalPath -> $newPath" -Index $ProgressIndex -Total $ProgressTotal -Status 'Audit'
            }
        }
    }
}

function Invoke-IbisUserArtifactStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory = $true)]
        [string]$UserName,

        [Parameter(Mandatory = $true)]
        [string]$Artifact,

        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [string]$ToolId,

        [string]$Description,

        [string]$SourcePath,

        [string]$SourceDirectory,

        [string]$OutputPath,

        [string]$OutputDirectory,

        [string]$ProgressPath,

        [int]$ProgressIndex = 0,

        [int]$ProgressTotal = 0
    )

    try {
        $result = & $ScriptBlock
        if ($null -eq $result) {
            $missingResult = New-IbisUserArtifactIssue -UserName $UserName -Artifact $Artifact -Operation $Operation -Status 'Failed' -ToolId $ToolId -Description $Description -SourcePath $SourcePath -SourceDirectory $SourceDirectory -OutputPath $OutputPath -OutputDirectory $OutputDirectory -Message "$Operation returned no result for user $UserName."
            Write-IbisUserArtifactProgressResult -Result $missingResult -ProgressPath $ProgressPath -ProgressIndex $ProgressIndex -ProgressTotal $ProgressTotal
            return $missingResult
        }

        $result = Add-IbisUserArtifactResultContext -Result $result -ToolId $ToolId -Description $Description -UserName $UserName -Artifact $Artifact -Operation $Operation
        Write-IbisUserArtifactProgressResult -Result $result -ProgressPath $ProgressPath -ProgressIndex $ProgressIndex -ProgressTotal $ProgressTotal
        $result
    }
    catch {
        $failure = New-IbisUserArtifactIssue -UserName $UserName -Artifact $Artifact -Operation $Operation -Status 'Failed' -ToolId $ToolId -Description $Description -SourcePath $SourcePath -SourceDirectory $SourceDirectory -OutputPath $OutputPath -OutputDirectory $OutputDirectory -Message "$Operation failed for user $UserName ($Artifact): $($_.Exception.Message)"
        Write-IbisUserArtifactProgressResult -Result $failure -ProgressPath $ProgressPath -ProgressIndex $ProgressIndex -ProgressTotal $ProgressTotal
        $failure
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

        [string]$Hostname = 'HOST',

        [string]$ProgressPath,

        [int]$ProgressIndex = 0,

        [int]$ProgressTotal = 0
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    $outputDirectory = Join-Path $hostOutputRoot 'Users'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $profiles = @(Get-IbisUserProfile -SourceRoot $SourceRoot)
    $progressSplat = @{
        ProgressPath = $ProgressPath
        ProgressIndex = $ProgressIndex
        ProgressTotal = $ProgressTotal
    }
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

        $preparedHives = @()
        $toolResults = @()
        $psReadLineResult = $null

        try {
            if (-not (Test-Path -LiteralPath $userOutputDirectory)) {
                New-Item -ItemType Directory -Path $userOutputDirectory -Force -ErrorAction Stop | Out-Null
            }
        }
        catch {
            $issue = New-IbisUserArtifactIssue -UserName $safeUserName -Artifact 'Profile' -Operation 'Create user output directory' -Status 'Failed' -ToolId "$safeUserName/Profile/Create output directory" -Description 'Create user output directory' -SourceDirectory $profile.ProfilePath -OutputDirectory $userOutputDirectory -Message "Unable to create user output directory for ${safeUserName}: $($_.Exception.Message)"
            $toolResults += $issue
            Write-IbisUserArtifactProgressResult -Result $issue @progressSplat
            $userResults += [pscustomobject]@{
                UserName = $profile.UserName
                ProfilePath = $profile.ProfilePath
                OutputDirectory = $userOutputDirectory
                PreparedHives = $preparedHives
                ToolResults = $toolResults
                PSReadLine = $psReadLineResult
            }
            continue
        }

        if (Test-Path -LiteralPath $profile.NtUserPath -PathType Leaf) {
            $ntUser = Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'Prepare hive' -ToolId "$safeUserName/NTUSER/Prepare hive" -Description 'Prepare NTUSER.dat' -SourcePath $profile.NtUserPath -OutputDirectory $workingsDirectory -ScriptBlock {
                Invoke-IbisPrepareRegistryHiveFile `
                    -ToolsRoot $ToolsRoot `
                    -ToolDefinitions $ToolDefinitions `
                    -SourceHivePath $profile.NtUserPath `
                    -OutputRoot $OutputRoot `
                    -Hostname $safeHost `
                    -HiveName 'NTUSER.dat' `
                    -CacheGroup 'Users' `
                    -CacheKey ($safeUserName + '-NTUSER')
            }
            $preparedHives += $ntUser

            if ($ntUser.Status -eq 'Failed' -or [string]::IsNullOrWhiteSpace($ntUser.PreparedHivePath)) {
                $skipMessage = 'NTUSER.dat could not be prepared, so dependent RegRipper operations were skipped.'
                $ntUserSkips = @(
                    New-IbisUserArtifactIssue -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'RegRipper all' -Status 'Skipped' -ToolId "$safeUserName/NTUSER/RegRipper all" -Description 'RegRipper NTUSER all' -SourcePath $profile.NtUserPath -OutputDirectory $userOutputDirectory -Message $skipMessage
                    New-IbisUserArtifactIssue -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'RegRipper timeline' -Status 'Skipped' -ToolId "$safeUserName/NTUSER/RegRipper timeline" -Description 'RegRipper NTUSER timeline' -SourcePath $profile.NtUserPath -OutputDirectory $userOutputDirectory -Message $skipMessage
                    New-IbisUserArtifactIssue -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'RegRipper run plugin' -Status 'Skipped' -ToolId "$safeUserName/NTUSER/RegRipper run" -Description 'RegRipper NTUSER run plugin' -SourcePath $profile.NtUserPath -OutputDirectory $userOutputDirectory -Message $skipMessage
                    New-IbisUserArtifactIssue -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'RegRipper userassist plugin' -Status 'Skipped' -ToolId "$safeUserName/NTUSER/RegRipper userassist" -Description 'RegRipper NTUSER userassist plugin' -SourcePath $profile.NtUserPath -OutputDirectory $userOutputDirectory -Message $skipMessage
                )
                $toolResults += $ntUserSkips
                Write-IbisUserArtifactProgressResult -Result $ntUserSkips @progressSplat
            }
            else {
                $ntUserAllOutput = Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-NTUSER.txt' -ArgumentList @($safeUserName))
                $ntUserTimelineOutput = Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-NTUSER-TLN.txt' -ArgumentList @($safeUserName))
                $ntUserRunOutput = Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-NTUSER-Run-AutoStart.txt' -ArgumentList @($safeUserName))
                $ntUserUserAssistOutput = Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-NTUSER-UserAssist.txt' -ArgumentList @($safeUserName))

                $toolResults += Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'RegRipper all' -ToolId "$safeUserName/NTUSER/RegRipper all" -Description 'RegRipper NTUSER all' -SourcePath $ntUser.PreparedHivePath -OutputPath $ntUserAllOutput -ScriptBlock {
                    Invoke-IbisRegRipperHiveMode -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $ntUser.PreparedHivePath -Mode 'All' -OutputPath $ntUserAllOutput
                }
                $toolResults += Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'RegRipper timeline' -ToolId "$safeUserName/NTUSER/RegRipper timeline" -Description 'RegRipper NTUSER timeline' -SourcePath $ntUser.PreparedHivePath -OutputPath $ntUserTimelineOutput -ScriptBlock {
                    Invoke-IbisRegRipperHiveMode -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $ntUser.PreparedHivePath -Mode 'Timeline' -OutputPath $ntUserTimelineOutput
                }
                $toolResults += Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'RegRipper run plugin' -ToolId "$safeUserName/NTUSER/RegRipper run" -Description 'RegRipper NTUSER run plugin' -SourcePath $ntUser.PreparedHivePath -OutputPath $ntUserRunOutput -ScriptBlock {
                    Invoke-IbisRegRipperPlugin -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $ntUser.PreparedHivePath -Plugin 'run' -OutputPath $ntUserRunOutput
                }
                $toolResults += Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'RegRipper userassist plugin' -ToolId "$safeUserName/NTUSER/RegRipper userassist" -Description 'RegRipper NTUSER userassist plugin' -SourcePath $ntUser.PreparedHivePath -OutputPath $ntUserUserAssistOutput -ScriptBlock {
                    Invoke-IbisRegRipperPlugin -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $ntUser.PreparedHivePath -Plugin 'userassist' -OutputPath $ntUserUserAssistOutput
                }
            }
        }
        else {
            $issue = New-IbisUserArtifactIssue -UserName $safeUserName -Artifact 'NTUSER.dat' -Operation 'Prepare hive' -Status 'Skipped' -ToolId "$safeUserName/NTUSER/Prepare hive" -Description 'Prepare NTUSER.dat' -SourcePath $profile.NtUserPath -OutputDirectory $workingsDirectory -Message "NTUSER.dat was not found for user $safeUserName."
            $preparedHives += $issue
            Write-IbisUserArtifactProgressResult -Result $issue @progressSplat
        }

        if (Test-Path -LiteralPath $profile.UsrClassPath -PathType Leaf) {
            $usrClass = Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'UsrClass.dat' -Operation 'Prepare hive' -ToolId "$safeUserName/UsrClass/Prepare hive" -Description 'Prepare UsrClass.dat' -SourcePath $profile.UsrClassPath -OutputDirectory $workingsDirectory -ScriptBlock {
                Invoke-IbisPrepareRegistryHiveFile `
                    -ToolsRoot $ToolsRoot `
                    -ToolDefinitions $ToolDefinitions `
                    -SourceHivePath $profile.UsrClassPath `
                    -OutputRoot $OutputRoot `
                    -Hostname $safeHost `
                    -HiveName 'UsrClass.dat' `
                    -CacheGroup 'Users' `
                    -CacheKey ($safeUserName + '-UsrClass')
            }
            $preparedHives += $usrClass

            if ($usrClass.Status -eq 'Failed' -or [string]::IsNullOrWhiteSpace($usrClass.PreparedHivePath)) {
                $skipMessage = 'UsrClass.dat could not be prepared, so dependent RegRipper operations were skipped.'
                $usrClassSkips = @(
                    New-IbisUserArtifactIssue -UserName $safeUserName -Artifact 'UsrClass.dat' -Operation 'RegRipper all' -Status 'Skipped' -ToolId "$safeUserName/UsrClass/RegRipper all" -Description 'RegRipper UsrClass all' -SourcePath $profile.UsrClassPath -OutputDirectory $userOutputDirectory -Message $skipMessage
                    New-IbisUserArtifactIssue -UserName $safeUserName -Artifact 'UsrClass.dat' -Operation 'RegRipper timeline' -Status 'Skipped' -ToolId "$safeUserName/UsrClass/RegRipper timeline" -Description 'RegRipper UsrClass timeline' -SourcePath $profile.UsrClassPath -OutputDirectory $userOutputDirectory -Message $skipMessage
                )
                $toolResults += $usrClassSkips
                Write-IbisUserArtifactProgressResult -Result $usrClassSkips @progressSplat
            }
            else {
                $usrClassAllOutput = Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-UsrClass.txt' -ArgumentList @($safeUserName))
                $usrClassTimelineOutput = Join-Path $userOutputDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'RR-{0}-UsrClass-TLN.txt' -ArgumentList @($safeUserName))

                $toolResults += Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'UsrClass.dat' -Operation 'RegRipper all' -ToolId "$safeUserName/UsrClass/RegRipper all" -Description 'RegRipper UsrClass all' -SourcePath $usrClass.PreparedHivePath -OutputPath $usrClassAllOutput -ScriptBlock {
                    Invoke-IbisRegRipperHiveMode -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $usrClass.PreparedHivePath -Mode 'All' -OutputPath $usrClassAllOutput
                }
                $toolResults += Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'UsrClass.dat' -Operation 'RegRipper timeline' -ToolId "$safeUserName/UsrClass/RegRipper timeline" -Description 'RegRipper UsrClass timeline' -SourcePath $usrClass.PreparedHivePath -OutputPath $usrClassTimelineOutput -ScriptBlock {
                    Invoke-IbisRegRipperHiveMode -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -HivePath $usrClass.PreparedHivePath -Mode 'Timeline' -OutputPath $usrClassTimelineOutput
                }
            }
        }
        else {
            $issue = New-IbisUserArtifactIssue -UserName $safeUserName -Artifact 'UsrClass.dat' -Operation 'Prepare hive' -Status 'Skipped' -ToolId "$safeUserName/UsrClass/Prepare hive" -Description 'Prepare UsrClass.dat' -SourcePath $profile.UsrClassPath -OutputDirectory $workingsDirectory -Message "UsrClass.dat was not found for user $safeUserName."
            $preparedHives += $issue
            Write-IbisUserArtifactProgressResult -Result $issue @progressSplat
        }

        $jumpListOutput = Join-Path $userOutputDirectory 'JumpLists'
        $recentLnkOutput = Join-Path $userOutputDirectory 'RecentLNKs'
        $shellBagOutput = Join-Path $userOutputDirectory 'ShellBags'
        $toolResults += Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'JumpLists' -Operation 'Run JLECmd' -ToolId 'zimmerman-jlecmd' -Description 'JLECmd' -SourceDirectory $profile.RecentPath -OutputDirectory $jumpListOutput -ScriptBlock {
            Invoke-IbisUserDirectoryTool -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -ToolId 'zimmerman-jlecmd' -SourceDirectory $profile.RecentPath -OutputDirectory $jumpListOutput -ArgumentList @('-d', $profile.RecentPath, '--all', '--csv', $jumpListOutput, '--html', $jumpListOutput, '-q', '--fd') -Description 'JLECmd' -Hostname $safeHost -UserName $safeUserName
        }
        $toolResults += Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'RecentLNKs' -Operation 'Run LECmd' -ToolId 'zimmerman-lecmd' -Description 'LECmd' -SourceDirectory $profile.RecentPath -OutputDirectory $recentLnkOutput -ScriptBlock {
            Invoke-IbisUserDirectoryTool -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -ToolId 'zimmerman-lecmd' -SourceDirectory $profile.RecentPath -OutputDirectory $recentLnkOutput -ArgumentList @('-d', $profile.RecentPath, '--all', '--csv', $recentLnkOutput, '-q') -Description 'LECmd' -Hostname $safeHost -UserName $safeUserName
        }
        $toolResults += Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'ShellBags' -Operation 'Run SBECmd' -ToolId 'zimmerman-sbecmd' -Description 'SBECmd' -SourceDirectory $profile.ProfilePath -OutputDirectory $shellBagOutput -ScriptBlock {
            Invoke-IbisUserDirectoryTool -ToolsRoot $ToolsRoot -ToolDefinitions $ToolDefinitions -ToolId 'zimmerman-sbecmd' -SourceDirectory $profile.ProfilePath -OutputDirectory $shellBagOutput -ArgumentList @('-d', $profile.ProfilePath, '--csv', $shellBagOutput) -Description 'SBECmd' -Hostname $safeHost -UserName $safeUserName
        }

        $psReadLineOutput = Join-Path $userOutputDirectory 'PSReadLine'
        $psReadLineResult = Invoke-IbisUserArtifactStep @progressSplat -UserName $safeUserName -Artifact 'PSReadLine' -Operation 'Copy history' -ToolId 'psreadline' -Description 'PSReadLine' -SourceDirectory $profile.PSReadLinePath -OutputDirectory $psReadLineOutput -ScriptBlock {
            Copy-IbisPSReadLineHistory -SourceDirectory $profile.PSReadLinePath -OutputDirectory $psReadLineOutput -Hostname $safeHost -UserName $safeUserName
        }

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
    $psReadLineAll = @($userResults | ForEach-Object { $_.PSReadLine } | Where-Object { $null -ne $_ })
    $allOperationResults = @($toolResultsAll + $preparedAll + $psReadLineAll)
    $failed = @($toolResultsAll | Where-Object { $_.Status -eq 'Failed' })
    $failed += @($preparedAll | Where-Object { $_.Status -eq 'Failed' })
    $failed += @($psReadLineAll | Where-Object { $_.Status -eq 'Failed' })
    $warnings = @($toolResultsAll | Where-Object { $_.Status -match 'Warnings' })
    $warnings += @($preparedAll | Where-Object { $_.Status -match 'Warnings' })
    $warnings += @($psReadLineAll | Where-Object { $_.Status -match 'Warnings' })
    $skipped = @($allOperationResults | Where-Object { $_.Status -eq 'Skipped' })
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
        Users = $userResults
        FailedItems = $failed
        SkippedItems = $skipped
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

function Get-IbisEventLogToolOutputDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [string]$ToolFolder
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    Join-Path (Join-Path $hostOutputRoot 'EventLogs') $ToolFolder
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
    $outputDirectory = Get-IbisEventLogToolOutputDirectory -OutputRoot $OutputRoot -Hostname $safeHost -ToolFolder 'EvtxECmd'
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

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'EvtxECmd-EventLogs.json')
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
    Join-Path (Get-IbisEventLogToolOutputDirectory -OutputRoot $OutputRoot -Hostname $safeHost -ToolFolder 'EvtxECmd') (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'EvtxECmd-EventLogs-Output.csv')
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
            OutputFileNameFormat = 'DuckDB-Event-Log-Time-Span-Info.csv'
        }
        [pscustomobject]@{
            Id = 'logons'
            Name = 'Event log user-session activity'
            QueryPath = Join-Path $queryRoot 'logons.sql'
            OutputFileNameFormat = 'DuckDB-Event-Log-User-Logons.csv'
        }
        [pscustomobject]@{
            Id = 'outbound-rdp'
            Name = 'Event log outbound RDP'
            QueryPath = Join-Path $queryRoot 'outbound-rdp.sql'
            OutputFileNameFormat = 'DuckDB-Event-Log-Outbound-RDP.csv'
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
    $outputDirectory = Get-IbisEventLogToolOutputDirectory -OutputRoot $OutputRoot -Hostname $safeHost -ToolFolder 'DuckDB'
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
    $runToken = [System.Guid]::NewGuid().ToString('N')
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
                $stagedOutputPath = Join-Path $workingsDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'DuckDB-{0}-{1}.staged.csv' -ArgumentList @($queryDefinition.Id, $runToken))
                if (-not (Test-Path -LiteralPath $queryDefinition.QueryPath -PathType Leaf)) {
                    $toolResults += [pscustomobject]@{ ToolId = $tool.id; QueryId = $queryDefinition.Id; QueryPath = $queryDefinition.QueryPath; OutputPath = $outputPath; StagedOutputPath = $null; Published = $false; RenderedSqlPath = $null; Status = 'Failed'; ExitCode = $null; Message = "DuckDB SQL template was not found: $($queryDefinition.QueryPath)" }
                    continue
                }

                $sql = Expand-IbisDuckDbSqlTemplate -TemplatePath $queryDefinition.QueryPath -InputCsvPath $evtxCsvPath -OutputCsvPath $stagedOutputPath
                $sql | Out-File -LiteralPath $renderedSqlPath -Encoding UTF8
                $processResult = Invoke-IbisProcessCapture -FilePath $toolPath -ArgumentList @('-c', $sql) -WorkingDirectory (Split-Path -Path $toolPath -Parent)
                $published = $false
                $publishError = $null
                if ($processResult.ExitCode -eq 0 -and (Test-Path -LiteralPath $stagedOutputPath -PathType Leaf)) {
                    try {
                        Move-Item -LiteralPath $stagedOutputPath -Destination $outputPath -Force -ErrorAction Stop
                        $published = $true
                    }
                    catch {
                        $publishError = $_.Exception.Message
                    }
                }
                if ($processResult.ExitCode -ne 0) {
                    $processResult.StandardError | Out-File -LiteralPath (Join-Path $workingsDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'DuckDB-{0}.stderr.txt' -ArgumentList @($queryDefinition.Id))) -Encoding UTF8
                }
                elseif (-not [string]::IsNullOrWhiteSpace($publishError)) {
                    $publishError | Out-File -LiteralPath (Join-Path $workingsDirectory (Format-IbisHostPrefixedValue -Hostname $safeHost -Format 'DuckDB-{0}.stderr.txt' -ArgumentList @($queryDefinition.Id))) -Encoding UTF8
                }

                $status = 'Completed'
                $message = "$($queryDefinition.Name) completed."
                if ($processResult.ExitCode -ne 0) {
                    $status = 'Failed'
                    $message = "$($queryDefinition.Name) exited with code $($processResult.ExitCode); any previous final output was preserved."
                }
                elseif (-not [string]::IsNullOrWhiteSpace($publishError)) {
                    $status = 'Failed'
                    $message = "$($queryDefinition.Name) completed, but its staged output could not be published: $publishError"
                }
                elseif (-not $published) {
                    $status = 'Completed With Warnings'
                    $message = "$($queryDefinition.Name) completed, but the expected staged CSV was not found; any previous final output was preserved."
                }

                $toolResults += [pscustomobject]@{
                    ToolId = $tool.id
                    QueryId = $queryDefinition.Id
                    QueryPath = $queryDefinition.QueryPath
                    InputPath = $evtxCsvPath
                    OutputPath = $outputPath
                    StagedOutputPath = $stagedOutputPath
                    Published = $published
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
    $outputDirectory = Get-IbisEventLogToolOutputDirectory -OutputRoot $OutputRoot -Hostname $safeHost -ToolFolder 'Hayabusa'
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
    Join-Path (Get-IbisEventLogToolOutputDirectory -OutputRoot $OutputRoot -Hostname $safeHost -ToolFolder 'Hayabusa') (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Hayabusa-EventLogs-SuperVerbose.jsonl')
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
    $takajoOutputDirectory = Get-IbisEventLogToolOutputDirectory -OutputRoot $OutputRoot -Hostname $safeHost -ToolFolder 'Takajo'
    $eventLogsDirectory = Split-Path -Path $takajoOutputDirectory -Parent
    # Takajo requires a non-existent output directory. Keep a backup workspace outside the
    # result directory while it runs and remove it when it has no material to preserve.
    $temporaryWorkspaceRoot = Join-Path $eventLogsDirectory '_Working'
    $temporaryWorkingsDirectory = Join-Path $temporaryWorkspaceRoot 'Takajo'
    $workingsDirectory = Join-Path $takajoOutputDirectory '_Working'
    $hayabusaJsonlPath = Get-IbisHayabusaJsonlPath -OutputRoot $OutputRoot -Hostname $safeHost
    $hasExistingTakajoDirectory = Test-Path -LiteralPath $takajoOutputDirectory -PathType Container
    $hasExistingTakajoOutput = $hasExistingTakajoDirectory -and @((Get-ChildItem -LiteralPath $takajoOutputDirectory -Force)).Count -gt 0

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

    if (-not (Test-Path -LiteralPath $temporaryWorkingsDirectory)) { New-Item -ItemType Directory -Path $temporaryWorkingsDirectory -Force | Out-Null }
    if ($hasExistingTakajoDirectory -and -not $hasExistingTakajoOutput) {
        Remove-Item -LiteralPath $takajoOutputDirectory -Force
    }

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
            $takajoBackup = if ($hasExistingTakajoOutput) { Move-IbisExistingDirectoryToBackup -DirectoryPath $takajoOutputDirectory -BackupRoot (Join-Path $temporaryWorkingsDirectory 'Takajo-Backups') } else { $null }
            $takajoResult = Invoke-IbisProcessCapture -FilePath $takajoPath -ArgumentList @('automagic', '-t', $hayabusaJsonlPath, '-o', $takajoOutputDirectory, '-s') -WorkingDirectory (Split-Path -Path $takajoPath -Parent)
            $takajoStatus = 'Completed'
            $takajoMessage = 'Takajo automagic completed.'
            if ($takajoResult.ExitCode -ne 0) { $takajoStatus = 'Failed'; $takajoMessage = "Takajo automagic exited with code $($takajoResult.ExitCode)." }
            elseif (-not (Test-Path -LiteralPath $takajoOutputDirectory -PathType Container)) { $takajoStatus = 'Completed With Warnings'; $takajoMessage = 'Takajo automagic completed, but the output folder was not found.' }
            $renamedAutomagicOutputs = @(Rename-IbisToolOutputFiles -SourceDirectory $takajoOutputDirectory -OutputDirectory $takajoOutputDirectory -Hostname $safeHost -ToolName 'Takajo')
            $automagicToolResult = [pscustomobject]@{ ToolId = $takajo.id; Mode = 'automagic'; OutputPath = $takajoOutputDirectory; BackupPath = $takajoBackup; RenamedOutputs = $renamedAutomagicOutputs; Status = $takajoStatus; ExitCode = $takajoResult.ExitCode; CommandLine = $takajoResult.CommandLine; Message = $takajoMessage }
            $toolResults += $automagicToolResult

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
            $renamedFinalOutputs = @(Rename-IbisToolOutputFiles -SourceDirectory $takajoOutputDirectory -OutputDirectory $takajoOutputDirectory -Hostname $safeHost -ToolName 'Takajo')
            $automagicToolResult.RenamedOutputs = @($automagicToolResult.RenamedOutputs) + $renamedFinalOutputs
            $renameFailures = @($automagicToolResult.RenamedOutputs | Where-Object { $_.Status -eq 'Failed' })
            if ($renameFailures.Count -gt 0 -and $automagicToolResult.Status -ne 'Failed') {
                $automagicToolResult.Status = 'Completed With Warnings'
                $automagicToolResult.Message = "Takajo automagic completed, but $($renameFailures.Count) file(s) could not be renamed. Check Microsoft Defender exclusions or the file lock details in the summary JSON."
            }
        }
    }

    if (-not (Test-Path -LiteralPath $takajoOutputDirectory)) {
        New-Item -ItemType Directory -Path $takajoOutputDirectory -Force | Out-Null
    }
    $hasTemporaryWorkspaceContent = (Test-Path -LiteralPath $temporaryWorkingsDirectory) -and @((Get-ChildItem -LiteralPath $temporaryWorkingsDirectory -Force)).Count -gt 0
    $movedWorkspaceToFinalDirectory = $false
    if ($hasTemporaryWorkspaceContent) {
        if (-not (Test-Path -LiteralPath $workingsDirectory)) {
            Move-Item -LiteralPath $temporaryWorkingsDirectory -Destination $workingsDirectory -Force
            $movedWorkspaceToFinalDirectory = $true
        }
        else {
            Get-ChildItem -LiteralPath $temporaryWorkingsDirectory -Force | ForEach-Object {
                $destinationPath = Join-Path $workingsDirectory $_.Name
                if (Test-Path -LiteralPath $destinationPath) {
                    $destinationPath = Join-Path $workingsDirectory ('Ibis-{0}-{1}' -f $_.Name, (Get-Date).ToString('yyyyMMdd-HHmmss'))
                }
                Move-Item -LiteralPath $_.FullName -Destination $destinationPath -Force
            }
            Remove-Item -LiteralPath $temporaryWorkingsDirectory -Force
        }
    }
    elseif (Test-Path -LiteralPath $temporaryWorkingsDirectory) {
        Remove-Item -LiteralPath $temporaryWorkingsDirectory -Force
    }
    if ($movedWorkspaceToFinalDirectory) {
        $temporaryBackupRoot = Join-Path $temporaryWorkingsDirectory 'Takajo-Backups'
        $finalBackupRoot = Join-Path $workingsDirectory 'Takajo-Backups'
        foreach ($toolResult in $toolResults) {
            if ($toolResult.PSObject.Properties['BackupPath'] -and $toolResult.BackupPath -and $toolResult.BackupPath.StartsWith($temporaryBackupRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                $toolResult.BackupPath = $finalBackupRoot + $toolResult.BackupPath.Substring($temporaryBackupRoot.Length)
            }
        }
    }
    if ((Test-Path -LiteralPath $temporaryWorkspaceRoot) -and @((Get-ChildItem -LiteralPath $temporaryWorkspaceRoot -Force)).Count -eq 0) {
        Remove-Item -LiteralPath $temporaryWorkspaceRoot -Force
    }

    $summaryPath = Join-Path $takajoOutputDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'Takajo.json')
    $payload = [pscustomobject]@{
        ModuleId = 'takajo'
        Created = (Get-Date).ToString('s')
        HayabusaJsonlPath = $hayabusaJsonlPath
        ToolsRoot = $ToolsRoot
        HostOutputRoot = $hostOutputRoot
        OutputDirectory = $takajoOutputDirectory
        WorkingsDirectory = if (Test-Path -LiteralPath $workingsDirectory) { $workingsDirectory } else { $null }
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

function Rename-IbisToolOutputFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [string]$ToolName
    )

    if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
        return @()
    }
    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $filePrefix = New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix ($ToolName + '-')
    $sameDirectory = (Resolve-IbisComparablePath -Path $SourceDirectory) -eq (Resolve-IbisComparablePath -Path $OutputDirectory)
    $moved = @()
    $files = @(Get-ChildItem -LiteralPath $SourceDirectory -File -Recurse -Force -ErrorAction SilentlyContinue | Where-Object {
        $relativePath = $_.FullName.Substring($SourceDirectory.Length).TrimStart('\', '/')
        $pathParts = $relativePath -split '[\\/]'
        -not ($pathParts -contains '_Working')
    })
    foreach ($file in $files) {
        $relativeDirectory = $file.DirectoryName.Substring($SourceDirectory.Length).TrimStart('\', '/')
        $destinationDirectory = if ([string]::IsNullOrWhiteSpace($relativeDirectory)) { $OutputDirectory } else { Join-Path $OutputDirectory $relativeDirectory }
        if (-not (Test-Path -LiteralPath $destinationDirectory)) {
            New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
        }
        $destinationName = if ($file.Name.StartsWith($filePrefix, [System.StringComparison]::OrdinalIgnoreCase)) { $file.Name } else { $filePrefix + $file.Name }
        $destinationPath = Join-Path $destinationDirectory $destinationName
        if ($sameDirectory -and (Resolve-IbisComparablePath -Path $file.FullName) -eq (Resolve-IbisComparablePath -Path $destinationPath)) {
            continue
        }
        try {
            Move-Item -LiteralPath $file.FullName -Destination $destinationPath -Force -ErrorAction Stop
            $moved += [pscustomobject]@{ OriginalPath = $file.FullName; NewPath = $destinationPath; Status = 'Renamed'; Error = $null }
        }
        catch {
            # Defender or another process can temporarily block one result file. Preserve it
            # and let the calling module finish its remaining output and cleanup.
            $moved += [pscustomobject]@{ OriginalPath = $file.FullName; NewPath = $destinationPath; Status = 'Failed'; Error = $_.Exception.Message }
        }
    }

    if (-not $sameDirectory) {
        $remaining = @(Get-ChildItem -LiteralPath $SourceDirectory -Force -ErrorAction SilentlyContinue)
        if ($remaining.Count -eq 0) {
            Remove-Item -LiteralPath $SourceDirectory -Force
        }
    }

    $moved
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

    Rename-IbisToolOutputFiles -SourceDirectory $StagingDirectory -OutputDirectory $OutputDirectory -Hostname $Hostname -ToolName 'Chainsaw'
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
    $outputDirectory = Get-IbisEventLogToolOutputDirectory -OutputRoot $OutputRoot -Hostname $safeHost -ToolFolder 'Chainsaw'
    $workingsDirectory = Join-Path $outputDirectory '_Working'
    $chainsawStagingDirectory = Join-Path $workingsDirectory 'Staging'

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

function Get-IbisWebHistoryToolOutputDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Hostname,

        [Parameter(Mandatory = $true)]
        [string]$ToolFolder
    )

    $safeHost = ConvertTo-IbisSafeFileName -Value $Hostname -DefaultValue ''
    $hostOutputRoot = Get-IbisHostOutputRoot -OutputRoot $OutputRoot -Hostname $safeHost
    Join-Path (Join-Path $hostOutputRoot 'WebHistory') $ToolFolder
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
    $outputDirectory = Get-IbisWebHistoryToolOutputDirectory -OutputRoot $OutputRoot -Hostname $safeHost -ToolFolder 'BrowsingHistoryView'
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
            elseif ((Get-Item -LiteralPath $outputPath -Force).Length -eq 0) {
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
    $outputDirectory = Get-IbisWebHistoryToolOutputDirectory -OutputRoot $OutputRoot -Hostname $safeHost -ToolFolder 'ForensicWebHistory'
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

function Get-IbisParseUsbArgumentList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Staging,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory
    )

    # Volume mode only.  ParseUSBs 1.8 crashes in explicit hive mode (-s/-w/-u) with
    # 'NameError: name ... is not defined', even with the tool's own documented syntax,
    # so that pass could never produce output.  Volume mode covers the same registry
    # hives and adds the USB event logs and user LNK files.
    if (-not (Test-Path -LiteralPath $Staging.StagingDirectory -PathType Container)) {
        throw 'The ParseUSBs staging directory was not found.'
    }
    if (-not (Test-Path -LiteralPath $Staging.SystemHivePath -PathType Leaf)) {
        throw 'The staged SYSTEM hive was not found.'
    }

    @('-v', $Staging.StagingDirectory.TrimEnd('\', '/'), '-o', 'csv', '-d', $OutputDirectory)
}

function Move-IbisParseUsbOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string]$StagingDirectory,
        [Parameter(Mandatory = $true)] [string]$OutputDirectory,
        [Parameter(Mandatory = $true)] [AllowEmptyString()] [string]$Hostname
    )

    if (-not (Test-Path -LiteralPath $StagingDirectory -PathType Container)) { return @() }
    if (-not (Test-Path -LiteralPath $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null }
    $moved = @()
    foreach ($file in @(Get-ChildItem -LiteralPath $StagingDirectory -File -Recurse -Force -ErrorAction SilentlyContinue)) {
        $newName = New-IbisHostPrefixedFileName -Hostname $Hostname -Suffix ("ParseUSBs-{0}" -f $file.Name)
        $destinationPath = Join-Path $OutputDirectory $newName
        Move-Item -LiteralPath $file.FullName -Destination $destinationPath -Force
        $moved += [pscustomobject]@{ OriginalPath = $file.FullName; NewPath = $destinationPath }
    }
    $moved
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
    $stagingDirectory = Join-Path $workingsDirectory 'ParseUSBs-Evidence-Staging'
    $stagingBackupRoot = Join-Path $workingsDirectory 'ParseUSBs-Evidence-Staging-Backups'
    $stagingResult = $null
    $stagingBackupPath = $null

    if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
        return [pscustomobject]@{ ModuleId = 'usb'; Status = 'Skipped'; SourceRoot = $SourceRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; JsonPath = $null; Message = 'Evidence source root was not found.' }
    }

    if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $workingsDirectory)) { New-Item -ItemType Directory -Path $workingsDirectory -Force | Out-Null }

    $tool = Get-IbisToolDefinitionById -ToolDefinitions $ToolDefinitions -Id 'parseusbs'
    if ($null -eq $tool) {
        $toolResult = [pscustomobject]@{ ToolId = 'parseusbs'; SourceRoot = $SourceRoot; OutputDirectory = $outputDirectory; Status = 'Failed'; ExitCode = $null; Message = 'parseusbs is not configured.' }
    }
    else {
        $toolPath = Get-IbisToolExpectedPath -ToolsRoot $ToolsRoot -ToolDefinition $tool
        if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
            $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceRoot = $SourceRoot; OutputDirectory = $outputDirectory; Status = 'Failed'; ExitCode = $null; Message = "parseusbs is missing at: $toolPath" }
        }
        else {
            try {
                $stagingBackupPath = Move-IbisExistingDirectoryToBackup -DirectoryPath $stagingDirectory -BackupRoot $stagingBackupRoot
                $stagingResult = Copy-IbisParseUsbEvidenceToStaging -SourceRoot $SourceRoot -StagingDirectory $stagingDirectory
                if ($stagingResult.Status -ne 'Staged') {
                    throw $stagingResult.Message
                }

                $stagedOutputDirectory = Join-Path $workingsDirectory 'ParseUSBs-Output'
                $stagedOutputBackupPath = Move-IbisExistingDirectoryToBackup -DirectoryPath $stagedOutputDirectory -BackupRoot (Join-Path $workingsDirectory 'ParseUSBs-Output-Backups')
                if (-not (Test-Path -LiteralPath $stagedOutputDirectory)) { New-Item -ItemType Directory -Path $stagedOutputDirectory -Force | Out-Null }

                $arguments = Get-IbisParseUsbArgumentList -Staging $stagingResult -OutputDirectory $stagedOutputDirectory
                $processResult = Invoke-IbisProcessCapture -FilePath $toolPath -ArgumentList $arguments -WorkingDirectory (Split-Path -Path $toolPath -Parent)
                $logPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'ParseUSBs-Log.txt')
                $errorPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'ParseUSBs.stderr.txt')
                $processResult.StandardOutput | Out-File -LiteralPath $logPath -Encoding UTF8
                if ($processResult.ExitCode -ne 0 -or -not [string]::IsNullOrWhiteSpace($processResult.StandardError)) { $processResult.StandardError | Out-File -LiteralPath $errorPath -Encoding UTF8 }
                $movedOutputs = @(Move-IbisParseUsbOutput -StagingDirectory $stagedOutputDirectory -OutputDirectory $outputDirectory -Hostname $safeHost)

                # ParseUSBs writes no CSV when it observed no devices.  That is a finding
                # about the evidence, not a processing problem, so report it as a result.
                $noDevicesObserved = ($processResult.StandardOutput -match 'No USB device connections found')
                $toolStatus = 'Completed'
                $toolMessage = 'parseusbs completed against staged evidence.'
                if ($processResult.ExitCode -ne 0) {
                    $toolStatus = 'Failed'
                    $toolMessage = "parseusbs exited with code $($processResult.ExitCode)."
                }
                elseif ($movedOutputs.Count -eq 0 -and $noDevicesObserved) {
                    $toolMessage = 'parseusbs completed. No USB device connections were observed in the available artefacts.'
                }
                elseif ($movedOutputs.Count -eq 0) {
                    $toolStatus = 'Completed With Warnings'
                    $toolMessage = 'parseusbs completed, but no output files were found.'
                }

                $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceRoot = $stagingDirectory; OriginalSourceRoot = $SourceRoot; OutputDirectory = $outputDirectory; StagedOutputDirectory = $stagedOutputDirectory; StagedOutputBackupPath = $stagedOutputBackupPath; LogPath = $logPath; StandardErrorPath = $errorPath; StagingDirectory = $stagingDirectory; StagingBackupPath = $stagingBackupPath; Status = $toolStatus; ExitCode = $processResult.ExitCode; CommandLine = $processResult.CommandLine; NoDevicesObserved = [bool]$noDevicesObserved; MovedOutputs = $movedOutputs; Message = $toolMessage }
            }
            catch {
                $toolResult = [pscustomobject]@{ ToolId = $tool.id; SourceRoot = $SourceRoot; OutputDirectory = $outputDirectory; StagingDirectory = $stagingDirectory; StagingBackupPath = $stagingBackupPath; Status = 'Failed'; ExitCode = $null; Message = "ParseUSBs evidence staging failed: $($_.Exception.Message)" }
            }
        }
    }

    $summaryPath = Join-Path $workingsDirectory (New-IbisHostPrefixedFileName -Hostname $safeHost -Suffix 'USB.json')
    $payload = [pscustomobject]@{ ModuleId = 'usb'; Created = (Get-Date).ToString('s'); SourceRoot = $SourceRoot; ToolsRoot = $ToolsRoot; HostOutputRoot = $hostOutputRoot; OutputDirectory = $outputDirectory; WorkingsDirectory = $workingsDirectory; StagingDirectory = $stagingDirectory; StagingBackupPath = $stagingBackupPath; Staging = $stagingResult; ToolResults = @($toolResult) }
    $payload | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $summaryPath -Encoding UTF8

    $status = 'Completed'
    $message = 'USB artefact processing completed.'
    if ($toolResult.Status -eq 'Failed') {
        $status = 'Failed'
        $message = 'parseusbs failed. See USB summary JSON for details.'
    }
    elseif ($toolResult.Status -match 'Warnings') {
        $status = 'Completed With Warnings'
        $message = 'USB artefact processing completed with warning(s). See summary JSON for details.'
    }
    elseif ($toolResult.NoDevicesObserved) {
        $message = 'USB artefact processing completed. No USB device connections were observed in the available artefacts.'
    }

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
Export-ModuleMember -Function Get-IbisSevenZipPath
Export-ModuleMember -Function Expand-IbisArchive
Export-ModuleMember -Function Get-IbisLongPathsEnabled
Export-ModuleMember -Function Set-IbisLongPathsEnabled
Export-ModuleMember -Function Get-IbisVisualCppRedistributableStatus
Export-ModuleMember -Function Get-IbisPowerShellReadiness
Export-ModuleMember -Function Unblock-IbisProjectScriptFile
Export-ModuleMember -Function New-IbisToolInstallWorkspace
Export-ModuleMember -Function Remove-IbisEmptyToolWorkspaceDirectory
Export-ModuleMember -Function Get-IbisToolPublishSource
Export-ModuleMember -Function Backup-IbisToolInstallDirectory
Export-ModuleMember -Function Publish-IbisStagedToolInstall
Export-ModuleMember -Function New-IbisToolBackupPath
Export-ModuleMember -Function Clear-IbisToolPreviousInstall
Export-ModuleMember -Function Get-IbisToolInstallAssessment
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
Export-ModuleMember -Function Copy-IbisEvidenceFile
Export-ModuleMember -Function Copy-IbisRegistryHiveToCache
Export-ModuleMember -Function Find-IbisUsbEventLogCandidate
Export-ModuleMember -Function Copy-IbisParseUsbEvidenceToStaging
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
Export-ModuleMember -Function Get-IbisEventLogToolOutputDirectory
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
Export-ModuleMember -Function Rename-IbisToolOutputFiles
Export-ModuleMember -Function Rename-IbisChainsawOutput
Export-ModuleMember -Function Invoke-IbisChainsawEventLogs
Export-ModuleMember -Function Get-IbisUserAccessLogPath
Export-ModuleMember -Function Rename-IbisSumECmdOutput
Export-ModuleMember -Function Invoke-IbisUserAccessLogsSum
Export-ModuleMember -Function Get-IbisBrowserHistoryUsersPath
Export-ModuleMember -Function Get-IbisWebHistoryToolOutputDirectory
Export-ModuleMember -Function Invoke-IbisBrowsingHistoryView
Export-ModuleMember -Function Move-IbisForensicWebHistoryOutput
Export-ModuleMember -Function Invoke-IbisForensicWebHistory
Export-ModuleMember -Function Get-IbisParseUsbArgumentList
Export-ModuleMember -Function Move-IbisParseUsbOutput
Export-ModuleMember -Function Invoke-IbisParseUsbArtifacts
Export-ModuleMember -Function Get-IbisRegexValue
Export-ModuleMember -Function ConvertFrom-IbisSystemSummaryRegRipperOutput
Export-ModuleMember -Function Format-IbisSystemSummaryText
Export-ModuleMember -Function Invoke-IbisExtractHostName
Export-ModuleMember -Function Invoke-IbisSystemSummary
Export-ModuleMember -Function New-IbisRunSummary
Export-ModuleMember -Function Save-IbisRunSummary


