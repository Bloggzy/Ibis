$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Import-Module (Join-Path $projectRoot 'modules\Ibis.Core.psm1') -Force
Import-Module (Join-Path $projectRoot 'modules\Ibis.Gui.psm1') -Force

Describe 'Ibis core configuration' {
    It 'loads the main configuration' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $config.name | Should Be 'Ibis'
        $config.version | Should Be '0.7.5'
    }

    It 'records release history in the changelog' {
        $changelogPath = Join-Path $projectRoot 'CHANGELOG.md'
        Test-Path -LiteralPath $changelogPath -PathType Leaf | Should Be $true

        $changelog = Get-Content -LiteralPath $changelogPath -Raw
        $changelog | Should Match 'v0\.5\.4'
        $changelog | Should Match 'v0\.5\.3'
        $changelog | Should Match 'v0\.5\.2'
        $changelog | Should Match 'v0\.5\.1'
        $changelog | Should Match 'v0\.5\.0'
    }

    It 'saves path settings locally without changing the shared config' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $projectRoot 'config.json') -Destination (Join-Path $tempRoot 'config.json') -Force

        try {
            [void](Save-IbisConfigPathSetting `
                -ProjectRoot $tempRoot `
                -ToolsRoot 'C:\Tools\Saved' `
                -SourceRoot 'E:\Evidence' `
                -OutputRoot 'D:\Output' `
                -CompletionBeepEnabled $false)

            $saved = Get-IbisConfig -ProjectRoot $tempRoot
            $saved.defaultToolsRoot | Should Be 'C:\Tools\Saved'
            $saved.defaultSourceRoot | Should Be 'E:\Evidence'
            $saved.defaultOutputRoot | Should Be 'D:\Output'
            $saved.completionBeepEnabled | Should Be $false
            Test-Path -LiteralPath (Join-Path $tempRoot 'config.local.json') | Should Be $true
            $sharedConfig = Get-Content -LiteralPath (Join-Path $tempRoot 'config.json') -Raw | ConvertFrom-Json
            $sharedConfig.defaultSourceRoot | Should Be 'C:\Source\Evidence'
            $sharedConfig.defaultOutputRoot | Should Be 'C:\Export\Output'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'loads tool definitions' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $tools = @(Get-IbisToolDefinition -ProjectRoot $projectRoot -Config $config)
        ($tools.Count -gt 0) | Should Be $true
    }

    It 'enables processing modules by default now that initial testing is complete' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $disabledByDefault = @($config.modules | Where-Object { $_.enabledByDefault -ne $true })

        $disabledByDefault.Count | Should Be 0
    }

    It 'provides GUI hover hints for processing modules' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $modulesWithoutHints = @($config.modules | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.hint) })

        $modulesWithoutHints.Count | Should Be 0
    }

    It 'normalizes GUI display text line endings and ANSI escapes' {
        $escape = [string][char]27
        $text = "one`ntwo`rthree${escape}[38;2;255;175;0mfour${escape}[0m"

        $normalized = ConvertTo-IbisGuiDisplayText -Text $text -StripAnsi

        $normalized | Should Be "one`r`ntwo`r`nthreefour"
    }

    It 'checks for the Visual C++ Redistributable prerequisite' {
        $status = Get-IbisVisualCppRedistributableStatus -Architecture 'x64'

        $status.Architecture | Should Be 'x64'
        $status.MicrosoftUrl | Should Match 'latest-supported-vc-redist'
        [string]::IsNullOrWhiteSpace([string]$status.Message) | Should Be $false
        $status.PSObject.Properties.Name -contains 'Present' | Should Be $true
    }

    It 'checks that a fresh background PowerShell runspace can import Ibis core' {
        $readiness = Get-IbisPowerShellReadiness -ProjectRoot $projectRoot

        $readiness.Files.Count | Should Be 3
        @($readiness.Files | Where-Object { -not $_.Exists }).Count | Should Be 0
        $readiness.BackgroundImport.Passed | Should Be $true
        $readiness.CanStartBackgroundWork | Should Be $true
        $readiness.Status | Should Not Be 'Blocked'
    }

    It 'blocks background work when the Ibis core module is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        try {
            $readiness = Get-IbisPowerShellReadiness -ProjectRoot $tempRoot

            $readiness.Status | Should Be 'Blocked'
            $readiness.CanStartBackgroundWork | Should Be $false
            $readiness.BackgroundImport.ErrorMessage | Should Match 'Ibis core module was not found'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'formats processing run summaries with failed module details' {
        $moduleResults = @(
            [pscustomobject]@{
                ModuleId = 'amcache'
                ModuleName = 'Amcache'
                ErrorMessage = $null
                Result = [pscustomobject]@{
                    Status = 'Completed'
                    Message = 'Amcache completed.'
                }
            },
            [pscustomobject]@{
                ModuleId = 'srum'
                ModuleName = 'SRUM'
                ErrorMessage = $null
                Result = [pscustomobject]@{
                    Status = 'Failed'
                    Message = 'SrumECmd failed.'
                    ToolResults = @(
                        [pscustomobject]@{
                            ToolId = 'zimmerman-srumecmd'
                            Status = 'Failed'
                            Message = 'SrumECmd is missing.'
                        },
                        [pscustomobject]@{
                            ToolId = 'srum-dependent-step'
                            Status = 'Skipped'
                            Message = 'Skipped because the required SRUM input was unavailable.'
                        }
                    )
                }
            },
            [pscustomobject]@{
                ModuleId = 'prefetch'
                ModuleName = 'Prefetch'
                ErrorMessage = $null
                Result = [pscustomobject]@{
                    Status = 'Skipped'
                    Message = 'Prefetch folder was not found.'
                }
            }
        )

        $summary = @(Format-IbisProcessingRunSummary -ModuleResults $moduleResults)

        $summary[0] | Should Be 'Processing summary: 1 worked, 1 failed, 1 skipped.'
        ($summary -join "`n") | Should Match 'SRUM: SrumECmd failed'
        ($summary -join "`n") | Should Match 'Tool zimmerman-srumecmd: SrumECmd is missing'
        ($summary -join "`n") | Should Match 'Skipped items:'
        ($summary -join "`n") | Should Match 'Tool srum-dependent-step: Skipped because the required SRUM input was unavailable'
        ($summary -join "`n") | Should Match 'Prefetch: Prefetch folder was not found'
    }
}

Describe 'Ibis path comparison' {
    It 'exposes exactly one path comparison implementation' {
        $moduleText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Core.psm1') -Raw
        ([regex]::Matches($moduleText, 'function Resolve-IbisComparablePath')).Count | Should Be 1
    }

    It 'returns nothing for a blank path' {
        Resolve-IbisComparablePath -Path '' | Should BeNullOrEmpty
        Resolve-IbisComparablePath -Path '   ' | Should BeNullOrEmpty
    }

    It 'compares paths that differ only by case or trailing separator' {
        (Resolve-IbisComparablePath -Path 'C:\DFIR\Tools\') -eq (Resolve-IbisComparablePath -Path 'c:\dfir\tools') | Should Be $true
    }
}

Describe 'Ibis evidence checks' {
    It 'classifies an empty temporary folder as not Windows evidence' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempRoot | Out-Null
        try {
            $result = Test-IbisEvidenceRoot -SourceRoot $tempRoot
            $result.LooksLikeWindowsEvidence | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'detects writable folders inside the evidence source root' {
        $sourceRoot = 'C:\Evidence\Case01'
        $writablePaths = @(
            [pscustomobject]@{ Name = 'Output folder'; Path = 'C:\Evidence\Case01\Ibis-Output' },
            [pscustomobject]@{ Name = 'Tools folder'; Path = 'C:\DFIR\Tools' }
        )

        $result = Test-IbisSourceWriteBoundary -SourceRoot $sourceRoot -WritablePaths $writablePaths

        $result.Passed | Should Be $false
        $result.Violations.Count | Should Be 1
        $result.Violations[0].Name | Should Be 'Output folder'
    }

    It 'does not confuse sibling paths with folders inside the evidence source root' {
        Test-IbisPathInsideRoot -RootPath 'C:\Evidence\Case01' -Path 'C:\Evidence\Case010\Output' | Should Be $false
        Test-IbisPathInsideRoot -RootPath 'C:\Evidence\Case01' -Path 'C:\Evidence\Case01\Output' | Should Be $true
    }
}

Describe 'Ibis tool acquisition guidance' {
    It 'creates guidance for missing tools' {
        $statuses = @(
            [pscustomobject]@{
                Id = 'missing-tool'
                Name = 'Missing Tool'
                Present = $false
                ExpectedPath = 'C:\Tools\Missing\missing.exe'
                DownloadUrl = 'https://example.test/tool.zip'
                ManualUrl = 'https://example.test/manual'
                Notes = 'Test only'
            },
            [pscustomobject]@{
                Id = 'present-tool'
                Name = 'Present Tool'
                Present = $true
                ExpectedPath = 'C:\Tools\Present\present.exe'
                DownloadUrl = 'https://example.test/present.zip'
                ManualUrl = 'https://example.test/present'
                Notes = 'Test only'
            }
        )

        $plan = @(Get-IbisToolAcquisitionPlan -ToolStatuses $statuses)
        $plan.Count | Should Be 1
        $plan[0].Name | Should Be 'Missing Tool'
    }

    It 'formats an empty missing tool guidance plan' {
        Format-IbisToolAcquisitionPlan -AcquisitionPlan @() | Should Be 'All configured tools are present.'
    }
}

Describe 'Ibis progress logging' {
    It 'writes progress events as JSON lines' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $progressPath = Join-Path $tempRoot 'progress.jsonl'

        try {
            Write-IbisProgressEvent `
                -ProgressPath $progressPath `
                -ToolId 'example' `
                -ToolName 'Example Tool' `
                -Stage 'Download' `
                -Message 'Downloading example.zip.' `
                -Index 2 `
                -Total 5 `
                -Status 'Info'

            $line = Get-Content -LiteralPath $progressPath | Select-Object -First 1
            $progressEvent = $line | ConvertFrom-Json

            $progressEvent.ToolId | Should Be 'example'
            $progressEvent.ToolName | Should Be 'Example Tool'
            $progressEvent.Stage | Should Be 'Download'
            $progressEvent.Index | Should Be 2
            $progressEvent.Total | Should Be 5
        }
        finally {
            if (Test-Path -LiteralPath $tempRoot) {
                Remove-Item -LiteralPath $tempRoot -Recurse -Force
            }
        }
    }
}

Describe 'Ibis GUI processing runspace' {
    It 'allows processing to start with a blank hostname prefix' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $tools = @(Get-IbisToolDefinition -ProjectRoot $projectRoot -Config $config)
        $module = [pscustomobject]@{
            id = 'usb'
            name = 'USB artefacts'
            status = 'implemented'
        }
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        $operation = $null
        try {
            $operation = Start-IbisProcessingRunspace `
                -ProjectRoot $projectRoot `
                -ToolsRoot $tempRoot `
                -ToolDefinitions $tools `
                -Modules @($module) `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname ''

            $operation.Handle.AsyncWaitHandle.WaitOne(30000) | Should Be $true
            $result = @($operation.PowerShell.EndInvoke($operation.Handle))

            $result.Count | Should Be 1
            $result[0].ModuleId | Should Be 'usb'
            $result[0].Result.HostOutputRoot | Should Be $outputRoot
        }
        finally {
            if ($operation) {
                Stop-IbisToolInstallRunspace -Operation $operation
            }
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'does not adopt an extracted hostname during a run when the prefix starts blank' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $tools = @(Get-IbisToolDefinition -ProjectRoot $projectRoot -Config $config)
        $modules = @(
            [pscustomobject]@{ id = 'system-summary'; name = 'System summary'; status = 'implemented' },
            [pscustomobject]@{ id = 'usb'; name = 'USB artefacts'; status = 'implemented' }
        )
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $hiveRoot -Force | Out-Null
        'system' | Out-File -LiteralPath (Join-Path $hiveRoot 'SYSTEM') -Encoding ASCII

        $operation = $null
        try {
            $operation = Start-IbisProcessingRunspace `
                -ProjectRoot $projectRoot `
                -ToolsRoot $tempRoot `
                -ToolDefinitions $tools `
                -Modules $modules `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname ''

            $operation.Handle.AsyncWaitHandle.WaitOne(30000) | Should Be $true
            $result = @($operation.PowerShell.EndInvoke($operation.Handle))

            $result.Count | Should Be 2
            $result[0].Hostname | Should Be ''
            $result[0].OutputRoot | Should Be $outputRoot
            $result[1].Hostname | Should Be ''
            $result[1].OutputRoot | Should Be $outputRoot
        }
        finally {
            if ($operation) {
                Stop-IbisToolInstallRunspace -Operation $operation
            }
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Describe 'Ibis tool download metadata' {
    It 'has install metadata for each configured tool' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $tools = @(Get-IbisToolDefinition -ProjectRoot $projectRoot -Config $config)

        foreach ($tool in $tools) {
            [string]::IsNullOrWhiteSpace($tool.executablePath) | Should Be $false
            [string]::IsNullOrWhiteSpace($tool.packageType) | Should Be $false
            [string]::IsNullOrWhiteSpace($tool.downloadUrl) | Should Be $false
        }
    }

    It 'has GitHub metadata for latest-release tools' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $tools = @(Get-IbisToolDefinition -ProjectRoot $projectRoot -Config $config)
        $latestReleaseTools = @($tools | Where-Object { $_.downloadUrl -eq 'latest-release' })

        $latestReleaseTools.Count | Should Be 6
        foreach ($tool in $latestReleaseTools) {
            [string]::IsNullOrWhiteSpace($tool.githubRepo) | Should Be $false
            [string]::IsNullOrWhiteSpace($tool.assetPattern) | Should Be $false
        }
    }

    It 'resolves direct download URLs without network access' {
        $tool = [pscustomobject]@{
            name = 'Direct Tool'
            downloadUrl = 'https://example.test/direct.zip'
        }

        Resolve-IbisToolDownloadUrl -ToolDefinition $tool | Should Be 'https://example.test/direct.zip'
    }

    It 'calculates expected paths and install directories' {
        $tool = [pscustomobject]@{
            executablePath = 'Example\tool.exe'
            installDirectory = 'Example'
        }

        Get-IbisToolExpectedPath -ToolsRoot 'C:\DFIR\Tools' -ToolDefinition $tool | Should Be 'C:\DFIR\Tools\Example\tool.exe'
        Get-IbisToolInstallDirectory -ToolsRoot 'C:\DFIR\Tools' -ToolDefinition $tool | Should Be 'C:\DFIR\Tools\Example'
    }

    It 'keeps EvtxECmd in its own EZTools subfolder' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $tools = @(Get-IbisToolDefinition -ProjectRoot $projectRoot -Config $config)
        $evtx = Get-IbisToolDefinitionById -ToolDefinitions $tools -Id 'zimmerman-evtxecmd'

        $evtx.executablePath | Should Be 'EZTools\net9\EvtxECmd\EvtxECmd.exe'
        $evtx.installDirectory | Should Be 'EZTools\net9\EvtxECmd'
        Get-IbisToolInstallDirectory -ToolsRoot 'C:\DFIR\Tools' -ToolDefinition $evtx | Should Be 'C:\DFIR\Tools\EZTools\net9\EvtxECmd'
    }
}

Describe 'Ibis Defender exclusion metadata' {
    It 'returns an administrator status as a boolean' {
        (Test-IbisIsAdministrator) -is [bool] | Should Be $true
    }

    It 'recommends exclusions for rule-heavy tools' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $tools = @(Get-IbisToolDefinition -ProjectRoot $projectRoot -Config $config)
        $recommendations = @(Get-IbisDefenderExclusionRecommendation -ToolsRoot 'C:\DFIR\Tools' -ToolDefinitions $tools)

        $recommendations.Count | Should Be 3
        @($recommendations | Where-Object { $_.Id -eq 'chainsaw' }).Count | Should Be 1
        @($recommendations | Where-Object { $_.Id -eq 'hayabusa' }).Count | Should Be 1
        @($recommendations | Where-Object { $_.Id -eq 'takajo' }).Count | Should Be 1
    }

    It 'calculates recommended exclusion paths from install directories' {
        $tool = [pscustomobject]@{
            id = 'example'
            name = 'Example'
            installDirectory = 'ExampleTool'
            executablePath = 'ExampleTool\example.exe'
            defenderExclusionRecommended = $true
            defenderExclusionReason = 'Test'
        }

        $recommendation = Get-IbisDefenderExclusionRecommendation -ToolsRoot 'C:\DFIR\Tools' -ToolDefinitions @($tool)
        $recommendation.Path | Should Be 'C:\DFIR\Tools\ExampleTool'
    }
}

Describe 'Ibis staged tool installs' {
    It 'detects a partial install when files exist but the executable is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $tool = [pscustomobject]@{
            id = 'example'
            name = 'Example'
            executablePath = 'Example\example.exe'
            installDirectory = 'Example'
        }

        New-Item -ItemType Directory -Path (Join-Path $tempRoot 'Example') | Out-Null
        'partial' | Out-File -LiteralPath (Join-Path $tempRoot 'Example\leftover.txt') -Encoding ASCII

        try {
            $state = Test-IbisToolInstallState -ToolsRoot $tempRoot -ToolDefinition $tool
            $state.Status | Should Be 'Partial'
            $state.Present | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'assesses a single versioned executable as needing normalization' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $installRoot = Join-Path $tempRoot 'Hayabusa'
        New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
        'versioned' | Out-File -LiteralPath (Join-Path $installRoot 'hayabusa-3.10.0-win-x64.exe') -Encoding ASCII
        $tool = [pscustomobject]@{
            id = 'hayabusa'
            name = 'Hayabusa'
            executablePath = 'Hayabusa\hayabusa.exe'
            installDirectory = 'Hayabusa'
            renameExecutablePattern = 'hayabusa*x64.exe'
            renameExecutableTo = 'hayabusa.exe'
        }

        try {
            $assessment = Get-IbisToolInstallAssessment -ToolsRoot $tempRoot -ToolDefinition $tool
            $assessment.Status | Should Be 'Needs Normalization'
            $assessment.VersionedExecutables.Count | Should Be 1
            $assessment.VersionedExecutables[0].Name | Should Be 'hayabusa-3.10.0-win-x64.exe'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'reports multiple versioned executables as ambiguous' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $installRoot = Join-Path $tempRoot 'Hayabusa'
        New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
        'old' | Out-File -LiteralPath (Join-Path $installRoot 'hayabusa-3.9.0-win-x64.exe') -Encoding ASCII
        'new' | Out-File -LiteralPath (Join-Path $installRoot 'hayabusa-3.10.0-win-x64.exe') -Encoding ASCII
        $tool = [pscustomobject]@{
            id = 'hayabusa'
            name = 'Hayabusa'
            executablePath = 'Hayabusa\hayabusa.exe'
            installDirectory = 'Hayabusa'
            renameExecutablePattern = 'hayabusa*x64.exe'
            renameExecutableTo = 'hayabusa.exe'
        }

        try {
            (Get-IbisToolInstallAssessment -ToolsRoot $tempRoot -ToolDefinition $tool).Status | Should Be 'Ambiguous Existing Versions'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'creates Defender-sensitive staging under the install directory' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $tool = [pscustomobject]@{
            id = 'chainsaw'
            name = 'Chainsaw'
            executablePath = 'Chainsaw\chainsaw.exe'
            installDirectory = 'Chainsaw'
            defenderExclusionRecommended = $true
        }

        try {
            $workspace = New-IbisToolInstallWorkspace -ToolsRoot $tempRoot -ToolDefinition $tool
            $expectedPrefix = Join-Path $tempRoot 'Chainsaw\_s'
            $workspace.Root.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase) | Should Be $true
            Test-Path -LiteralPath $workspace.DownloadDirectory | Should Be $true
            Test-Path -LiteralPath $workspace.ExtractDirectory | Should Be $true
            (Split-Path -Path $workspace.DownloadDirectory -Leaf) | Should Be 'd'
            (Split-Path -Path $workspace.ExtractDirectory -Leaf) | Should Be 'x'
        }
        finally {
            if (Test-Path -LiteralPath $tempRoot) {
                Remove-Item -LiteralPath $tempRoot -Recurse -Force
            }
        }
    }

    It 'removes only empty legacy and Defender-aware workspace folders' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $installRoot = Join-Path $tempRoot 'Hayabusa'
        New-Item -ItemType Directory -Path (Join-Path $installRoot '_s') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $installRoot '_ibis-staging') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $installRoot '_ibis-backup\keep') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $installRoot 'rules') -Force | Out-Null
        'rule' | Out-File -LiteralPath (Join-Path $installRoot 'rules\rule.yml') -Encoding ASCII

        try {
            Remove-IbisEmptyToolWorkspaceDirectory -InstallDirectory $installRoot
            Test-Path -LiteralPath (Join-Path $installRoot '_s') | Should Be $false
            Test-Path -LiteralPath (Join-Path $installRoot '_ibis-staging') | Should Be $false
            Test-Path -LiteralPath (Join-Path $installRoot '_ibis-backup') | Should Be $true
            Test-Path -LiteralPath (Join-Path $installRoot 'rules\rule.yml') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'flattens a single archive root folder unless a tool needs directory rename handling' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $extractRoot = Join-Path $tempRoot 'extract'
        $singleRoot = Join-Path $extractRoot 'tool-root'
        New-Item -ItemType Directory -Path $singleRoot -Force | Out-Null

        try {
            $normalTool = [pscustomobject]@{ id = 'normal'; name = 'Normal' }
            Get-IbisToolPublishSource -ExtractDirectory $extractRoot -ToolDefinition $normalTool | Should Be $singleRoot

            $renameTool = [pscustomobject]@{ id = 'rename'; name = 'Rename'; renameExtractedDirectoryFrom = 'tool-root' }
            Get-IbisToolPublishSource -ExtractDirectory $extractRoot -ToolDefinition $renameTool | Should Be $extractRoot
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'backs up only the named tool directory when install directory is the tools root' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path (Join-Path $tempRoot 'RegRipper') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $tempRoot 'OtherTool') -Force | Out-Null
        $tool = [pscustomobject]@{
            id = 'regripper'
            name = 'RegRipper'
            renameExtractedDirectoryTo = 'RegRipper'
        }

        try {
            $backup = Backup-IbisToolInstallDirectory -InstallDirectory $tempRoot -ToolsRoot $tempRoot -ToolDefinition $tool
            Test-Path -LiteralPath (Join-Path $backup 'RegRipper') | Should Be $true
            Test-Path -LiteralPath (Join-Path $tempRoot 'OtherTool') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'publishes staged files into a shared install directory without backing up unrelated tools' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $installRoot = Join-Path $tempRoot 'EZTools\net9'
        $stageRoot = Join-Path $tempRoot 'stage'
        New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null
        'amcache' | Out-File -LiteralPath (Join-Path $installRoot 'AmcacheParser.exe') -Encoding ASCII
        'appcompat' | Out-File -LiteralPath (Join-Path $stageRoot 'AppCompatCacheParser.exe') -Encoding ASCII

        try {
            $backup = Publish-IbisStagedToolInstall -StagedSourcePath $stageRoot -InstallDirectory $installRoot -ToolsRoot $tempRoot -ToolDefinition ([pscustomobject]@{ id = 'appcompat' })
            $backup | Should Be $null
            Test-Path -LiteralPath (Join-Path $installRoot 'AmcacheParser.exe') | Should Be $true
            Test-Path -LiteralPath (Join-Path $installRoot 'AppCompatCacheParser.exe') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'backs up only conflicting staged items when publishing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $installRoot = Join-Path $tempRoot 'EZTools\net9'
        $stageRoot = Join-Path $tempRoot 'stage'
        New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null
        'old' | Out-File -LiteralPath (Join-Path $installRoot 'AmcacheParser.exe') -Encoding ASCII
        'new' | Out-File -LiteralPath (Join-Path $stageRoot 'AmcacheParser.exe') -Encoding ASCII

        try {
            $backup = Publish-IbisStagedToolInstall -StagedSourcePath $stageRoot -InstallDirectory $installRoot -ToolsRoot $tempRoot -ToolDefinition ([pscustomobject]@{ id = 'amcache' })
            Test-Path -LiteralPath (Join-Path $backup 'AmcacheParser.exe') | Should Be $true
            (Get-Content -LiteralPath (Join-Path $installRoot 'AmcacheParser.exe')) | Should Be 'new'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'selects rename candidates outside Ibis backup and staging folders' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $installRoot = Join-Path $tempRoot 'Chainsaw'
        New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $installRoot '_ibis-backup\old') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $installRoot '_ibis-staging\current') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $installRoot '_s\current') -Force | Out-Null
        'live' | Out-File -LiteralPath (Join-Path $installRoot 'chainsaw_x86_64-pc-windows-msvc.exe') -Encoding ASCII
        'backup' | Out-File -LiteralPath (Join-Path $installRoot '_ibis-backup\old\chainsaw_x86_64-pc-windows-msvc.exe') -Encoding ASCII
        'staging' | Out-File -LiteralPath (Join-Path $installRoot '_ibis-staging\current\chainsaw_x86_64-pc-windows-msvc.exe') -Encoding ASCII
        'workspace' | Out-File -LiteralPath (Join-Path $installRoot '_s\current\chainsaw_x86_64-pc-windows-msvc.exe') -Encoding ASCII

        try {
            $renameCandidates = @(Get-IbisExecutableRenameCandidate -InstallDirectory $installRoot -Pattern 'chainsaw_x86_64-pc-windows-msvc.exe')
            $renameCandidates.Count | Should Be 1
            $renameCandidates[0].FullName | Should Be (Join-Path $installRoot 'chainsaw_x86_64-pc-windows-msvc.exe')
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'renames a selected executable to the expected root executable name' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $installRoot = Join-Path $tempRoot 'Chainsaw'
        New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
        'live' | Out-File -LiteralPath (Join-Path $installRoot 'chainsaw_x86_64-pc-windows-msvc.exe') -Encoding ASCII
        $tool = [pscustomobject]@{
            id = 'chainsaw'
            name = 'Chainsaw'
            executablePath = 'Chainsaw\chainsaw.exe'
            installDirectory = 'Chainsaw'
            renameExecutablePattern = 'chainsaw_x86_64-pc-windows-msvc.exe'
            renameExecutableTo = 'chainsaw.exe'
        }

        try {
            Invoke-IbisToolPostInstall -ToolsRoot $tempRoot -ToolDefinition $tool
            Test-Path -LiteralPath (Join-Path $installRoot 'chainsaw.exe') | Should Be $true
            Test-Path -LiteralPath (Join-Path $installRoot 'chainsaw_x86_64-pc-windows-msvc.exe') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'renames the executable from the staged release when older versions remain' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $installRoot = Join-Path $tempRoot 'Hayabusa'
        New-Item -ItemType Directory -Path $installRoot -Force | Out-Null
        'old' | Out-File -LiteralPath (Join-Path $installRoot 'hayabusa-3.9.0-win-x64.exe') -Encoding ASCII
        'new' | Out-File -LiteralPath (Join-Path $installRoot 'hayabusa-3.10.0-win-x64.exe') -Encoding ASCII
        $tool = [pscustomobject]@{
            id = 'hayabusa'
            name = 'Hayabusa'
            executablePath = 'Hayabusa\hayabusa.exe'
            installDirectory = 'Hayabusa'
            renameExecutablePattern = 'hayabusa*x64.exe'
            renameExecutableTo = 'hayabusa.exe'
        }

        try {
            Invoke-IbisToolPostInstall -ToolsRoot $tempRoot -ToolDefinition $tool -PublishedExecutableRelativePath 'hayabusa-3.10.0-win-x64.exe'
            (Get-Content -LiteralPath (Join-Path $installRoot 'hayabusa.exe')) | Should Be 'new'
            Test-Path -LiteralPath (Join-Path $installRoot 'hayabusa-3.9.0-win-x64.exe') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'archives a dedicated tool folder before reinstalling it' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $installRoot = Join-Path $tempRoot 'Hayabusa'
        New-Item -ItemType Directory -Path (Join-Path $installRoot 'rules') -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $installRoot '_s\active') -Force | Out-Null
        'old' | Out-File -LiteralPath (Join-Path $installRoot 'hayabusa.exe') -Encoding ASCII
        'rule' | Out-File -LiteralPath (Join-Path $installRoot 'rules\rule.yml') -Encoding ASCII
        'workspace' | Out-File -LiteralPath (Join-Path $installRoot '_s\active\keep.txt') -Encoding ASCII
        $tool = [pscustomobject]@{
            id = 'hayabusa'
            name = 'Hayabusa'
            executablePath = 'Hayabusa\hayabusa.exe'
            installDirectory = 'Hayabusa'
        }

        try {
            $backup = Clear-IbisToolPreviousInstall -ToolsRoot $tempRoot -ToolDefinition $tool
            Test-Path -LiteralPath (Join-Path $backup 'hayabusa.exe') | Should Be $true
            Test-Path -LiteralPath (Join-Path $backup 'rules\rule.yml') | Should Be $true
            Test-Path -LiteralPath (Join-Path $installRoot '_s\active\keep.txt') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Describe 'Ibis command specs' {
    It 'renders a command line without executing it' {
        $command = New-IbisCommandSpec `
            -ToolId 'example' `
            -FilePath 'C:\Program Files\Tool\tool.exe' `
            -ArgumentList @('-f', 'C:\Evidence Path\file.evtx') `
            -Description 'Test command'

        $line = ConvertTo-IbisCommandLine -CommandSpec $command
        $line | Should Be '"C:\Program Files\Tool\tool.exe" -f "C:\Evidence Path\file.evtx"'
    }

    It 'reports Hayabusa rule update as failed when Hayabusa is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $tool = [pscustomobject]@{
            id = 'hayabusa'
            name = 'Hayabusa'
            executablePath = 'Hayabusa\hayabusa.exe'
        }

        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        try {
            $result = Invoke-IbisHayabusaRuleUpdate -ToolsRoot $tempRoot -ToolDefinitions @($tool)

            $result.ModuleId | Should Be 'hayabusa-rule-update'
            $result.Status | Should Be 'Failed'
            $result.Message | Should Match 'missing'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'captures large stdout and stderr streams without blocking the child process' {
        $hostExe = (Get-Process -Id $PID).Path
        if ([string]::IsNullOrWhiteSpace($hostExe)) {
            $hostExe = (Get-Command powershell.exe).Source
        }
        $script = '$chunk = "x" * 4096; 1..80 | ForEach-Object { [Console]::Out.WriteLine($chunk); [Console]::Error.WriteLine($chunk) }'
        $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($script))

        $result = Invoke-IbisProcessCapture -FilePath $hostExe -ArgumentList @('-NoProfile', '-EncodedCommand', $encoded)

        $result.ExitCode | Should Be 0
        $result.StandardOutput.Length -gt 100000 | Should Be $true
        $result.StandardError.Length -gt 100000 | Should Be $true
    }
}

Describe 'Ibis tool download readiness' {
    It 'is ready when every recommended exclusion is present' {
        $readiness = Get-IbisToolDownloadReadiness -ExclusionStatus @([pscustomobject]@{ Name = 'Hayabusa'; Path = 'C:\T\H'; Status = 'Present'; Present = $true; IsAdministrator = $true; Message = 'x' })

        $readiness.Status | Should Be 'Ready'
        $readiness.ShouldWarn | Should Be $false
    }

    It 'warns when an exclusion is missing' {
        $readiness = Get-IbisToolDownloadReadiness -ExclusionStatus @(
            [pscustomobject]@{ Name = 'Hayabusa'; Path = 'C:\T\H'; Status = 'Missing'; Present = $false; IsAdministrator = $true; Message = 'x' },
            [pscustomobject]@{ Name = 'Chainsaw'; Path = 'C:\T\C'; Status = 'Present'; Present = $true; IsAdministrator = $true; Message = 'x' }
        )

        $readiness.Status | Should Be 'ExclusionsMissing'
        $readiness.ShouldWarn | Should Be $true
        $readiness.MissingCount | Should Be 1
        $readiness.Headline | Should Match '1 of 2'
        ($readiness.Detail -join ' ') | Should Match 'Hayabusa'
    }

    It 'warns that a standard-user check cannot be trusted' {
        $readiness = Get-IbisToolDownloadReadiness -ExclusionStatus @([pscustomobject]@{ Name = 'Hayabusa'; Path = 'C:\T\H'; Status = 'Missing'; Present = $false; IsAdministrator = $false; Message = 'x' })

        $readiness.Status | Should Be 'NotAuthoritative'
        $readiness.ShouldWarn | Should Be $true
        $readiness.Headline | Should Match 'Administrator'
    }

    It 'does not warn when Defender is not the active scanner' {
        # Nothing to exclude, so nagging on every download would be noise.
        $readiness = Get-IbisToolDownloadReadiness -ExclusionStatus @([pscustomobject]@{ Name = 'Hayabusa'; Path = 'C:\T\H'; Status = 'Unavailable'; Present = $false; IsAdministrator = $false; Message = 'x' })

        $readiness.Status | Should Be 'DefenderUnavailable'
        $readiness.ShouldWarn | Should Be $false
    }

    It 'warns when the exclusion check itself failed' {
        $readiness = Get-IbisToolDownloadReadiness -ExclusionStatus @([pscustomobject]@{ Name = 'Hayabusa'; Path = 'C:\T\H'; Status = 'Failed'; Present = $false; IsAdministrator = $true; Message = 'x' })

        $readiness.Status | Should Be 'CheckFailed'
        $readiness.ShouldWarn | Should Be $true
    }

    It 'does not warn when no exclusions are recommended' {
        $readiness = Get-IbisToolDownloadReadiness -ExclusionStatus @()

        $readiness.Status | Should Be 'NoneRecommended'
        $readiness.ShouldWarn | Should Be $false
    }
}

Describe 'Ibis setup step indicators' {
    It 'marks a finished step with a green tick' {
        $indicator = Get-IbisSetupStepIndicator -State 'Ready'

        [int][char]$indicator.Symbol | Should Be 0x2714
        $indicator.Green | Should BeGreaterThan $indicator.Red
    }

    It 'marks a step that needs action with an amber warning' {
        $indicator = Get-IbisSetupStepIndicator -State 'Attention'

        [int][char]$indicator.Symbol | Should Be 0x26A0
        $indicator.Red | Should BeGreaterThan 128
        $indicator.Blue | Should Be 0
    }

    It 'marks a blocked step in red' {
        $indicator = Get-IbisSetupStepIndicator -State 'Blocked'

        [int][char]$indicator.Symbol | Should Be 0x2716
        $indicator.Red | Should BeGreaterThan $indicator.Green
    }

    It 'uses a neutral marker before anything has been checked' {
        $indicator = Get-IbisSetupStepIndicator -State 'Unknown'

        $indicator.State | Should Be 'Unknown'
        $indicator.Red | Should Be $indicator.Green
    }

    It 'builds symbols from code points so Windows PowerShell cannot mangle them' {
        # A .psm1 without a byte order mark is read as ANSI by Windows PowerShell,
        # which would turn a literal tick into nonsense.
        $moduleText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Core.psm1') -Raw
        $start = $moduleText.IndexOf('function Get-IbisSetupStepIndicator {')
        $end = $moduleText.IndexOf('function ', $start + 40)
        $body = $moduleText.Substring($start, $end - $start)

        $body | Should Match '\[char\]0x2714'
        $body | Should Match '\[char\]0x26A0'
    }

    It 'returns one symbol character per state' {
        foreach ($state in @('Ready', 'Attention', 'Blocked', 'Unknown')) {
            (Get-IbisSetupStepIndicator -State $state).Symbol.Length | Should Be 1
        }
    }
}

Describe 'Ibis USB artefact tools' {
    It 'gives each USB tool its own folder beneath one USB parent' {
        $parseUsbs = Get-IbisUsbToolOutputDirectory -OutputRoot 'C:\Export' -Hostname 'HOST01' -ToolFolder 'ParseUSBs'
        $boobook = Get-IbisUsbToolOutputDirectory -OutputRoot 'C:\Export' -Hostname 'HOST01' -ToolFolder 'Boobook'

        $parseUsbs | Should Be 'C:\Export\HOST01\USB\ParseUSBs'
        $boobook | Should Be 'C:\Export\HOST01\USB\Boobook'
    }

    It 'writes both USB tools under the same parent as the event log tools do' {
        $usbParent = Split-Path -Path (Get-IbisUsbToolOutputDirectory -OutputRoot 'C:\Export' -Hostname 'HOST01' -ToolFolder 'Boobook') -Parent
        $eventLogParent = Split-Path -Path (Get-IbisEventLogToolOutputDirectory -OutputRoot 'C:\Export' -Hostname 'HOST01' -ToolFolder 'Hayabusa') -Parent

        Split-Path -Path $usbParent -Leaf | Should Be 'USB'
        Split-Path -Path $usbParent -Parent | Should Be (Split-Path -Path $eventLogParent -Parent)
    }

    It 'asks boobook to write into the folder Ibis chose' {
        $arguments = Get-IbisBoobookArgumentList -SourceRoot 'E:\Triage' -OutputDirectory 'C:\Export\HOST01\USB\Boobook' -Hostname 'HOST01'

        $arguments -contains '-in-place' | Should Be $true
        $arguments[($arguments.IndexOf('-output') + 1)] | Should Be 'C:\Export\HOST01\USB\Boobook'
        $arguments[($arguments.IndexOf('-evidence') + 1)] | Should Be 'E:\Triage'
    }

    It 'never asks boobook for a run directory of its own' {
        # -run-id and -in-place contradict each other, and a run directory
        # Ibis did not name would hide the results from the output layout.
        $arguments = Get-IbisBoobookArgumentList -SourceRoot 'E:\Triage' -OutputDirectory 'C:\Out' -Hostname 'HOST01'

        $arguments -contains '-run-id' | Should Be $false
    }

    It 'silences the progress narration, which Ibis redirects' {
        $arguments = Get-IbisBoobookArgumentList -SourceRoot 'E:\Triage' -OutputDirectory 'C:\Out'

        $arguments -contains '-quiet' | Should Be $true
    }

    It 'passes the hostname only when there is one' {
        $named = Get-IbisBoobookArgumentList -SourceRoot 'E:\Triage' -OutputDirectory 'C:\Out' -Hostname 'HOST01'
        $blank = Get-IbisBoobookArgumentList -SourceRoot 'E:\Triage' -OutputDirectory 'C:\Out' -Hostname ''

        $named[($named.IndexOf('-host') + 1)] | Should Be 'HOST01'
        $blank -contains '-host' | Should Be $false
    }

    It 'skips cleanly when the evidence source is not there' {
        $outputRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ibis-boobook-' + [guid]::NewGuid().ToString('N'))
        try {
            $result = Invoke-IbisBoobookUsbArtifacts -ToolsRoot 'C:\DFIR\Tools' -ToolDefinitions @() -SourceRoot (Join-Path $outputRoot 'missing') -OutputRoot $outputRoot -Hostname 'HOST01'

            $result.ModuleId | Should Be 'boobook'
            $result.Status | Should Be 'Skipped'
        }
        finally {
            if (Test-Path -LiteralPath $outputRoot) { Remove-Item -LiteralPath $outputRoot -Recurse -Force }
        }
    }

    It 'records a failure when boobook is not configured' {
        $outputRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ibis-boobook-' + [guid]::NewGuid().ToString('N'))
        $sourceRoot = Join-Path $outputRoot 'source'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        try {
            $result = Invoke-IbisBoobookUsbArtifacts -ToolsRoot 'C:\DFIR\Tools' -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'HOST01'

            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
            $payload = Get-Content -LiteralPath $result.JsonPath -Raw | ConvertFrom-Json
            $payload.ToolResults[0].Message | Should Match 'not configured'
        }
        finally {
            if (Test-Path -LiteralPath $outputRoot) { Remove-Item -LiteralPath $outputRoot -Recurse -Force }
        }
    }
}

Describe 'Ibis module grouping' {
    $config = Get-IbisConfig -ProjectRoot $projectRoot

    It 'gives every module a group' {
        foreach ($module in $config.modules) {
            [string]::IsNullOrWhiteSpace($module.group) | Should Be $false
        }
    }

    It 'keeps each group together in the file, so run order matches the display' {
        $seen = @()
        $previous = ''
        foreach ($module in $config.modules) {
            $group = Get-IbisModuleGroupName -Module $module
            if ($group -ne $previous) {
                $seen -contains $group | Should Be $false
                $seen += $group
                $previous = $group
            }
        }
    }

    It 'keeps the host artefacts in one group' {
        # Several of these overlap: user artefacts are largely registry hives
        # and SRUM is a system database, so splitting them drew a line that
        # does not exist in the evidence.
        $hostModules = @($config.modules | Where-Object { (Get-IbisModuleGroupName -Module $_) -eq 'System and registry' } | ForEach-Object { $_.id })

        foreach ($id in @('ntfs-metadata', 'srum', 'user-artifacts', 'ual', 'registry', 'amcache')) {
            $hostModules -contains $id | Should Be $true
        }
    }

    It 'reads left to right as host, then event logs, then the rest' {
        $groups = @(@($config.modules | ForEach-Object { Get-IbisModuleGroupName -Module $_ }) | Select-Object -Unique)

        $groups[0] | Should Be 'System and registry'
        $groups[1] | Should Be 'Windows Event Logs'
    }

    It 'groups both USB tools together' {
        $usbModules = @($config.modules | Where-Object { (Get-IbisModuleGroupName -Module $_) -eq 'USB artefacts' })

        @($usbModules | ForEach-Object { $_.id }) -contains 'usb' | Should Be $true
        @($usbModules | ForEach-Object { $_.id }) -contains 'boobook' | Should Be $true
    }

    It 'keeps the tools that depend on another tool in the same group' {
        # Takajo needs Hayabusa and DuckDB needs EvtxECmd. Showing a module
        # apart from the one it depends on invites selecting it alone.
        $byId = @{}
        foreach ($module in $config.modules) { $byId[$module.id] = Get-IbisModuleGroupName -Module $module }

        $byId['takajo'] | Should Be $byId['hayabusa']
        $byId['duckdb-eventlogs'] | Should Be $byId['eventlogs']
    }

    It 'has a dispatch case for every implemented module' {
        # The run used to filter on a hand-written list of module ids, so a
        # new module was dropped from every run until someone remembered it.
        $guiText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Gui.psm1') -Raw
        foreach ($module in @($config.modules | Where-Object { $_.status -eq 'implemented' })) {
            $guiText | Should Match ("'" + [regex]::Escape($module.id) + "' " + [regex]::Escape('{ $result = Invoke-Ibis'))
        }
    }
}

Describe 'Ibis module column layout' {
    It 'balances the columns rather than filling each in turn' {
        # Filling to a quota put six of the nineteen modules below the fold.
        (Split-IbisModuleBlockColumn -BlockRows @(7, 5, 6, 3, 3) -ColumnCount 3) -join ',' | Should Be '0,1,1,2,2'
    }

    It 'settles a tie on the most even spread' {
        # 11,6,3,3 is eleven rows tall whichever way the last three are split.
        # Leaving the event logs a column of their own reads better than
        # filling the middle column and stranding one group on the right.
        (Split-IbisModuleBlockColumn -BlockRows @(11, 6, 3, 3) -ColumnCount 3) -join ',' | Should Be '0,1,2,2'
    }

    It 'never splits a group across two columns' {
        $assignment = @(Split-IbisModuleBlockColumn -BlockRows @(4, 4, 4, 4) -ColumnCount 2)

        $assignment.Count | Should Be 4
        $assignment -join ',' | Should Be '0,0,1,1'
    }

    It 'uses no more columns than it was given' {
        $assignment = @(Split-IbisModuleBlockColumn -BlockRows @(2, 2, 2, 2, 2, 2, 2) -ColumnCount 3)

        (($assignment | Measure-Object -Maximum).Maximum) | Should BeLessThan 3
    }

    It 'copes with one group and with none' {
        (Split-IbisModuleBlockColumn -BlockRows @(4) -ColumnCount 3) -join ',' | Should Be '0'
        @(Split-IbisModuleBlockColumn -BlockRows @() -ColumnCount 3).Count | Should Be 0
    }

    It 'fits every module without scrolling' {
        $guiText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Gui.psm1') -Raw
        $config = Get-IbisConfig -ProjectRoot $projectRoot

        $panelHeight = [int][regex]::Match($guiText, '\$modulePanel\.Size = New-Object System\.Drawing\.Size\(\d+, (\d+)\)').Groups[1].Value
        $rowHeight = [int][regex]::Match($guiText, '\$moduleRowHeight = (\d+)').Groups[1].Value
        $columnCount = [int][regex]::Match($guiText, '\$moduleColumnCount = (\d+)').Groups[1].Value

        $blockRows = @()
        foreach ($group in (@($config.modules | ForEach-Object { Get-IbisModuleGroupName -Module $_ }) | Select-Object -Unique)) {
            $blockRows += (@($config.modules | Where-Object { (Get-IbisModuleGroupName -Module $_) -eq $group }).Count + 1)
        }

        $assignment = @(Split-IbisModuleBlockColumn -BlockRows $blockRows -ColumnCount $columnCount)
        $tallest = 0
        for ($column = 0; $column -lt $columnCount; $column++) {
            $rows = 0
            for ($index = 0; $index -lt $blockRows.Count; $index++) {
                if ($assignment[$index] -ne $column) { continue }
                if ($rows -gt 0) { $rows++ }
                $rows += $blockRows[$index]
            }
            if ($rows -gt $tallest) { $tallest = $rows }
        }

        ($tallest * $rowHeight) | Should BeLessThan ($panelHeight + 1)
    }
}

Describe 'Ibis About tab support link' {
    $guiText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Gui.psm1') -Raw

    It 'holds the support address in one place' {
        ([regex]::Matches($guiText, 'buymeacoffee\.com/bloggz')).Count | Should Be 1
        $guiText | Should Match "\`$supportUrl = 'https://buymeacoffee\.com/bloggz'"
    }

    It 'shows the link on the About tab' {
        $guiText | Should Match '\$aboutTab\.Controls\.Add\(\$supportLinkLabel\)'
    }

    It 'opens the address from the link rather than a hard-coded copy' {
        $guiText | Should Match '\$supportLinkLabel\.Tag = \$supportUrl'
        $guiText | Should Match '\$supportLinkLabel\.Add_LinkClicked'
    }

    It 'keeps the link clear of the changelog label' {
        $linkTop = [int][regex]::Match($guiText, '\$supportLinkLabel\.Location = New-Object System\.Drawing\.Point\(\d+, (\d+)\)').Groups[1].Value
        $linkHeight = [int][regex]::Match($guiText, '\$supportLinkLabel\.Size = New-Object System\.Drawing\.Size\(\d+, (\d+)\)').Groups[1].Value
        $changelogTop = [int][regex]::Match($guiText, '\$changelogLabel\.Location = New-Object System\.Drawing\.Point\(\d+, (\d+)\)').Groups[1].Value

        ($linkTop + $linkHeight) | Should BeLessThan $changelogTop
    }
}

Describe 'Ibis setup tab layout' {
    $guiText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Gui.psm1') -Raw

    function Get-IbisTestControlPoint {
        param([string]$Text, [string]$VariableName, [string]$Property)

        $pattern = '\$' + $VariableName + '\.' + $Property + ' = New-Object System\.Drawing\.(?:Point|Size)\((\d+), (\d+)\)'
        $match = [regex]::Match($Text, $pattern)
        if (-not $match.Success) { throw "Could not find $VariableName.$Property" }
        [pscustomobject]@{ X = [int]$match.Groups[1].Value; Y = [int]$match.Groups[2].Value }
    }

    It 'puts the preparation steps in the left column, in order' {
        # The analyst reads left to right, so everything that must happen before
        # tools are downloaded belongs on the left, top to bottom.
        # Nothing else can be done until Ibis itself is allowed to run, so the
        # PowerShell check is the first thing the analyst should see.
        $groups = @('powerShellReadinessGroup', 'defenderActionsGroup', 'longPathsGroup', 'runtimePrereqGroup')
        $previousBottom = 0
        foreach ($group in $groups) {
            $location = Get-IbisTestControlPoint -Text $guiText -VariableName $group -Property 'Location'
            $size = Get-IbisTestControlPoint -Text $guiText -VariableName $group -Property 'Size'

            $location.X | Should Be 12
            ($location.Y -ge $previousBottom) | Should Be $true
            $previousBottom = $location.Y + $size.Y
        }

        # Nothing may run into the tool list below.
        ($previousBottom -le 326) | Should Be $true
    }

    It 'puts tool management in the right column' {
        $tools = Get-IbisTestControlPoint -Text $guiText -VariableName 'toolActionsGroup' -Property 'Location'
        $firstStep = Get-IbisTestControlPoint -Text $guiText -VariableName 'powerShellReadinessGroup' -Property 'Location'

        ($tools.X -gt $firstStep.X) | Should Be $true
        $tools.Y | Should Be $firstStep.Y
    }

    It 'numbers the setup steps so the order is obvious' {
        $guiText | Should Match "PowerShell = '1\. "
        $guiText | Should Match "Defender = '2\. "
        $guiText | Should Match "LongPaths = '3\. "
        $guiText | Should Match "Runtime = '4\. "
        $guiText | Should Match "Tools = '5\. "
    }

    It 'gives every setup step a status symbol' {
        foreach ($group in @('powerShellReadinessGroup', 'defenderActionsGroup', 'longPathsGroup', 'runtimePrereqGroup', 'toolActionsGroup')) {
            $guiText | Should Match ('setSetupStepState \$' + $group)
        }
    }

    It 'checks Defender exclusions before downloading tools' {
        $start = $guiText.IndexOf('$downloadMissingToolsButton.Add_Click(')
        $handler = $guiText.Substring($start, 4000)

        $handler | Should Match 'updateDownloadReadiness'
        $handler | Should Match 'Defender Exclusions Are Not Confirmed'
    }
}

Describe 'Ibis Hayabusa version handling' {
    $v4Help = "Hayabusa v4.0.0 - Black Hat USA Arsenal 2026 Release`n`nCommands:`n  computer-metrics         Output the total number of events`n  dfir-timeline            Create a DFIR timeline`n  update-rules             Update to the latest rules"
    $v3Help = "Hayabusa v3.10.0 - Independence Day Release`n`nCommands:`n  csv-timeline             Save the timeline in CSV format`n  json-timeline            Save the timeline in JSON format`n  update-rules             Update to the latest rules"

    It 'detects the Hayabusa 4 command set' {
        $capability = Get-IbisHayabusaCapability -HayabusaPath 'C:\Tools\hayabusa.exe' -HelpText $v4Help

        $capability.Version | Should Be '4.0.0'
        $capability.TimelineCommand | Should Be 'dfir-timeline'
        $capability.UsesOutputType | Should Be $true
    }

    It 'detects the Hayabusa 3 command set' {
        $capability = Get-IbisHayabusaCapability -HayabusaPath 'C:\Tools\hayabusa.exe' -HelpText $v3Help

        $capability.Version | Should Be '3.10.0'
        $capability.TimelineCommand | Should Be 'csv-timeline'
        $capability.UsesOutputType | Should Be $false
    }

    It 'assumes the current command set when the help text is unusable' {
        # Ibis installs the latest release, so the newer layout is the safer guess.
        $capability = Get-IbisHayabusaCapability -HayabusaPath 'C:\Tools\hayabusa.exe' -HelpText 'nonsense'

        $capability.TimelineCommand | Should Be 'dfir-timeline'
        $capability.DetectionMethod | Should Match 'Assumed'
    }

    It 'builds Hayabusa 4 timeline arguments' {
        $capability = Get-IbisHayabusaCapability -HayabusaPath 'C:\Tools\hayabusa.exe' -HelpText $v4Help

        $csv = (Get-IbisHayabusaTimelineArgumentList -Capability $capability -Format 'csv' -SourceDirectory 'D:\evtx' -OutputPath 'D:\out.csv') -join ' '
        $csv | Should Be 'dfir-timeline -d D:\evtx -t csv --iso-8601 -o D:\out.csv -w -C'

        $jsonl = (Get-IbisHayabusaTimelineArgumentList -Capability $capability -Format 'jsonl' -SourceDirectory 'D:\evtx' -OutputPath 'D:\out.jsonl') -join ' '
        $jsonl | Should Be 'dfir-timeline -d D:\evtx -t jsonl -x -U -a -A -w -p super-verbose -o D:\out.jsonl -C'
    }

    It 'builds Hayabusa 3 timeline arguments' {
        $capability = Get-IbisHayabusaCapability -HayabusaPath 'C:\Tools\hayabusa.exe' -HelpText $v3Help

        $csv = (Get-IbisHayabusaTimelineArgumentList -Capability $capability -Format 'csv' -SourceDirectory 'D:\evtx' -OutputPath 'D:\out.csv') -join ' '
        $csv | Should Be 'csv-timeline -d D:\evtx --ISO-8601 -o D:\out.csv -w -C'

        $jsonl = (Get-IbisHayabusaTimelineArgumentList -Capability $capability -Format 'jsonl' -SourceDirectory 'D:\evtx' -OutputPath 'D:\out.jsonl') -join ' '
        $jsonl | Should Be 'json-timeline -d D:\evtx -x -U -a -A -w -p super-verbose -L -o D:\out.jsonl -C'
    }

    It 'always overwrites an existing timeline' {
        # Without -C Hayabusa refuses to write, still exits 0, and leaves the old
        # timeline in place, which Ibis would report as a successful run.
        foreach ($helpText in @($v4Help, $v3Help)) {
            $capability = Get-IbisHayabusaCapability -HayabusaPath 'C:\Tools\hayabusa.exe' -HelpText $helpText
            foreach ($format in @('csv', 'jsonl')) {
                $arguments = Get-IbisHayabusaTimelineArgumentList -Capability $capability -Format $format -SourceDirectory 'D:\evtx' -OutputPath 'D:\out'
                ($arguments -contains '-C') | Should Be $true
            }
        }
    }

    It 'honours a custom output profile' {
        $capability = Get-IbisHayabusaCapability -HayabusaPath 'C:\Tools\hayabusa.exe' -HelpText $v4Help
        $arguments = Get-IbisHayabusaTimelineArgumentList -Capability $capability -Format 'jsonl' -SourceDirectory 'D:\evtx' -OutputPath 'D:\out.jsonl' -OutputProfile 'verbose'

        ($arguments -join ' ') | Should Match '-p verbose'
    }

    It 'no longer hard-codes a single timeline command' {
        $moduleText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Core.psm1') -Raw
        $start = $moduleText.IndexOf('function Invoke-IbisHayabusaEventLogs {')
        $end = $moduleText.IndexOf('function ', $start + 40)
        $moduleBody = $moduleText.Substring($start, $end - $start)

        $moduleBody | Should Match 'Get-IbisHayabusaTimelineArgumentList'
        $moduleBody | Should Not Match "'csv-timeline'"
        $moduleBody | Should Not Match "'json-timeline'"
    }
}

Describe 'Ibis processing module dispatch' {
    $guiPath = Join-Path $projectRoot 'modules\Ibis.Gui.psm1'
    $guiText = Get-Content -LiteralPath $guiPath -Raw

    It 'calls each processing module from exactly one place' {
        # The GUI once carried a second, unreachable dispatch chain behind a bare
        # return. Editing it had no effect, which is a trap worth keeping shut.
        $moduleFunctions = @(
            'Invoke-IbisSystemSummary',
            'Invoke-IbisVelociraptorResultsCopy',
            'Invoke-IbisWindowsRegistryHives',
            'Invoke-IbisAmcache',
            'Invoke-IbisAppCompatCache',
            'Invoke-IbisPrefetch',
            'Invoke-IbisNtfsMetadata',
            'Invoke-IbisSrum',
            'Invoke-IbisUserArtifacts',
            'Invoke-IbisEvtxECmdEventLogs',
            'Invoke-IbisDuckDbEventLogSummary',
            'Invoke-IbisHayabusaEventLogs',
            'Invoke-IbisTakajoEventLogs',
            'Invoke-IbisChainsawEventLogs',
            'Invoke-IbisUserAccessLogsSum',
            'Invoke-IbisBrowsingHistoryView',
            'Invoke-IbisForensicWebHistory',
            'Invoke-IbisParseUsbArtifacts',
            'Invoke-IbisBoobookUsbArtifacts'
        )

        $duplicated = @()
        foreach ($moduleFunction in $moduleFunctions) {
            $count = ([regex]::Matches($guiText, [regex]::Escape($moduleFunction))).Count
            if ($count -ne 1) { $duplicated += ('{0} ({1})' -f $moduleFunction, $count) }
        }

        ($duplicated -join ', ') | Should Be ''
    }

    It 'has no statements stranded after a bare return' {
        $guiText | Should Not Match '(?m)^\s+return\s*\r?\n\s*\r?\n\s+try \{'
    }
}

Describe 'Ibis module failure summaries' {
    It 'writes a failure summary when a module throws' {
        $outputRoot = Join-Path $TestDrive 'failure-output'
        New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null

        $errorRecord = $null
        try { throw 'MFTECmd could not read the volume.' } catch { $errorRecord = $_ }

        $path = Write-IbisModuleFailureSummary -OutputRoot $outputRoot -ModuleId 'ntfs-metadata' -ModuleName 'NTFS metadata' -Hostname 'HOST1' -ErrorRecord $errorRecord -Context @{ SourceRoot = 'C:\Source' }

        $path | Should Not BeNullOrEmpty
        Test-Path -LiteralPath $path | Should Be $true
        (Split-Path -Path $path -Leaf) | Should Be 'HOST1-ntfs-metadata-Failure.json'
        (Split-Path -Path (Split-Path -Path $path -Parent) -Leaf) | Should Be '_Ibis-Failures'

        $payload = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $payload.ModuleId | Should Be 'ntfs-metadata'
        $payload.ModuleName | Should Be 'NTFS metadata'
        $payload.Status | Should Be 'Failed'
        $payload.Message | Should Match 'could not read the volume'
        $payload.ScriptStackTrace | Should Not BeNullOrEmpty
        $payload.Context.SourceRoot | Should Be 'C:\Source'
    }

    It 'falls back to the module id when no name is supplied' {
        $outputRoot = Join-Path $TestDrive 'failure-noname'
        New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null

        $path = Write-IbisModuleFailureSummary -OutputRoot $outputRoot -ModuleId 'usb' -Message 'Staging failed.'
        $payload = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json

        $payload.ModuleName | Should Be 'usb'
        $payload.Message | Should Be 'Staging failed.'
        (Split-Path -Path $path -Leaf) | Should Be 'usb-Failure.json'
    }

    It 'records a message even when the error carried none' {
        $outputRoot = Join-Path $TestDrive 'failure-blank'
        New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null

        $path = Write-IbisModuleFailureSummary -OutputRoot $outputRoot -ModuleId 'srum'
        $payload = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json

        $payload.Message | Should Be 'The module failed without reporting a message.'
    }

    It 'returns nothing rather than throwing when the output root is blank' {
        # The recorder runs from inside a catch block, so it must never raise a
        # second error and take down the rest of the module run.
        Write-IbisModuleFailureSummary -OutputRoot '' -ModuleId 'srum' | Should BeNullOrEmpty
    }

    It 'places failures in a predictable folder beside the results' {
        (Split-Path -Path (Get-IbisModuleFailureDirectory -OutputRoot 'C:\Export\HOST1') -Leaf) | Should Be '_Ibis-Failures'
    }

    It 'records a failure summary from the processing runspace' {
        $guiText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Gui.psm1') -Raw
        $guiText | Should Match 'Write-IbisModuleFailureSummary'
    }
}

Describe 'Ibis tool error reporting' {
    It 'returns nothing for empty stderr' {
        Get-IbisToolErrorExcerpt -StandardError '' | Should Be ''
        Get-IbisToolErrorExcerpt -StandardError $null | Should Be ''
    }

    It 'keeps the last meaningful lines rather than the banner' {
        $stderr = "Started the command`n`nCounting lines`noserrors.nim(92) raiseOSError`nError: unhandled exception: The handle is invalid.`n [OSError]"
        $excerpt = Get-IbisToolErrorExcerpt -StandardError $stderr

        $excerpt | Should Match 'The handle is invalid'
        $excerpt | Should Not Match 'Started the command'
    }

    It 'truncates a very long excerpt' {
        $excerpt = Get-IbisToolErrorExcerpt -StandardError ('x' * 900) -MaximumLength 50
        $excerpt.Length | Should Be 53
        $excerpt | Should Match '\.\.\.$'
    }

    It 'writes tool stderr into the working directory' {
        $workingDirectory = Join-Path $TestDrive 'takajo-working'
        $result = [pscustomobject]@{ StandardError = 'Error: unhandled exception' }
        $path = Save-IbisTakajoStandardError -WorkingDirectory $workingDirectory -Hostname 'HOST1' -Mode 'stack-users' -ProcessResult $result

        $path | Should Not BeNullOrEmpty
        Test-Path -LiteralPath $path | Should Be $true
        (Split-Path -Path $path -Leaf) | Should Be 'HOST1-Takajo-stack-users.stderr.txt'
        (Get-Content -LiteralPath $path -Raw) | Should Match 'unhandled exception'
    }

    It 'writes no stderr file when the tool was quiet' {
        $workingDirectory = Join-Path $TestDrive 'takajo-quiet'
        $result = [pscustomobject]@{ StandardError = '' }
        $path = Save-IbisTakajoStandardError -WorkingDirectory $workingDirectory -Hostname 'HOST1' -Mode 'automagic' -ProcessResult $result

        $path | Should BeNullOrEmpty
        Test-Path -LiteralPath $workingDirectory | Should Be $false
    }

    It 'asks Takajo to skip the progress bar by long option name' {
        # Takajo maps '-s' to skipProgressBar for most stack commands but to
        # sourceUsers for stack-users, where the progress bar then crashes under a
        # redirected console.
        $moduleText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Core.psm1') -Raw
        $takajoStart = $moduleText.IndexOf('function Invoke-IbisTakajoEventLogs {')
        $takajoEnd = $moduleText.IndexOf('function ', $takajoStart + 40)
        $takajoText = $moduleText.Substring($takajoStart, $takajoEnd - $takajoStart)

        $takajoText | Should Match '--skipProgressBar'
        $takajoText | Should Not Match ", '-s'"
    }
}

Describe 'Ibis external process control' {
    $powerShellPath = (Get-Process -Id $PID).Path
    if ([string]::IsNullOrWhiteSpace($powerShellPath)) { $powerShellPath = 'powershell.exe' }

    AfterEach {
        Clear-IbisProcessCancellationCheck
    }

    It 'reports no cancellation when no check is registered' {
        Clear-IbisProcessCancellationCheck
        Test-IbisProcessCancellationRequested | Should Be $false
    }

    It 'uses a registered cancellation check' {
        Set-IbisProcessCancellationCheck -ScriptBlock { $true }
        Test-IbisProcessCancellationRequested | Should Be $true
    }

    It 'treats a broken cancellation check as no cancellation' {
        Set-IbisProcessCancellationCheck -ScriptBlock { throw 'control file unreadable' }
        Test-IbisProcessCancellationRequested | Should Be $false
    }

    It 'resolves a cancellation check that calls back into the caller scope' {
        # The processing runspace registers a scriptblock that calls its own
        # Get-IbisProcessingControlState. Ibis invokes it from inside the module,
        # so the scriptblock must still resolve commands from where it was defined.
        function Get-IbisTestControlStateProbe { 'CancelRequested' }
        Set-IbisProcessCancellationCheck -ScriptBlock { (Get-IbisTestControlStateProbe) -eq 'CancelRequested' }

        Test-IbisProcessCancellationRequested | Should Be $true
    }

    It 'stops a long running tool when the timeout is reached' {
        $started = Get-Date
        $result = Invoke-IbisProcessCapture -FilePath $powerShellPath -ArgumentList @('-NoProfile', '-Command', 'Start-Sleep -Seconds 60') -TimeoutSeconds 2
        $elapsed = ((Get-Date) - $started).TotalSeconds

        $result.TimedOut | Should Be $true
        $result.Cancelled | Should Be $false
        $elapsed | Should BeLessThan 40
        $result.StandardError | Should Match 'exceeded the configured timeout'
    }

    It 'stops a running tool when the analyst cancels the run' {
        Set-IbisProcessCancellationCheck -ScriptBlock { $true }
        $started = Get-Date
        $result = Invoke-IbisProcessCapture -FilePath $powerShellPath -ArgumentList @('-NoProfile', '-Command', 'Start-Sleep -Seconds 60')
        $elapsed = ((Get-Date) - $started).TotalSeconds

        $result.Cancelled | Should Be $true
        $result.TimedOut | Should Be $false
        $elapsed | Should BeLessThan 40
        $result.StandardError | Should Match 'analyst cancelled the run'
    }

    It 'leaves a normally completing tool untouched' {
        $result = Invoke-IbisProcessCapture -FilePath $powerShellPath -ArgumentList @('-NoProfile', '-Command', 'Write-Output ibis-ok') -TimeoutSeconds 120

        $result.ExitCode | Should Be 0
        $result.TimedOut | Should Be $false
        $result.Cancelled | Should Be $false
        $result.StandardOutput.Trim() | Should Be 'ibis-ok'
    }
}

Describe 'Ibis system summary' {
    It 'marks the system summary module as implemented' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $module = $config.modules | Where-Object { $_.id -eq 'system-summary' } | Select-Object -First 1

        $module.status | Should Be 'implemented'
    }

    It 'resolves host output roots without duplicating the hostname' {
        Get-IbisHostOutputRoot -OutputRoot 'C:\Export' -Hostname 'TESTHOST' | Should Be 'C:\Export\TESTHOST'
        Get-IbisHostOutputRoot -OutputRoot 'C:\Export\TESTHOST' -Hostname 'TESTHOST' | Should Be 'C:\Export\TESTHOST'
        Get-IbisHostOutputRoot -OutputRoot 'C:\Export\TESTHOST\' -Hostname 'TESTHOST' | Should Be 'C:\Export\TESTHOST'
    }

    It 'does not add a host folder or filename prefix when hostname is blank' {
        Get-IbisHostOutputRoot -OutputRoot 'C:\Export' -Hostname '' | Should Be 'C:\Export'
        New-IbisHostPrefixedFileName -Hostname '' -Suffix 'Tool-Output.csv' | Should Be 'Tool-Output.csv'
    }

    It 'calculates standard offline registry hive paths' {
        Get-IbisSystemHivePath -SourceRoot 'E:\Evidence' -HiveName 'SYSTEM' | Should Be 'E:\Evidence\Windows\System32\config\SYSTEM'
        Get-IbisSystemHivePath -SourceRoot 'E:\Evidence' -HiveName 'SOFTWARE' | Should Be 'E:\Evidence\Windows\System32\config\SOFTWARE'
    }

    It 'parses RegRipper plugin output into a summary' {
        $summary = ConvertFrom-IbisSystemSummaryRegRipperOutput `
            -CompNameOutput @('ComputerName    =    TESTHOST') `
            -WinVerOutput @(
                'ProductName    Windows 10 Pro',
                'BuildLab    22631.ni_release.220506-1250',
                'InstallDate    2024-01-02 03:04:05Z'
            ) `
            -IpsOutput @(
                'Header',
                'IPAddress    Domain',
                '10.0.0.5    example.local',
                '192.168.1.10    lab.local'
            ) `
            -ShutdownOutput @('LastWrite time: 2024-02-03 04:05:06Z') `
            -TimeZoneOutput @('TimeZoneKeyName->AUS Eastern Standard Time')

        $summary.HostName | Should Be 'TESTHOST'
        $summary.OperatingSystem | Should Match 'Windows 10 Pro'
        $summary.OperatingSystem | Should Match '22631'
        $summary.InstallDate | Should Be '2024-01-02 03:04:05Z'
        $summary.LastShutdown | Should Be '2024-02-03 04:05:06Z'
        $summary.TimeZone | Should Be 'AUS Eastern Standard Time'
        $summary.IpAddressSummary.Count | Should Be 2
    }

    It 'formats a concise text summary' {
        $summary = [pscustomobject]@{
            HostName = 'TESTHOST'
            OperatingSystem = 'Windows 10 Pro (Build: 19045)'
            TimeZone = 'W. Australia Standard Time'
            InstallDate = '2024-01-02 03:04:05Z'
            LastShutdown = '2024-02-03 04:05:06Z'
            IpAddressSummary = @('10.0.0.5 example.local')
        }

        $text = Format-IbisSystemSummaryText -Summary $summary

        $text | Should Match 'System Information'
        $text | Should Match 'Host name: TESTHOST'
        $text | Should Match 'Operating system: Windows 10 Pro'
        $text | Should Match 'IP Address\(es\) / Domain\(s\):'
    }

    It 'skips a RegRipper plugin when the hive is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $ripPath = Join-Path $tempRoot 'RegRipper\rip.exe'
        $outputPath = Join-Path $tempRoot 'out\compname.txt'
        $tool = [pscustomobject]@{
            id = 'regripper'
            name = 'RegRipper'
            executablePath = 'RegRipper\rip.exe'
        }

        New-Item -ItemType Directory -Path (Split-Path -Path $ripPath -Parent) -Force | Out-Null
        'fake' | Out-File -LiteralPath $ripPath -Encoding ASCII

        try {
            $result = Invoke-IbisRegRipperPlugin `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @($tool) `
                -HivePath (Join-Path $tempRoot 'missing-SYSTEM') `
                -Plugin 'compname' `
                -OutputPath $outputPath

            $result.Status | Should Be 'Skipped'
            Test-Path -LiteralPath $outputPath | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'extracts hostname status from the RegRipper compname module path' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $ripPath = Join-Path $tempRoot 'RegRipper\rip.exe'
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        $tool = [pscustomobject]@{
            id = 'regripper'
            name = 'RegRipper'
            executablePath = 'RegRipper\rip.exe'
        }

        New-Item -ItemType Directory -Path (Split-Path -Path $ripPath -Parent) -Force | Out-Null
        'fake' | Out-File -LiteralPath $ripPath -Encoding ASCII

        try {
            $result = Invoke-IbisExtractHostName `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @($tool) `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'HOST'

            $result.ModuleId | Should Be 'extract-hostname'
            $result.Status | Should Be 'Skipped'
            $result.HostName | Should Be 'Unknown'
            $result.HostOutputRoot | Should Be (Join-Path $outputRoot 'HOST')
            $result.SystemHive | Should Be (Join-Path $sourceRoot 'Windows\System32\config\SYSTEM')
            $result.OutputPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'HOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'writes system summary directly under the output root when hostname is blank' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisSystemSummary `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname ''

            $result.HostOutputRoot | Should Be $outputRoot
            $result.OutputPath | Should Be (Join-Path $outputRoot 'System-Summary\RR-System-Summary.txt')
            Test-Path -LiteralPath $result.OutputPath | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'HOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Describe 'Ibis Windows Registry hives' {
    It 'marks the registry module as implemented' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $module = $config.modules | Where-Object { $_.id -eq 'registry' } | Select-Object -First 1

        $module.status | Should Be 'implemented'
    }

    It 'uses the standard Windows system hive set' {
        $hives = @(Get-IbisWindowsRegistryHiveName)

        $hives -join ',' | Should Be 'SAM,SECURITY,SOFTWARE,SYSTEM'
    }

    It 'copies a hive and its transaction logs into the processing cache' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'
        $cacheRoot = Join-Path $tempRoot 'Cache'
        $sourceHive = Join-Path $hiveRoot 'SYSTEM'

        New-Item -ItemType Directory -Path $hiveRoot -Force | Out-Null
        'hive' | Out-File -LiteralPath $sourceHive -Encoding ASCII
        'log1' | Out-File -LiteralPath (Join-Path $hiveRoot 'SYSTEM.LOG1') -Encoding ASCII
        'log2' | Out-File -LiteralPath (Join-Path $hiveRoot 'SYSTEM.LOG2') -Encoding ASCII

        try {
            $result = Copy-IbisRegistryHiveToCache -SourceHivePath $sourceHive -CacheDirectory $cacheRoot

            Test-Path -LiteralPath (Join-Path $cacheRoot 'SYSTEM') | Should Be $true
            Test-Path -LiteralPath (Join-Path $cacheRoot 'SYSTEM.LOG1') | Should Be $true
            Test-Path -LiteralPath (Join-Path $cacheRoot 'SYSTEM.LOG2') | Should Be $true
            $result.TransactionLogCount | Should Be 2
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'stages ParseUSBs registry files without changing the evidence copies' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'
        $stagingRoot = Join-Path $tempRoot 'Staging'
        $sourceHive = Join-Path $hiveRoot 'SYSTEM'
        $sourceLog = Join-Path $hiveRoot 'SYSTEM.LOG1'

        New-Item -ItemType Directory -Path $hiveRoot -Force | Out-Null
        'original hive' | Out-File -LiteralPath $sourceHive -Encoding ASCII
        'transaction log' | Out-File -LiteralPath $sourceLog -Encoding ASCII

        try {
            $result = Copy-IbisParseUsbEvidenceToStaging -SourceRoot $sourceRoot -StagingDirectory $stagingRoot
            'staged change' | Out-File -LiteralPath (Join-Path $stagingRoot 'Windows\System32\config\SYSTEM') -Encoding ASCII

            $result.Status | Should Be 'Staged'
            $result.FileCount | Should Be 2
            Test-Path -LiteralPath (Join-Path $stagingRoot 'Windows\System32\config\SYSTEM.LOG1') | Should Be $true
            (Get-Content -LiteralPath $sourceHive -Raw).Trim() | Should Be 'original hive'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'stages ParseUSBs user hives, LNK files, and supported event logs without changing source evidence' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'
        $profileRoot = Join-Path $sourceRoot 'Users\Analyst'
        $recentRoot = Join-Path $profileRoot 'AppData\Roaming\Microsoft\Windows\Recent'
        $eventLogRoot = Join-Path $sourceRoot 'Windows\System32\winevt\Logs'
        $stagingRoot = Join-Path $tempRoot 'Staging'

        New-Item -ItemType Directory -Path $hiveRoot, $recentRoot, $eventLogRoot -Force | Out-Null
        'system' | Out-File -LiteralPath (Join-Path $hiveRoot 'SYSTEM') -Encoding ASCII
        'software' | Out-File -LiteralPath (Join-Path $hiveRoot 'SOFTWARE') -Encoding ASCII
        'user hive' | Out-File -LiteralPath (Join-Path $profileRoot 'NTUSER.dat') -Encoding ASCII
        'user log' | Out-File -LiteralPath (Join-Path $profileRoot 'NTUSER.dat.LOG1') -Encoding ASCII
        'lnk evidence' | Out-File -LiteralPath (Join-Path $recentRoot 'usb.lnk') -Encoding ASCII
        'partition log' | Out-File -LiteralPath (Join-Path $eventLogRoot 'Microsoft-Windows-Partition%4Diagnostic.evtx') -Encoding ASCII

        try {
            $result = Copy-IbisParseUsbEvidenceToStaging -SourceRoot $sourceRoot -StagingDirectory $stagingRoot
            'changed stage' | Out-File -LiteralPath (Join-Path $stagingRoot 'Users\Analyst\NTUSER.dat') -Encoding ASCII

            $result.UserHiveCount | Should Be 1
            $result.LnkFileCount | Should Be 1
            $result.EventLogCount | Should Be 1
            Test-Path -LiteralPath (Join-Path $stagingRoot 'Users\Analyst\NTUSER.dat.LOG1') | Should Be $true
            Test-Path -LiteralPath (Join-Path $stagingRoot 'Users\Analyst\AppData\Roaming\Microsoft\Windows\Recent\usb.lnk') | Should Be $true
            Test-Path -LiteralPath (Join-Path $stagingRoot 'Windows\System32\winevt\Logs\Microsoft-Windows-Partition%4Diagnostic.evtx') | Should Be $true
            (Get-Content -LiteralPath (Join-Path $profileRoot 'NTUSER.dat') -Raw).Trim() | Should Be 'user hive'
        }
        finally { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }

    It 'clears the read-only attribute on staged copies of read-only evidence' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'
        $profileRoot = Join-Path $sourceRoot 'Users\Analyst'
        $stagingRoot = Join-Path $tempRoot 'Staging'

        New-Item -ItemType Directory -Path $hiveRoot, $profileRoot -Force | Out-Null
        'system' | Out-File -LiteralPath (Join-Path $hiveRoot 'SYSTEM') -Encoding ASCII
        'system log' | Out-File -LiteralPath (Join-Path $hiveRoot 'SYSTEM.LOG1') -Encoding ASCII
        'user hive' | Out-File -LiteralPath (Join-Path $profileRoot 'NTUSER.dat') -Encoding ASCII

        # A read-only mount is the recommended workflow, so staged copies must not
        # inherit the attribute or the tool cannot replay transactions into them.
        Get-ChildItem -LiteralPath $sourceRoot -File -Recurse -Force | ForEach-Object { $_.IsReadOnly = $true }

        try {
            $result = Copy-IbisParseUsbEvidenceToStaging -SourceRoot $sourceRoot -StagingDirectory $stagingRoot

            $result.Status | Should Be 'Staged'
            (Get-Item -LiteralPath (Join-Path $stagingRoot 'Windows\System32\config\SYSTEM') -Force).IsReadOnly | Should Be $false
            (Get-Item -LiteralPath (Join-Path $stagingRoot 'Windows\System32\config\SYSTEM.LOG1') -Force).IsReadOnly | Should Be $false
            (Get-Item -LiteralPath (Join-Path $stagingRoot 'Users\Analyst\NTUSER.dat') -Force).IsReadOnly | Should Be $false
            (Get-Item -LiteralPath (Join-Path $hiveRoot 'SYSTEM') -Force).IsReadOnly | Should Be $true
        }
        finally {
            Get-ChildItem -LiteralPath $tempRoot -File -Recurse -Force | ForEach-Object { $_.IsReadOnly = $false }
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
    It 'builds a volume-mode ParseUSBs argument list from staged evidence' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $stageRoot = Join-Path $tempRoot 'Staging'
        $configRoot = Join-Path $stageRoot 'Windows\System32\config'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $configRoot -Force | Out-Null
        'system' | Out-File -LiteralPath (Join-Path $configRoot 'SYSTEM') -Encoding ASCII
        'software' | Out-File -LiteralPath (Join-Path $configRoot 'SOFTWARE') -Encoding ASCII
        try {
            $staging = [pscustomobject]@{ StagingDirectory = $stageRoot; SystemHivePath = (Join-Path $configRoot 'SYSTEM'); SoftwareHivePath = (Join-Path $configRoot 'SOFTWARE') }
            $arguments = @(Get-IbisParseUsbArgumentList -Staging $staging -OutputDirectory $outputRoot)

            # Explicit hive mode is not used: ParseUSBs 1.8 crashes in that mode.
            $arguments[0] | Should Be '-v'
            $arguments[1] | Should Be $stageRoot
            ($arguments -join ' ') | Should Not Match '(^| )-s( |$)'
            ($arguments -join ' ') | Should Not Match '(^| )-u( |$)'
            ($arguments -join ' ') | Should Match '(^| )-o csv( |$)'
        }
        finally { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }

    It 'fails argument construction when the staged SYSTEM hive is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $stageRoot = Join-Path $tempRoot 'Staging'
        New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null
        try {
            $staging = [pscustomobject]@{ StagingDirectory = $stageRoot; SystemHivePath = (Join-Path $stageRoot 'SYSTEM') }

            # Pester 3 'Should Throw' does not work under PowerShell 7, so catch it here.
            $thrownMessage = $null
            try { Get-IbisParseUsbArgumentList -Staging $staging -OutputDirectory (Join-Path $tempRoot 'Output') }
            catch { $thrownMessage = $_.Exception.Message }

            $thrownMessage | Should Be 'The staged SYSTEM hive was not found.'
        }
        finally { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }
    It 'publishes ParseUSBs output with a hostname prefix' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $stagingRoot = Join-Path $tempRoot 'Staging'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null
        'header,value' | Out-File -LiteralPath (Join-Path $stagingRoot 'usb-info.csv') -Encoding ASCII
        try {
            $moved = @(Move-IbisParseUsbOutput -StagingDirectory $stagingRoot -OutputDirectory $outputRoot -Hostname 'TESTHOST')

            $moved.Count | Should Be 1
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST-ParseUSBs-usb-info.csv') | Should Be $true
        }
        finally { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }

    It 'stages Velociraptor percent-encoded USB event logs under their canonical name' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'
        $eventLogRoot = Join-Path $sourceRoot 'Windows\System32\winevt\Logs'
        $stagingRoot = Join-Path $tempRoot 'Staging'

        New-Item -ItemType Directory -Path $hiveRoot, $eventLogRoot -Force | Out-Null
        'system' | Out-File -LiteralPath (Join-Path $hiveRoot 'SYSTEM') -Encoding ASCII

        # Velociraptor escapes the '%' in the channel name, so %4 is collected as %254.
        'partition' | Out-File -LiteralPath (Join-Path $eventLogRoot 'Microsoft-Windows-Partition%254Diagnostic.evtx') -Encoding ASCII
        'storsvc' | Out-File -LiteralPath (Join-Path $eventLogRoot 'Microsoft-Windows-Storsvc%4Diagnostic.evtx') -Encoding ASCII

        try {
            $result = Copy-IbisParseUsbEvidenceToStaging -SourceRoot $sourceRoot -StagingDirectory $stagingRoot
            $stagedLogRoot = Join-Path $stagingRoot 'Windows\System32\winevt\Logs'

            $result.EventLogCount | Should Be 2
            Test-Path -LiteralPath (Join-Path $stagedLogRoot 'Microsoft-Windows-Partition%4Diagnostic.evtx') | Should Be $true
            Test-Path -LiteralPath (Join-Path $stagedLogRoot 'Microsoft-Windows-Storsvc%4Diagnostic.evtx') | Should Be $true

            $partition = @($result.EventLogs | Where-Object { $_.Name -eq 'Microsoft-Windows-Partition%4Diagnostic.evtx' })[0]
            $partition.SourceName | Should Be 'Microsoft-Windows-Partition%254Diagnostic.evtx'
            $partition.DiscoveryMethod | Should Be 'Velociraptor encoded name'

            $storsvc = @($result.EventLogs | Where-Object { $_.Name -eq 'Microsoft-Windows-Storsvc%4Diagnostic.evtx' })[0]
            $storsvc.DiscoveryMethod | Should Be 'Canonical name'
        }
        finally { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }

    It 'prefers the canonical USB event log name over the encoded form' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $eventLogRoot = Join-Path $tempRoot 'Logs'
        New-Item -ItemType Directory -Path $eventLogRoot -Force | Out-Null
        'canonical' | Out-File -LiteralPath (Join-Path $eventLogRoot 'Microsoft-Windows-Partition%4Diagnostic.evtx') -Encoding ASCII
        'encoded' | Out-File -LiteralPath (Join-Path $eventLogRoot 'Microsoft-Windows-Partition%254Diagnostic.evtx') -Encoding ASCII
        try {
            $candidate = Find-IbisUsbEventLogCandidate -EventLogDirectory $eventLogRoot -CanonicalName 'Microsoft-Windows-Partition%4Diagnostic.evtx'

            $candidate.SourceName | Should Be 'Microsoft-Windows-Partition%4Diagnostic.evtx'
            $candidate.DiscoveryMethod | Should Be 'Canonical name'
        }
        finally { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }

    It 'returns nothing when no USB event log candidate exists' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        try {
            Find-IbisUsbEventLogCandidate -EventLogDirectory $tempRoot -CanonicalName 'Microsoft-Windows-Storsvc%4Diagnostic.evtx' | Should BeNullOrEmpty
        }
        finally { Remove-Item -LiteralPath $tempRoot -Recurse -Force }
    }
    It 'caches prepared hive metadata and reuses it on repeated preparation' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'
        $sourceHive = Join-Path $hiveRoot 'SYSTEM'

        New-Item -ItemType Directory -Path $hiveRoot -Force | Out-Null
        'hive' | Out-File -LiteralPath $sourceHive -Encoding ASCII

        try {
            $first = Invoke-IbisPrepareRegistryHive `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST' `
                -HiveName 'SYSTEM'
            $second = Invoke-IbisPrepareRegistryHive `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST' `
                -HiveName 'SYSTEM'

            $first.Status | Should Be 'Prepared With Warnings'
            $first.CacheHit | Should Be $false
            $second.CacheHit | Should Be $true
            $second.PreparedHivePath | Should Be $first.PreparedHivePath
            Test-Path -LiteralPath $second.PreparedHivePath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips a missing hive without creating a host output folder' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisPrepareRegistryHive `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST' `
                -HiveName 'SYSTEM'

            $result.Status | Should Be 'Skipped'
            $result.PreparedHivePath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues registry module processing with warnings when hive state cannot be checked' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'

        New-Item -ItemType Directory -Path $hiveRoot -Force | Out-Null
        foreach ($hive in @(Get-IbisWindowsRegistryHiveName)) {
            $hive | Out-File -LiteralPath (Join-Path $hiveRoot $hive) -Encoding ASCII
        }

        try {
            $result = Invoke-IbisWindowsRegistryHives `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'registry'
            $result.Status | Should Be 'Failed'
            $result.PreparedHiveCount | Should Be 4
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Registry-Hives\_Working\Prepared-Hives\SYSTEM\SYSTEM') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'processes registry hives without adding a host folder when hostname is blank' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'

        New-Item -ItemType Directory -Path $hiveRoot -Force | Out-Null
        foreach ($hive in @(Get-IbisWindowsRegistryHiveName)) {
            $hive | Out-File -LiteralPath (Join-Path $hiveRoot $hive) -Encoding ASCII
        }

        try {
            $result = Invoke-IbisWindowsRegistryHives `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname ''

            $result.HostOutputRoot | Should Be $outputRoot
            $result.PreparedHiveCount | Should Be 4
            Test-Path -LiteralPath (Join-Path $outputRoot 'Registry-Hives\_Working\Registry-Hives.json') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'Registry-Hives\_Working\Prepared-Hives\SYSTEM\SYSTEM') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'HOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Describe 'Ibis Amcache' {
    It 'marks the Amcache module as implemented' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $module = $config.modules | Where-Object { $_.id -eq 'amcache' } | Select-Object -First 1

        $module.status | Should Be 'implemented'
    }

    It 'calculates the standard Amcache hive path' {
        Get-IbisAmcacheHivePath -SourceRoot 'E:\Evidence' | Should Be 'E:\Evidence\Windows\appcompat\Programs\Amcache.hve'
    }

    It 'prepares arbitrary registry hive files for feature-specific modules' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $amcacheRoot = Join-Path $sourceRoot 'Windows\appcompat\Programs'
        $sourceHive = Join-Path $amcacheRoot 'Amcache.hve'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $amcacheRoot -Force | Out-Null
        'amcache' | Out-File -LiteralPath $sourceHive -Encoding ASCII
        'log1' | Out-File -LiteralPath (Join-Path $amcacheRoot 'Amcache.hve.LOG1') -Encoding ASCII

        try {
            $result = Invoke-IbisPrepareRegistryHiveFile `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceHivePath $sourceHive `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST' `
                -HiveName 'Amcache.hve' `
                -CacheGroup 'Amcache' `
                -CacheKey 'Amcache'

            $result.Status | Should Be 'Prepared With Warnings'
            $result.CacheGroup | Should Be 'Amcache'
            $result.CacheKey | Should Be 'Amcache'
            $result.TransactionLogCount | Should Be 1
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Amcache\_Working\Prepared-Hives\Amcache\Amcache.hve') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Amcache\_Working\Prepared-Hives\Amcache\Amcache.hve.LOG1') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips Amcache cleanly when the source hive is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisAmcache `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues Amcache processing with failures noted when tools are missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $amcacheRoot = Join-Path $sourceRoot 'Windows\appcompat\Programs'
        $sourceHive = Join-Path $amcacheRoot 'Amcache.hve'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $amcacheRoot -Force | Out-Null
        'amcache' | Out-File -LiteralPath $sourceHive -Encoding ASCII

        try {
            $result = Invoke-IbisAmcache `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'amcache'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Amcache\_Working\Prepared-Hives\Amcache\Amcache.hve') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Describe 'Ibis AppCompatCache, Prefetch, and NTFS metadata' {
    It 'marks AppCompatCache, Prefetch, and NTFS metadata modules as implemented' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $appCompat = $config.modules | Where-Object { $_.id -eq 'appcompatcache' } | Select-Object -First 1
        $prefetch = $config.modules | Where-Object { $_.id -eq 'prefetch' } | Select-Object -First 1
        $ntfsMetadata = $config.modules | Where-Object { $_.id -eq 'ntfs-metadata' } | Select-Object -First 1

        $appCompat.status | Should Be 'implemented'
        $prefetch.status | Should Be 'implemented'
        $prefetch.name | Should Be 'Prefetch'
        $prefetch.hint | Should Match 'Windows Server'
        $ntfsMetadata.status | Should Be 'implemented'
    }

    It 'skips AppCompatCache cleanly when the SYSTEM hive is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisAppCompatCache `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues AppCompatCache processing with failures noted when the parser is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $hiveRoot -Force | Out-Null
        'system' | Out-File -LiteralPath (Join-Path $hiveRoot 'SYSTEM') -Encoding ASCII

        try {
            $result = Invoke-IbisAppCompatCache `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'appcompatcache'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Registry-Hives\_Working\Prepared-Hives\SYSTEM\SYSTEM') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'calculates the standard Prefetch source path' {
        Get-IbisPrefetchPath -SourceRoot 'E:\Evidence' | Should Be 'E:\Evidence\Windows\Prefetch'
    }

    It 'renames PECmd timestamp-prefixed output files to the hostname' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        'csv' | Out-File -LiteralPath (Join-Path $tempRoot '20260425_PECmd_Output.csv') -Encoding ASCII
        'other' | Out-File -LiteralPath (Join-Path $tempRoot 'unchanged.csv') -Encoding ASCII

        try {
            $renamed = @(Rename-IbisPrefetchOutput -OutputDirectory $tempRoot -Hostname 'TESTHOST')

            $renamed.Count | Should Be 1
            Test-Path -LiteralPath (Join-Path $tempRoot 'TESTHOST_PECmd_Output.csv') | Should Be $true
            Test-Path -LiteralPath (Join-Path $tempRoot 'unchanged.csv') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips Prefetch cleanly when the folder is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisPrefetch `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues Prefetch processing with failures noted when PECmd is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $prefetchRoot = Join-Path $sourceRoot 'Windows\Prefetch'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $prefetchRoot -Force | Out-Null
        'pf' | Out-File -LiteralPath (Join-Path $prefetchRoot 'APP.EXE-12345678.pf') -Encoding ASCII

        try {
            $result = Invoke-IbisPrefetch `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'prefetch'
            $result.Status | Should Be 'Failed'
            $result.SourcePrefetchFileCount | Should Be 1
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'finds $MFT at the mounted source root' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'MountedRoot'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        'mft' | Out-File -LiteralPath (Join-Path $sourceRoot '$MFT') -Encoding ASCII

        try {
            Find-IbisNtfsArtifactPath -SourceRoot $sourceRoot -ArtifactName '$MFT' | Should Be (Join-Path $sourceRoot '$MFT')
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'finds $MFT in a nearby Velociraptor ntfs upload folder' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Collection\uploads\auto\C%3A'
        $ntfsRoot = Join-Path $tempRoot 'Collection\uploads\ntfs\%5C%5C.%5CC%3A'
        $mftPath = Join-Path $ntfsRoot '$MFT'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $ntfsRoot -Force | Out-Null
        'mft' | Out-File -LiteralPath $mftPath -Encoding ASCII

        try {
            Find-IbisNtfsArtifactPath -SourceRoot $sourceRoot -ArtifactName '$MFT' | Should Be $mftPath
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'finds the USN Journal $J in a nearby Velociraptor ntfs upload folder' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Collection\uploads\auto\C%3A'
        $extendRoot = Join-Path $tempRoot 'Collection\uploads\ntfs\%5C%5C.%5CC%3A\$Extend'
        $usnPath = Join-Path $extendRoot '$UsnJrnl%3A$J'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $extendRoot -Force | Out-Null
        'usn' | Out-File -LiteralPath $usnPath -Encoding ASCII

        try {
            Find-IbisUsnJournalPath -SourceRoot $sourceRoot | Should Be $usnPath
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'finds renamed and standalone extracted USN Journal $J files under $Extend' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Collection\uploads\auto\C%3A'
        $extendRoot = Join-Path $tempRoot 'Collection\uploads\ntfs\%5C%5C.%5CC%3A\$Extend'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $extendRoot -Force | Out-Null

        try {
            foreach ($name in @('$UsnJrnl_$J', '$UsnJrnl-$J', '$J', '$UsnJrnl')) {
                $filePath = Join-Path $extendRoot $name
                'usn' | Out-File -LiteralPath $filePath -Encoding ASCII
                Find-IbisUsnJournalPath -SourceRoot $sourceRoot | Should Be $filePath
                Remove-Item -LiteralPath $filePath -Force
            }
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'prefers the canonical encoded USN Journal name and records discovery details' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Collection\uploads\auto\C%3A'
        $ntfsRoot = Join-Path $tempRoot 'Collection\uploads\ntfs\%5C%5C.%5CC%3A'
        $extendRoot = Join-Path $ntfsRoot '$Extend'
        $outputRoot = Join-Path $tempRoot 'Output'
        $canonicalPath = Join-Path $extendRoot '$UsnJrnl%3A$J'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $extendRoot -Force | Out-Null
        'mft' | Out-File -LiteralPath (Join-Path $ntfsRoot '$MFT') -Encoding ASCII
        'usn' | Out-File -LiteralPath $canonicalPath -Encoding ASCII
        'alternate' | Out-File -LiteralPath (Join-Path $extendRoot '$UsnJrnl_$J') -Encoding ASCII

        try {
            Find-IbisUsnJournalPath -SourceRoot $sourceRoot | Should Be $canonicalPath
            $result = Invoke-IbisNtfsMetadata -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'
            $summary = Get-Content -LiteralPath $result.JsonPath -Raw | ConvertFrom-Json
            $usn = $summary.LocatedArtifacts | Where-Object { $_.Name -eq '$UsnJrnl:$J' } | Select-Object -First 1

            $usn.DiscoveryMethod | Should Be 'NtfsUploadEncodedJournalStream'
            $usn.DiscoveryConfidence | Should Be 'High'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'treats a USN Journal stream path as present when its base file exists' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $extendRoot = Join-Path $tempRoot '$Extend'
        $basePath = Join-Path $extendRoot '$UsnJrnl'
        New-Item -ItemType Directory -Path $extendRoot -Force | Out-Null
        'usn' | Out-File -LiteralPath $basePath -Encoding ASCII

        try {
            Test-IbisNtfsSpecialFilePath -Path ($basePath + ':$J') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips NTFS metadata cleanly when no supported files are found' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisNtfsMetadata -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'ntfs-metadata'
            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues NTFS metadata processing with failures noted when MFTECmd is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        'mft' | Out-File -LiteralPath (Join-Path $sourceRoot '$MFT') -Encoding ASCII

        try {
            $result = Invoke-IbisNtfsMetadata -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'ntfs-metadata'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues NTFS metadata processing without a host folder when hostname is blank' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        'mft' | Out-File -LiteralPath (Join-Path $sourceRoot '$MFT') -Encoding ASCII

        try {
            $result = Invoke-IbisNtfsMetadata -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname ''

            $result.ModuleId | Should Be 'ntfs-metadata'
            $result.Status | Should Be 'Failed'
            $result.HostOutputRoot | Should Be $outputRoot
            Test-Path -LiteralPath (Join-Path $outputRoot 'NTFS-Metadata\_Working\NTFS-Metadata.json') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'HOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'marks USN Journal processing ready only when both $J and $MFT are found' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Collection\uploads\auto\C%3A'
        $ntfsRoot = Join-Path $tempRoot 'Collection\uploads\ntfs\%5C%5C.%5CC%3A'
        $extendRoot = Join-Path $ntfsRoot '$Extend'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $extendRoot -Force | Out-Null
        'mft' | Out-File -LiteralPath (Join-Path $ntfsRoot '$MFT') -Encoding ASCII
        'usn' | Out-File -LiteralPath (Join-Path $extendRoot '$UsnJrnl%3A$J') -Encoding ASCII

        try {
            $result = Invoke-IbisNtfsMetadata -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'
            $summary = Get-Content -LiteralPath $result.JsonPath -Raw | ConvertFrom-Json
            $usn = $summary.LocatedArtifacts | Where-Object { $_.Name -eq '$UsnJrnl:$J' } | Select-Object -First 1

            $result.Status | Should Be 'Failed'
            $usn.Found | Should Be $true
            $usn.ReadyToProcess | Should Be $true
            $usn.OutputFileName | Should Be 'TESTHOST-MFTECmd-UsnJrnl-J-Output.csv'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Describe 'Ibis SRUM and user artefacts' {
    It 'marks SRUM and user artefact modules as implemented' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        $srum = $config.modules | Where-Object { $_.id -eq 'srum' } | Select-Object -First 1
        $users = $config.modules | Where-Object { $_.id -eq 'user-artifacts' } | Select-Object -First 1

        $srum.status | Should Be 'implemented'
        $users.status | Should Be 'implemented'
    }

    It 'calculates the standard SRUM database path' {
        Get-IbisSrumDatabasePath -SourceRoot 'E:\Evidence' | Should Be 'E:\Evidence\Windows\System32\sru\SRUDB.dat'
    }

    It 'renames SrumECmd timestamped CSV outputs to the host naming format' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        'csv' | Out-File -LiteralPath (Join-Path $tempRoot '20260427005500_SrumECmd_AppResourceUseInfo_Output.csv') -Encoding ASCII

        try {
            $renamed = @(Rename-IbisSrumECmdOutput -OutputDirectory $tempRoot -Hostname 'TESTHOST')

            $renamed.Count | Should Be 1
            Test-Path -LiteralPath (Join-Path $tempRoot 'TESTHOST-SrumECmd-AppResourceUseInfo_Output.csv') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips SRUM cleanly when required source files are absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisSrum `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues SRUM processing with failures noted when SrumECmd is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $srumRoot = Join-Path $sourceRoot 'Windows\System32\sru'
        $hiveRoot = Join-Path $sourceRoot 'Windows\System32\config'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $srumRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $hiveRoot -Force | Out-Null
        'srum' | Out-File -LiteralPath (Join-Path $srumRoot 'SRUDB.dat') -Encoding ASCII
        'software' | Out-File -LiteralPath (Join-Path $hiveRoot 'SOFTWARE') -Encoding ASCII

        try {
            $result = Invoke-IbisSrum `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'srum'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Registry-Hives\_Working\Prepared-Hives\SOFTWARE\SOFTWARE') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'discovers user profile artefact paths' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $profileRoot = Join-Path $sourceRoot 'Users\Alice'
        $defaultRoot = Join-Path $sourceRoot 'Users\Default'

        New-Item -ItemType Directory -Path $profileRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $defaultRoot -Force | Out-Null

        try {
            $profiles = @(Get-IbisUserProfile -SourceRoot $sourceRoot)

            $profiles.Count | Should Be 2
            @($profiles.UserName) -contains 'Alice' | Should Be $true
            @($profiles.UserName) -contains 'Default' | Should Be $true
            $aliceProfile = $profiles | Where-Object { $_.UserName -eq 'Alice' } | Select-Object -First 1
            $aliceProfile.NtUserPath | Should Be (Join-Path $profileRoot 'NTUSER.dat')
            $aliceProfile.UsrClassPath | Should Be (Join-Path $profileRoot 'AppData\Local\Microsoft\Windows\UsrClass.dat')
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'copies PSReadLine history for a user profile' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'PSReadLine'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null
        'Get-Process' | Out-File -LiteralPath (Join-Path $sourceRoot 'ConsoleHost_history.txt') -Encoding ASCII

        try {
            $result = Copy-IbisPSReadLineHistory -SourceDirectory $sourceRoot -OutputDirectory $outputRoot -Hostname 'TESTHOST' -UserName 'Alice'

            $result.Status | Should Be 'Completed'
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST-Alice-PSReadLine-ConsoleHost_history.txt') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'PSReadLine\ConsoleHost_history.txt') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'normalizes user artefact tool output names' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $jumpListRoot = Join-Path $tempRoot 'JumpLists'
        $lnkRoot = Join-Path $tempRoot 'RecentLNKs'
        $shellBagRoot = Join-Path $tempRoot 'ShellBags'

        New-Item -ItemType Directory -Path $jumpListRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $lnkRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $shellBagRoot -Force | Out-Null
        'csv' | Out-File -LiteralPath (Join-Path $jumpListRoot '20260425062602_AutomaticDestinations.csv') -Encoding ASCII
        New-Item -ItemType Directory -Path (Join-Path $jumpListRoot '20260425062602_CustomDestinations') | Out-Null
        'csv' | Out-File -LiteralPath (Join-Path $lnkRoot '20260425062603_LECmd_Output.csv') -Encoding ASCII
        'csv' | Out-File -LiteralPath (Join-Path $shellBagRoot 'Alice_NTUSER.csv') -Encoding ASCII
        'csv' | Out-File -LiteralPath (Join-Path $shellBagRoot 'Alice_UsrClass.csv') -Encoding ASCII

        try {
            @(Rename-IbisUserArtifactToolOutput -OutputDirectory $jumpListRoot -Hostname 'TESTHOST' -UserName 'Alice' -ToolName 'JLECmd').Count | Should Be 2
            @(Rename-IbisUserArtifactToolOutput -OutputDirectory $lnkRoot -Hostname 'TESTHOST' -UserName 'Alice' -ToolName 'LECmd').Count | Should Be 1
            @(Rename-IbisUserArtifactToolOutput -OutputDirectory $shellBagRoot -Hostname 'TESTHOST' -UserName 'Alice' -ToolName 'SBECmd').Count | Should Be 2

            Test-Path -LiteralPath (Join-Path $jumpListRoot 'TESTHOST-Alice-JLECmd-AutomaticDestinations.csv') | Should Be $true
            Test-Path -LiteralPath (Join-Path $jumpListRoot 'TESTHOST-Alice-JLECmd-CustomDestinations') | Should Be $true
            Test-Path -LiteralPath (Join-Path $lnkRoot 'TESTHOST-Alice-LECmd-Output.csv') | Should Be $true
            Test-Path -LiteralPath (Join-Path $shellBagRoot 'TESTHOST-Alice-SBECmd-NTUSER.csv') | Should Be $true
            Test-Path -LiteralPath (Join-Path $shellBagRoot 'TESTHOST-Alice-SBECmd-UsrClass.csv') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips user artefacts cleanly when no user profiles exist' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisUserArtifacts `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.Status | Should Be 'Skipped'
            $result.UserCount | Should Be 0
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues user artefact processing with failures noted when tools are missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $profileRoot = Join-Path $sourceRoot 'Users\Alice'
        $psReadLineRoot = Join-Path $profileRoot 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $psReadLineRoot -Force | Out-Null
        'ntuser' | Out-File -LiteralPath (Join-Path $profileRoot 'NTUSER.dat') -Encoding ASCII
        'Get-ChildItem' | Out-File -LiteralPath (Join-Path $psReadLineRoot 'ConsoleHost_history.txt') -Encoding ASCII

        try {
            $result = Invoke-IbisUserArtifacts `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'user-artifacts'
            $result.Status | Should Be 'Failed'
            $result.UserCount | Should Be 1
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Users\_Working\Prepared-Hives\Alice-NTUSER\NTUSER.dat') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Users\Alice\PSReadLine\TESTHOST-Alice-PSReadLine-ConsoleHost_history.txt') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Users\Alice\PSReadLine\PSReadLine\ConsoleHost_history.txt') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'does not create a host folder when processing user profiles with a blank hostname' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $aliceRoot = Join-Path $sourceRoot 'Users\Alice'
        $defaultRoot = Join-Path $sourceRoot 'Users\Default'
        $psReadLineRoot = Join-Path $aliceRoot 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $psReadLineRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $defaultRoot -Force | Out-Null
        'Get-Date' | Out-File -LiteralPath (Join-Path $psReadLineRoot 'ConsoleHost_history.txt') -Encoding ASCII

        try {
            $result = Invoke-IbisUserArtifacts `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname ''

            $result.UserCount | Should Be 2
            $result.HostOutputRoot | Should Be $outputRoot
            Test-Path -LiteralPath (Join-Path $outputRoot 'Users\Alice\PSReadLine\Alice-PSReadLine-ConsoleHost_history.txt') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'Users\Default') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'HOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues processing later user artefacts and profiles when one user tool throws' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        $toolsRoot = Join-Path $tempRoot 'Tools'
        $badToolPath = Join-Path $toolsRoot 'bad-tool.txt'
        $aliceRoot = Join-Path $sourceRoot 'Users\Alice'
        $bobRoot = Join-Path $sourceRoot 'Users\Bob'
        $aliceRecent = Join-Path $aliceRoot 'AppData\Roaming\Microsoft\Windows\Recent'
        $bobRecent = Join-Path $bobRoot 'AppData\Roaming\Microsoft\Windows\Recent'
        $alicePsReadLine = Join-Path $aliceRoot 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine'
        $bobPsReadLine = Join-Path $bobRoot 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine'

        New-Item -ItemType Directory -Path $aliceRecent,$bobRecent,$alicePsReadLine,$bobPsReadLine,$toolsRoot -Force | Out-Null
        'not an executable' | Out-File -LiteralPath $badToolPath -Encoding ASCII
        'Get-Date' | Out-File -LiteralPath (Join-Path $alicePsReadLine 'ConsoleHost_history.txt') -Encoding ASCII
        'Get-Process' | Out-File -LiteralPath (Join-Path $bobPsReadLine 'ConsoleHost_history.txt') -Encoding ASCII

        $toolDefinitions = @(
            [pscustomobject]@{
                id = 'zimmerman-jlecmd'
                name = 'Bad JLECmd'
                executablePath = 'bad-tool.txt'
            }
        )

        try {
            $result = Invoke-IbisUserArtifacts `
                -ToolsRoot $toolsRoot `
                -ToolDefinitions $toolDefinitions `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.Status | Should Be 'Failed'
            $result.UserCount | Should Be 2
            @($result.Users).Count | Should Be 2

            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Users\Alice\PSReadLine\TESTHOST-Alice-PSReadLine-ConsoleHost_history.txt') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Users\Bob\PSReadLine\TESTHOST-Bob-PSReadLine-ConsoleHost_history.txt') | Should Be $true

            $alice = @($result.Users | Where-Object { $_.UserName -eq 'Alice' } | Select-Object -First 1)[0]
            $bob = @($result.Users | Where-Object { $_.UserName -eq 'Bob' } | Select-Object -First 1)[0]
            @($alice.ToolResults | Where-Object { $_.ToolId -eq 'zimmerman-jlecmd' -and $_.Status -eq 'Failed' }).Count | Should Be 1
            @($alice.ToolResults | Where-Object { $_.ToolId -eq 'zimmerman-lecmd' }).Count | Should Be 1
            @($alice.ToolResults | Where-Object { $_.ToolId -eq 'zimmerman-sbecmd' }).Count | Should Be 1
            $alice.PSReadLine.Status | Should Be 'Completed'
            $bob.PSReadLine.Status | Should Be 'Completed'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues processing later artefacts and profiles when the first user is missing a target artefact path' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        $aliceRoot = Join-Path $sourceRoot 'Users\Alice'
        $bobRoot = Join-Path $sourceRoot 'Users\Bob'
        $alicePsReadLine = Join-Path $aliceRoot 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine'
        $bobRecent = Join-Path $bobRoot 'AppData\Roaming\Microsoft\Windows\Recent'
        $bobPsReadLine = Join-Path $bobRoot 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine'

        New-Item -ItemType Directory -Path $alicePsReadLine,$bobRecent,$bobPsReadLine -Force | Out-Null
        'Get-Date' | Out-File -LiteralPath (Join-Path $alicePsReadLine 'ConsoleHost_history.txt') -Encoding ASCII
        'Get-Process' | Out-File -LiteralPath (Join-Path $bobPsReadLine 'ConsoleHost_history.txt') -Encoding ASCII

        try {
            $result = Invoke-IbisUserArtifacts `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST'

            $result.UserCount | Should Be 2

            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Users\Alice\PSReadLine\TESTHOST-Alice-PSReadLine-ConsoleHost_history.txt') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Users\Bob\PSReadLine\TESTHOST-Bob-PSReadLine-ConsoleHost_history.txt') | Should Be $true

            $alice = @($result.Users | Where-Object { $_.UserName -eq 'Alice' } | Select-Object -First 1)[0]
            $bob = @($result.Users | Where-Object { $_.UserName -eq 'Bob' } | Select-Object -First 1)[0]

            @($alice.ToolResults | Where-Object { $_.ToolId -eq 'zimmerman-jlecmd' -and $_.Status -eq 'Skipped' }).Count | Should Be 1
            @($alice.ToolResults | Where-Object { $_.ToolId -eq 'zimmerman-lecmd' -and $_.Status -eq 'Skipped' }).Count | Should Be 1
            @($alice.ToolResults | Where-Object { $_.ToolId -eq 'zimmerman-sbecmd' }).Count | Should Be 1
            $alice.PSReadLine.Status | Should Be 'Completed'
            $bob.PSReadLine.Status | Should Be 'Completed'
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'detects hidden user artefact paths and writes per-user progress in order' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        $progressPath = Join-Path $tempRoot 'user-progress.jsonl'
        $aliceRoot = Join-Path $sourceRoot 'Users\Alice'
        $bobRoot = Join-Path $sourceRoot 'Users\Bob'
        $aliceRecent = Join-Path $aliceRoot 'AppData\Roaming\Microsoft\Windows\Recent'
        $bobRecent = Join-Path $bobRoot 'AppData\Roaming\Microsoft\Windows\Recent'
        $alicePsReadLine = Join-Path $aliceRoot 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine'
        $bobPsReadLine = Join-Path $bobRoot 'AppData\Roaming\Microsoft\Windows\PowerShell\PSReadLine'

        New-Item -ItemType Directory -Path $aliceRecent,$bobRecent,$alicePsReadLine,$bobPsReadLine -Force | Out-Null
        'Get-Date' | Out-File -LiteralPath (Join-Path $alicePsReadLine 'ConsoleHost_history.txt') -Encoding ASCII
        'Get-Process' | Out-File -LiteralPath (Join-Path $bobPsReadLine 'ConsoleHost_history.txt') -Encoding ASCII
        (Get-Item -LiteralPath $aliceRecent -Force).Attributes = (Get-Item -LiteralPath $aliceRecent -Force).Attributes -bor [System.IO.FileAttributes]::Hidden

        try {
            $result = Invoke-IbisUserArtifacts `
                -ToolsRoot $tempRoot `
                -ToolDefinitions @() `
                -SourceRoot $sourceRoot `
                -OutputRoot $outputRoot `
                -Hostname 'TESTHOST' `
                -ProgressPath $progressPath `
                -ProgressIndex 3 `
                -ProgressTotal 7

            $alice = @($result.Users | Where-Object { $_.UserName -eq 'Alice' } | Select-Object -First 1)[0]
            @($alice.ToolResults | Where-Object { $_.ToolId -eq 'zimmerman-jlecmd' -and $_.Status -eq 'Failed' }).Count | Should Be 1
            @($alice.ToolResults | Where-Object { $_.ToolId -eq 'zimmerman-lecmd' -and $_.Status -eq 'Failed' }).Count | Should Be 1

            $events = @(Get-Content -LiteralPath $progressPath | ConvertFrom-Json)
            $aliceEventIndexes = @()
            $bobEventIndexes = @()
            for ($eventIndex = 0; $eventIndex -lt $events.Count; $eventIndex++) {
                if ([string]$events[$eventIndex].Stage -like 'Alice / *') {
                    $aliceEventIndexes += $eventIndex
                }
                elseif ([string]$events[$eventIndex].Stage -like 'Bob / *') {
                    $bobEventIndexes += $eventIndex
                }
            }

            ($aliceEventIndexes.Count -gt 0) | Should Be $true
            ($bobEventIndexes.Count -gt 0) | Should Be $true
            (($aliceEventIndexes | Measure-Object -Maximum).Maximum -lt ($bobEventIndexes | Measure-Object -Minimum).Minimum) | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Describe 'Ibis Windows Event Log modules' {
    It 'marks EventLogs, DuckDB summaries, Hayabusa, Takajo, and Chainsaw modules as implemented' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        ($config.modules | Where-Object { $_.id -eq 'eventlogs' } | Select-Object -First 1).status | Should Be 'implemented'
        ($config.modules | Where-Object { $_.id -eq 'duckdb-eventlogs' } | Select-Object -First 1).status | Should Be 'implemented'
        ($config.modules | Where-Object { $_.id -eq 'hayabusa' } | Select-Object -First 1).status | Should Be 'implemented'
        ($config.modules | Where-Object { $_.id -eq 'takajo' } | Select-Object -First 1).status | Should Be 'implemented'
        ($config.modules | Where-Object { $_.id -eq 'chainsaw' } | Select-Object -First 1).status | Should Be 'implemented'
    }

    It 'calculates the standard Windows Event Log source path' {
        Get-IbisEventLogPath -SourceRoot 'E:\Evidence' | Should Be 'E:\Evidence\Windows\System32\winevt\Logs'
    }

    It 'creates a separate EventLogs output directory for each tool' {
        Get-IbisEventLogToolOutputDirectory -OutputRoot 'C:\Export' -Hostname 'TESTHOST' -ToolFolder 'Hayabusa' | Should Be 'C:\Export\TESTHOST\EventLogs\Hayabusa'
    }

    It 'skips EvtxECmd cleanly when the event log folder is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisEvtxECmdEventLogs -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues EvtxECmd processing with failures noted when the tool is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $eventLogRoot = Join-Path $sourceRoot 'Windows\System32\winevt\Logs'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $eventLogRoot -Force | Out-Null
        'evtx' | Out-File -LiteralPath (Join-Path $eventLogRoot 'Security.evtx') -Encoding ASCII

        try {
            $result = Invoke-IbisEvtxECmdEventLogs -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'eventlogs'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'calculates the EvtxECmd CSV path consumed by DuckDB summaries' {
        Get-IbisEvtxECmdCsvPath -OutputRoot 'C:\Export' -Hostname 'TESTHOST' | Should Be 'C:\Export\TESTHOST\EventLogs\EvtxECmd\TESTHOST-EvtxECmd-EventLogs-Output.csv'
    }

    It 'loads DuckDB event log query templates from editable SQL files' {
        $queries = @(Get-IbisDuckDbEventLogQueryDefinition -ProjectRoot $projectRoot)

        $queries.Count | Should Be 3
        @($queries | Where-Object { $_.Id -eq 'time-span' }).Count | Should Be 1
        @($queries | Where-Object { $_.Id -eq 'logons' }).Count | Should Be 1
        @($queries | Where-Object { $_.Id -eq 'outbound-rdp' }).Count | Should Be 1
        foreach ($query in $queries) {
            Test-Path -LiteralPath $query.QueryPath | Should Be $true
        }
    }

    It 'renders DuckDB SQL templates with escaped CSV paths' {
        $queries = @(Get-IbisDuckDbEventLogQueryDefinition -ProjectRoot $projectRoot)
        $sql = Expand-IbisDuckDbSqlTemplate -TemplatePath $queries[0].QueryPath -InputCsvPath "C:\Evidence\Bob's.csv" -OutputCsvPath "C:\Output\Result's.csv"

        $sql | Should Match "Bob''s.csv"
        $sql | Should Match "Result''s.csv"
        $sql | Should Not Match '\{\{INPUT_CSV\}\}'
        $sql | Should Not Match '\{\{OUTPUT_CSV\}\}'
    }

    It 'defines the supported user-session activity events and corrected labels' {
        $sql = Get-Content -LiteralPath (Join-Path $projectRoot 'queries\eventlogs\logons.sql') -Raw

        $sql | Should Match '4624, 4625, 4634, 4647, 4648, 4672, 4776, 4778, 4779, 4800, 4801, 4802, 4803'
        $sql | Should Match "4800 THEN 'WorkstationLock'"
        $sql | Should Match "4801 THEN 'WorkstationUnlock'"
        $sql | Should Match "4802 THEN 'ScreenSaverInvoked'"
        $sql | Should Match "4803 THEN 'ScreenSaverDismissed'"
        $sql | Should Match "4778 THEN 'WindowStationReconnect'"
        $sql | Should Match "4779 THEN 'WindowStationDisconnect'"
        $sql | Should Match "4776 THEN 'NTLMCredentialValidation'"
        $sql | Should Not Match "4776 THEN 'DCAuthenticationAttempt'"
    }

    It 'retains user-session correlation and source provenance fields' {
        $sql = Get-Content -LiteralPath (Join-Path $projectRoot 'queries\eventlogs\logons.sql') -Raw

        foreach ($fieldName in @('SecurityLogonID', 'TerminalSessionID', 'RelatedSessionID', 'SessionName', 'DisconnectReason', 'RecordNumber', 'EventRecordId', 'Level', 'Provider', 'ProcessId', 'ThreadId', 'Computer', 'ChunkNumber', 'UserId', 'HiddenRecord', 'SourceFile', 'Keywords', 'ExtraDataOffset', 'Payload')) {
            $sql | Should Match ([regex]::Escape($fieldName))
        }
        $sql | Should Match 'all_varchar = true'
        $sql | Should Match 'TRY_CAST\(TRIM\(EventId\) AS INTEGER\)'
    }

    It 'stages DuckDB results before replacing a final analyst-facing output' {
        $moduleText = Get-Content -LiteralPath (Join-Path $projectRoot 'modules\Ibis.Core.psm1') -Raw

        $moduleText | Should Match 'StagedOutputPath'
        $moduleText | Should Match 'Move-Item -LiteralPath \$stagedOutputPath -Destination \$outputPath -Force'
        $moduleText | Should Match 'any previous final output was preserved'
    }

    It 'includes representative synthetic rows for each newly covered workstation state event' {
        $events = @(Import-Csv -LiteralPath (Join-Path $projectRoot 'tests\fixtures\eventlogs\user-session-events.csv'))

        foreach ($eventId in @('4647', '4800', '4801', '4802', '4803')) {
            @($events | Where-Object { $_.EventId -eq $eventId }).Count | Should Be 1
        }
    }

    It 'skips DuckDB summaries cleanly when the EvtxECmd CSV is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $outputRoot = Join-Path $tempRoot 'Output'

        try {
            $result = Invoke-IbisDuckDbEventLogSummary -ToolsRoot $tempRoot -ToolDefinitions @() -OutputRoot $outputRoot -Hostname 'TESTHOST' -ProjectRoot $projectRoot

            $result.ModuleId | Should Be 'duckdb-eventlogs'
            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            if (Test-Path -LiteralPath $tempRoot) {
                Remove-Item -LiteralPath $tempRoot -Recurse -Force
            }
        }
    }

    It 'continues DuckDB summary processing with failures noted when DuckDB is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $outputRoot = Join-Path $tempRoot 'Output'
        $eventLogRoot = Join-Path $outputRoot 'TESTHOST\EventLogs\EvtxECmd'
        New-Item -ItemType Directory -Path $eventLogRoot -Force | Out-Null
        'TimeCreated,Channel,EventId' | Out-File -LiteralPath (Join-Path $eventLogRoot 'TESTHOST-EvtxECmd-EventLogs-Output.csv') -Encoding ASCII

        try {
            $result = Invoke-IbisDuckDbEventLogSummary -ToolsRoot $tempRoot -ToolDefinitions @() -OutputRoot $outputRoot -Hostname 'TESTHOST' -ProjectRoot $projectRoot

            $result.ModuleId | Should Be 'duckdb-eventlogs'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'moves an existing tool output directory to a timestamped backup' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $existing = Join-Path $tempRoot 'Takajo'
        $backupRoot = Join-Path $tempRoot 'Backups'
        New-Item -ItemType Directory -Path $existing -Force | Out-Null
        'old' | Out-File -LiteralPath (Join-Path $existing 'old.txt') -Encoding ASCII

        try {
            $backupPath = Move-IbisExistingDirectoryToBackup -DirectoryPath $existing -BackupRoot $backupRoot

            Test-Path -LiteralPath $existing | Should Be $false
            Test-Path -LiteralPath (Join-Path $backupPath 'old.txt') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues Hayabusa processing with failures noted when Hayabusa is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $eventLogRoot = Join-Path $sourceRoot 'Windows\System32\winevt\Logs'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $eventLogRoot -Force | Out-Null
        'evtx' | Out-File -LiteralPath (Join-Path $eventLogRoot 'Security.evtx') -Encoding ASCII

        try {
            $result = Invoke-IbisHayabusaEventLogs -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'hayabusa'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'calculates the Hayabusa JSONL path consumed by Takajo' {
        Get-IbisHayabusaJsonlPath -OutputRoot 'C:\Export' -Hostname 'TESTHOST' | Should Be 'C:\Export\TESTHOST\EventLogs\Hayabusa\TESTHOST-Hayabusa-EventLogs-SuperVerbose.jsonl'
    }

    It 'skips Takajo cleanly when the Hayabusa JSONL is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $outputRoot = Join-Path $tempRoot 'Output'

        try {
            $result = Invoke-IbisTakajoEventLogs -ToolsRoot $tempRoot -ToolDefinitions @() -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'takajo'
            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            if (Test-Path -LiteralPath $tempRoot) {
                Remove-Item -LiteralPath $tempRoot -Recurse -Force
            }
        }
    }

    It 'continues Takajo processing with failures noted when Takajo is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $outputRoot = Join-Path $tempRoot 'Output'
        $eventLogRoot = Join-Path $outputRoot 'TESTHOST\EventLogs\Hayabusa'
        New-Item -ItemType Directory -Path $eventLogRoot -Force | Out-Null
        '{}' | Out-File -LiteralPath (Join-Path $eventLogRoot 'TESTHOST-Hayabusa-EventLogs-SuperVerbose.jsonl') -Encoding ASCII

        try {
            $result = Invoke-IbisTakajoEventLogs -ToolsRoot $tempRoot -ToolDefinitions @() -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'takajo'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
            $result.JsonPath | Should Match ([regex]::Escape('TESTHOST\EventLogs\Takajo\TESTHOST-Takajo.json'))
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\EventLogs\_Working\Takajo') | Should Be $false
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\EventLogs\_Working') | Should Be $false
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\EventLogs\Takajo\_Working') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'renames Chainsaw staged CSV outputs into its own EventLogs folder' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $stagingRoot = Join-Path $tempRoot 'Chainsaw'
        $outputRoot = Join-Path $tempRoot 'EventLogs'
        New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
        'csv' | Out-File -LiteralPath (Join-Path $stagingRoot 'detections.csv') -Encoding ASCII

        try {
            $moved = @(Rename-IbisChainsawOutput -StagingDirectory $stagingRoot -OutputDirectory $outputRoot -Hostname 'TESTHOST')

            $moved.Count | Should Be 1
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST-Chainsaw-detections.csv') | Should Be $true
            Test-Path -LiteralPath $stagingRoot | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'normalizes all Takajo result files, including nested automagic output files' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $takajoRoot = Join-Path $tempRoot 'Takajo'
        $nestedRoot = Join-Path $takajoRoot 'automagic\summary'
        $workingRoot = Join-Path $takajoRoot '_Working'
        New-Item -ItemType Directory -Path $nestedRoot -Force | Out-Null
        New-Item -ItemType Directory -Path $workingRoot -Force | Out-Null
        'csv' | Out-File -LiteralPath (Join-Path $takajoRoot 'timeline.csv') -Encoding ASCII
        'json' | Out-File -LiteralPath (Join-Path $nestedRoot 'findings.json') -Encoding ASCII
        'summary' | Out-File -LiteralPath (Join-Path $nestedRoot 'Summary.csv') -Encoding ASCII
        'work' | Out-File -LiteralPath (Join-Path $workingRoot 'TESTHOST-Takajo.json') -Encoding ASCII

        try {
            $renamed = @(Rename-IbisToolOutputFiles -SourceDirectory $takajoRoot -OutputDirectory $takajoRoot -Hostname 'TESTHOST' -ToolName 'Takajo')

            $renamed.Count | Should Be 3
            Test-Path -LiteralPath (Join-Path $takajoRoot 'TESTHOST-Takajo-timeline.csv') | Should Be $true
            Test-Path -LiteralPath (Join-Path $nestedRoot 'TESTHOST-Takajo-findings.json') | Should Be $true
            Test-Path -LiteralPath (Join-Path $nestedRoot 'TESTHOST-Takajo-Summary.csv') | Should Be $true
            Test-Path -LiteralPath (Join-Path $workingRoot 'TESTHOST-Takajo.json') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues Chainsaw processing with failures noted when Chainsaw is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $eventLogRoot = Join-Path $sourceRoot 'Windows\System32\winevt\Logs'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $eventLogRoot -Force | Out-Null
        'evtx' | Out-File -LiteralPath (Join-Path $eventLogRoot 'Security.evtx') -Encoding ASCII

        try {
            $result = Invoke-IbisChainsawEventLogs -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'chainsaw'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Describe 'Ibis remaining artefact modules' {
    It 'marks UAL, Browser History, Forensic webhistory, and USB modules as implemented' {
        $config = Get-IbisConfig -ProjectRoot $projectRoot
        ($config.modules | Where-Object { $_.id -eq 'ual' } | Select-Object -First 1).status | Should Be 'implemented'
        ($config.modules | Where-Object { $_.id -eq 'ual' } | Select-Object -First 1).name | Should Be 'User Access Logs / SUM'
        ($config.modules | Where-Object { $_.id -eq 'ual' } | Select-Object -First 1).hint | Should Match 'Windows Server'
        ($config.modules | Where-Object { $_.id -eq 'browser-history' } | Select-Object -First 1).status | Should Be 'implemented'
        ($config.modules | Where-Object { $_.id -eq 'forensic-webhistory' } | Select-Object -First 1).status | Should Be 'implemented'
        ($config.modules | Where-Object { $_.id -eq 'usb' } | Select-Object -First 1).status | Should Be 'implemented'
    }

    It 'calculates the User Access Logs / SUM source path' {
        Get-IbisUserAccessLogPath -SourceRoot 'E:\Evidence' | Should Be 'E:\Evidence\Windows\System32\LogFiles\Sum'
    }

    It 'calculates the browser history users source path' {
        Get-IbisBrowserHistoryUsersPath -SourceRoot 'E:\Evidence' | Should Be 'E:\Evidence\Users'
    }

    It 'groups browser-history tool output beneath WebHistory' {
        Get-IbisWebHistoryToolOutputDirectory -OutputRoot 'D:\Output' -Hostname 'TESTHOST' -ToolFolder 'BrowsingHistoryView' | Should Be 'D:\Output\TESTHOST\WebHistory\BrowsingHistoryView'
        Get-IbisWebHistoryToolOutputDirectory -OutputRoot 'D:\Output' -Hostname '' -ToolFolder 'ForensicWebHistory' | Should Be 'D:\Output\WebHistory\ForensicWebHistory'
    }

    It 'renames SumECmd timestamped CSV outputs to the host naming format' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        'csv' | Out-File -LiteralPath (Join-Path $tempRoot '20260425062602_SumECmd_Output.csv') -Encoding ASCII

        try {
            $renamed = @(Rename-IbisSumECmdOutput -OutputDirectory $tempRoot -Hostname 'TESTHOST')

            $renamed.Count | Should Be 1
            Test-Path -LiteralPath (Join-Path $tempRoot 'TESTHOST-SumECmd-Output.csv') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips User Access Logs / SUM cleanly when the source folder is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisUserAccessLogsSum -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'ual'
            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues User Access Logs / SUM processing with failures noted when SumECmd is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $sumRoot = Join-Path $sourceRoot 'Windows\System32\LogFiles\Sum'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sumRoot -Force | Out-Null
        'ual' | Out-File -LiteralPath (Join-Path $sumRoot 'SystemIdentity.mdb') -Encoding ASCII

        try {
            $result = Invoke-IbisUserAccessLogsSum -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'ual'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips browser history cleanly when the Users folder is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisBrowsingHistoryView -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'browser-history'
            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'continues browser history processing with failures noted when BrowsingHistoryView is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $usersRoot = Join-Path $sourceRoot 'Users'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $usersRoot -Force | Out-Null

        try {
            $result = Invoke-IbisBrowsingHistoryView -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'browser-history'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
            $result.OutputPath | Should Be (Join-Path $outputRoot 'TESTHOST\WebHistory\BrowsingHistoryView\TESTHOST-BrowsingHistoryView-All-Users.csv')
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'does not prefix BrowsingHistoryView output filenames when hostname is blank' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $usersRoot = Join-Path $sourceRoot 'Users'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $usersRoot -Force | Out-Null

        try {
            $result = Invoke-IbisBrowsingHistoryView -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname ''

            $result.OutputPath | Should Be (Join-Path $outputRoot 'WebHistory\BrowsingHistoryView\BrowsingHistoryView-All-Users.csv')
            $result.OutputPath.Contains('\HOST\') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'renames Forensic webhistory staged outputs to the host naming format' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $stagingRoot = Join-Path $tempRoot 'Staging'
        $outputRoot = Join-Path $tempRoot 'Output'
        $nestedRoot = Join-Path $stagingRoot 'chrome'
        New-Item -ItemType Directory -Path $nestedRoot -Force | Out-Null
        'csv' | Out-File -LiteralPath (Join-Path $nestedRoot 'history.csv') -Encoding ASCII

        try {
            $moved = @(Move-IbisForensicWebHistoryOutput -StagingDirectory $stagingRoot -OutputDirectory $outputRoot -Hostname 'TESTHOST')

            $moved.Count | Should Be 1
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST-ForensicWebHistory-chrome-history.csv') | Should Be $true
            Test-Path -LiteralPath $stagingRoot | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'renames Forensic webhistory staged outputs without a hostname prefix when hostname is blank' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $stagingRoot = Join-Path $tempRoot 'Staging'
        $outputRoot = Join-Path $tempRoot 'Output'
        $nestedRoot = Join-Path $stagingRoot 'chrome'
        New-Item -ItemType Directory -Path $nestedRoot -Force | Out-Null
        'csv' | Out-File -LiteralPath (Join-Path $nestedRoot 'history.csv') -Encoding ASCII

        try {
            $moved = @(Move-IbisForensicWebHistoryOutput -StagingDirectory $stagingRoot -OutputDirectory $outputRoot -Hostname '')

            $moved.Count | Should Be 1
            Test-Path -LiteralPath (Join-Path $outputRoot 'ForensicWebHistory-chrome-history.csv') | Should Be $true
            Test-Path -LiteralPath $stagingRoot | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips Forensic webhistory cleanly when the source root is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'

        try {
            $result = Invoke-IbisForensicWebHistory -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'forensic-webhistory'
            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            if (Test-Path -LiteralPath $tempRoot) {
                Remove-Item -LiteralPath $tempRoot -Recurse -Force
            }
        }
    }

    It 'continues Forensic webhistory processing with failures noted when the tool is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisForensicWebHistory -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'forensic-webhistory'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips USB artefacts cleanly when the source root is absent' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'

        try {
            $result = Invoke-IbisParseUsbArtifacts -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'usb'
            $result.Status | Should Be 'Skipped'
            $result.JsonPath | Should Be $null
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST') | Should Be $false
        }
        finally {
            if (Test-Path -LiteralPath $tempRoot) {
                Remove-Item -LiteralPath $tempRoot -Recurse -Force
            }
        }
    }

    It 'continues USB artefact processing with failures noted when parseusbs is missing' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Evidence'
        $outputRoot = Join-Path $tempRoot 'Output'
        New-Item -ItemType Directory -Path $sourceRoot -Force | Out-Null

        try {
            $result = Invoke-IbisParseUsbArtifacts -ToolsRoot $tempRoot -ToolDefinitions @() -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'

            $result.ModuleId | Should Be 'usb'
            $result.Status | Should Be 'Failed'
            Test-Path -LiteralPath $result.JsonPath | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Describe 'Ibis Velociraptor Results copy-out' {
    It 'finds and copies Results near the source root' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Collection\Uploads\auto\C'
        $resultsRoot = Join-Path $tempRoot 'Collection\Results'
        $outputRoot = Join-Path $tempRoot 'Output'
        $resultFile = Join-Path $resultsRoot 'artifact.csv'

        New-Item -ItemType Directory -Path $sourceRoot | Out-Null
        New-Item -ItemType Directory -Path $resultsRoot | Out-Null
        'header,value' | Out-File -LiteralPath $resultFile -Encoding ASCII

        try {
            $found = Find-IbisVelociraptorResultsPath -SourceRoot $sourceRoot
            $found | Should Be $resultsRoot

            $result = Invoke-IbisVelociraptorResultsCopy -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'
            $result.Status | Should Be 'Completed'
            $result.SourceItemCount | Should Be 1
            $result.CopiedItemCount | Should Be 1
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Velociraptor-Results\artifact.csv') | Should Be $true
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'copies Results directly under the output root when hostname is blank' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Collection\Uploads\auto\C'
        $resultsRoot = Join-Path $tempRoot 'Collection\Results'
        $outputRoot = Join-Path $tempRoot 'Output'
        $resultFile = Join-Path $resultsRoot 'artifact.csv'

        New-Item -ItemType Directory -Path $sourceRoot | Out-Null
        New-Item -ItemType Directory -Path $resultsRoot | Out-Null
        'header,value' | Out-File -LiteralPath $resultFile -Encoding ASCII

        try {
            $result = Invoke-IbisVelociraptorResultsCopy -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname ''
            $result.Status | Should Be 'Completed'
            $result.OutputPath | Should Be (Join-Path $outputRoot 'Velociraptor-Results')
            Test-Path -LiteralPath (Join-Path $outputRoot 'Velociraptor-Results\artifact.csv') | Should Be $true
            Test-Path -LiteralPath (Join-Path $outputRoot 'HOST') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips cleanly when no Results folder exists near the source root' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Collection\Uploads\auto\C'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $sourceRoot | Out-Null

        try {
            $found = Find-IbisVelociraptorResultsPath -SourceRoot $sourceRoot
            $found | Should Be $null

            $result = Invoke-IbisVelociraptorResultsCopy -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'
            $result.Status | Should Be 'Skipped'
            $result.SourcePath | Should Be $null
            $result.SourceItemCount | Should Be 0
            $result.CopiedItemCount | Should Be 0
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Velociraptor-Results') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }

    It 'skips cleanly when a Results folder exists but contains no files' {
        $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
        $sourceRoot = Join-Path $tempRoot 'Collection\Uploads\auto\C'
        $resultsRoot = Join-Path $tempRoot 'Collection\Results'
        $emptyArtifactRoot = Join-Path $resultsRoot 'EmptyArtifact'
        $outputRoot = Join-Path $tempRoot 'Output'

        New-Item -ItemType Directory -Path $sourceRoot | Out-Null
        New-Item -ItemType Directory -Path $emptyArtifactRoot | Out-Null

        try {
            $result = Invoke-IbisVelociraptorResultsCopy -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'
            $result.Status | Should Be 'Skipped'
            $result.SourcePath | Should Be $resultsRoot
            $result.SourceItemCount | Should Be 1
            $result.CopiedItemCount | Should Be 0
            Test-Path -LiteralPath (Join-Path $outputRoot 'TESTHOST\Velociraptor-Results') | Should Be $false
        }
        finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

