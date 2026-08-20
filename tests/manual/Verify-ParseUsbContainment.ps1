<#
.SYNOPSIS
    Manual evidence-safety check for the ParseUSBs processing module.

.DESCRIPTION
    Proves that Invoke-IbisParseUsbArtifacts cannot alter source evidence.

    The script builds a synthetic Windows evidence tree, marks every source file
    read-only the way a mounted image behaves, compiles a stand-in parseusbs.exe
    that records the command line it is given, then runs the real module and
    compares a SHA-256 and last-write snapshot of the source tree taken before
    and after the run.

    This is not part of the Pester suite for two reasons: it compiles a stand-in
    executable with Add-Type -OutputAssembly, which Windows PowerShell 5.1
    supports but PowerShell 7 does not, and it is a slower whole-module check
    rather than a unit test.  Run it by hand after changing ParseUSBs staging,
    argument construction, or evidence copy behaviour.

    No real forensic tools, evidence, or network access are needed.  Everything
    is created under, and removed from, the temporary directory.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\manual\Verify-ParseUsbContainment.ps1

.NOTES
    Requires Windows PowerShell 5.1.  Writes a PASS/FAIL line per check and
    exits with code 1 if any check fails.
#>
[CmdletBinding()]
param(
    [string]$ProjectRoot
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ProjectRoot = Split-Path -Parent (Split-Path -Parent $scriptDirectory)
}
Import-Module (Join-Path $ProjectRoot 'modules\Ibis.Core.psm1') -Force

$tempRoot   = Join-Path ([System.IO.Path]::GetTempPath()) ("ibis-usb-verify-" + [guid]::NewGuid().ToString('N'))
$sourceRoot = Join-Path $tempRoot 'Evidence'
$outputRoot = Join-Path $tempRoot 'Output'
$toolsRoot  = Join-Path $tempRoot 'Tools'

function New-EvidenceFile {
    param([string]$Path, [string]$Content)
    $directory = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $directory)) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
    Set-Content -LiteralPath $Path -Value $Content -Encoding Ascii
}

function Get-TreeSnapshot {
    param([string]$Root)
    Get-ChildItem -LiteralPath $Root -File -Recurse -Force |
        Sort-Object FullName |
        ForEach-Object {
            [pscustomobject]@{
                Path         = $_.FullName.Substring($Root.Length)
                Hash         = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                LastWriteUtc = $_.LastWriteTimeUtc.ToString('o')
            }
        }
}

$failures = @()
function Assert-True {
    param([bool]$Condition, [string]$Label)
    if ($Condition) { Write-Host ("  [PASS] " + $Label) -ForegroundColor Green }
    else { Write-Host ("  [FAIL] " + $Label) -ForegroundColor Red; $script:failures += $Label }
}

try {
    # ---- Fake evidence -----------------------------------------------------
    $configRoot = Join-Path $sourceRoot 'Windows\System32\config'
    New-EvidenceFile (Join-Path $configRoot 'SYSTEM')        'SYSTEM hive body'
    New-EvidenceFile (Join-Path $configRoot 'SYSTEM.LOG1')   'SYSTEM dirty transaction log'
    New-EvidenceFile (Join-Path $configRoot 'SYSTEM.LOG2')   'SYSTEM dirty transaction log 2'
    New-EvidenceFile (Join-Path $configRoot 'SOFTWARE')      'SOFTWARE hive body'
    New-EvidenceFile (Join-Path $configRoot 'SOFTWARE.LOG1') 'SOFTWARE dirty transaction log'

    $profileRoot = Join-Path $sourceRoot 'Users\Analyst'
    New-EvidenceFile (Join-Path $profileRoot 'NTUSER.dat')      'NTUSER hive body'
    New-EvidenceFile (Join-Path $profileRoot 'NTUSER.dat.LOG1') 'NTUSER dirty transaction log'
    New-EvidenceFile (Join-Path $profileRoot 'AppData\Roaming\Microsoft\Windows\Recent\usb-device.lnk') 'LNK body'

    $eventLogRoot = Join-Path $sourceRoot 'Windows\System32\winevt\Logs'
    New-EvidenceFile (Join-Path $eventLogRoot 'Microsoft-Windows-Partition%4Diagnostic.evtx') 'partition diagnostic'
    New-EvidenceFile (Join-Path $eventLogRoot 'Security.evtx') 'security log that ParseUSBs does not need'

    # Make the source tree read-only, the way a mounted image behaves.
    Get-ChildItem -LiteralPath $sourceRoot -File -Recurse -Force | ForEach-Object { $_.IsReadOnly = $true }

    # ---- Stand-in parseusbs.exe -------------------------------------------
    # Records the command line it was given and writes one CSV, so argument
    # construction and output-move behaviour are both exercised.
    $toolDirectory = Join-Path $toolsRoot 'ParseUSBs'
    New-Item -ItemType Directory -Path $toolDirectory -Force | Out-Null
    $toolPath = Join-Path $toolDirectory 'parseusbs.exe'
    $callLogPath = Join-Path $tempRoot 'fake-parseusbs-calls.txt'

    $fakeToolSource = @"
using System;
using System.IO;
public static class FakeParseUsbs {
    public static int Main(string[] args) {
        File.AppendAllText(@"$callLogPath", string.Join(" ", args) + Environment.NewLine);
        string outputDirectory = null;
        for (int i = 0; i < args.Length - 1; i++) { if (args[i] == "-d") { outputDirectory = args[i + 1]; } }
        if (outputDirectory != null) {
            Directory.CreateDirectory(outputDirectory);
            File.WriteAllText(Path.Combine(outputDirectory, "usb-devices.csv"), "serial,friendly_name\r\n0123,Test USB\r\n");
        }
        Console.WriteLine("fake parseusbs ran");
        return 0;
    }
}
"@
    Add-Type -TypeDefinition $fakeToolSource -OutputAssembly $toolPath -OutputType ConsoleApplication

    $toolDefinitions = @([pscustomobject]@{
        id = 'parseusbs'
        name = 'parseusbs'
        executablePath = 'ParseUSBs\parseusbs.exe'
        installDirectory = 'ParseUSBs'
        packageType = 'file'
    })

    # ---- Run ---------------------------------------------------------------
    $before = Get-TreeSnapshot -Root $sourceRoot
    $result = Invoke-IbisParseUsbArtifacts -ToolsRoot $toolsRoot -ToolDefinitions $toolDefinitions -SourceRoot $sourceRoot -OutputRoot $outputRoot -Hostname 'TESTHOST'
    $after = Get-TreeSnapshot -Root $sourceRoot

    $summary = Get-Content -LiteralPath $result.JsonPath -Raw | ConvertFrom-Json
    $calls = @(Get-Content -LiteralPath $callLogPath)
    $stagingDirectory = $summary.StagingDirectory

    Write-Host ''
    Write-Host ("Module status : " + $result.Status)
    Write-Host ("Output folder : " + $result.OutputDirectory)
    Write-Host ("Staging folder: " + $stagingDirectory)
    Write-Host ''
    Write-Host 'Command lines the tool actually received:'
    foreach ($call in $calls) { Write-Host ("  " + $call) }
    Write-Host ''
    Write-Host 'Checks:'

    # 1. Source evidence is byte-for-byte unchanged.
    $difference = Compare-Object -ReferenceObject $before -DifferenceObject $after -Property Path, Hash, LastWriteUtc
    Assert-True ($null -eq $difference) 'Source evidence tree is unchanged (hash + last write time)'
    if ($null -ne $difference) { $difference | Format-Table -AutoSize | Out-String | Write-Host }

    # 2. Nothing new was created anywhere under the source root.
    Assert-True ($before.Count -eq $after.Count) ("Source file count unchanged (" + $before.Count + " files)")

    # 3. The tool was never handed a path inside the source root.
    $sourceReferences = @($calls | Where-Object { $_ -like ("*" + $sourceRoot + "*") })
    Assert-True ($sourceReferences.Count -eq 0) 'Tool was never given a path inside the evidence source root'

    # 4. Every tool invocation pointed at the staging area.
    $stagedReferences = @($calls | Where-Object { $_ -like ("*" + $stagingDirectory + "*") })
    Assert-True ($calls.Count -eq 2) ("Both phases ran (" + $calls.Count + " invocations)")
    Assert-True ($stagedReferences.Count -eq $calls.Count) 'Every invocation pointed at the staging area'

    # 5. Phases used the intended argument shapes.
    Assert-True ($calls[0] -like '*-s *' -and $calls[0] -like '*-w *' -and $calls[0] -like '*-u *') 'Registry phase used explicit -s / -w / -u hive arguments'
    Assert-True ($calls[1] -like '-v *') 'Volume phase used -v against the staged volume'

    # 6. Staging actually holds copies, including transaction logs.
    Assert-True (Test-Path -LiteralPath (Join-Path $stagingDirectory 'Windows\System32\config\SYSTEM.LOG1')) 'Staged SYSTEM transaction log present'
    Assert-True (Test-Path -LiteralPath (Join-Path $stagingDirectory 'Users\Analyst\NTUSER.dat.LOG1')) 'Staged NTUSER transaction log present'
    Assert-True (Test-Path -LiteralPath (Join-Path $stagingDirectory 'Users\Analyst\AppData\Roaming\Microsoft\Windows\Recent\usb-device.lnk')) 'Staged user LNK present'
    Assert-True (Test-Path -LiteralPath (Join-Path $stagingDirectory 'Windows\System32\winevt\Logs\Microsoft-Windows-Partition%4Diagnostic.evtx')) 'Staged USB event log present'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $stagingDirectory 'Windows\System32\winevt\Logs\Security.evtx'))) 'Unrelated event log was not staged'

    # 7. Staged copies are writable, so the tool can replay into them.
    $stagedSystem = Get-Item -LiteralPath (Join-Path $stagingDirectory 'Windows\System32\config\SYSTEM') -Force
    Assert-True (-not $stagedSystem.IsReadOnly) 'Staged SYSTEM hive is writable (read-only flag not carried over)'

    # 8. Results were published with distinct per-phase names.
    Assert-True (Test-Path -LiteralPath (Join-Path $result.OutputDirectory 'TESTHOST-ParseUSBs-RegistryHives-usb-devices.csv')) 'Registry-phase CSV published'
    Assert-True (Test-Path -LiteralPath (Join-Path $result.OutputDirectory 'TESTHOST-ParseUSBs-VolumeEnrichment-usb-devices.csv')) 'Volume-phase CSV published'
    Assert-True ($result.Status -eq 'Completed') 'Module reported Completed'

    Write-Host ''
    if ($failures.Count -eq 0) { Write-Host 'ALL CHECKS PASSED' -ForegroundColor Green }
    else { Write-Host ("FAILED CHECKS: " + $failures.Count) -ForegroundColor Red; $host.SetShouldExit(1) }
}
finally {
    if (Test-Path -LiteralPath $tempRoot) {
        Get-ChildItem -LiteralPath $tempRoot -File -Recurse -Force | ForEach-Object { $_.IsReadOnly = $false }
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
