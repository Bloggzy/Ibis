[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Import-Module (Join-Path $projectRoot 'modules\Ibis.Core.psm1') -Force
Import-Module (Join-Path $projectRoot 'modules\Ibis.Gui.psm1') -Force

Show-IbisGui -ProjectRoot $projectRoot
