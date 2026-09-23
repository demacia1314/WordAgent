$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$node = Join-Path $PSScriptRoot 'runtime\node.exe'
$env:PATH = (Join-Path $PSScriptRoot 'runtime') + ';' + $env:PATH
if (-not (Test-Path -LiteralPath $node)) { throw 'Extract the complete ZIP before running Install.cmd.' }
Write-Host 'WordAgent local trial: Microsoft 365 / Office 2021 / Office 2024 on Windows x64.'
Write-Host 'Setup trusts a localhost certificate for this user and sideloads the Word add-in.'
Write-Host 'No model keys are included. Office.js and AI require Internet access.'
$answer = Read-Host 'Continue? [y/N]'
if ($answer -notin @('y', 'Y')) { exit 0 }
& $node 'node_modules\office-addin-dev-certs\lib\cli.js' install --days 365
if ($LASTEXITCODE -ne 0) { throw 'Certificate setup failed.' }
& (Join-Path $PSScriptRoot 'start.ps1')
if (-not $?) { throw 'Server startup failed.' }
& $node 'node_modules\office-addin-debugging\lib\cli.js' start manifest.xml desktop --app word --no-debug --no-live-reload --dev-server-port 3001
if ($LASTEXITCODE -ne 0) { throw 'Word sideload failed. Check Office installation and organization policy.' }
Write-Host 'Ready. Configure your own model in the sidebar settings.'
