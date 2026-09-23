$ErrorActionPreference = 'Stop'
$pidFile = Join-Path $PSScriptRoot '.local\server.pid'
if (-not (Test-Path -LiteralPath $pidFile)) { Write-Host 'No saved WordAgent process.'; return }
$savedPid = 0
if (-not [int]::TryParse((Get-Content -LiteralPath $pidFile -Raw).Trim(), [ref]$savedPid)) { throw 'Invalid server.pid.' }
$process = Get-Process -Id $savedPid -ErrorAction SilentlyContinue
if ($process) {
  if ($process.Path -ne (Join-Path $PSScriptRoot 'runtime\node.exe')) { throw 'Process identity changed; refusing to stop another application.' }
  Stop-Process -Id $savedPid
}
Remove-Item -LiteralPath $pidFile
Write-Host 'WordAgent stopped. Documents and model settings were preserved.'
