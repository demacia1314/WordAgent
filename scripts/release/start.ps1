$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$node = Join-Path $PSScriptRoot 'runtime\node.exe'
$env:PATH = (Join-Path $PSScriptRoot 'runtime') + ';' + $env:PATH
if (-not (Test-Path -LiteralPath $node)) { throw 'Bundled runtime missing. Extract the whole ZIP.' }
New-Item -ItemType Directory -Path (Join-Path $PSScriptRoot '.local') -Force | Out-Null
$pidFile = Join-Path $PSScriptRoot '.local\server.pid'
if (Test-Path -LiteralPath $pidFile) {
  $savedPid = 0
  if ([int]::TryParse((Get-Content -LiteralPath $pidFile -Raw).Trim(), [ref]$savedPid)) {
    $existing = Get-Process -Id $savedPid -ErrorAction SilentlyContinue
    if ($existing -and $existing.Path -eq $node) {
      & $node 'health.mjs'
      if ($LASTEXITCODE -eq 0) { Write-Host 'WordAgent is already running.'; return }
      throw 'The saved process is not healthy. Run Stop.cmd, then Start.cmd.'
    }
  }
}
if (Get-NetTCPConnection -LocalPort 3001 -State Listen -ErrorAction SilentlyContinue) {
  throw 'Port 3001 is occupied. Close the other server first; it will not be stopped automatically.'
}
$env:NODE_ENV = 'production'
$env:PORT = '3001'
$env:WORDAGENT_DATA_DIR = Join-Path $PSScriptRoot '.local'
$env:HTTPS_KEY = Join-Path $env:USERPROFILE '.office-addin-dev-certs\localhost.key'
$env:HTTPS_CERT = Join-Path $env:USERPROFILE '.office-addin-dev-certs\localhost.crt'
$process = Start-Process -FilePath $node -ArgumentList 'dist-server/index.js' -WorkingDirectory $PSScriptRoot -WindowStyle Hidden -RedirectStandardOutput '.local\server.log' -RedirectStandardError '.local\server-error.log' -PassThru
$process.Id | Set-Content -LiteralPath $pidFile
& $node 'health.mjs'
if ($LASTEXITCODE -ne 0) {
  if (-not $process.HasExited) { $process.Kill() }
  Remove-Item -LiteralPath $pidFile -ErrorAction SilentlyContinue
  throw 'Server did not become ready. See .local/server-error.log.'
}
Write-Host 'Running at https://localhost:3001. Use Stop.cmd before moving the folder.'
