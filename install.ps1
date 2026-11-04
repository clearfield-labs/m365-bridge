# M365 Bridge installer for Windows (PowerShell).
# Usage: irm <repo>/install.ps1 | iex     or     .\install.ps1 from a clone
$ErrorActionPreference = 'Stop'
$RepoUrl = if ($env:M365_BRIDGE_REPO) { $env:M365_BRIDGE_REPO } else { 'https://github.com/irfad7/m365-bridge.git' }
$Dest = if ($env:M365_BRIDGE_DIR) { $env:M365_BRIDGE_DIR } else { Join-Path $HOME '.m365-bridge-app' }

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  Write-Host 'Node.js is not installed. Installing the LTS version with winget...'
  winget install --id OpenJS.NodeJS.LTS -e --accept-source-agreements --accept-package-agreements
  $env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
}

$here = if ($PSScriptRoot) { $PSScriptRoot } else { '' }
if ($here -and (Test-Path (Join-Path $here 'package.json')) -and (Select-String -Path (Join-Path $here 'package.json') -Pattern '"m365-bridge"' -Quiet)) {
  $Dest = $here
} elseif (Test-Path (Join-Path $Dest '.git')) {
  git -C $Dest pull --ff-only
} else {
  if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    winget install --id Git.Git -e --accept-source-agreements --accept-package-agreements
    $env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
  }
  git clone --depth 1 $RepoUrl $Dest
}

Set-Location $Dest
npm install --omit=dev --no-fund --no-audit
node bin/m365-bridge.js install
