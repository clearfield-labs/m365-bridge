# M365 Bridge installer for Windows. Nothing to install first, no admin rights:
# it downloads the code as a zip (no Git), and a private copy of Node.js if this computer
# has none (no winget, no system install). Run it again to update.
#   irm https://raw.githubusercontent.com/clearfield-labs/m365-bridge/main/install.ps1 | iex
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # makes downloads much faster in Windows PowerShell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Repo    = if ($env:M365_BRIDGE_REPO_SLUG) { $env:M365_BRIDGE_REPO_SLUG } else { 'clearfield-labs/m365-bridge' }
$Ref     = if ($env:M365_BRIDGE_REF) { $env:M365_BRIDGE_REF } else { 'main' }
$Dest    = if ($env:M365_BRIDGE_DIR) { $env:M365_BRIDGE_DIR } else { Join-Path $HOME '.m365-bridge-app' }
$NodeDir = if ($env:M365_BRIDGE_NODE_DIR) { $env:M365_BRIDGE_NODE_DIR } else { Join-Path $HOME '.m365-bridge-node' }
$Tmp     = Join-Path ([IO.Path]::GetTempPath()) ("m365b-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $Tmp | Out-Null

function Say($m) { Write-Host $m -ForegroundColor Cyan }

try {
  Say 'Installing M365 Bridge...'

  # 1. Node.js: use this computer's if it is 18 or newer, otherwise a private copy.
  $Node = $null
  $sys = Get-Command node -ErrorAction SilentlyContinue
  if ($sys -and -not $env:M365_FORCE_PRIVATE_NODE) {
    $major = [int]((& $sys.Source -p "process.versions.node.split('.')[0]") 2>$null)
    if ($major -ge 18) { $Node = $sys.Source }
  }
  if (-not $Node) {
    $private = Join-Path $NodeDir 'node.exe'
    if (-not (Test-Path $private) -or $env:M365_FORCE_PRIVATE_NODE) {
      Say 'Downloading Node.js (only for M365 Bridge)...'
      $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
      $sums = (Invoke-WebRequest -UseBasicParsing 'https://nodejs.org/dist/latest-v22.x/SHASUMS256.txt').Content -split "`n"
      $line = $sums | Where-Object { $_ -match "node-v[0-9.]+-win-$arch\.zip$" } | Select-Object -First 1
      if (-not $line) { throw "Could not find a Node.js download for win-$arch." }
      $hash, $file = ($line.Trim() -split '\s+')
      $zip = Join-Path $Tmp $file
      Invoke-WebRequest -UseBasicParsing "https://nodejs.org/dist/latest-v22.x/$file" -OutFile $zip
      if ((Get-FileHash $zip -Algorithm SHA256).Hash.ToLower() -ne $hash.ToLower()) { throw 'The Node.js download did not match its checksum. Try again.' }
      Expand-Archive $zip -DestinationPath $Tmp -Force
      if (Test-Path $NodeDir) { Remove-Item $NodeDir -Recurse -Force }
      Move-Item (Join-Path $Tmp ($file -replace '\.zip$', '')) $NodeDir
    }
    $Node = $private
  }
  $NodeHome = Split-Path $Node
  $env:Path = "$NodeHome;$env:Path"

  # 2. The code, as a zip from GitHub (no Git needed).
  Say 'Downloading M365 Bridge...'
  $appZip = Join-Path $Tmp 'app.zip'
  Invoke-WebRequest -UseBasicParsing "https://codeload.github.com/$Repo/zip/$Ref" -OutFile $appZip
  Expand-Archive $appZip -DestinationPath (Join-Path $Tmp 'app') -Force
  $src = Get-ChildItem (Join-Path $Tmp 'app') | Select-Object -First 1
  # An older copy may be running; ask it to stop so its files can be replaced.
  if (Test-Path (Join-Path $Dest 'bin\m365-bridge.js')) { & $Node (Join-Path $Dest 'bin\m365-bridge.js') stop 2>$null | Out-Null; Start-Sleep -Seconds 1 }
  if (Test-Path $Dest) { Remove-Item $Dest -Recurse -Force }
  Move-Item $src.FullName $Dest

  # 3. Dependencies, then set up: background service, Claude, Start menu shortcut.
  Set-Location $Dest
  $npm = Join-Path $NodeHome 'node_modules\npm\bin\npm-cli.js'
  if (Test-Path $npm) { & $Node $npm install --omit=dev --no-fund --no-audit --loglevel=error } else { & npm.cmd install --omit=dev --no-fund --no-audit --loglevel=error }
  if ($LASTEXITCODE -ne 0) { throw 'Installing dependencies failed.' }
  if ($env:M365_INSTALL_SKIP_RUN) { & $Node bin\m365-bridge.js status | Out-Null; Say 'Installed (setup skipped).'; return }
  & $Node bin\m365-bridge.js install
} catch {
  Write-Host "`n$($_.Exception.Message)" -ForegroundColor Red
  Write-Host 'Nothing was changed in Microsoft 365. Try again, or send this message to whoever set up M365 Bridge.'
} finally {
  Remove-Item $Tmp -Recurse -Force -ErrorAction SilentlyContinue
}
