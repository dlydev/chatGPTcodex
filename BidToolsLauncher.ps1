<#
.SYNOPSIS
  Clickable launcher that ensures Python + dependencies, then runs BidTools.py.
.NOTES
  - Installs Python per-user (no admin) when missing.
  - Installs openpyxl into the user site-packages.
#>

$ErrorActionPreference = "Stop"

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BidToolsPath = Join-Path $ScriptRoot "BidTools.py"

function Get-PythonCommand {
  $py = Get-Command "python" -ErrorAction SilentlyContinue
  if ($null -ne $py) { return $py.Source }
  $pyLauncher = Get-Command "py" -ErrorAction SilentlyContinue
  if ($null -ne $pyLauncher) { return $pyLauncher.Source }
  return $null
}

function Install-PythonUserScope {
  $pythonCmd = Get-PythonCommand
  if ($null -ne $pythonCmd) { return $pythonCmd }

  $winget = Get-Command "winget" -ErrorAction SilentlyContinue
  if ($null -ne $winget) {
    Write-Host "Python not found. Installing via winget (user scope)..." -ForegroundColor Yellow
    & $winget.Source install -e --id Python.Python.3.11 --scope user --silent
    Start-Sleep -Seconds 2
    $pythonCmd = Get-PythonCommand
    if ($null -ne $pythonCmd) { return $pythonCmd }
  }

  Write-Host "Python not found. Downloading per-user installer..." -ForegroundColor Yellow
  $installerUrl = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
  $installerPath = Join-Path $env:TEMP "python-3.11.9-amd64.exe"
  Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
  & $installerPath /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1
  Start-Sleep -Seconds 2
  $pythonCmd = Get-PythonCommand
  if ($null -eq $pythonCmd) {
    throw "Python installation did not complete. Please restart the launcher."
  }
  return $pythonCmd
}

function Ensure-Dependency([string]$pythonCmd, [string]$package) {
  Write-Host "Ensuring dependency: $package" -ForegroundColor Cyan
  & $pythonCmd -m pip install --user --upgrade $package
}

if (!(Test-Path $BidToolsPath)) {
  throw "BidTools.py not found at: $BidToolsPath"
}

$pythonCmd = Install-PythonUserScope
Ensure-Dependency -pythonCmd $pythonCmd -package "pip"
Ensure-Dependency -pythonCmd $pythonCmd -package "openpyxl"

Write-Host "Launching BidTools.py..." -ForegroundColor Green
& $pythonCmd $BidToolsPath
