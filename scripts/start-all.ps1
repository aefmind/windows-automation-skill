param(
    [string]$PythonDir = ""
)

$ErrorActionPreference = "Stop"

Write-Host "[windows-desktop-automation] Starting bridges and agent..." -ForegroundColor Cyan

# Roots
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$agentProj = Join-Path $root "src\src\MainAgentService\MainAgentService.csproj"
$pyBridge = Join-Path $root "src\bridge_python\bridge.py"

# Resolve Python from system install (preferred)
function Resolve-Python {
    # Prefer explicit -PythonDir if user passes one
    param([string]$Dir)

    if ($Dir -and (Test-Path (Join-Path $Dir "python.exe"))) {
        return (Join-Path $Dir "python.exe")
    }

    # Prefer Windows 'py' launcher if available (stable + versionable)
    if (Get-Command py -ErrorAction SilentlyContinue) {
        return "py"
    }

    return "python"
}

$PythonExe = Resolve-Python -Dir $PythonDir

Write-Host "Using Python: $PythonExe" -ForegroundColor Yellow

# If we're using the Windows launcher, target Python 3 explicitly
$PythonArgsPrefix = @()
if ($PythonExe -eq "py") { $PythonArgsPrefix = @("-3") }

# Start Python bridge (includes pywinauto, vision, WinRT OCR, context manager)
Write-Host "[python] launching bridge on :5001 (pywinauto + vision + WinRT OCR)" -ForegroundColor Yellow
$pyArgs = $PythonArgsPrefix + @("`"$pyBridge`"")
$py = Start-Process -FilePath $PythonExe -ArgumentList $pyArgs -WorkingDirectory (Split-Path $pyBridge) -PassThru -WindowStyle Hidden

# Start MainAgentService
Write-Host "[dotnet] launching MainAgentService (FlaUI)" -ForegroundColor Yellow
$agent = Start-Process -FilePath "dotnet" -ArgumentList "run --project `"$agentProj`"" -WorkingDirectory (Split-Path $agentProj) -PassThru -WindowStyle Hidden

Write-Host "All processes started." -ForegroundColor Green
Write-Host "PIDs => python: $($py.Id) | dotnet: $($agent.Id)" -ForegroundColor Green
Write-Host "Use stop-all.ps1 to terminate them."
