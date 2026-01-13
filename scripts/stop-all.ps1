Write-Host "Stopping windows-desktop-automation processes..." -ForegroundColor Cyan

$procNames = @("python", "python.exe", "dotnet", "dotnet.exe")
foreach ($p in Get-Process | Where-Object { $procNames -contains $_.Name }) {
    try {
        if ($p.Path -like "*python*" -or $p.Path -like "*dotnet*") {
            Write-Host "Killing $($p.Name) PID=$($p.Id)" -ForegroundColor Yellow
            $p.Kill()
        }
    } catch {}
}

Write-Host "Done." -ForegroundColor Green
