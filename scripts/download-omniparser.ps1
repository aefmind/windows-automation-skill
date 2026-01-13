# Download OmniParser Model Script
# Downloads the OmniParser icon detection ONNX model for the Vision Layer

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

Write-Host ">>> OmniParser Model Download <<<" -ForegroundColor Cyan

# Define paths
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$skillRoot = Resolve-Path (Join-Path $scriptDir "..")
$modelsDir = Join-Path $skillRoot "src\bridge_python\models"
$modelPath = Join-Path $modelsDir "omniparser-icon_detect.onnx"

# Model URL (Microsoft OmniParser from HuggingFace)
$modelUrl = "https://huggingface.co/microsoft/OmniParser/resolve/main/icon_detect/model.onnx"

# Create models directory if needed
if (-not (Test-Path $modelsDir)) {
    Write-Host "[INFO] Creating models directory..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Force -Path $modelsDir | Out-Null
}

# Download model
if (-not (Test-Path $modelPath)) {
    Write-Host "[1/1] Downloading OmniParser model (~6MB)..." -ForegroundColor Yellow
    Write-Host "      URL: $modelUrl" -ForegroundColor Gray
    Write-Host "      Destination: $modelPath" -ForegroundColor Gray
    
    try {
        # Use Invoke-WebRequest with longer timeout for HuggingFace
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($modelUrl, $modelPath)
        
        # Verify file size (should be ~6MB)
        $fileSize = (Get-Item $modelPath).Length / 1MB
        if ($fileSize -lt 1) {
            throw "Downloaded file is too small ($fileSize MB). Expected ~6MB."
        }
        
        Write-Host "[OK] OmniParser model downloaded successfully ($([math]::Round($fileSize, 2)) MB)" -ForegroundColor Green
    } catch {
        Write-Host "[ERROR] Failed to download OmniParser model: $_" -ForegroundColor Red
        Write-Host "[INFO] Please download manually from:" -ForegroundColor Yellow
        Write-Host "       $modelUrl" -ForegroundColor White
        Write-Host "       Save to: $modelPath" -ForegroundColor White
        exit 1
    }
} else {
    $fileSize = (Get-Item $modelPath).Length / 1MB
    Write-Host "[SKIP] OmniParser model already exists ($([math]::Round($fileSize, 2)) MB)" -ForegroundColor Gray
}

Write-Host "`n>>> OmniParser Download Complete <<<" -ForegroundColor Cyan
