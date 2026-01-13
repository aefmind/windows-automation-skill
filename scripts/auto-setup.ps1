# Auto-Setup Script for Windows Automation Skill
# Downloads and configures EVERYTHING automatically.

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

Write-Host ">>> Iniciando Auto-Setup Windows Automation Skill (system Python) <<<" -ForegroundColor Cyan

# 0. Definir rutas
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$skillRoot = Resolve-Path (Join-Path $scriptDir "..") # skill root
$javaLibDir = Join-Path $skillRoot "src\bridge_java\lib"
$modelsDir = Join-Path $skillRoot "src\bridge_python\models"
$sikuliUrl = "https://launchpad.net/sikuli/sikulix/2.0.5/+download/sikulixapi-2.0.5.jar"
$omniparserUrl = "https://huggingface.co/microsoft/OmniParser/resolve/main/icon_detect/model.onnx"

# 1. Verificar Python (system installed)
Write-Host "[1/4] Verificando Python instalado en el sistema..." -ForegroundColor Yellow

$PythonExe = $null
if (Get-Command py -ErrorAction SilentlyContinue) {
    $PythonExe = "py"
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $PythonExe = "python"
}

if (-not $PythonExe) {
    Write-Host "[ERROR] Python no encontrado. Instala Python 3.11+ y asegfarate de tener 'py' o 'python' en PATH." -ForegroundColor Red
    throw "Python missing"
}

Write-Host "[OK] Python detectado: $PythonExe" -ForegroundColor Green

# 2. Crear venv local (recomendado)
Write-Host "[2/4] Creando entorno virtual (.venv)..." -ForegroundColor Yellow
$venvDir = Join-Path $skillRoot ".venv"
if (-not (Test-Path $venvDir)) {
    if ($PythonExe -eq "py") {
        & py -3 -m venv $venvDir
    } else {
        & python -m venv $venvDir
    }
} else {
    Write-Host "[SKIP] .venv ya existe." -ForegroundColor Gray
}

$VenvPython = Join-Path $venvDir "Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    throw "No se pudo crear .venv (faltfa $VenvPython)"
}

# 3. Instalar dependencias Python
Write-Host "[3/4] Instalando dependencias Python (Flask, pywinauto, etc.)..." -ForegroundColor Yellow
try {
    & $VenvPython -m pip install --upgrade pip --quiet
    & $VenvPython -m pip install flask pywinauto comtypes onnxruntime numpy Pillow --quiet
    Write-Host "[OK] Dependencias Python instaladas." -ForegroundColor Green
} catch {
    Write-Host "[WARN] No se pudieron instalar dependencias Python: $_" -ForegroundColor Yellow
}

# 4. Instalar Vision Layer Dependencies (OmniParser is optional)
Write-Host "[4/4] Vision Layer: dependencias listas (modelo OmniParser opcional)." -ForegroundColor Yellow

# 3. Descargar OmniParser Model
if (-not (Test-Path $modelsDir)) {
    New-Item -ItemType Directory -Force -Path $modelsDir | Out-Null
}
$omniparserPath = Join-Path $modelsDir "omniparser-icon_detect.onnx"
if (-not (Test-Path $omniparserPath)) {
    Write-Host "[3/5] Descargando OmniParser model (~6MB)..." -ForegroundColor Yellow
    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($omniparserUrl, $omniparserPath)
        $fileSize = (Get-Item $omniparserPath).Length / 1MB
        Write-Host "[OK] OmniParser descargado ($([math]::Round($fileSize, 2)) MB)." -ForegroundColor Green
    } catch {
        Write-Host "[WARN] Falló descarga de OmniParser. La Vision Layer estará deshabilitada." -ForegroundColor Yellow
        Write-Host "       Descarga manual: $omniparserUrl" -ForegroundColor Gray
    }
} else {
    Write-Host "[SKIP] OmniParser ya existe." -ForegroundColor Gray
}

# 4. Configurar SikuliX
if (-not (Test-Path $javaLibDir)) {
    New-Item -ItemType Directory -Force -Path $javaLibDir | Out-Null
}
$sikuliJar = Join-Path $javaLibDir "sikulixapi.jar"
if (-not (Test-Path $sikuliJar)) {
    Write-Host "[4/5] Descargando SikuliX API JAR..." -ForegroundColor Yellow
    # SikuliX launchpad redirects can be tricky, using direct link logic or fallback
    # Using a known reliable direct link structure or user prompt if fails.
    # Trying direct download:
    try {
        Invoke-WebRequest -Uri $sikuliUrl -OutFile $sikuliJar
        Write-Host "[OK] SikuliX descargado." -ForegroundColor Green
    } catch {
        Write-Host "[ERROR] Falló descarga automática de SikuliX. Por favor descarga $sikuliUrl manualmente a $javaLibDir" -ForegroundColor Red
    }
} else {
    Write-Host "[SKIP] SikuliX ya existe." -ForegroundColor Gray
}

# 5. Compilar .NET Project
Write-Host "[5/5] Compilando MainAgentService (.NET)..." -ForegroundColor Yellow
$projPath = Join-Path $skillRoot "src\src\MainAgentService\MainAgentService.csproj"
dotnet build $projPath
if ($LASTEXITCODE -eq 0) {
    Write-Host "[OK] Compilación exitosa." -ForegroundColor Green
} else {
    Write-Host "[ERROR] Falló la compilación." -ForegroundColor Red
}

Write-Host "`n>>> Instalación v4.0 COMPLETADA <<<" -ForegroundColor Cyan
Write-Host "Componentes instalados/configurados:" -ForegroundColor White
Write-Host "  - Python del sistema + .venv" -ForegroundColor Gray
Write-Host "  - Dependencias Python (Flask, pywinauto, ONNX Runtime, etc.)" -ForegroundColor Gray
Write-Host "  - OmniParser ONNX model (~6MB, opcional)" -ForegroundColor Gray
Write-Host "  - SikuliX API JAR (si aplica)" -ForegroundColor Gray
Write-Host "  - MainAgentService (.NET 9)" -ForegroundColor Gray
Write-Host "`nAhora ejecuta: .\scripts\start-all.ps1" -ForegroundColor White
