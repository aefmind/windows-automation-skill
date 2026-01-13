# Setup Environment Script for Windows Automation Skill

Write-Host "Iniciando configuración del entorno para Windows Automation Skill..." -ForegroundColor Cyan

# 1. Verificar .NET SDK
if (Get-Command dotnet -ErrorAction SilentlyContinue) {
    Write-Host "[OK] .NET SDK detectado." -ForegroundColor Green
} else {
    Write-Host "[ERROR] .NET SDK no encontrado. Descárgalo de https://dotnet.microsoft.com/" -ForegroundColor Red
}

# 2. Configurar Python y Dependencias
Write-Host "Configurando dependencias de Python..." -ForegroundColor Yellow
try {
    python -m pip install flask pywinauto
    Write-Host "[OK] Dependencias de Python instaladas." -ForegroundColor Green
} catch {
    Write-Host "[WARN] No se pudo instalar dependencias de Python automáticamente. Ejecuta: pip install flask pywinauto" -ForegroundColor Yellow
}

# 3. Configurar Java y SikuliX
Write-Host "Configurando entorno Java para SikuliX..." -ForegroundColor Yellow
$javaBridgeDir = "bridge_java"
$libDir = "$javaBridgeDir/lib"

if (!(Test-Path $libDir)) {
    New-Item -ItemType Directory -Path $libDir
}

# Nota: El usuario debe descargar sikulixapi.jar manualmente debido a restricciones de red
Write-Host "[INFO] Por favor, descarga 'sikulixapi.jar' de https://raiman.github.io/SikuliX1/downloads.html y colócalo en $libDir" -ForegroundColor Blue

# 4. Compilar Servicio Principal
Write-Host "Compilando MainAgentService..." -ForegroundColor Yellow
dotnet build "src/MainAgentService/MainAgentService.csproj"

Write-Host "`nConfiguración completada con advertencias de dependencias externas." -ForegroundColor Cyan
Write-Host "Para iniciar:" -ForegroundColor White
Write-Host "1. Inicia el bridge de Python: python bridge_python/bridge.py"
Write-Host "2. Inicia el bridge de Java: java -cp 'bridge_java/lib/*;bridge_java/src' SikuliBridge"
Write-Host "3. Ejecuta el agente: dotnet run --project src/MainAgentService"
