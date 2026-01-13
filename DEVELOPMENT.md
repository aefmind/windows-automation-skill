# Development

## Prerequisites
- Windows 10/11
- .NET SDK 9+
- Python 3.11+ (recommended) available as `py` or `python`
- Node.js (for `scripts/*.cjs` helpers/tests)

## Setup (recommended)
```powershell
# from repo root
powershell -File scripts/auto-setup.ps1
```

This will:
- create `.venv/`
- install Python dependencies into `.venv/`
- (optionally) download OmniParser model via `scripts/download-omniparser.ps1`
- build the .NET orchestrator

## Manual Python setup
```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install flask pywinauto comtypes onnxruntime numpy Pillow
```

## Run
```powershell
powershell -File scripts/start-all.ps1
```

## Stop
```powershell
powershell -File scripts/stop-all.ps1
```

## Build (.NET)
```powershell
dotnet build src\src\MainAgentService\MainAgentService.csproj
```

## Tests (Node)
```powershell
node scripts\test-commands.cjs --list
node scripts\test-commands.cjs --category errorHandling
node scripts\test-commands.cjs
```

## Troubleshooting
### Python not found
- Install Python 3.11+ from https://www.python.org/downloads/
- On Windows, ensure the **Python Launcher** is installed so `py -3` works.

### Port 5001 already in use
Run `scripts/stop-all.ps1` and retry.
