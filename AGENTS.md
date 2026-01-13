# AGENTS.md

This repository is designed for collaborative development by humans and AI coding agents.

## Project Overview

**Windows Desktop Automation Skill** - A .NET 9 Windows application that provides GUI automation capabilities via FlaUI and Win32 APIs, with a Python bridge for vision/OCR functionality.

## Repository Structure

```
windows-desktop-automation/
├── src/
│   ├── src/MainAgentService/     # .NET 9 orchestrator (FlaUI + Win32)
│   └── bridge_python/            # Python Flask service (vision + OCR)
├── scripts/
│   ├── send-command.cjs          # Primary command interface
│   ├── get-command-help.cjs      # Command documentation helper
│   ├── test-commands.cjs         # Test runner
│   ├── start-all.ps1             # Start all services
│   ├── stop-all.ps1              # Stop all services
│   └── auto-setup.ps1            # One-time environment setup
├── docs/                         # Additional documentation
├── SKILL.md                      # Command reference (for AI agents)
├── README.md                     # User-facing documentation
├── DEVELOPMENT.md                # Developer setup guide
└── AGENTS.md                     # This file
```

## Key Files

| File | Lines | Purpose |
|------|-------|---------|
| `src/src/MainAgentService/Program.cs` | ~6,400 | Main orchestrator (all command handlers) |
| `src/bridge_python/server.py` | ~800 | Python Flask server (vision/OCR) |
| `SKILL.md` | ~350 | Command reference for AI consumption |

## Scope & Boundaries

### In Scope

- **Orchestrator**: `src/src/MainAgentService/` (.NET + FlaUI + Win32)
- **Python Bridge**: `src/bridge_python/` (Flask + pywinauto + WinRT OCR)
- **Scripts**: `scripts/*.ps1` and `scripts/*.cjs`
- **Documentation**: `SKILL.md`, `README.md`, `DEVELOPMENT.md`

### Out of Scope / Safety Boundaries

- **Do NOT** automate destructive OS actions without explicit user intent
- **Do NOT** run UI automation against sensitive apps (password managers, banking) by default
- **Prefer** idempotent scripts (safe to re-run)
- **Avoid** modifying system settings or registry

## Development Guidelines

### Prerequisites

- Windows 10/11
- .NET SDK 9+
- Python 3.11+ (available as `py` or `python`)
- Node.js (for scripts)

### Setup

```powershell
# One-time setup
powershell -File scripts/auto-setup.ps1

# Start services
powershell -File scripts/start-all.ps1

# Stop services
powershell -File scripts/stop-all.ps1
```

### Build & Test

```powershell
# Build .NET orchestrator
dotnet build src/src/MainAgentService/MainAgentService.csproj

# Run tests
node scripts/test-commands.cjs --list      # List all tests
node scripts/test-commands.cjs             # Run all tests
node scripts/test-commands.cjs --category mouse  # Run category
```

### Code Style

- **.NET**: Follow standard C# conventions
- **Python**: Use `ruff` for formatting/linting
- **Scripts**: Node.js CommonJS (`.cjs` extension)

## Definition of Done

A task is complete when:

1. `scripts/start-all.ps1` works with **system Python** (`py` or `python`)
2. No embedded Python runtime is required or referenced
3. Documentation is updated if behavior changes
4. `git status` shows only intentional changes
5. Build passes: `dotnet build` with 0 errors, 0 warnings

## Git Hygiene

### .gitignore Coverage

The following are excluded from version control:

```
.venv/              # Python virtual environment
__pycache__/        # Python bytecode
*.pyc, *.pyo        # Python compiled files
node_modules/       # Node dependencies
bin/, obj/          # .NET build output
*.log               # Log files
src/python-embedded/  # Embedded Python (not used)
```

### Commit Guidelines

- Write clear commit messages describing "why" not "what"
- Keep commits atomic (one logical change per commit)
- Verify clean build before committing

## Agent-Specific Instructions

### When Exploring the Codebase

1. Start with `SKILL.md` for command reference
2. Use `Program.cs` for implementation details (~6,400 lines, well-structured)
3. Check `scripts/send-command.cjs` for the command interface

### When Making Changes

1. **Build after every change**: `dotnet build src/src/MainAgentService/`
2. **Make small, surgical edits** - avoid large refactors
3. **Test changes**: Use `scripts/send-command.cjs` or `test-commands.cjs`
4. **Update docs** if adding/modifying commands

### When Adding Commands

1. Add handler method in `Program.cs`
2. Add case in the main switch statement
3. Update `SKILL.md` with command documentation
4. Update `docs/command-schemas.json` with the new command schema
5. Add test case in `scripts/test-commands.cjs` (optional)

### Keeping Documentation in Sync

When modifying commands, ensure these files stay synchronized:

| File | What to Update |
|------|----------------|
| `SKILL.md` | Command reference tables, version number |
| `docs/command-schemas.json` | Full JSON schema, version number, categories |
| `README.md` | User-facing docs if behavior changes |

**Version Numbers**: When adding new commands or significant features:
1. Update `version` in `SKILL.md` front matter
2. Update `version` in `docs/command-schemas.json`
3. Update `description` in `docs/command-schemas.json` to mention new features

### Common Patterns in Program.cs

```csharp
// Handler method pattern
private void HandleMyCommand(JsonElement root)
{
    // 1. Parse required parameters
    if (!root.TryGetProperty("param", out var paramEl))
    {
        WriteMissingParam("param");
        return;
    }
    
    // 2. Execute action
    // ...
    
    // 3. Write response
    WriteSuccess(new { result = "value" });
}
```

## Review Checklist

Before submitting changes, verify:

- [ ] `dotnet build` passes with 0 errors, 0 warnings
- [ ] `git status` shows only intentional files
- [ ] Entry points work: `scripts/start-all.ps1`, `scripts/auto-setup.ps1`
- [ ] Documentation updated if needed (`SKILL.md`, `README.md`)
- [ ] No secrets or credentials committed

## Troubleshooting

### Python not found

```powershell
# Verify Python is installed
py -3 --version
# or
python --version
```

### Port 5001 already in use

```powershell
scripts/stop-all.ps1
# Then retry
scripts/start-all.ps1
```

### Build errors

```powershell
# Clean and rebuild
dotnet clean src/src/MainAgentService/
dotnet build src/src/MainAgentService/
```
