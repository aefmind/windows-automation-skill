# Agents

This repository is designed to be worked on by both humans and coding agents.

## Scope
- **Orchestrator**: `src/src/MainAgentService` (.NET + FlaUI)
- **Python bridge**: `src/bridge_python` (Flask service for legacy + vision + OCR)
- **Scripts**: `scripts/*.ps1` and `scripts/*.cjs`

## Safety / Boundaries
- Do not automate destructive OS actions without explicit user intent.
- Avoid running UI automation against sensitive apps by default.
- Prefer idempotent scripts (safe to re-run).

## Definition of Done
- `scripts/start-all.ps1` works with **system Python** (`py` or `python`) and does not require `src/python-embedded`.
- Documentation clearly explains setup using `.venv`.
- No embedded Python runtime is required or referenced.

## Review Checklist
- Entry points updated: `scripts/start-all.ps1`, `scripts/auto-setup.ps1`.
- `.gitignore` prevents committing `.venv/`, `__pycache__/`, `*.pyc`, `bin/`, `obj/`.
- Repo starts clean: `git status` shows only intentional files.
