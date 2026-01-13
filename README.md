# Windows Desktop Automation Skill

This project provides a robust, hybrid automation engine for Windows, designed to be driven by an AI agent or external scripts. It uses a 2-engine architecture to ensure high compatibility across modern (UWP/WPF), legacy (Win32), and visual-only applications (Flutter, Electron, Canvas).

## Version 4.3.2 - Agent Vision Enhancements

**New in v4.3.2:**
- **Screenshot Caching** - LRU cache with TTL expiration for rapid successive operations
- **Agent Vision Commands** - 6 new JSON commands for AI agent vision workflows
- **Smart Fallback with Screenshot** - `agent_fallback` param returns screenshot when element not found

**New in v4.3:**
- **Simplified to 2-Engine Architecture** - Removed SikuliX/Java Bridge dependency
- **All OCR via WinRT** - 3.2x faster than previous PowerShell-based OCR
- **Reduced footprint** - No Java runtime required, ~50MB smaller

**v4.0 Features:**
- **Vision Layer** - OmniParser ONNX + Windows OCR for apps where UI Automation fails
- **Token-Efficient Commands** - 80% reduction in AI agent token consumption
- **Smart Fallback** - Automatic FlaUI → Vision fallback chain
- **Context Manager** - Intelligent caching with 75%+ hit rates

## Architecture

```
                    +---------------------------+
                    |      AI Agent / User      |
                    +-------------+-------------+
                                  | JSON Commands (STDIN)
                                  v
                    +---------------------------+
                    |    MainAgentService       |
                    |    (.NET 9 + FlaUI)       |
                    |    Primary Orchestrator   |
                    +--+--------------------+---+
                       |                    |
              Layer 1  |           Layer 2  |
              (FlaUI)  |                    |
                       v                    v
                 +----------+        +------------------+
                 |  Win32   |        |  Python Bridge   |
                 |   Apps   |        |  (localhost:5001)|
                 +----------+        +------------------+
                                            |
                              +------+------+------+
                              |      |      |      |
                          pywinauto Vision Context WinRT
                                     |     Manager  OCR
                                OmniParser
                                  ONNX
```

| Component | Technology | Role | Port |
|-----------|------------|------|------|
| **Orchestrator** | .NET 9 + FlaUI | Main entry point. Handles UIA3 (Modern Windows). | N/A (STDIN) |
| **Python Bridge** | Python 3.11 + pywinauto | Legacy Win32 + Vision Layer + WinRT OCR + Context Manager. | `localhost:5001` |

### Fallback Logic
When a command (e.g., `smart_click`) is received:
1.  **Try FlaUI**: Attempts to find the element via UI Automation. Fast and reliable.
2.  **Try pywinauto**: If FlaUI fails, calls the Python bridge. Good for older apps.
3.  **Try Vision**: If both fail, uses OmniParser + WinRT OCR to find element visually.

## Installation & Setup

### Prerequisites
*   Windows 10/11
*   .NET SDK 9.0+
*   Node.js (for the `send-command.cjs` helper and test suite)

### One-Time Setup
This project uses **system-installed Python** (recommended via a local `.venv`). The auto-setup script will create `.venv`, install Python deps, optionally download the OmniParser model, and compile the C# agent.

```powershell
powershell -File scripts/auto-setup.ps1
```

### Vision Layer Setup

The vision layer requires ONNX Runtime and the OmniParser model. `auto-setup.ps1` installs the Python deps into `.venv`; you can also set them up manually:

```powershell
# Create venv (recommended)
py -3 -m venv .venv

# Install deps
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install onnxruntime numpy Pillow

# Download OmniParser model (6MB)
powershell -File scripts/download-omniparser.ps1

# Or download manually:
# URL: https://huggingface.co/microsoft/OmniParser/resolve/main/icon_detect/model.onnx
# Save to: src/bridge_python/models/omniparser-icon_detect.onnx
```


**Verify Vision Layer:**
```bash
# Start Python bridge
powershell -File scripts/start-all.ps1

# Test vision endpoints
curl http://127.0.0.1:5001/health
curl http://127.0.0.1:5001/vision/detect -X POST -H "Content-Type: application/json" -d "{}"
```

## Usage

### 1. Start Background Services
Before sending commands, the bridges must be running.

```powershell
powershell -File scripts/start-all.ps1
```
*This starts the .NET Agent and the Python bridge in background mode.*

### 2. Send Commands
Use the helper script to talk to the running agent.

```bash
node scripts/send-command.cjs '{"action":"explore"}'
```

### 3. Stop Services
Always stop the services when finished to release file locks and ports.

```powershell
powershell -File scripts/stop-all.ps1
```

## Testing

### Test Suite
A comprehensive test suite is included to verify all 25 commands work correctly.

```bash
# Run all tests
node scripts/test-commands.cjs

# Run specific category
node scripts/test-commands.cjs --category discovery
node scripts/test-commands.cjs --category windowManagement
node scripts/test-commands.cjs --category keyboard

# List available tests and categories
node scripts/test-commands.cjs --list
```

### Test Categories

| Category | Tests | Description |
|----------|-------|-------------|
| `discovery` | 5 | explore, explore_window, find_element, get_window_info, wait_for_element |
| `windowManagement` | 10 | focus, close, wait, move, resize, minimize, maximize, restore, move_to_monitor |
| `appProcess` | 3 | launch_app, list_processes, kill_process |
| `mouse` | 4 | click_at, mouse_move, drag_and_drop, scroll |
| `keyboard` | 3 | type, hotkey, key_press |
| `clipboard` | 2 | get_clipboard, set_clipboard |
| `textAndScreenshot` | 3 | read_text, screenshot, screenshot_monitor |
| `multiMonitor` | 3 | list_monitors, screenshot_monitor, move_to_monitor |
| `fileDialog` | 5 | detect, set_path, set_filename, confirm, cancel |
| `batch` | 2 | batch execution, batch with errors |
| `system` | 1 | health check |
| `errorHandling` | 3 | missing_param, window_not_found, unknown_action |

### Running Tests After Changes
After making code changes, always run the test suite:

```bash
# Build the agent
cd src/src/MainAgentService && dotnet build

# Run error handling tests (quick verification)
node scripts/test-commands.cjs --category errorHandling

# Run full suite
node scripts/test-commands.cjs
```

## Available Commands (v4.0)

### Discovery
| Command | Description |
|---------|-------------|
| `explore` | List all visible windows on desktop |
| `explore_window` | List UI elements inside a specific window |
| `find_element` | Find specific element and get coordinates |
| `get_window_info` | Get detailed window information |
| `wait_for_element` | Wait for UI element to appear |
| `wait_for_state` | Wait for element state change |

### Token-Efficient Discovery (v4.0)
| Command | Description | Token Cost |
|---------|-------------|------------|
| `element_exists` | Check if element exists (boolean only) | ~30 |
| `get_interactive_elements` | Get only clickable/typeable elements | ~500 |
| `get_window_summary` | Structured window overview | ~300 |
| `get_element_brief` | Minimal element info | ~100 |
| `fuzzy_find_element` | Find element with fuzzy matching | ~200 |
| `describe_element` | Vision-based description | ~100 |

### Smart Fallback Commands (v4.0)
| Command | Description |
|---------|-------------|
| `smart_click` | Click with automatic FlaUI → Vision fallback. Supports `agent_fallback` param (v4.3) |
| `smart_type` | Type with automatic FlaUI → Vision fallback. Supports `agent_fallback` param (v4.3) |
| `vision_click` | Click element by OCR text (vision only) |

### Agent Vision Commands (v4.3)
| Command | Description |
|---------|-------------|
| `vision_screenshot` | Get optimized base64 JPEG screenshot for AI vision |
| `vision_screenshot_region` | Get cropped region screenshot (token-efficient) |
| `vision_config` | Get/set vision mode (local/agent/auto) |
| `vision_analyze_smart` | Smart analysis respecting current vision mode |
| `vision_screenshot_cache_stats` | Get screenshot cache statistics |
| `vision_screenshot_cache_clear` | Clear screenshot cache |
| `vision_stream` | Get WebSocket URL for real-time screenshot streaming |

### Window Management
| Command | Description |
|---------|-------------|
| `focus_window` | Bring window to foreground (required before interactions) |
| `close_window` | Close a window gracefully |
| `wait_for_window` | Wait for window to appear (with timeout) |
| `move_window` | Move window to X,Y coordinates |
| `resize_window` | Resize window to width x height |
| `minimize_window` | Minimize a window |
| `maximize_window` | Maximize a window |
| `restore_window` | Restore window from minimized/maximized |
| `move_to_monitor` | Move window to specific monitor (NEW) |

### Application & Process Control
| Command | Description |
|---------|-------------|
| `launch_app` | Start an application |
| `list_processes` | List running processes |
| `kill_process` | Kill a process by PID or name |

### Mouse Actions
| Command | Description |
|---------|-------------|
| `click` | Click on element |
| `double_click` | Double-click on element |
| `right_click` | Right-click on element |
| `long_press` | Press and hold on element |
| `click_at` | Click at X,Y coordinates |
| `mouse_move` | Move cursor to position |
| `drag_and_drop` | Drag from one element/position to another |
| `scroll` | Scroll mouse wheel up/down |
| `mouse_path` | Move mouse along waypoints (NEW v3.3) |
| `mouse_bezier` | Move mouse along bezier curve (NEW v3.3) |
| `draw` | Hold, move along path, release (NEW v3.3) |
| `draw_bezier` | Hold, move along bezier curve, release (NEW) |
| `mouse_down` | Press and hold mouse button (NEW v3.3) |
| `mouse_up` | Release mouse button (NEW v3.3) |
| `click_relative` | Click relative to element anchor (NEW v3.3) |

### Keyboard Actions
| Command | Description |
|---------|-------------|
| `type` | Type text into focused element |
| `type_here` | Type text at current cursor position (NEW v3.3) |
| `hotkey` | Send key combinations (Ctrl+C, Alt+F4, etc.) |
| `key_press` | Send single key press |
| `key_down` | Press and hold key (NEW v3.3) |
| `key_up` | Release key (NEW v3.3) |

### Clipboard
| Command | Description |
|---------|-------------|
| `get_clipboard` | Read clipboard text |
| `set_clipboard` | Write text to clipboard |

### Text / OCR
| Command | Description |
|---------|-------------|
| `read_text` | Read text from UI element or via OCR |
| `ocr_region` | OCR on explicit screen region (NEW v3.3) |

### Verification / Screenshot
| Command | Description |
|---------|-------------|
| `screenshot` | Capture screen image |
| `screenshot_monitor` | Capture specific monitor (NEW) |

### Multi-Monitor Support (NEW in v3.2)
| Command | Description |
|---------|-------------|
| `list_monitors` | Enumerate all connected displays |
| `screenshot_monitor` | Capture specific monitor by index |
| `move_to_monitor` | Move window to specific display |

### File Dialog Automation (NEW in v3.2)
| Command | Description |
|---------|-------------|
| `file_dialog` | Automate Open/Save/Browse dialogs (detect, set_path, set_filename, confirm, cancel) |

### Batch Operations (NEW in v3.2)
| Command | Description |
|---------|-------------|
| `batch` | Execute multiple commands sequentially |

### System (NEW in v3.2)
| Command | Description |
|---------|-------------|
| `health` | System diagnostics and bridge status |

## Example Commands

```bash
# Discover windows
node scripts/send-command.cjs '{"action":"explore"}'

# Focus and explore a window
node scripts/send-command.cjs '{"action":"focus_window", "selector":"Notepad"}'
node scripts/send-command.cjs '{"action":"explore_window", "selector":"Notepad", "max_depth":3}'

# Launch an application
node scripts/send-command.cjs '{"action":"launch_app", "path":"notepad.exe", "wait_for_window":"Notepad"}'

# Type and save
node scripts/send-command.cjs '{"action":"type", "selector":"Notepad", "text":"Hello World"}'
node scripts/send-command.cjs '{"action":"hotkey", "keys":"ctrl+s"}'

# Window management
node scripts/send-command.cjs '{"action":"move_window", "selector":"Notepad", "x":100, "y":100}'
node scripts/send-command.cjs '{"action":"resize_window", "selector":"Notepad", "width":800, "height":600}'
node scripts/send-command.cjs '{"action":"minimize_window", "selector":"Notepad"}'

# Process management
node scripts/send-command.cjs '{"action":"list_processes", "filter":"notepad"}'
node scripts/send-command.cjs '{"action":"kill_process", "name":"notepad"}'

# Clipboard operations
node scripts/send-command.cjs '{"action":"hotkey", "keys":"ctrl+a"}'
node scripts/send-command.cjs '{"action":"hotkey", "keys":"ctrl+c"}'
node scripts/send-command.cjs '{"action":"get_clipboard"}'

# Mouse operations
node scripts/send-command.cjs '{"action":"click_at", "x":500, "y":300}'
node scripts/send-command.cjs '{"action":"drag_and_drop", "from_x":100, "from_y":200, "to_x":300, "to_y":400}'

# Screenshot
node scripts/send-command.cjs '{"action":"screenshot", "filename":"my_screenshot.png"}'
```

### New v3.2 Commands

```bash
# Health check - verify system status
node scripts/send-command.cjs '{"action":"health"}'

# Wait for element with timeout
node scripts/send-command.cjs '{"action":"wait_for_element", "selector":"Save", "window":"Notepad", "timeout":5000}'

# Multi-monitor support
node scripts/send-command.cjs '{"action":"list_monitors"}'
node scripts/send-command.cjs '{"action":"screenshot_monitor", "monitor":1, "filename":"secondary.png"}'
node scripts/send-command.cjs '{"action":"move_to_monitor", "selector":"Notepad", "monitor":0, "position":"center"}'

# File dialog automation
node scripts/send-command.cjs '{"action":"file_dialog", "dialog_action":"detect"}'
node scripts/send-command.cjs '{"action":"file_dialog", "dialog_action":"set_path", "path":"C:\\Users\\marke\\Documents"}'
node scripts/send-command.cjs '{"action":"file_dialog", "dialog_action":"set_filename", "filename":"test.txt"}'
node scripts/send-command.cjs '{"action":"file_dialog", "dialog_action":"confirm"}'

# Batch execution
node scripts/send-command.cjs '{"action":"batch", "commands":[{"action":"launch_app","path":"notepad.exe"},{"action":"wait_for_window","selector":"Notepad"},{"action":"type","selector":"Notepad","text":"Hello!"}], "stop_on_error":true}'
```

### New v3.3 Commands

```bash
# Mouse path - move along waypoints
node scripts/send-command.cjs '{"action":"mouse_path", "points":[[100,100],[200,150],[300,100]], "duration":500}'

# Bezier curve movement
node scripts/send-command.cjs '{"action":"mouse_bezier", "start":[100,100], "control1":[150,50], "control2":[250,50], "end":[300,100], "steps":50, "duration":500}'

# Draw (hold + move + release) - for drawing applications
node scripts/send-command.cjs '{"action":"draw", "points":[[100,100],[200,200],[300,100]], "button":"left", "duration":500}'

# Draw bezier curve (hold + bezier path + release) - smooth curves in drawing apps
node scripts/send-command.cjs '{"action":"draw_bezier", "start":[100,100], "control1":[150,50], "control2":[250,150], "end":[300,100], "button":"left", "duration":500}'

# Mouse hold/release
node scripts/send-command.cjs '{"action":"mouse_down", "button":"left", "x":100, "y":200}'
node scripts/send-command.cjs '{"action":"mouse_up", "button":"left"}'

# Key hold/release - for modifier key combinations
node scripts/send-command.cjs '{"action":"key_down", "key":"shift"}'
node scripts/send-command.cjs '{"action":"key_up", "key":"shift"}'

# Type at current cursor position (no selector needed)
node scripts/send-command.cjs '{"action":"type_here", "text":"Hello World"}'

# Wait for element state
node scripts/send-command.cjs '{"action":"wait_for_state", "selector":"SaveButton", "state":"enabled", "timeout":5000}'

# OCR on explicit region
node scripts/send-command.cjs '{"action":"ocr_region", "x":100, "y":200, "width":300, "height":50}'

# Click relative to element anchor
node scripts/send-command.cjs '{"action":"click_relative", "selector":"Window", "anchor":"center", "offset_x":50, "offset_y":30}'
```

### New v4.0 Commands

```bash
# Token-efficient commands (AI agent optimized)
node scripts/send-command.cjs '{"action":"element_exists", "selector":"Save", "window":"Notepad"}'
# Returns: {"exists": true} (~30 tokens)

node scripts/send-command.cjs '{"action":"get_window_summary", "window":"Calculator"}'
# Returns: structured overview (~300 tokens)

node scripts/send-command.cjs '{"action":"get_interactive_elements", "window":"Notepad"}'
# Returns: only clickable/typeable elements (~500 tokens)

node scripts/send-command.cjs '{"action":"fuzzy_find_element", "text":"clos", "window":"Calculator"}'
# Returns: elements matching "clos" with similarity scores

# Smart fallback commands (FlaUI → Vision automatic fallback)
node scripts/send-command.cjs '{"action":"smart_click", "selector":"Submit", "window":"MyApp"}'
# Returns: {"status":"success", "method_used":"flaui", "x":450, "y":320}

node scripts/send-command.cjs '{"action":"smart_type", "selector":"Email", "text":"user@example.com", "window":"MyApp", "clear":true}'
# Returns: {"status":"success", "method_used":"vision"}

# Vision-only commands (for Flutter/Electron/Canvas apps)
node scripts/send-command.cjs '{"action":"vision_click", "text":"Login", "window":"ElectronApp"}'

# Vision API (via Python bridge)
curl http://127.0.0.1:5001/vision/find_text -X POST -H "Content-Type: application/json" -d '{"text":"Submit"}'
curl http://127.0.0.1:5001/vision/click_text -X POST -H "Content-Type: application/json" -d '{"text":"Cancel"}'

# Context manager
curl http://127.0.0.1:5001/context/stats
# Returns: {"cache_entries":5, "hits":12, "misses":4, "hit_rate":0.75}

curl http://127.0.0.1:5001/context/clear -X POST
curl http://127.0.0.1:5001/context/invalidate -X POST -H "Content-Type: application/json" -d '{"window":"Calculator"}'
```

### New v4.3 Commands

```bash
# Agent Vision Commands (via MainAgentService)
node scripts/send-command.cjs '{"action":"vision_screenshot"}'
# Returns: {"screenshot": {"data": "base64...", "width": 1920, "height": 1080, ...}}

node scripts/send-command.cjs '{"action":"vision_screenshot_region", "x":100, "y":100, "width":600, "height":400}'
# Returns: cropped region as base64 JPEG

node scripts/send-command.cjs '{"action":"vision_config"}'
# Returns: {"mode": "auto", "jpeg_quality": 75, "max_width": 1920, ...}

node scripts/send-command.cjs '{"action":"vision_screenshot_cache_stats"}'
# Returns: {"entries": 2, "hits": 15, "misses": 3, "hit_rate": 0.833, ...}

node scripts/send-command.cjs '{"action":"vision_screenshot_cache_clear"}'
# Clears all cached screenshots

# Smart click/type with agent_fallback - returns screenshot when element not found
node scripts/send-command.cjs '{"action":"smart_click", "search_text":"Submit", "window":"MyApp", "agent_fallback":true}'
# If element found: {"status": "success", "method_used": "flaui", "x": 450, "y": 320}
# If not found: {"status": "agent_fallback", "screenshot": {...}, "suggestion": "Use click_at with coordinates"}

node scripts/send-command.cjs '{"action":"smart_type", "search_text":"Email", "text":"user@example.com", "window":"MyApp", "agent_fallback":true, "clear":true}'

# Screenshot cache via HTTP (Python bridge)
curl http://127.0.0.1:5001/vision/screenshot_cache/stats
curl -X POST http://127.0.0.1:5001/vision/screenshot_cache/clear
curl -X POST http://127.0.0.1:5001/vision/screenshot_cache/config \
  -H "Content-Type: application/json" \
  -d '{"ttl_seconds": 3.0, "max_entries": 10}'
curl -X POST http://127.0.0.1:5001/vision/screenshot_cached \
  -H "Content-Type: application/json" \
  -d '{"use_cache": true, "jpeg_quality": 75}'

# WebSocket streaming - get URL for real-time screenshots
node scripts/send-command.cjs '{"action":"vision_stream", "fps":5, "quality":70}'
# Returns: {"websocket_url": "ws://localhost:5001/vision/stream?fps=5&quality=70", ...}

# Test streaming with dedicated script
node scripts/test-stream.cjs --fps 10 --duration 5 --save
```

## Bridge API Reference

### Python Bridge (pywinauto + Vision) - Port 5001

The Python bridge provides legacy Win32 automation, vision layer, and context manager via Flask REST API.

**Core Endpoints:**
| Endpoint | Method | Description |
|----------|--------|-------------|
| `/health` | GET | Health check |
| `/explore` | GET | List all windows (cached) |
| `/click` | POST | Click on element |
| `/double_click` | POST | Double-click |
| `/right_click` | POST | Right-click |
| `/click_at` | POST | Click at coordinates |
| `/mouse_move` | POST | Move mouse |
| `/scroll` | POST | Scroll wheel |
| `/drag_and_drop` | POST | Drag and drop |
| `/type` | POST | Type text |
| `/hotkey` | POST | Send keyboard shortcut |
| `/key_press` | POST | Single key press |
| `/focus_window` | POST | Focus window |
| `/close_window` | POST | Close window |

**Vision Endpoints (v4.0):**
| Endpoint | Method | Description |
|----------|--------|-------------|
| `/vision/detect` | POST | Detect UI elements via OmniParser |
| `/vision/ocr` | POST | OCR text recognition |
| `/vision/find_text` | POST | Find text location using OCR |
| `/vision/click_text` | POST | Find and click text via OCR |
| `/vision/analyze` | POST | Combined detection + OCR |

**Context Manager Endpoints (v4.0):**
| Endpoint | Method | Description |
|----------|--------|-------------|
| `/context/stats` | GET | Get cache statistics |
| `/context/clear` | POST | Clear all caches |
| `/context/enable` | POST | Enable caching |
| `/context/disable` | POST | Disable caching |
| `/context/invalidate` | POST | Invalidate cache for window |

**Agent Vision Endpoints (v4.2):**
| Endpoint | Method | Description |
|----------|--------|-------------|
| `/vision/config` | GET | Get current vision mode and settings |
| `/vision/config` | POST | Set vision mode (local/agent/auto) and compression settings |
| `/vision/screenshot` | POST | Get optimized base64 JPEG screenshot for AI vision |
| `/vision/screenshot_region` | POST | Get cropped region screenshot (token-efficient) |
| `/vision/analyze_or_screenshot` | POST | Smart endpoint: local analysis with screenshot fallback |

**Screenshot Cache Endpoints (v4.3):**
| Endpoint | Method | Description |
|----------|--------|-------------|
| `/vision/screenshot_cache/stats` | GET | Get cache statistics (hits, misses, hit_rate, entries) |
| `/vision/screenshot_cache/clear` | POST | Clear all cached screenshots |
| `/vision/screenshot_cache/config` | POST | Update TTL and max_entries settings |
| `/vision/screenshot_cached` | POST | Get screenshot with caching support (use_cache param) |

### Agent Vision Mode (v4.2)

The agent vision mode allows AI agents to perform their own vision analysis instead of relying on local OmniParser + OCR. This is useful when:

- Local vision models aren't available or installed
- The AI agent has superior vision capabilities
- You want to reduce local processing overhead
- Local detection produces limited results

**Vision Modes:**

| Mode | Description |
|------|-------------|
| `local` | Full local processing via OmniParser + WinRT OCR (default) |
| `agent` | Skip local processing, return optimized screenshots for AI |
| `auto` | Try local first; if results are limited, include screenshot as fallback |

**Configuration:**

Set mode via environment variable:
```bash
set VISION_MODE=agent  # Windows
export VISION_MODE=agent  # Linux/Mac
```

Or via API at runtime:
```bash
# Get current configuration
curl http://127.0.0.1:5001/vision/config

# Set to agent mode with custom quality
curl -X POST http://127.0.0.1:5001/vision/config \
  -H "Content-Type: application/json" \
  -d '{"mode": "agent", "jpeg_quality": 75, "max_width": 1920}'
```

**Using Agent Vision:**

```bash
# Get optimized screenshot (compressed JPEG, ~150-300KB for 1080p)
curl -X POST http://127.0.0.1:5001/vision/screenshot \
  -H "Content-Type: application/json" \
  -d '{"jpeg_quality": 70}'

# Get specific region (more token-efficient)
curl -X POST http://127.0.0.1:5001/vision/screenshot_region \
  -H "Content-Type: application/json" \
  -d '{"x": 100, "y": 100, "width": 600, "height": 400}'

# Smart endpoint - local analysis with automatic screenshot fallback
curl -X POST http://127.0.0.1:5001/vision/analyze_or_screenshot \
  -H "Content-Type: application/json" \
  -d '{"force_screenshot": true}'
```

**Response format:**
```json
{
  "status": "success",
  "screenshot": {
    "data": "base64-encoded-jpeg...",
    "width": 1920,
    "height": 1080,
    "original_width": 2560,
    "original_height": 1440,
    "size_bytes": 185000,
    "compression_ratio": 35.2,
    "format": "jpeg",
    "timestamp": 1234567890.123
  }
}
```

### Response Format

All commands return consistent JSON responses:

```json
// Success
{
  "status": "success",
  "action": "click",
  "engine": "pywinauto|flaui|vision"
}

// Error
{
  "status": "error",
  "code": "ERROR_CODE",
  "message": "Human readable message",
  "engine": "pywinauto|flaui|vision"
}
```

## Project Structure

```text
skill/windows-desktop-automation/
├── SKILL.md                # AI Agent documentation (command API)
├── README.md               # Human developer documentation
├── scripts/                # Helper scripts for lifecycle management
│   ├── auto-setup.ps1      # Installs dependencies
│   ├── start-all.ps1       # Launches bridges and agent
│   ├── stop-all.ps1        # Kills processes
│   ├── send-command.cjs    # Client to talk to the agent
│   └── test-commands.cjs   # Automated test suite
├── src/                    # Source code
│   ├── src/MainAgentService # The C# Orchestrator
│   ├── bridge_python/      # The Python Flask Server (Vision + OCR + Context)
│   └── .venv/              # Local virtualenv (not committed)
```

## Troubleshooting

### Build & Setup Issues

| Problem | Solution |
|---------|----------|
| **Build Fails (File Locked)** | Run `scripts/stop-all.ps1` to release locks |
| **Python Bridge Offline** | Ensure `start-all.ps1` ran successfully, check port 5001 |
| **Python Bridge Fails** | Ensure `.venv` deps are installed (run `scripts/auto-setup.ps1`) |

### Command Issues

| Problem | Solution |
|---------|----------|
| **Window Not Found** | Use partial matching - "Note" matches "Notepad - Untitled" |
| **Element Click Fails** | First `focus_window`, then `explore_window` to find element names |
| **Hotkey Not Working** | Ensure target window is focused first |

### Test Failures

| Test Failure | Common Cause | Solution |
|--------------|--------------|----------|
| `explore` fails | Agent not running | Run `scripts/start-all.ps1` |
| `focus_window` fails | Window doesn't exist | Launch the target app first |
| `type` fails | No element selected | Focus window and click element first |
| Timeout errors | System under load | Increase timeout in test script |
| JSON parse errors | Bridge warning messages | Already handled in test runner |

## Documentation

- **SKILL.md**: Comprehensive API documentation for AI agents (command formats, parameters, error codes)
- **README.md**: This file - setup and usage guide for developers

## Version History

- **v4.3.2**: Agent Vision Enhancements - Screenshot caching with LRU/TTL (ScreenshotCache class), 7 new JSON commands for agent vision workflows (vision_screenshot, vision_screenshot_region, vision_config, vision_analyze_smart, vision_screenshot_cache_stats, vision_screenshot_cache_clear, vision_stream), WebSocket streaming for real-time screenshots, smart_click/smart_type agent_fallback parameter returns screenshot when element not found
- **v4.3**: Architecture Simplification - Removed SikuliX/Java Bridge, all OCR via Python bridge WinRT (3.2x faster), no Java runtime required
- **v4.2**: Agent Vision Mode - New endpoints (`/vision/screenshot`, `/vision/screenshot_region`, `/vision/config`, `/vision/analyze_or_screenshot`) for AI agent vision. Supports local/agent/auto modes with configurable JPEG compression. OmniParser detection always uses tiling (no global resize)
- **v4.1**: WinRT Native OCR - Direct Python → WinRT bindings (166ms full screen, 60ms region)
- **v4.0**: AI Integration - Vision layer (OmniParser + Windows OCR), token-efficient commands (`element_exists`, `get_window_summary`, `get_interactive_elements`, `fuzzy_find_element`), smart fallback commands (`smart_click`, `smart_type`), context manager with caching
- **v3.6**: Human-like behavior (`set_human_mode`, `human_click`, `human_type`, `human_move`), Bezier curves, typo simulation, fatigue system
- **v3.5**: Performance diagnostics (`get_metrics`), element caching, event subscriptions
- **v3.4**: Enhanced mouse (`hover`, `swipe`, easing functions), wait commands (`wait_for_color`, `wait_for_idle`, `wait_for_text`)
- **v3.3**: Advanced mouse controls (`mouse_path`, `mouse_bezier`, `draw`, `mouse_down`, `mouse_up`, `click_relative`), keyboard hold/release (`key_down`, `key_up`), type at cursor (`type_here`), element state waiting (`wait_for_state`), region OCR (`ocr_region`)
- **v3.2**: Multi-monitor support (`list_monitors`, `screenshot_monitor`, `move_to_monitor`), file dialog automation (`file_dialog`), batch command execution (`batch`), element waiting (`wait_for_element`), health check endpoint (`health`)
- **v3.1**: Added automated test suite (33 tests), refactored all handlers to use helper methods, simplified Python/Java bridges with consistent response format, added comprehensive endpoint coverage to all bridges
- **v3.0**: Added window management (move/resize/min/max), process control (list/kill), drag-and-drop, mouse_move, find_element, get_window_info, read_text
- **v2.0**: Added focus_window, close_window, launch_app, hotkey, explore_window, wait_for_window, click_at, clipboard operations, key_press
- **v1.0**: Initial release with explore, click, type, scroll, screenshot
