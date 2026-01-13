---
name: windows-desktop-automation
description: AI agent capability for Windows desktop GUI automation via multi-engine orchestrator.
version: 4.4.0
---

# Windows Desktop Automation

Control Windows GUI: explore, click, type, manage windows. Vision layer for Flutter/Electron/Canvas apps. Agent vision mode for AI-driven screen analysis.

## Quick Start

```bash
node scripts/send-command.cjs '{"action": "explore"}'
node scripts/get-command-help.cjs <command>          # Full command details
node scripts/get-command-help.cjs --list             # All commands
node scripts/get-command-help.cjs --category mouse   # By category
```

## Essential Workflow

```
1. explore              -> List windows
2. focus_window         -> Activate target
3. explore_window       -> Find elements (or get_window_summary for ~300 tokens)
4. click/type/hotkey    -> Interact
5. screenshot           -> Verify
```

---

## Command Reference (98 JSON + 18 HTTP endpoints)

### Discovery (9)

| Command | Description |
|---------|-------------|
| `explore` | **START HERE.** List all visible windows |
| `explore_window` | Explore elements inside a window (`max_depth` param) |
| `find_element` | Find element with detailed info + clickable coordinates |
| `get_window_info` | Get window details (title, bounds, state) |
| `get_element_bounds` | Get element bounding box and center point |
| `wait_for_element` | Wait for element to appear (`timeout` param) |
| `wait_for_state` | Wait for state: `enabled`/`disabled`/`visible`/`hidden`/`exists` |
| `wait_for_color` | Wait for pixel color at coordinates |
| `wait_for_idle` | Wait for UI to become responsive |

### Window Management (9)

| Command | Description |
|---------|-------------|
| `focus_window` | **REQUIRED** before interaction. Bring to foreground |
| `close_window` | Close window gracefully |
| `wait_for_window` | Wait for window to appear after launch |
| `move_window` | Move window to x,y coordinates |
| `resize_window` | Resize window to width,height |
| `minimize_window` | Minimize window |
| `maximize_window` | Maximize window |
| `restore_window` | Restore from minimized/maximized |
| `move_to_monitor` | Move to specific monitor (multi-display) |

### Application Control (3)

| Command | Description |
|---------|-------------|
| `launch_app` | Start app by path, optional `wait_for_window` |
| `list_processes` | List running processes, filter by name |
| `kill_process` | Kill by PID or name, `force` option |

### Mouse Actions (19)

| Command | Description |
|---------|-------------|
| `click` | Click element by selector |
| `double_click` | Double-click element |
| `right_click` | Right-click for context menu |
| `long_press` | Press and hold (`duration` param) |
| `click_at` | Click at absolute x,y coordinates |
| `click_relative` | Click at offset from element anchor |
| `mouse_move` | Move cursor to x,y or selector |
| `mouse_move_eased` | Move with easing: `linear`, `ease-in-out`, `bounce`, `elastic` |
| `hover` | Hover over element for duration |
| `drag_and_drop` | Drag between selectors or coordinates |
| `swipe` | Touch-style swipe gesture |
| `scroll` | Scroll wheel up/down by amount |
| `mouse_path` | Move through array of waypoints |
| `mouse_bezier` | Move along cubic bezier curve |
| `draw` | Hold button, move path, release (for drawing apps) |
| `draw_bezier` | Hold button, bezier curve path, release |
| `mouse_down` | Press and hold button |
| `mouse_up` | Release held button |
| `get_cursor_position` | Get current cursor x,y |

### Keyboard Actions (6)

| Command | Description |
|---------|-------------|
| `type` | Type text into element by selector |
| `type_here` | Type at current cursor position |
| `hotkey` | Send combo: `ctrl+c`, `alt+f4`, `ctrl+shift+s` |
| `key_press` | Single key: `enter`, `tab`, `escape`, `f1`-`f12`, arrows |
| `key_down` | Press and hold key |
| `key_up` | Release held key |

### Clipboard (3)

| Command | Description |
|---------|-------------|
| `get_clipboard` | Read clipboard text |
| `set_clipboard` | Write text to clipboard |
| `clipboard_image` | Copy image to/from clipboard (`set`/`get` operation) |

### Text & OCR (3)

| Command | Description |
|---------|-------------|
| `read_text` | Read text from element or region |
| `ocr_region` | OCR screen region (WinRT native, ~60ms) |
| `wait_for_text` | Wait for OCR text to appear |

### Screenshots (4)

| Command | Description |
|---------|-------------|
| `screenshot` | Capture full screen |
| `element_screenshot` | Capture specific element only |
| `list_monitors` | List connected displays |
| `screenshot_monitor` | Capture specific monitor by index |

### File Dialogs (1)

| Command | Description |
|---------|-------------|
| `file_dialog` | Automate Open/Save: `detect`, `set_path`, `set_filename`, `confirm`, `cancel` |

### Events (3)

| Command | Description |
|---------|-------------|
| `subscribe` | Subscribe to element events (`focus`, `invoke`, `property_change`) |
| `unsubscribe` | Unsubscribe by `subscription_id` |
| `get_subscriptions` | List active subscriptions |

### Diagnostics (5)

| Command | Description |
|---------|-------------|
| `get_metrics` | Get command performance stats |
| `clear_metrics` | Reset performance metrics |
| `set_debug_mode` | Enable/disable verbose logging |
| `get_cache_stats` | Get element cache stats |
| `clear_cache` | Clear element cache |

### Human-like Behavior (5)

| Command | Description |
|---------|-------------|
| `set_human_mode` | Configure: `jitter`, `typing_error_rate`, `fatigue`, `thinking_delay` |
| `get_human_mode` | Get current human mode settings |
| `human_click` | Click with bezier mouse path + random pauses |
| `human_type` | Type with variable speed, typos, auto-corrections |
| `human_move` | Move mouse with curves and micro-jitter |

### Token-Efficient Commands (6)

| Command | Tokens | Description |
|---------|--------|-------------|
| `element_exists` | ~30 | Boolean check only |
| `get_element_brief` | ~100 | Just name, type, coords |
| `get_window_summary` | ~300 | Overview vs ~2000 for `explore_window` |
| `get_interactive_elements` | ~500 | Only clickable/typeable items |
| `describe_element` | ~100 | Vision-based description at x,y |
| `fuzzy_find_element` | ~200 | Levenshtein fuzzy matching |

### Smart Fallback (3)

| Command | Description |
|---------|-------------|
| `smart_click` | FlaUI first -> Vision fallback. Returns `method_used`. Supports `agent_fallback` param |
| `smart_type` | FlaUI first -> Vision fallback. Supports `clear` and `agent_fallback` params |
| `vision_click` | Click by OCR text (Flutter/Electron/Canvas) |

### Agent Vision Commands (7)

| Command | Description |
|---------|-------------|
| `vision_screenshot` | Get optimized base64 JPEG screenshot for AI vision |
| `vision_screenshot_region` | Get cropped region screenshot (token-efficient) |
| `vision_config` | Get/set vision mode (`local`/`agent`/`auto`) |
| `vision_analyze_smart` | Smart analysis respecting current vision mode |
| `vision_screenshot_cache_stats` | Get screenshot cache statistics |
| `vision_screenshot_cache_clear` | Clear screenshot cache |
| `vision_stream` | Get WebSocket URL for real-time screenshot streaming |

### Win32 Low-Level Commands (7)

| Command | Description |
|---------|-------------|
| `scroll_sendinput` | Low-level scroll using Win32 SendInput (bypasses UI Automation) |
| `type_sendinput` | Type Unicode text via Win32 SendInput (direct keyboard events) |
| `set_always_on_top` | Pin/unpin window always on top using SetWindowPos |
| `flash_window` | Flash window taskbar/caption for user attention |
| `set_window_opacity` | Set window transparency (0-255 alpha) |
| `fast_screenshot` | Fast GDI BitBlt screenshot (faster than WinForms) |
| `clipboard_image` | Copy image to/from clipboard (`set`/`get` operations) |

### Batch & System (2)

| Command | Description |
|---------|-------------|
| `batch` | Execute command array, `stop_on_error` option |
| `health` | System health check, bridge status |

---

## Vision/Context HTTP Endpoints (Python Bridge)

The Python bridge runs on `http://127.0.0.1:5001` and provides vision, OCR, and context management capabilities.

### Vision - Local Processing

```bash
POST /vision/detect      # OmniParser UI detection
POST /vision/ocr         # WinRT OCR (166ms full, 60ms region)
POST /vision/find_text   # Locate text on screen
POST /vision/click_text  # Find and click text
POST /vision/analyze     # Combined detection + OCR
```

### Vision - Agent Mode

Returns screenshots for AI vision analysis instead of local processing.

```bash
GET  /vision/config                  # Get current mode/settings
POST /vision/config                  # Set mode: local|agent|auto
POST /vision/screenshot              # Optimized base64 JPEG for AI
POST /vision/screenshot_region       # Cropped region (token-efficient)
POST /vision/analyze_or_screenshot   # Smart: local first, screenshot fallback
```

**Vision Modes:**

| Mode | Behavior |
|------|----------|
| `local` | Full OmniParser + OCR processing (default) |
| `agent` | Skip local processing, return optimized screenshots |
| `auto` | Try local first; include screenshot if results are limited |

### Screenshot Caching

Reduces redundant screen captures for rapid successive operations.

```bash
GET  /vision/screenshot_cache/stats  # Cache statistics
POST /vision/screenshot_cache/clear  # Clear cached screenshots
POST /vision/screenshot_cache/config # Update TTL/max_entries settings
POST /vision/screenshot_cached       # Screenshot with cache support
```

| Environment Variable | Default | Description |
|----------------------|---------|-------------|
| `SCREENSHOT_CACHE_TTL` | 2.0 | Seconds before cached screenshots expire |
| `SCREENSHOT_CACHE_MAX_ENTRIES` | 5 | Maximum cached screenshots (LRU eviction) |

### WebSocket Streaming

```bash
WS /vision/stream?fps=5&quality=70   # Real-time JPEG stream
```

### Context/Cache Management

```bash
GET  /context/stats      # Cache hit rate
POST /context/clear      # Clear all caches
POST /context/enable     # Enable caching
POST /context/disable    # Disable caching
POST /context/invalidate # Invalidate specific window
```

---

## Usage Examples

### Agent Vision Mode

```bash
# Switch to agent mode
curl -X POST http://127.0.0.1:5001/vision/config \
  -H "Content-Type: application/json" \
  -d '{"mode": "agent", "jpeg_quality": 75}'

# Get optimized screenshot for AI vision
curl -X POST http://127.0.0.1:5001/vision/screenshot \
  -H "Content-Type: application/json" \
  -d '{"max_width": 1920, "jpeg_quality": 70}'
```

**Screenshot Response:**

```json
{
  "screenshot": {
    "data": "base64...",
    "width": 1920,
    "height": 1080,
    "size_bytes": 185000,
    "compression_ratio": 35.2,
    "format": "jpeg"
  }
}
```

### Smart Fallback with Agent Vision

When `smart_click` or `smart_type` fails to find an element, enable `agent_fallback` to get a screenshot for AI analysis.

```bash
node scripts/send-command.cjs '{
  "action": "smart_click",
  "search_text": "Submit Button",
  "window": "MyApp",
  "agent_fallback": true,
  "jpeg_quality": 75
}'
```

**Success Response:**

```json
{
  "status": "success",
  "action": "smart_click",
  "method_used": "flaui",
  "x": 450,
  "y": 320
}
```

**Agent Fallback Response (element not found):**

```json
{
  "status": "agent_fallback",
  "code": "ELEMENT_NOT_FOUND",
  "message": "Element 'Submit Button' not found with any method",
  "suggestion": "Analyze the screenshot and provide click coordinates using click_at command",
  "screenshot": {
    "data": "base64...",
    "width": 1920,
    "height": 1080,
    "format": "jpeg"
  }
}
```

### Win32 Low-Level Commands

```bash
# Low-level scroll (bypasses UI Automation)
node scripts/send-command.cjs '{
  "action": "scroll_sendinput",
  "delta": -120,
  "x": 500,
  "y": 400
}'

# Unicode typing including emoji
node scripts/send-command.cjs '{
  "action": "type_sendinput",
  "text": "Hello World! 🎉"
}'

# Pin window on top
node scripts/send-command.cjs '{
  "action": "set_always_on_top",
  "selector": "Notepad",
  "enable": true
}'

# Set window transparency
node scripts/send-command.cjs '{
  "action": "set_window_opacity",
  "selector": "Paint",
  "alpha": 180
}'

# Fast GDI screenshot
node scripts/send-command.cjs '{
  "action": "fast_screenshot",
  "region": [0, 0, 800, 600],
  "format": "png"
}'
```

---

## Error Recovery

| Code | Recovery Action |
|------|-----------------|
| `WINDOW_NOT_FOUND` | Run `explore` to list available windows |
| `ELEMENT_NOT_FOUND` | Try `explore_window`, `vision_click`, or `click_at` |
| `TIMEOUT` | Increase `timeout` parameter |
| `VISION_NOT_AVAILABLE` | Start Python bridge with `scripts/start-all.ps1` |
| `ALL_METHODS_FAILED` | Verify element exists, try screenshot + `click_at` |

**Agent Fallback:** When using `smart_click`/`smart_type` with `agent_fallback: true`, a `status: "agent_fallback"` response means the element wasn't found but a screenshot is included. Analyze it and call `click_at` with coordinates.

Full error reference: `node scripts/get-command-help.cjs --errors`

---

## Architecture

```
AI Agent --> send-command.cjs --> MainAgentService (C#/.NET 9)
                                        |
                                 +------+------+
                                 |             |
                              FlaUI        Python Bridge
                              (UIA3)       (localhost:5001)
                                 |             |
                           Windows UI    +-----+-----+
                           Automation    |     |     |
                                     pywinauto Vision Context
                                              |      Manager
                                         +----+----+
                                         |         |
                                      WinRT    OmniParser
                                       OCR       ONNX
```

| Layer | Technology | Purpose |
|-------|------------|---------|
| Primary | FlaUI (UIA3) | Windows UI Automation - native, fast |
| Secondary | Python Bridge | pywinauto + Vision + WinRT OCR + Context |
| Vision | OmniParser (ONNX) | UI element detection for non-accessible apps |
| OCR | WinRT OCR | Fast native text recognition (~60ms) |

---

## Tips

1. **Always `explore` first** - Never assume window state
2. **`focus_window` before interact** - Required for clicks/typing
3. **Use `wait_for_*` commands** - Apps need time to respond
4. **Token-efficient first** - Prefer `get_window_summary` over `explore_window`
5. **Use `smart_click`/`smart_type`** - Auto FlaUI -> Vision fallback
6. **Check `health`** - Diagnose bridge issues
7. **Use `agent_fallback: true`** - Get screenshots when elements aren't found

---

*Full command schema and examples: `node scripts/get-command-help.cjs <command>`*
