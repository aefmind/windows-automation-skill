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

## All Commands (98 JSON + 18 HTTP endpoints)

### Discovery (9)
| Command | Description |
|---------|-------------|
| `explore` | **START HERE.** List all visible windows |
| `explore_window` | Explore elements inside a window (max_depth param) |
| `find_element` | Find element with detailed info + clickable coordinates |
| `get_window_info` | Get window details (title, bounds, state) |
| `get_element_bounds` | Get element bounding box and center point |
| `wait_for_element` | Wait for element to appear (timeout param) |
| `wait_for_state` | Wait for state: enabled/disabled/visible/hidden/exists |
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
| `launch_app` | Start app by path, optional wait_for_window |
| `list_processes` | List running processes, filter by name |
| `kill_process` | Kill by PID or name, force option |

### Mouse Actions (19)
| Command | Description |
|---------|-------------|
| `click` | Click element by selector |
| `double_click` | Double-click element |
| `right_click` | Right-click for context menu |
| `long_press` | Press and hold (duration param) |
| `click_at` | Click at absolute x,y coordinates |
| `click_relative` | Click at offset from element anchor |
| `mouse_move` | Move cursor to x,y or selector |
| `mouse_move_eased` | Move with easing: linear, ease-in-out, bounce, elastic |
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
| `hotkey` | Send combo: ctrl+c, alt+f4, ctrl+shift+s |
| `key_press` | Single key: enter, tab, escape, f1-f12, arrows |
| `key_down` | Press and hold key |
| `key_up` | Release held key |

### Clipboard (3)
| Command | Description |
|---------|-------------|
| `get_clipboard` | Read clipboard text |
| `set_clipboard` | Write text to clipboard |
| `clipboard_image` | Copy image to/from clipboard (set/get) |

### Text/OCR (3)
| Command | Description |
|---------|-------------|
| `read_text` | Read text from element or region |
| `ocr_region` | OCR screen region (WinRT native, 60ms) |
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
| `file_dialog` | Automate Open/Save: detect, set_path, set_filename, confirm, cancel |

### Events (3)
| Command | Description |
|---------|-------------|
| `subscribe` | Subscribe to element events (focus, invoke, property_change) |
| `unsubscribe` | Unsubscribe by subscription_id |
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
| `set_human_mode` | Configure: jitter, typing_error_rate, fatigue, thinking_delay |
| `get_human_mode` | Get current human mode settings |
| `human_click` | Click with bezier mouse path + random pauses |
| `human_type` | Type with variable speed, typos, auto-corrections |
| `human_move` | Move mouse with curves and micro-jitter |

### Token-Efficient (6)
| Command | Tokens | Description |
|---------|--------|-------------|
| `element_exists` | ~30 | Boolean check only |
| `get_element_brief` | ~100 | Just name, type, coords |
| `get_window_summary` | ~300 | Overview vs ~2000 for explore_window |
| `get_interactive_elements` | ~500 | Only clickable/typeable items |
| `describe_element` | ~100 | Vision-based description at x,y |
| `fuzzy_find_element` | ~200 | Levenshtein fuzzy matching |

### Smart Fallback (3)
| Command | Description |
|---------|-------------|
| `smart_click` | FlaUI first → Vision fallback. Returns method_used. Supports `agent_fallback` param |
| `smart_type` | FlaUI first → Vision fallback. Supports `clear` and `agent_fallback` params |
| `vision_click` | Click by OCR text (Flutter/Electron/Canvas) |

### Agent Vision Commands (7) - v4.3
| Command | Description |
|---------|-------------|
| `vision_screenshot` | Get optimized base64 JPEG screenshot for AI vision |
| `vision_screenshot_region` | Get cropped region screenshot (token-efficient) |
| `vision_config` | Get/set vision mode (local/agent/auto) |
| `vision_analyze_smart` | Smart analysis respecting current vision mode |
| `vision_screenshot_cache_stats` | Get screenshot cache statistics |
| `vision_screenshot_cache_clear` | Clear screenshot cache |
| `vision_stream` | Get WebSocket URL for real-time screenshot streaming |

### Win32 Low-Level Commands (7) - v4.4
| Command | Description |
|---------|-------------|
| `scroll_sendinput` | Low-level scroll using Win32 SendInput (bypasses UI Automation) |
| `type_sendinput` | Type Unicode text via Win32 SendInput (direct keyboard events) |
| `set_always_on_top` | Pin/unpin window always on top using SetWindowPos |
| `flash_window` | Flash window taskbar/caption for user attention |
| `set_window_opacity` | Set window transparency (0-255 alpha) |
| `fast_screenshot` | Fast GDI BitBlt screenshot (faster than WinForms) |
| `clipboard_image` | Copy image to/from clipboard (set/get operations) |

### Batch/System (2)
| Command | Description |
|---------|-------------|
| `batch` | Execute command array, stop_on_error option |
| `health` | System health check, bridge status |

---

## Vision/Context HTTP Endpoints (Python Bridge)

```bash
# Vision - Local Processing (OmniParser + WinRT OCR)
POST /vision/detect      # OmniParser UI detection
POST /vision/ocr         # WinRT OCR (166ms full, 60ms region)
POST /vision/find_text   # Locate text on screen
POST /vision/click_text  # Find and click text
POST /vision/analyze     # Combined detection + OCR

# Vision - Agent Mode (v4.2) - Returns screenshots for AI vision
GET  /vision/config                  # Get current mode/settings
POST /vision/config                  # Set mode: local|agent|auto
POST /vision/screenshot              # Optimized base64 JPEG for AI
POST /vision/screenshot_region       # Cropped region (token-efficient)
POST /vision/analyze_or_screenshot   # Smart: local first, screenshot fallback

# Screenshot Cache (v4.3) - Reduce redundant captures
GET  /vision/screenshot_cache/stats  # Cache statistics (hits, misses, hit_rate)
POST /vision/screenshot_cache/clear  # Clear cached screenshots
POST /vision/screenshot_cache/config # Update TTL/max_entries settings
POST /vision/screenshot_cached       # Screenshot with cache support (use_cache param)

# WebSocket Streaming (v4.3) - Real-time screenshots
WS   /vision/stream?fps=5&quality=70  # Connect for real-time JPEG stream

# Context/Cache
GET  /context/stats      # Cache hit rate
POST /context/clear      # Clear all caches
POST /context/enable     # Enable caching
POST /context/disable    # Disable caching
POST /context/invalidate # Invalidate specific window
```

### Agent Vision Mode (v4.2)

When local OmniParser/OCR isn't available or produces limited results, use agent vision mode to let your AI model analyze screenshots directly.

**Vision Modes:**
| Mode | Behavior |
|------|----------|
| `local` | Full OmniParser + OCR processing (default) |
| `agent` | Skip local processing, return optimized screenshots |
| `auto` | Try local first; include screenshot if results are limited |

**Configuration:**
```bash
# Get current config
curl http://127.0.0.1:5001/vision/config

# Switch to agent mode
curl -X POST http://127.0.0.1:5001/vision/config \
  -H "Content-Type: application/json" \
  -d '{"mode": "agent", "jpeg_quality": 75}'

# Get optimized screenshot for AI vision
curl -X POST http://127.0.0.1:5001/vision/screenshot \
  -H "Content-Type: application/json" \
  -d '{"max_width": 1920, "jpeg_quality": 70}'

# Get specific region (more token-efficient)
curl -X POST http://127.0.0.1:5001/vision/screenshot_region \
  -H "Content-Type: application/json" \
  -d '{"x": 0, "y": 0, "width": 800, "height": 600}'
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

### Screenshot Caching (v4.3)

Reduce redundant screen captures for rapid successive operations (e.g., find → click workflows).

**Cache Endpoints:**
```bash
# Get cache statistics
curl http://127.0.0.1:5001/vision/screenshot_cache/stats
# Returns: {"entries":2, "max_entries":5, "ttl_seconds":2.0, "hits":15, "misses":3, "hit_rate":0.833}

# Clear cache
curl -X POST http://127.0.0.1:5001/vision/screenshot_cache/clear

# Update configuration
curl -X POST http://127.0.0.1:5001/vision/screenshot_cache/config \
  -H "Content-Type: application/json" \
  -d '{"ttl_seconds": 3.0, "max_entries": 10}'

# Get screenshot with caching (default: use_cache=true)
curl -X POST http://127.0.0.1:5001/vision/screenshot_cached \
  -H "Content-Type: application/json" \
  -d '{"jpeg_quality": 75, "use_cache": true}'
```

**Environment Variables:**
| Variable | Default | Description |
|----------|---------|-------------|
| `SCREENSHOT_CACHE_TTL` | 2.0 | Seconds before cached screenshots expire |
| `SCREENSHOT_CACHE_MAX_ENTRIES` | 5 | Maximum cached screenshots (LRU eviction) |

### Smart Fallback with Agent Vision (v4.3)

When `smart_click` or `smart_type` fails to find an element, enable `agent_fallback` to get a screenshot for AI analysis instead of just an error.

**Usage:**
```bash
# Enable agent_fallback - returns screenshot on failure
node scripts/send-command.cjs '{
  "action": "smart_click",
  "search_text": "Submit Button",
  "window": "MyApp",
  "agent_fallback": true,
  "jpeg_quality": 75
}'
```

**Success Response (element found):**
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
  "action": "smart_click",
  "code": "ELEMENT_NOT_FOUND",
  "message": "Element 'Submit Button' not found with any method",
  "search_text": "Submit Button",
  "suggestion": "Analyze the screenshot and provide click coordinates using click_at command",
  "screenshot": {
    "data": "base64...",
    "width": 1920,
    "height": 1080,
    "size_bytes": 185000,
    "format": "jpeg"
  }
}
```

**Parameters for smart_click/smart_type:**
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `agent_fallback` | bool | false | Return screenshot on failure instead of error |
| `jpeg_quality` | int | 75 | Screenshot quality when fallback triggers (1-100) |

---

### Win32 Low-Level Commands (v4.4)

Direct Win32 API access for scenarios where UI Automation is too slow or doesn't work (games, custom controls, DirectX overlays).

**scroll_sendinput** - Low-level mouse scroll
```bash
node scripts/send-command.cjs '{
  "action": "scroll_sendinput",
  "delta": -120,
  "x": 500,
  "y": 400,
  "horizontal": false
}'
# delta: positive=up, negative=down (120 = 1 wheel notch)
# x,y: optional, scrolls at current cursor if omitted
# horizontal: optional, for horizontal scroll
```

**type_sendinput** - Unicode keyboard input
```bash
node scripts/send-command.cjs '{
  "action": "type_sendinput",
  "text": "Hello 世界! 🎉",
  "delay": 10
}'
# Supports full Unicode including emoji
# delay: optional ms between keystrokes (default 0)
```

**set_always_on_top** - Pin window above others
```bash
node scripts/send-command.cjs '{
  "action": "set_always_on_top",
  "selector": "Notepad",
  "enable": true
}'
# enable: true=pin on top, false=normal z-order
```

**flash_window** - Flash window for attention
```bash
node scripts/send-command.cjs '{
  "action": "flash_window",
  "selector": "Visual Studio Code",
  "count": 5,
  "timeout": 0,
  "flags": 3
}'
# count: number of flashes (0=until focused)
# timeout: ms between flashes (0=default blink rate)
# flags: 1=caption, 2=taskbar, 3=both (default)
```

**set_window_opacity** - Window transparency
```bash
node scripts/send-command.cjs '{
  "action": "set_window_opacity",
  "selector": "Paint",
  "alpha": 180
}'
# alpha: 0=invisible, 255=fully opaque
```

**fast_screenshot** - GDI BitBlt screenshot
```bash
node scripts/send-command.cjs '{
  "action": "fast_screenshot",
  "region": [0, 0, 800, 600],
  "path": "C:/screenshots/capture.png",
  "format": "png"
}'
# region: [x, y, width, height] or {"x":0,"y":0,"width":800,"height":600}
# path: optional, auto-generates temp path if omitted
# format: png (default), jpg, bmp
```

**clipboard_image** - Image clipboard operations
```bash
# Copy image to clipboard
node scripts/send-command.cjs '{
  "action": "clipboard_image",
  "operation": "set",
  "path": "C:/images/photo.png"
}'

# Get image from clipboard
node scripts/send-command.cjs '{
  "action": "clipboard_image",
  "operation": "get",
  "path": "C:/output/clipboard.png"
}'
# path for get: optional, auto-generates temp path if omitted
```

---

## Error Recovery

| Code | Recovery |
|------|----------|
| `WINDOW_NOT_FOUND` | Run `explore` |
| `ELEMENT_NOT_FOUND` | Try `explore_window`, `vision_click`, or `click_at` |
| `TIMEOUT` | Increase timeout parameter |
| `VISION_NOT_AVAILABLE` | Start Python bridge |
| `ALL_METHODS_FAILED` | Check element exists, try screenshot |

**Agent Fallback Status:**
When using `smart_click`/`smart_type` with `agent_fallback: true`, a `status: "agent_fallback"` response means the element wasn't found but a screenshot is included. Use your AI vision to analyze it and call `click_at` with coordinates.

Full errors: `node scripts/get-command-help.cjs --errors`

---

## Architecture

```
AI Agent --> send-command.cjs --> MainAgentService (C#)
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

**Primary:** FlaUI (Windows UI Automation - native, fast)
**Secondary:** Python Bridge (pywinauto + Vision + WinRT OCR + Context Manager)

---

## Tips

1. **Always `explore` first** - Never assume window state
2. **`focus_window` before interact** - Required for clicks/typing
3. **Use `wait_for_*`** - Apps need time to respond
4. **Token-efficient first** - `get_window_summary` over `explore_window`
5. **`smart_click`/`smart_type`** - Auto FlaUI→Vision fallback
6. **Check `health`** - Diagnose bridge issues

---

*Full command schema/example: `node scripts/get-command-help.cjs <command>`*
