# Windows Desktop Automation Skill

Hybrid automation engine for Windows using .NET 9 + FlaUI with Python vision fallback. Supports modern (UWP/WPF), legacy (Win32), and visual-only apps (Flutter, Electron).

## Quick Start

```powershell
# One-time setup
powershell -File scripts/auto-setup.ps1

# Start services
powershell -File scripts/start-all.ps1

# Send commands
node scripts/send-command.cjs '{"action":"explore"}'

# Stop services
powershell -File scripts/stop-all.ps1
```

## Architecture

```
AI Agent → JSON Commands → MainAgentService (.NET 9 + FlaUI)
                              ├── Layer 1: FlaUI (UIA3)
                              └── Layer 2: Python Bridge (:5001)
                                           ├── pywinauto (Win32)
                                           ├── OmniParser (Vision)
                                           └── WinRT OCR
```

**Fallback Chain:** FlaUI → pywinauto → Vision (OmniParser + OCR)

## Commands Reference

### Discovery
| Command | Description |
|---------|-------------|
| `explore` | List visible windows |
| `explore_window` | List UI elements in window |
| `find_element` | Find element with coordinates |
| `get_window_info` | Window details |
| `wait_for_element` | Wait for element (with timeout) |
| `element_exists` | Boolean check (~30 tokens) |
| `get_interactive_elements` | Clickable/typeable only |
| `get_window_summary` | Structured overview |
| `fuzzy_find_element` | Fuzzy text matching |

### Smart Commands (Auto-Fallback)
| Command | Description |
|---------|-------------|
| `smart_click` | Click with FlaUI→Vision fallback |
| `smart_type` | Type with FlaUI→Vision fallback |
| `vision_click` | Vision-only click by OCR text |

### Agent Vision (v4.3+)
| Command | Description |
|---------|-------------|
| `vision_screenshot` | Base64 JPEG for AI vision |
| `vision_screenshot_region` | Cropped region screenshot |
| `vision_config` | Get/set vision mode |
| `vision_analyze_smart` | Smart analysis with mode |
| `vision_screenshot_cache_stats` | Cache statistics |
| `vision_screenshot_cache_clear` | Clear cache |
| `vision_stream` | WebSocket URL for streaming |

### Win32 Low-Level (v5.0)
| Command | Description |
|---------|-------------|
| `scroll_sendinput` | Low-level scroll (bypasses UIA) |
| `type_sendinput` | Unicode + emoji via SendInput |
| `set_always_on_top` | Pin/unpin window on top |
| `flash_window` | Flash taskbar/caption |
| `set_window_opacity` | Window transparency (0-255) |
| `fast_screenshot` | Fast GDI BitBlt capture |
| `clipboard_image` | Image to/from clipboard |

### Window Management
| Command | Description |
|---------|-------------|
| `focus_window` | Bring to foreground |
| `close_window` | Close gracefully |
| `move_window` | Move to X,Y |
| `resize_window` | Resize to W×H |
| `minimize_window` / `maximize_window` / `restore_window` | State control |
| `move_to_monitor` | Move to specific display |
| `wait_for_window` | Wait for window to appear |

### Mouse
| Command | Description |
|---------|-------------|
| `click` / `double_click` / `right_click` | Element clicks |
| `click_at` | Click at X,Y coordinates |
| `mouse_move` | Move cursor |
| `drag_and_drop` | Drag between positions |
| `scroll` | Mouse wheel |
| `mouse_path` | Move along waypoints |
| `mouse_bezier` / `draw_bezier` | Bezier curve movement/drawing |
| `draw` | Hold + move + release |
| `mouse_down` / `mouse_up` | Button hold/release |
| `click_relative` | Click relative to anchor |

### Keyboard
| Command | Description |
|---------|-------------|
| `type` | Type into focused element |
| `type_here` | Type at cursor position |
| `hotkey` | Key combinations (Ctrl+C, etc.) |
| `key_press` | Single key |
| `key_down` / `key_up` | Key hold/release |

### Clipboard & OCR
| Command | Description |
|---------|-------------|
| `get_clipboard` / `set_clipboard` | Text clipboard |
| `clipboard_image` | Image clipboard (v5.0) |
| `read_text` | Read from element or OCR |
| `ocr_region` | OCR on screen region |

### Other
| Command | Description |
|---------|-------------|
| `screenshot` / `screenshot_monitor` | Screen capture |
| `list_monitors` | Enumerate displays |
| `launch_app` / `list_processes` / `kill_process` | Process control |
| `file_dialog` | Automate Open/Save dialogs |
| `batch` | Execute multiple commands |
| `health` | System diagnostics |

## Example Commands

```bash
# Window discovery and focus
node scripts/send-command.cjs '{"action":"explore"}'
node scripts/send-command.cjs '{"action":"focus_window", "selector":"Notepad"}'
node scripts/send-command.cjs '{"action":"explore_window", "selector":"Notepad", "max_depth":3}'

# Launch and interact
node scripts/send-command.cjs '{"action":"launch_app", "path":"notepad.exe", "wait_for_window":"Notepad"}'
node scripts/send-command.cjs '{"action":"type", "selector":"Notepad", "text":"Hello World"}'
node scripts/send-command.cjs '{"action":"hotkey", "keys":"ctrl+s"}'

# Smart commands with fallback
node scripts/send-command.cjs '{"action":"smart_click", "selector":"Submit", "window":"MyApp"}'
node scripts/send-command.cjs '{"action":"smart_click", "search_text":"Submit", "agent_fallback":true}'

# Win32 low-level (v5.0)
node scripts/send-command.cjs '{"action":"type_sendinput", "text":"Hello 🎉"}'
node scripts/send-command.cjs '{"action":"set_always_on_top", "selector":"Notepad", "enable":true}'
node scripts/send-command.cjs '{"action":"set_window_opacity", "selector":"Paint", "alpha":180}'

# Vision screenshots for AI
node scripts/send-command.cjs '{"action":"vision_screenshot"}'
node scripts/send-command.cjs '{"action":"vision_screenshot_region", "x":100, "y":100, "width":600, "height":400}'

# Batch execution
node scripts/send-command.cjs '{"action":"batch", "commands":[{"action":"launch_app","path":"notepad.exe"},{"action":"wait_for_window","selector":"Notepad"},{"action":"type","selector":"Notepad","text":"Hello!"}]}'
```

## Testing

```bash
node scripts/test-commands.cjs              # Run all tests
node scripts/test-commands.cjs --list       # List tests
node scripts/test-commands.cjs --category mouse  # Run category
```

## Python Bridge API (Port 5001)

| Category | Endpoints |
|----------|-----------|
| **Core** | `/health`, `/explore`, `/click`, `/type`, `/hotkey`, `/focus_window` |
| **Vision** | `/vision/detect`, `/vision/ocr`, `/vision/find_text`, `/vision/screenshot` |
| **Context** | `/context/stats`, `/context/clear`, `/context/invalidate` |
| **Cache** | `/vision/screenshot_cache/stats`, `/vision/screenshot_cache/clear` |

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Build fails (locked) | `scripts/stop-all.ps1` |
| Python bridge offline | Check port 5001, run `start-all.ps1` |
| Window not found | Use partial match: "Note" → "Notepad - Untitled" |
| Element click fails | `focus_window` first, then `explore_window` |

## Version History

| Version | Highlights |
|---------|------------|
| **5.0.0** | Win32 Low-Level Commands (7 new: `scroll_sendinput`, `type_sendinput`, `set_always_on_top`, `flash_window`, `set_window_opacity`, `fast_screenshot`, `clipboard_image`) |
| **4.3** | Agent Vision Commands, Screenshot Caching, Removed Java/SikuliX |
| **4.0** | Vision Layer (OmniParser + WinRT OCR), Token-efficient commands, Smart fallback |
| **3.3** | Advanced mouse (bezier, paths), Key hold/release, OCR region |
| **3.2** | Multi-monitor, File dialog automation, Batch commands |

---

**Full API details:** See [SKILL.md](SKILL.md) for complete command schemas and parameters.
